"""
mb_inbound.py
=============
Musical Box Response → Stibo (Inbound)

When a Musical Box response CSV (.txt/.csv) lands in S3 at
response/musicalbox/ (or any prefix configured via env vars), this module:

    1. Parses the pipe-delimited CSV (article|productSize|mb_compqty|mb_code)
    2. Groups rows by article code
    3. Resolves article codes → Stibo Product IDs via AT_Generic in archived XML
    4. Builds a STIBO-inbound XML  (PRD_GenericArticle per article,
       AT_MBVariantSizes / AT_MBVariantCompQty / AT_MBCode)
    5. Archives the XML to S3
    6. POSTs the XML to Stibo using OAuth2 Bearer token (client credentials)

Environment variables (mirrors by_inbound.py conventions):
    MAPPING_BUCKET                  – shared config bucket
    XML_ARCHIVE_BUCKET              – bucket that holds archived outbound XMLs
    XML_ARCHIVE_PREFIX              – prefix for archived XMLs   (default: xml-archive/)
    MB_STIBO_XML_OUTPUT_BUCKET      – bucket for archived generated XMLs
    MB_STIBO_XML_OUTPUT_PREFIX      – prefix for archived generated XMLs
                                       (default: stibo-response/musicalbox/)
    STIBO_CONTEXT_ID                – Stibo context   (default: Context1)
    STIBO_WORKSPACE_ID              – Stibo workspace (default: Main)
    STIBO_ENDPOINT_BY               – full Stibo IIEP upload URL (shared with BY)
    STIBO_TOKEN_URL / STIBO_CLIENT_ID / STIBO_CLIENT_SECRET / STIBO_GRANT_TYPE
    AUDIT_LOG_BUCKET                – bucket for JSON audit logs
    AUDIT_LOG_PREFIX                – prefix for audit logs
                                       (default: etl-audit-logs/musicalbox)
    ARTICLE_ID_MODE                 – "variant" (default) or "direct"
"""

import csv
import datetime
import json
import logging
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

import boto3
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
log = logging.getLogger("lambda_mb_etl")

# ─────────────────────────────────────────────────────────────────────────────
# AWS CLIENTS
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ─────────────────────────────────────────────────────────────────────────────
MAPPING_BUCKET             = os.environ.get("MAPPING_BUCKET", "")
XML_ARCHIVE_BUCKET         = os.environ.get("XML_ARCHIVE_BUCKET") or MAPPING_BUCKET
XML_ARCHIVE_PREFIX         = os.environ.get("XML_ARCHIVE_PREFIX", "xml-archive/")
MB_STIBO_XML_OUTPUT_BUCKET = os.environ.get("MB_STIBO_XML_OUTPUT_BUCKET") or MAPPING_BUCKET
MB_STIBO_XML_OUTPUT_PREFIX = os.environ.get("MB_STIBO_XML_OUTPUT_PREFIX", "stibo-response/musicalbox/")
STIBO_CONTEXT_ID           = os.environ.get("STIBO_CONTEXT_ID", "Context1")
STIBO_WORKSPACE_ID         = os.environ.get("STIBO_WORKSPACE_ID", "Main")
STIBO_ENDPOINT_BY          = os.environ.get(
    "STIBO_ENDPOINT_BY",
    "https://mapactive-dev.mdm.stibosystems.com/restapiv2/"
    "inbound-integration-endpoints/IIEP_BYToolsResponse/upload-direct"
    "?fileName=Send.xml&context=Context1&workspace=Main",
)
STIBO_TOKEN_URL            = os.environ.get("STIBO_TOKEN_URL", "")
STIBO_CLIENT_ID            = os.environ.get("STIBO_CLIENT_ID", "")
STIBO_CLIENT_SECRET        = os.environ.get("STIBO_CLIENT_SECRET", "")
STIBO_GRANT_TYPE           = os.environ.get("STIBO_GRANT_TYPE", "client_credentials")
AUDIT_LOG_BUCKET           = os.environ.get("AUDIT_LOG_BUCKET") or MAPPING_BUCKET
AUDIT_LOG_PREFIX           = os.environ.get("AUDIT_LOG_PREFIX", "etl-audit-logs/musicalbox")
ARTICLE_ID_MODE            = os.environ.get("ARTICLE_ID_MODE", "variant").lower()

TMP = Path("/tmp")

# ─────────────────────────────────────────────────────────────────────────────
# TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────
_TZ_GMT7 = datetime.timezone(datetime.timedelta(hours=7))

# ─────────────────────────────────────────────────────────────────────────────
# UPSERT KEYS  (used by audit logger identity blocks)
# ─────────────────────────────────────────────────────────────────────────────
UPSERT_KEYS = ["article"]


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────────────────────
def _utcnow_iso() -> str:
    return datetime.datetime.now(_TZ_GMT7).isoformat()


def _size_sort_key(s: str):
    """Sort sizes numerically where possible (4, 4.5, 5 …), alphabetically otherwise."""
    # Handle half-sizes written as e.g. "4H" → treat as 4.5
    half_match = re.match(r'^(\d+)H$', s.strip(), re.IGNORECASE)
    if half_match:
        return (0, float(half_match.group(1)) + 0.5)
    try:
        return (0, float(s.strip()))
    except ValueError:
        return (1, s.strip().lower())


# ─────────────────────────────────────────────────────────────────────────────
# AUDIT LOGGER
# ─────────────────────────────────────────────────────────────────────────────
class AuditLogger:
    def __init__(self, etl_id: int, raw_event: dict, context=None):
        self._start_ts = datetime.datetime.now(_TZ_GMT7)
        self.etl_id    = etl_id
        self.raw_event = raw_event
        self.files     = []
        self._cur_file = None
        self.context   = context

    def begin_file(self, s3_path: str):
        self._cur_file = {
            "s3_path":       s3_path,
            "started_at":    _utcnow_iso(),
            "rows_produced": 0,
            "rows_upserted": 0,
            "rows_failed":   0,
            "rows_fallback": 0,
            "status":        "processing",
            "BGPID":         None,
            "rows":          [],
        }
        self.files.append(self._cur_file)

    def end_file(self, status: str, rows_produced=0,
                 rows_upserted=0, rows_failed=0, rows_fallback=0):
        if self._cur_file is None:
            return
        self._cur_file.update({
            "finished_at":   _utcnow_iso(),
            "status":        status,
            "rows_produced": rows_produced,
            "rows_upserted": rows_upserted,
            "rows_failed":   rows_failed,
            "rows_fallback": rows_fallback,
        })
        self._cur_file = None

    def set_stibo_response_id(self, response_id: str):
        if self._cur_file is not None:
            self._cur_file["BGPID"] = response_id
            log.info("[MB] BGPID: %s", response_id)

    def record_row(self, article: str, warnings: list, db_status: str, db_error: str = None):
        if self._cur_file is None:
            return
        self._cur_file["rows"].append({
            "identity":   {"article": article},
            "db_status":  db_status,
            "db_error":   db_error,
            "warnings":   warnings,
            "recorded_at": _utcnow_iso(),
        })

    def record_issue(self, article: str, issue: str):
        if self._cur_file is None:
            return
        self._cur_file["rows"].append({
            "identity":    {"article": article},
            "issue":       issue,
            "recorded_at": _utcnow_iso(),
        })

    def flush(self, stem: str):
        if not AUDIT_LOG_BUCKET:
            log.warning("[MB] AUDIT_LOG_BUCKET not set — audit log NOT written")
            return

        end_ts   = datetime.datetime.now(_TZ_GMT7)
        duration = round((end_ts - self._start_ts).total_seconds(), 3)
        date_str = self._start_ts.strftime("%Y/%m/%d")

        document = {
            "schema_version":   "1.0",
            "etl_id":           self.etl_id,
            "lambda_function":  os.environ.get("AWS_LAMBDA_FUNCTION_NAME", "unknown"),
            "aws_request_id":   getattr(self.context, "aws_request_id", "unknown"),
            "started_at":       self._start_ts.isoformat(),
            "finished_at":      end_ts.isoformat(),
            "duration_seconds": duration,
            "input_event": {
                "bucket": self.raw_event.get("bucket"),
                "key":    self.raw_event.get("key"),
            },
            "files_processed": self.files,
            "totals": {
                "files":         len(self.files),
                "rows_produced": sum(f["rows_produced"] for f in self.files),
                "rows_upserted": sum(f["rows_upserted"] for f in self.files),
                "rows_failed":   sum(f["rows_failed"]   for f in self.files),
                "rows_fallback": sum(f["rows_fallback"] for f in self.files),
            },
        }

        key = (
            f"{AUDIT_LOG_PREFIX.rstrip('/')}/"
            f"{date_str}/"
            f"{self.etl_id}_{stem}.json"
        )

        try:
            s3.put_object(
                Bucket      = AUDIT_LOG_BUCKET,
                Key         = key,
                Body        = json.dumps(document, indent=2, default=str).encode("utf-8"),
                ContentType = "application/json",
            )
            log.info("✅ [MB] Audit log written → s3://%s/%s", AUDIT_LOG_BUCKET, key)
        except Exception as exc:
            log.error("❌ [MB] Audit log upload failed: %s", exc)


# ─────────────────────────────────────────────────────────────────────────────
# CSV PARSING
# Groups rows by article; returns dict[article_code → grouped_record]
# ─────────────────────────────────────────────────────────────────────────────
def parse_mb_csv(local_path: str) -> dict:
    """
    Parse pipe-delimited CSV with columns: article|productSize|mb_compqty|mb_code

    Returns a dict keyed by article code:
        {
            "ADI255588888": {
                "article":   "ADI255588888",
                "sizes":     [("4", "1", "0FA"), ("4H", "1", "0FA"), ...],
                             # list of (productSize, mb_compqty, mb_code) in original order
            },
            ...
        }
    """
    log.info("[MB] Parsing CSV: %s", local_path)
    grouped: dict = {}

    with open(local_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh, delimiter="|")
        header_skipped = False

        for raw_row in reader:
            row = [cell.strip() for cell in raw_row]

            if not any(row):
                continue

            if not header_skipped:
                first = row[0].lower() if row else ""
                if first in ("article", "s. no.", "s.no.", "#"):
                    header_skipped = True
                    continue
                if not re.match(r'^[A-Z0-9]', row[0], re.IGNORECASE) or row[0].lower() == "article":
                    header_skipped = True
                    continue
                header_skipped = True

            if len(row) < 4:
                log.warning("[MB] Skipping short row: %s", row)
                continue

            article   = row[0].strip()
            size      = row[1].strip()
            compqty   = row[2].strip()
            mb_code   = row[3].strip()

            if not article:
                log.warning("[MB] Skipping row with blank article: %s", row)
                continue

            if article not in grouped:
                grouped[article] = {
                    "article": article,
                    "sizes":   [],   # list of (size, compqty, mb_code) tuples
                }

            if size:
                grouped[article]["sizes"].append((size, compqty, mb_code))
            else:
                log.warning("[MB] Article %r has a row with blank productSize — row skipped.", article)

    log.info("[MB] Parsed %d unique article(s) from CSV.", len(grouped))
    return grouped


# ─────────────────────────────────────────────────────────────────────────────
# ARCHIVED XML LOOKUP  (AT_Generic → Stibo Product ID)
# ─────────────────────────────────────────────────────────────────────────────
def build_generic_lookup(xml_bytes: bytes) -> dict:
    """
    Build a dict mapping AT_Generic attribute value (upper-cased) → Stibo Product ID.
    Also indexes by raw Product ID so direct IDs still resolve.
    """
    lookup = {}
    try:
        root = etree.fromstring(xml_bytes)
    except Exception as exc:
        log.error("[MB] Failed to parse archived XML: %s", type(exc).__name__)
        return lookup

    for product in root.iter("Product"):
        stibo_id = (product.get("ID") or "").strip()
        if not stibo_id:
            continue

        # Index by raw Product ID
        lookup[stibo_id.upper()] = stibo_id

        values_el = product.find("Values")
        if values_el is None:
            continue

        for value_el in values_el.findall("Value"):
            if value_el.get("AttributeID") != "AT_Generic":
                continue
            # Check both text content and ID attribute
            for candidate in (value_el.text, value_el.get("ID")):
                if candidate and str(candidate).strip():
                    lookup[str(candidate).strip().upper()] = stibo_id

    log.info("[MB] Generic lookup built: %d entries.", len(lookup))
    return lookup


def load_latest_archived_xml() -> bytes | None:
    if not XML_ARCHIVE_BUCKET:
        log.warning("[MB] XML_ARCHIVE_BUCKET not set — cannot load archived XML.")
        return None

    prefix = XML_ARCHIVE_PREFIX.rstrip("/") + "/"
    try:
        paginator = s3.get_paginator("list_objects_v2")
        objects = []
        for page in paginator.paginate(Bucket=XML_ARCHIVE_BUCKET, Prefix=prefix):
            objects.extend(page.get("Contents", []))

        if not objects:
            log.warning("[MB] No archived XMLs found at s3://%s/%s", XML_ARCHIVE_BUCKET, prefix)
            return None

        latest = max(objects, key=lambda obj: obj["LastModified"])
        log.info("[MB] Loading archived XML: s3://%s/%s", XML_ARCHIVE_BUCKET, latest["Key"])
        response = s3.get_object(Bucket=XML_ARCHIVE_BUCKET, Key=latest["Key"])
        return response["Body"].read()

    except Exception as exc:
        log.error("[MB] Failed to load archived XML: %s", exc)
        return None


def resolve_stibo_id(article_code: str, lookup: dict) -> str:
    """Return Stibo Product ID for article_code, or article_code itself as fallback."""
    article_code = str(article_code or "").strip()
    if ARTICLE_ID_MODE == "direct":
        return article_code

    resolved = lookup.get(article_code.upper())
    if not resolved:
        log.warning(
            "[MB] Cannot resolve %r to a Stibo Product ID via AT_Generic — writing as-is.",
            article_code,
        )
        return article_code
    return resolved


# ─────────────────────────────────────────────────────────────────────────────
# BUILD STIBO XML
# ─────────────────────────────────────────────────────────────────────────────
def build_mb_stibo_xml(grouped: dict, lookup: dict) -> tuple[bytes, dict]:
    """
    Build the STIBO-inbound XML.

    One PRD_GenericArticle per article with:
        AT_MBVariantSizes            — sizes joined by ";"  (size-sorted)
        AT_MBVariantCompQty          — compqty joined by ";" (same sort order)
        AT_MBVariantAssortmentCode   — mb_code joined by ";" (same sort order,
                                        one value per row — supports multiple
                                        distinct assortment codes per article)
    """
    now_str = datetime.datetime.now(_TZ_GMT7).strftime("%Y-%m-%d %H:%M:%S")
    root = etree.Element(
        "STEP-ProductInformation",
        attrib={
            "ExportTime":       now_str,
            "ExportContext":    STIBO_CONTEXT_ID,
            "ContextID":        STIBO_CONTEXT_ID,
            "WorkspaceID":      STIBO_WORKSPACE_ID,
            "UseContextLocale": "false",
        },
    )
    products_el = etree.SubElement(root, "Products")

    resolved_count   = 0
    unresolved_count = 0
    skipped_count    = 0

    for article_code, rec in grouped.items():
        sizes_raw = rec["sizes"]   # list of (size, compqty, mb_code) tuples

        if not sizes_raw:
            log.debug("[MB] Skipping article %r — no valid size rows.", article_code)
            skipped_count += 1
            continue

        try:
            sizes_sorted = sorted(sizes_raw, key=lambda triple: _size_sort_key(triple[0]))
        except Exception:
            sizes_sorted = sizes_raw

        sizes_str    = ";".join(triple[0] for triple in sizes_sorted)
        compqty_str  = ";".join(triple[1] for triple in sizes_sorted)
        mb_code_str  = ";".join(triple[2] for triple in sizes_sorted)

        stibo_id = resolve_stibo_id(article_code, lookup)

        if ARTICLE_ID_MODE == "variant" and article_code.upper() in lookup:
            resolved_count += 1
        elif ARTICLE_ID_MODE == "direct":
            resolved_count += 1
        else:
            unresolved_count += 1

        product_el = etree.SubElement(
            products_el,
            "Product",
            attrib={"UserTypeID": "PRD_GenericArticle"},
        )

        kv_el = etree.SubElement(product_el, "KeyValue")
        kv_el.set("KeyID", "KEY_Article")
        kv_el.text = stibo_id

        values_el = etree.SubElement(product_el, "Values")

        v_sizes = etree.SubElement(
            values_el,
            "Value",
            attrib={"AttributeID": "AT_MBVariantSizes"},
        )
        v_sizes.text = sizes_str

        v_qty = etree.SubElement(
            values_el,
            "Value",
            attrib={"AttributeID": "AT_MBVariantCompQty"},
        )
        v_qty.text = compqty_str

        if mb_code_str.strip(";"):
            v_code = etree.SubElement(
                values_el,
                "Value",
                attrib={"AttributeID": "AT_MBVariantAssortmentCode"},
            )
            v_code.text = mb_code_str

    xml_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        pretty_print=True,
    )

    total = len(grouped)
    log.info(
        "[MB] XML generated: articles=%d resolved=%d unresolved=%d skipped=%d bytes=%d",
        total, resolved_count, unresolved_count, skipped_count, len(xml_bytes),
    )
    return xml_bytes, {
        "resolved":   resolved_count,
        "unresolved": unresolved_count,
        "skipped":    skipped_count,
    }


# ─────────────────────────────────────────────────────────────────────────────
# ARCHIVE & POST TO STIBO
# ─────────────────────────────────────────────────────────────────────────────
def archive_mb_xml_to_s3(xml_bytes: bytes, filename: str) -> str | None:
    if not MB_STIBO_XML_OUTPUT_BUCKET:
        log.warning("[MB] MB_STIBO_XML_OUTPUT_BUCKET not set — skipping XML archive.")
        return None

    key = "%s/%s" % (MB_STIBO_XML_OUTPUT_PREFIX.rstrip("/"), filename)
    log.info("[MB] Archiving generated XML → s3://%s/%s", MB_STIBO_XML_OUTPUT_BUCKET, key)
    s3.put_object(
        Bucket      = MB_STIBO_XML_OUTPUT_BUCKET,
        Key         = key,
        Body        = xml_bytes,
        ContentType = "application/xml",
    )
    return key


def _mb_endpoint_url(filename: str) -> str:
    if not STIBO_ENDPOINT_BY:
        return ""
    if "?" in STIBO_ENDPOINT_BY:
        return STIBO_ENDPOINT_BY
    query = urllib.parse.urlencode({
        "fileName":  filename,
        "context":   STIBO_CONTEXT_ID,
        "workspace": STIBO_WORKSPACE_ID,
    })
    return "%s?%s" % (STIBO_ENDPOINT_BY.rstrip("/"), query)


def _get_stibo_bearer_token() -> str:
    if not STIBO_TOKEN_URL or not STIBO_CLIENT_ID or not STIBO_CLIENT_SECRET:
        log.error("[MB] Stibo token credentials not fully configured.")
        return ""

    payload = urllib.parse.urlencode({
        "grant_type":    STIBO_GRANT_TYPE,
        "client_id":     STIBO_CLIENT_ID,
        "client_secret": STIBO_CLIENT_SECRET,
    }).encode("utf-8")

    request = urllib.request.Request(
        url     = STIBO_TOKEN_URL,
        data    = payload,
        headers = {"Content-Type": "application/x-www-form-urlencoded"},
        method  = "POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        return body.get("access_token", "")
    except Exception as exc:
        log.error("[MB] Failed to fetch Stibo token: %s", exc)
        return ""


def post_mb_xml_to_stibo(xml_bytes: bytes, filename: str) -> dict:
    url = _mb_endpoint_url(filename)
    if not url:
        log.warning("[MB] STIBO_ENDPOINT_BY not configured — skipping HTTP POST.")
        return {"posted": False, "status": None, "body": "STIBO_ENDPOINT_BY not configured"}

    token = _get_stibo_bearer_token()
    if not token:
        log.error("[MB] Could not obtain Stibo Bearer token — skipping POST.")
        return {"posted": False, "status": None, "body": "Could not obtain Stibo Bearer token"}

    log.info("[MB] Posting XML to Stibo: %s", url)
    request = urllib.request.Request(
        url     = url,
        data    = xml_bytes,
        headers = {
            "Authorization": "Bearer %s" % token,
            "Content-Type":  "application/octet-stream",
        },
        method = "POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as resp:
            status = resp.status
            body   = resp.read().decode("utf-8", errors="replace")

        if status in (200, 201, 202):
            log.info("[MB] Stibo upload success: HTTP %s", status)
            try:
                response_id = json.loads(body).get("id", "")
            except Exception:
                response_id = ""
            return {"posted": True, "status": status, "body": body, "id": response_id}

        log.error("[MB] Stibo upload failure: HTTP %s — %s", status, body)
        return {"posted": False, "status": status, "body": body}

    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        log.error("[MB] Stibo upload failure: HTTP %s — %s", exc.code, body)
        return {"posted": False, "status": exc.code, "body": body}
    except urllib.error.URLError as exc:
        log.error("[MB] Stibo upload failure: %s", exc.reason)
        return {"posted": False, "status": None, "body": str(exc.reason)}
    except Exception as exc:
        log.error("[MB] Stibo upload failure: %s", exc)
        return {"posted": False, "status": None, "body": str(exc)}


# ─────────────────────────────────────────────────────────────────────────────
# HANDLER  (called by router in lambda_function.py)
# ─────────────────────────────────────────────────────────────────────────────
def handle_mb_response(src_bucket: str, response_key: str, auditor: AuditLogger) -> dict:
    """
    Full Musical Box inbound pipeline:
        download CSV → parse & group → resolve IDs → build XML → archive → POST to Stibo.

    Called from the router in lambda_function.py with:
        handle_mb_response(bucket, key, auditor)
    """
    local_response = TMP / response_key.replace("/", "_")
    stem           = Path(response_key).stem
    grouped        = {}

    try:
        auditor.begin_file("s3://%s/%s" % (src_bucket, response_key))
        log.info("[MB] File received: s3://%s/%s", src_bucket, response_key)

        # ── 1. Download ───────────────────────────────────────────────────
        log.info("[MB] Downloading response file.")
        s3.download_file(src_bucket, response_key, str(local_response))

        # ── 2. Parse & group ──────────────────────────────────────────────
        grouped = parse_mb_csv(str(local_response))
        if not grouped:
            auditor.end_file("skipped")
            return {"file": response_key, "status": "Skipped", "reason": "Empty or unparseable CSV"}

        # ── 3. Resolve article IDs ────────────────────────────────────────
        lookup = {}
        if ARTICLE_ID_MODE == "variant":
            archived_xml = load_latest_archived_xml()
            if archived_xml:
                lookup = build_generic_lookup(archived_xml)
            else:
                log.warning("[MB] No archived XML available — article IDs will be written as-is.")
        else:
            log.info("[MB] ARTICLE_ID_MODE=direct — article codes used directly.")

        # Audit unresolved articles
        for article_code in grouped:
            if ARTICLE_ID_MODE == "variant" and article_code.upper() not in lookup:
                auditor.record_issue(article_code, "unresolved_identifier")

        # ── 4. Build XML ──────────────────────────────────────────────────
        xml_bytes, counts = build_mb_stibo_xml(grouped, lookup)

        # ── 5. Archive XML ────────────────────────────────────────────────
        timestamp    = datetime.datetime.now(_TZ_GMT7).strftime("%Y%m%d_%H%M%S")
        xml_filename = "mb_stibo_response_%s_%s.xml" % (stem, timestamp)
        archived_key = archive_mb_xml_to_s3(xml_bytes, xml_filename)
        log.info("[MB] XML archived: %s", archived_key or "(not archived)")

        # ── 6. POST to Stibo ──────────────────────────────────────────────
        post_result = post_mb_xml_to_stibo(xml_bytes, xml_filename)

        response_id = post_result.get("id", "")
        auditor.set_stibo_response_id(response_id)
        if response_id:
            log.info("[MB] BGP ID received and stored: %s", response_id)

        # Audit each article outcome
        for article_code in grouped:
            stibo_id  = resolve_stibo_id(article_code, lookup)
            resolved  = ARTICLE_ID_MODE == "direct" or article_code.upper() in lookup
            warnings  = [] if resolved else ["unresolved_identifier"]
            db_status = "upserted" if post_result.get("posted") and resolved else (
                "failed" if not post_result.get("posted") else "fallback"
            )
            auditor.record_row(article_code, warnings, db_status)

        total_articles = len(grouped)
        auditor.end_file(
            "ok" if post_result.get("posted") else "failed",
            rows_produced = total_articles,
            rows_upserted = counts["resolved"],
            rows_failed   = counts["unresolved"] + counts["skipped"],
        )

        return {
            "file":               response_key,
            "status":             "Success" if post_result.get("posted") else "PostFailed",
            "articles_total":     total_articles,
            "articles_resolved":  counts["resolved"],
            "articles_unresolved":counts["unresolved"],
            "articles_skipped":   counts["skipped"],
            "archived_xml":       archived_key,
            "stibo_posted":       post_result.get("posted"),
            "stibo_status":       post_result.get("status"),
            "stibo_response_id":  post_result.get("id", ""),
        }

    except Exception as exc:
        log.error("[MB] Exception while processing %s: %s", response_key, exc, exc_info=True)
        for article_code in grouped:
            auditor.record_issue(article_code, "exception: %s" % exc)
        auditor.end_file("failed")
        return {"file": response_key, "status": "Failed", "error": str(exc)}

    finally:
        local_response.unlink(missing_ok=True)