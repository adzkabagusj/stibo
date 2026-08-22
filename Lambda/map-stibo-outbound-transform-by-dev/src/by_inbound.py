"""
by_inbound.py
=============
Step 2 — Blue Yonder Response → Stibo (Inbound)

When a BY response file (.txt/.csv/.json) lands in S3 at response/blueyonder/,
this module:
    1. Parses the BY response file
    2. Loads the BY→Stibo mapping Excel
    3. Resolves article codes back to Stibo Product IDs via archived XML
    4. Builds a STIBO-inbound XML
    5. POSTs the XML to Stibo using OAuth2 Bearer token (client credentials)
"""

import csv
import datetime
import json
import logging
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import boto3
from lxml import etree
from openpyxl import load_workbook

_DATE_PARSE_FORMATS = (
    "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%m-%d-%Y",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S.%f",
    "%d%m%Y", "%Y%m%d", "%d %m %Y", "%m %d %Y", "%d/%m/%Y %H:%M:%S",
)

def _normalize_date_dd_mm_yyyy(raw: str, field: str) -> str:
    stripped = raw.strip()
    for fmt in _DATE_PARSE_FORMATS:
        try:
            return datetime.datetime.strptime(stripped, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    log.warning(
        "⚠️ [Step 2] Could not parse date value %r for field %r — keeping original value",
        stripped, field,
    )
    return stripped

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
log = logging.getLogger("lambda_by_etl")

# ─────────────────────────────────────────────────────────────────────────────
# AWS CLIENTS
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ─────────────────────────────────────────────────────────────────────────────
MAPPING_BUCKET              = os.environ.get("MAPPING_BUCKET", "")
MAPPING_KEY_BY_TO_STIBO     = os.environ.get("MAPPING_KEY_BY_TO_STIBO", "")
BY_TO_STIBO_MAPPING_SHEET   = os.environ.get("BY_TO_STIBO_MAPPING_SHEET", "Sheet1")
ARTICLE_ID_MODE             = os.environ.get("ARTICLE_ID_MODE", "variant").lower()
OUTPUT_BUCKET               = os.environ.get("OUTPUT_BUCKET") or MAPPING_BUCKET
XML_ARCHIVE_BUCKET          = os.environ.get("XML_ARCHIVE_BUCKET") or OUTPUT_BUCKET
XML_ARCHIVE_PREFIX          = os.environ.get("XML_ARCHIVE_PREFIX", "xml-archive/")
BY_STIBO_XML_OUTPUT_BUCKET  = os.environ.get("BY_STIBO_XML_OUTPUT_BUCKET") or OUTPUT_BUCKET
BY_STIBO_XML_OUTPUT_PREFIX  = os.environ.get("BY_STIBO_XML_OUTPUT_PREFIX", "stibo-response/blueyonder/")
STIBO_CONTEXT_ID            = os.environ.get("STIBO_CONTEXT_ID", "Context1")
STIBO_WORKSPACE_ID          = os.environ.get("STIBO_WORKSPACE_ID", "Main")
STIBO_ENDPOINT_BY           = os.environ.get(
    "STIBO_ENDPOINT_BY",
    "https://mapactive-dev.mdm.stibosystems.com/restapiv2/"
    "inbound-integration-endpoints/IIEP_BYToolsResponse/upload-direct"
    "?fileName=Send.xml&context=Context1&workspace=Main",
)
BY_GENERIC_PARENT_ID        = os.environ.get("BY_GENERIC_PARENT_ID", "PPH_F-TempSubCat")

AUDIT_LOG_BUCKET = os.environ.get("AUDIT_LOG_BUCKET") or OUTPUT_BUCKET
AUDIT_LOG_PREFIX = os.environ.get("AUDIT_LOG_PREFIX", "etl-audit-logs/blueyonder")

STIBO_TOKEN_URL     = os.environ.get("STIBO_TOKEN_URL", "")
STIBO_CLIENT_ID     = os.environ.get("STIBO_CLIENT_ID", "")
STIBO_CLIENT_SECRET = os.environ.get("STIBO_CLIENT_SECRET", "")
STIBO_GRANT_TYPE    = os.environ.get("STIBO_GRANT_TYPE", "client_credentials")

TMP = Path("/tmp")

HEADER_MARKER = "S. No."
UPSERT_KEYS   = ["Brand_Principal_Gen_Article", "BY_Sub_Category_ID"]

# ─────────────────────────────────────────────────────────────────────────────
# TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────
_TZ_GMT7 = datetime.timezone(datetime.timedelta(hours=7))


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────────────────────
def _utcnow_iso() -> str:
    return datetime.datetime.now(_TZ_GMT7).isoformat()


# ─────────────────────────────────────────────────────────────────────────────
# AUDIT LOGGER  (same class used by outbound — duplicated here to keep files
#                fully independent; could be extracted to shared.py later)
# ─────────────────────────────────────────────────────────────────────────────
class AuditLogger:
    def __init__(self, etl_id: int, raw_event: dict, context=None):
        self._start_ts  = datetime.datetime.now(_TZ_GMT7)
        self.etl_id     = etl_id
        self.raw_event  = raw_event
        self.files      = []
        self._cur_file  = None
        self.context    = context

    def begin_file(self, s3_path: str):
        self._cur_file = {
            "s3_path":        s3_path,
            "started_at":     _utcnow_iso(),
            "rows_produced":  0,
            "rows_upserted":  0,
            "rows_failed":    0,
            "rows_fallback":  0,
            "status":         "processing",
            "BGPID":          None,
            "rows":           [],
        }
        self.files.append(self._cur_file)

    def end_file(self, status: str, rows_produced=0,
                 rows_upserted=0, rows_failed=0, rows_fallback=0):
        if self._cur_file is None:
            return
        self._cur_file.update({
            "finished_at":    _utcnow_iso(),
            "status":         status,
            "rows_produced":  rows_produced,
            "rows_upserted":  rows_upserted,
            "rows_failed":    rows_failed,
            "rows_fallback":  rows_fallback,
        })
        self._cur_file = None

    def set_stibo_response_id(self, response_id: str):
        if self._cur_file is not None:
            self._cur_file["BGPID"] = response_id
            log.info("[Audit] BGPID: %s", response_id)

    def record_row(self, row: dict, warnings: list, db_status: str, db_error: str = None):
        if self._cur_file is None:
            return
        identity = {k: row.get(k) for k in UPSERT_KEYS + ["Brand", "Season", "Style"]}
        populated   = {k: v for k, v in row.items() if v is not None}
        null_fields = [k for k, v in row.items() if v is None]
        self._cur_file["rows"].append({
            "identity":         identity,
            "db_status":        db_status,
            "db_error":         db_error,
            "populated_fields": populated,
            "null_fields":      null_fields,
            "warnings":         warnings,
            "recorded_at":      _utcnow_iso(),
        })

    def flush(self, stem: str):
        if not AUDIT_LOG_BUCKET:
            log.warning("AUDIT_LOG_BUCKET not set — audit log NOT written")
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
            "files_processed":  self.files,
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
            log.info("✅ Audit log written → s3://%s/%s", AUDIT_LOG_BUCKET, key)
        except Exception as exc:
            log.error("❌ Audit log upload failed: %s", exc)


# ─────────────────────────────────────────────────────────────────────────────
# BY → STIBO MAPPING
# ─────────────────────────────────────────────────────────────────────────────
class ByToStiboRule:
    __slots__ = ("by_field", "field_key", "stibo_attr_id", "stibo_attr_type")

    def __init__(self, by_field, stibo_attr_id, stibo_attr_type):
        self.by_field        = str(by_field or "").strip()
        self.field_key       = _normalize_by_field(self.by_field)
        self.stibo_attr_id   = str(stibo_attr_id or "").strip()
        self.stibo_attr_type = str(stibo_attr_type or "").strip()


def _normalize_by_field(name: str) -> str:
    if name is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


def _is_lov_type(stibo_attr_type: str) -> bool:
    return "list" in (stibo_attr_type or "").lower() and "value" in (stibo_attr_type or "").lower()


def load_by_to_stibo_mapping(bucket: str, key: str):
    if not bucket or not key:
        raise ValueError("MAPPING_BUCKET or MAPPING_KEY_BY_TO_STIBO missing for BY->Stibo mapping.")

    local_path = TMP / "mapping_by_to_stibo.xlsx"
    log.info("[Step 2] Loading BY->Stibo mapping from s3://%s/%s", bucket, key)
    s3.download_file(bucket, key, str(local_path))

    wb = load_workbook(str(local_path), read_only=True, data_only=True)
    if BY_TO_STIBO_MAPPING_SHEET in wb.sheetnames:
        ws = wb[BY_TO_STIBO_MAPPING_SHEET]
    else:
        raise ValueError(
            "Sheet %r not found in BY->Stibo mapping. Available sheets: %s"
            % (BY_TO_STIBO_MAPPING_SHEET, wb.sheetnames)
        )

    rules = []
    data_started = False

    for row in ws.iter_rows(values_only=True):
        row_values = list(row or [])
        first_col = str(row_values[0]).strip() if len(row_values) > 0 and row_values[0] is not None else ""

        if not data_started:
            if first_col == HEADER_MARKER:
                data_started = True
                continue
            if len(row_values) > 7 and row_values[1] and row_values[7]:
                data_started = True
            else:
                continue

        by_field = str(row_values[1]).strip() if len(row_values) > 1 and row_values[1] else ""
        stibo_attr_id = str(row_values[7]).strip() if len(row_values) > 7 and row_values[7] else ""
        stibo_attr_type = str(row_values[8]).strip() if len(row_values) > 8 and row_values[8] else ""

        if not by_field or not stibo_attr_id:
            continue
        rules.append(ByToStiboRule(by_field, stibo_attr_id, stibo_attr_type))

    log.info("[Step 2] Mapping loaded: %d BY->Stibo rules from sheet %r", len(rules), BY_TO_STIBO_MAPPING_SHEET)
    return rules


# ─────────────────────────────────────────────────────────────────────────────
# RESPONSE PARSING
# ─────────────────────────────────────────────────────────────────────────────
def _by_identifier_rule(mapping_rules):
    for rule in mapping_rules:
        if rule.stibo_attr_id.upper() == "AT_VARIANT":
            return rule
    return None


def _by_size_field_key(mapping_rules) -> str:
    """Resolve the record key that carries the size list, by looking up
    whichever mapping rule targets AT_Size — rather than assuming the BY
    field is literally named 'productSize'. Different response templates
    use different BY field names (e.g. 'sap_size_code') for the same
    Stibo AT_Size attribute."""
    for rule in mapping_rules:
        if rule.stibo_attr_id.upper() == "AT_SIZE":
            return rule.field_key
    # Fallback to legacy hardcoded name if no AT_Size rule is mapped.
    return _normalize_by_field("productSize")


def _by_identifier_value(record: dict, mapping_rules) -> str:
    identifier_rule = _by_identifier_rule(mapping_rules)
    if identifier_rule:
        value = record.get(identifier_rule.field_key)
        if value:
            return str(value).strip()

    for fallback in ("Article", "Variant", "VariantArticle"):
        value = record.get(_normalize_by_field(fallback))
        if value:
            return str(value).strip()

    return ""


def _json_records_payload(data):
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("records", "data", "items", "responses", "results"):
            value = data.get(key)
            if isinstance(value, list):
                return value
        for value in data.values():
            if isinstance(value, list):
                return value
    return []


def parse_by_response(local_path: str, mapping_rules):
    log.info("[Step 2] Parsing started: %s", local_path)
    records = []
    suffix = local_path.lower()

    if suffix.endswith(".json"):
        with open(local_path, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        for item in _json_records_payload(data):
            if not isinstance(item, dict):
                continue
            record = {}
            for key, value in item.items():
                text = "" if value is None else str(value).strip()
                record[_normalize_by_field(key)] = text
            if any(record.values()):
                records.append(record)
        log.info("[Step 2] Parsing completed: %d JSON record(s)", len(records))
        return records

    # Detect delimiter: BY response files have historically been pipe-delimited,
    # but templates such as the MD Tools response are plain comma-delimited CSV
    # (with quoted fields, e.g. "OSM, 00M, 00L"). Sniff from the first line
    # rather than assuming pipe, or every field on comma-delimited files
    # collapses into a single garbage column.
    with open(local_path, "r", encoding="utf-8-sig", newline="") as fh:
        first_line = fh.readline()
    delimiter = "|" if "|" in first_line else ","
    log.info("[Step 2] Detected response delimiter: %r", delimiter)

    with open(local_path, "r", encoding="utf-8-sig", newline="") as fh:
        rows = [
            [cell.strip() for cell in row]
            for row in csv.reader(fh, delimiter=delimiter)
            if row and any(cell.strip() for cell in row)
        ]

    if not rows:
        log.warning("[Step 2] BY response file is empty.")
        return records

    mapping_keys = {rule.field_key for rule in mapping_rules if rule.field_key}
    first_keys = [_normalize_by_field(cell) for cell in rows[0]]
    has_header = bool(mapping_keys.intersection(first_keys))

    if has_header:
        headers = first_keys
        data_rows = rows[1:]
        log.info("[Step 2] Response header detected (delimiter=%r).", delimiter)
    else:
        headers = [rule.field_key for rule in mapping_rules]
        data_rows = rows
        log.info("[Step 2] No response header detected; using mapping order (delimiter=%r).", delimiter)

    for row in data_rows:
        record = {}
        for idx, value in enumerate(row):
            if idx >= len(headers):
                break
            header = headers[idx]
            if header:
                record[header] = value.strip()
        if any(record.values()):
            records.append(record)

    log.info("[Step 2] Parsing completed: %d pipe-delimited record(s)", len(records))
    return records


# ─────────────────────────────────────────────────────────────────────────────
# ARTICLE ID RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────
def build_variant_lookup(xml_bytes: bytes) -> dict:
    lookup = {}
    try:
        root = etree.fromstring(xml_bytes)
    except Exception as exc:
        log.error("[Step 2] Failed to parse archived XML: %s", type(exc).__name__)  # ← fixed log message
        return lookup

    lookup_attrs = (
        "AT_Variant",
        "AT_Generic",
        "MATNR",
        "AT_InternalBarcode",
        "AT_PrincipalBarcode",
        "AT_Barcode",
    )

    for product in root.iter("Product"):
        stibo_id = (product.get("ID") or "").strip()
        if not stibo_id:
            continue

        lookup[stibo_id.upper()] = stibo_id
        values_el = product.find("Values")
        if values_el is None:
            continue

        for attr_id in lookup_attrs:
            for value_el in values_el.findall("Value"):
                if value_el.get("AttributeID") != attr_id:
                    continue
                for candidate in (value_el.text, value_el.get("ID")):
                    if candidate and str(candidate).strip():
                        lookup[str(candidate).strip().upper()] = stibo_id

    log.info("[Step 2] Variant lookup built: %d entries.", len(lookup))
    return lookup


def load_latest_archived_xml():
    if not XML_ARCHIVE_BUCKET:
        log.warning("[Step 2] XML_ARCHIVE_BUCKET not set - cannot load archived XML.")
        return None

    prefix = XML_ARCHIVE_PREFIX.rstrip("/") + "/"
    try:
        paginator = s3.get_paginator("list_objects_v2")
        objects = []
        for page in paginator.paginate(Bucket=XML_ARCHIVE_BUCKET, Prefix=prefix):
            objects.extend(page.get("Contents", []))

        if not objects:
            log.warning("[Step 2] No archived XMLs found at s3://%s/%s", XML_ARCHIVE_BUCKET, prefix)
            return None

        latest = max(objects, key=lambda obj: obj["LastModified"])
        log.info("[Step 2] Loading archived XML: s3://%s/%s", XML_ARCHIVE_BUCKET, latest["Key"])
        response = s3.get_object(Bucket=XML_ARCHIVE_BUCKET, Key=latest["Key"])
        return response["Body"].read()
    except Exception as exc:
        log.error("[Step 2] Failed to load archived XML: %s", exc)
        return None


def resolve_stibo_id(article_code: str, lookup: dict) -> str:
    article_code = str(article_code or "").strip()
    if ARTICLE_ID_MODE == "direct":
        return article_code

    resolved = lookup.get(article_code.upper())
    if not resolved:
        log.warning(
            "[Step 2] Cannot resolve %r to a STIBO Product ID. Writing as-is.",
            article_code,
        )
        return article_code
    return resolved


# ─────────────────────────────────────────────────────────────────────────────
# BUILD STIBO XML
# ─────────────────────────────────────────────────────────────────────────────
GENERIC_ARTICLE_FIELDS = {
    "AT_SuggestedRetailPrice",
    "AT_ProposedRetailPrice",
    "AT_EstimatedLandedCost",
    "AT_BYIndicator",
    "AT_BYErrorMessage",
    "AT_Generic",
    "AT_IncomingMonth",
    "AT_BYGender",
    "AT_FOB",
    "AT_EcomIndicator",   # NEW — generic-level, Y/N normalized like AT_BYIndicator
}

VARIANT_ARTICLE_FIELDS = {
    "AT_Variant",
    "AT_IncomingMonth",
}

# Stibo attribute IDs that use Y/N LOV normalization (Yes/Y/True/1 -> "Y", else "N")
_YN_LOV_ATTRIBUTES = {
    "AT_BYIndicator",
    "AT_EcomIndicator",
}

# BY response field that carries the classification ID for CPL_BYHierarchy.
# Both "by_subcategory_id" and "by_subcategory_name" map to the same Stibo
# attribute ID in the mapping sheet (CPL_BYHierarchy), but only the ID
# column should be used as the ClassificationReference value.
_BY_SUBCATEGORY_ID_FIELD_KEY = _normalize_by_field("by_subcategory_id")
_CLASSIFICATION_TYPE_BY_SUBCATEGORY = "CPL_BYHierarchy"

_NO_SIZE_CARRY_FIELDS = GENERIC_ARTICLE_FIELDS - {"AT_Generic"}
_NO_SIZE_VALUES = frozenset({"no size", "000"})

def _is_no_size(size_str: str) -> bool:
    """Return True when the size string maps to the 'no size' LOV entry
    (display value = 'no size'  OR  LOV ID = '000')."""
    return str(size_str).strip().lower() in _NO_SIZE_VALUES

def _size_sort_key(s: str):
    """Sort numerically where possible, alphabetically otherwise."""
    try:
        return (0, float(s))
    except ValueError:
        return (1, s.lower())

def _transform_size_for_stibo(size: str) -> str:
    size = size.strip()
    half_size_match = re.match(r'^(\d+)H$', size)
    if half_size_match:
        return str(int(half_size_match.group(1)) + 0.5)
    neg_half_match = re.match(r'^-(\d+)H$', size)
    if neg_half_match:
        return str(-(int(neg_half_match.group(1)) + 0.5))
    return size

_GENDER_LOV_MAP = {
    "female": "F",
    "male":   "M",
    "unisex": "U",
}

def _normalize_gender_lov(raw: str) -> str:
    """Map incoming BY gender display value → Stibo AT_BYGender LOV ID."""
    return _GENDER_LOV_MAP.get(str(raw).strip().lower(), str(raw).strip())


def _normalize_yn_lov(raw: str) -> str:
    """Map incoming BY yes/no-style display value → Stibo Y/N LOV ID.
    Used for AT_BYIndicator and AT_EcomIndicator."""
    return "Y" if str(raw).strip().lower() in ("yes", "y", "true", "1") else "N"


def _add_by_subcategory_classification(product_el, record: dict):
    """Append a <ClassificationReference .../> directly under product_el
    (as a sibling of KeyValue/Values, no wrapping <Classifications> element)
    when a BY Subcategory ID is present on the record. CPL_BYHierarchy is
    a classification reference, not an attribute Value, so it's handled
    separately from the GENERIC_ARTICLE_FIELDS Value loop."""
    subcategory_id = record.get(_BY_SUBCATEGORY_ID_FIELD_KEY)
    if subcategory_id is None or str(subcategory_id).strip() == "":
        return

    etree.SubElement(
        product_el,
        "ClassificationReference",
        attrib={
            "ClassificationID": str(subcategory_id).strip(),
            "Type": _CLASSIFICATION_TYPE_BY_SUBCATEGORY,
        },
    )

def build_by_stibo_xml(records, lookup: dict, mapping_rules):
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
    size_field_key   = _by_size_field_key(mapping_rules)

    for record in records:
        identifier = _by_identifier_value(record, mapping_rules)
        if not identifier:
            skipped_count += 1
            continue

        # ── Size extraction — done BEFORE ID resolution so we can early-skip ──
        raw_sizes = record.get(size_field_key, "")

        # ── Rule 2: blank productSize → discard ───────────────────────────
        if not str(raw_sizes).strip():
            log.debug("[Step 2] Discarding record %r — blank productSize.", identifier)
            skipped_count += 1
            continue

        sizes = sorted(
            [s.strip() for s in re.split(r"[,\s]+", str(raw_sizes)) if s.strip()],
            key=_size_sort_key,
        )
        if not sizes:
            log.debug("[Step 2] Discarding record %r — no valid sizes after parsing.", identifier)
            skipped_count += 1
            continue

        # ── ID resolution ─────────────────────────────────────────────────
        stibo_id = resolve_stibo_id(identifier, lookup)

        if ARTICLE_ID_MODE == "variant" and identifier.upper() in lookup:
            resolved_count += 1
        elif ARTICLE_ID_MODE == "direct":
            resolved_count += 1
        else:
            unresolved_count += 1

        # ── Rule 1: no-size product → PRD_SingleArticle, NO generic ───────
        if len(sizes) == 1 and _is_no_size(sizes[0]):
            log.debug("[Step 2] No-size product %r → single PRD_SingleArticle.", stibo_id)

            single_el = etree.SubElement(
                products_el,
                "Product",
                attrib={"UserTypeID": "PRD_SingleArticle"},         # ← PRD_SingleArticle
            )
            skv_el = etree.SubElement(single_el, "KeyValue")
            skv_el.set("KeyID", "KEY_Article")                      # ← KEY_Article
            skv_el.text = stibo_id

            # CPL_BYHierarchy classification reference (if present on the record)
            # — placed right after KeyValue, before Values
            _add_by_subcategory_classification(single_el, record)

            ns_values_el = etree.SubElement(single_el, "Values")

            # Carry generic-level attributes (excluding AT_Generic) onto this single article
            for rule in mapping_rules:
                if rule.stibo_attr_id not in _NO_SIZE_CARRY_FIELDS:  # ← excludes AT_Generic
                    continue
                value = record.get(rule.field_key)
                if value is None or str(value).strip() == "":
                    continue
                if rule.stibo_attr_id in _YN_LOV_ATTRIBUTES:
                    etree.SubElement(
                        ns_values_el,
                        "Value",
                        attrib={"AttributeID": rule.stibo_attr_id, "ID": _normalize_yn_lov(value)},
                    )
                    continue
                
                if rule.stibo_attr_id == "AT_BYGender":
                    etree.SubElement(
                        ns_values_el,
                        "Value",
                        attrib={"AttributeID": "AT_BYGender", "ID": _normalize_gender_lov(value)},
                    )
                    continue
                
                if rule.stibo_attr_id == "AT_IncomingMonth":
                    value = _normalize_date_dd_mm_yyyy(str(value).strip(), rule.stibo_attr_id)
                v_el = etree.SubElement(
                    ns_values_el,
                    "Value",
                    attrib={"AttributeID": rule.stibo_attr_id},
                )
                v_el.text = str(value).strip()

            # AT_Size = LOV ID "000" (no size)
            etree.SubElement(
                ns_values_el,
                "Value",
                attrib={"AttributeID": "AT_Size", "ID": "000"},
            )
            # ← AT_Variant removed entirely

            continue   # ← skip generic-article block below

        # ── Generic article ───────────────────────────────────────────────
        generic_el = etree.SubElement(
            products_el,
            "Product",
            attrib={"UserTypeID": "PRD_GenericArticle"},
        )
        kv_el = etree.SubElement(generic_el, "KeyValue")
        kv_el.set("KeyID", "KEY_Article")
        kv_el.text = stibo_id

        # CPL_BYHierarchy classification reference (if present on the record)
        # — placed right after KeyValue, before Values
        _add_by_subcategory_classification(generic_el, record)

        values_el = etree.SubElement(generic_el, "Values")

        for rule in mapping_rules:
            if rule.stibo_attr_id not in GENERIC_ARTICLE_FIELDS:
                continue
            value = record.get(rule.field_key)
            if value is None or str(value).strip() == "":
                continue
            if rule.stibo_attr_id in _YN_LOV_ATTRIBUTES:
                etree.SubElement(
                    values_el,
                    "Value",
                    attrib={
                        "AttributeID": rule.stibo_attr_id,
                        "ID":          _normalize_yn_lov(value),
                    },
                )
                continue
            
            if rule.stibo_attr_id == "AT_BYGender":
                etree.SubElement(
                    values_el,
                    "Value",
                    attrib={"AttributeID": "AT_BYGender", "ID": _normalize_gender_lov(value)},
                )
                continue
            if rule.stibo_attr_id == "AT_IncomingMonth":
                value = _normalize_date_dd_mm_yyyy(str(value).strip(), rule.stibo_attr_id)
            value_el = etree.SubElement(
                values_el,
                "Value",
                attrib={"AttributeID": rule.stibo_attr_id},
            )
            value_el.text = str(value).strip()

        # ── Variant articles — one per size (sorted numerically) ──────────
        for size in sizes:
            lov_size        = _transform_size_for_stibo(size)
            safe_size       = lov_size.replace("/", "").replace(" ", "")
            variant_id      = "%s%s" % (stibo_id, safe_size)
            at_variant_text = "%s%s" % (stibo_id, lov_size.replace(" ", ""))
            variant_name    = "Size %s" % lov_size

            variant_el = etree.SubElement(
                generic_el,
                "Product",
                attrib={"UserTypeID": "PRD_VariantArticle"},
            )
            vkv_el = etree.SubElement(variant_el, "KeyValue")
            vkv_el.set("KeyID", "KEY_Variant")
            vkv_el.text = variant_id

            name_el = etree.SubElement(variant_el, "Name")
            name_el.text = variant_name

            variant_values_el = etree.SubElement(variant_el, "Values")

            etree.SubElement(
                variant_values_el,
                "Value",
                attrib={"AttributeID": "AT_Size", "ID": lov_size},
            )

            at_variant_el = etree.SubElement(
                variant_values_el,
                "Value",
                attrib={"AttributeID": "AT_Variant"},
            )
            at_variant_el.text = at_variant_text

            # AT_PrincipalSize — this variant's own size only, not the full list
            principal_size_el = etree.SubElement(
                variant_values_el,
                "Value",
                attrib={"AttributeID": "AT_PrincipalSize"},
            )
            principal_size_el.text = lov_size

            for rule in mapping_rules:
                if rule.stibo_attr_id not in VARIANT_ARTICLE_FIELDS:
                    continue
                value = record.get(rule.field_key)
                if value is None or str(value).strip() == "":
                    continue
                
                if rule.stibo_attr_id == "AT_IncomingMonth":
                    value = _normalize_date_dd_mm_yyyy(str(value).strip(), rule.stibo_attr_id)
                value_el = etree.SubElement(
                    variant_values_el,
                    "Value",
                    attrib={"AttributeID": rule.stibo_attr_id},
                )
                value_el.text = str(value).strip()

    xml_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        pretty_print=True,
    )
    log.info(
        "[Step 2] XML generated: records=%d resolved=%d unresolved=%d skipped=%d bytes=%d",
        len(records), resolved_count, unresolved_count, skipped_count, len(xml_bytes),
    )
    return xml_bytes, {
        "resolved":   resolved_count,
        "unresolved": unresolved_count,
        "skipped":    skipped_count,
    }

# ─────────────────────────────────────────────────────────────────────────────
# ARCHIVE & POST TO STIBO
# ─────────────────────────────────────────────────────────────────────────────
def archive_by_stibo_xml_to_s3(xml_bytes: bytes, filename: str):
    if not BY_STIBO_XML_OUTPUT_BUCKET:
        log.warning("[Step 2] BY_STIBO_XML_OUTPUT_BUCKET not set - skipping XML archive.")
        return None

    key = "%s/%s" % (BY_STIBO_XML_OUTPUT_PREFIX.rstrip("/"), filename)
    log.info("[Step 2] Archiving generated XML -> s3://%s/%s", BY_STIBO_XML_OUTPUT_BUCKET, key)
    s3.put_object(
        Bucket=BY_STIBO_XML_OUTPUT_BUCKET,
        Key=key,
        Body=xml_bytes,
        ContentType="application/xml",
    )
    return key


def _stibo_endpoint_by_url(filename: str) -> str:
    if not STIBO_ENDPOINT_BY:
        return ""
    if "?" in STIBO_ENDPOINT_BY:
        return STIBO_ENDPOINT_BY
    query = urllib.parse.urlencode({
        "fileName": filename,
        "context": STIBO_CONTEXT_ID,
        "workspace": STIBO_WORKSPACE_ID,
    })
    return "%s?%s" % (STIBO_ENDPOINT_BY.rstrip("/"), query)


def _get_stibo_bearer_token() -> str:
    if not STIBO_TOKEN_URL or not STIBO_CLIENT_ID or not STIBO_CLIENT_SECRET:
        log.error("[Step 2] STIBO token credentials not fully configured.")
        return ""

    payload = urllib.parse.urlencode({
        "grant_type":    STIBO_GRANT_TYPE,
        "client_id":     STIBO_CLIENT_ID,
        "client_secret": STIBO_CLIENT_SECRET,
    }).encode("utf-8")

    request = urllib.request.Request(
        url=STIBO_TOKEN_URL,
        data=payload,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = json.loads(response.read().decode("utf-8"))
        return body.get("access_token", "")
    except Exception as exc:
        log.error("[Step 2] Failed to fetch Stibo token: %s", exc)
        return ""


def post_by_xml_to_stibo(xml_bytes: bytes, filename: str):
    url = _stibo_endpoint_by_url(filename)
    if not url:
        log.warning("[Step 2] STIBO_ENDPOINT_BY not configured - skipping HTTP POST.")
        return {"posted": False, "status": None, "body": "STIBO_ENDPOINT_BY not configured"}

    token = _get_stibo_bearer_token()
    if not token:
        log.error("[Step 2] Could not obtain Stibo Bearer token - skipping Stibo POST.")
        return {"posted": False, "status": None, "body": "Could not obtain Stibo Bearer token"}

    log.info("[Step 2] Posting generated XML to Stibo: %s", url)
    request = urllib.request.Request(
        url=url,
        data=xml_bytes,
        headers={
            "Authorization": "Bearer %s" % token,
            "Content-Type": "application/octet-stream",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            status = response.status
            body = response.read().decode("utf-8", errors="replace")

        if status in (200, 201, 202):
            log.info("[Step 2] STIBO API upload success: HTTP %s", status)
            response_id = ""
            try:
                response_id = json.loads(body).get("id", "")
            except Exception:
                response_id = ""
            return {"posted": True, "status": status, "body": body, "id": response_id}

        log.error("[Step 2] STIBO API upload failure: HTTP %s - %s", status, body)
        return {"posted": False, "status": status, "body": body}

    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        log.error("[Step 2] STIBO API upload failure: HTTP %s - %s", exc.code, body)
        return {"posted": False, "status": exc.code, "body": body}
    except urllib.error.URLError as exc:
        log.error("[Step 2] STIBO API upload failure: %s", exc.reason)
        return {"posted": False, "status": None, "body": str(exc.reason)}
    except Exception as exc:
        log.error("[Step 2] STIBO API upload failure: %s", exc)
        return {"posted": False, "status": None, "body": str(exc)}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _record_by_response_issue(auditor: AuditLogger, record: dict, issue: str, mapping_rules):
    if auditor._cur_file is None:
        return
    auditor._cur_file["rows"].append({
        "identifier": _by_identifier_value(record, mapping_rules),
        "issue": issue,
        "record": record,
        "recorded_at": _utcnow_iso(),
    })


# ─────────────────────────────────────────────────────────────────────────────
# HANDLER  (called by router)
# ─────────────────────────────────────────────────────────────────────────────
def handle_by_response(src_bucket: str, response_key: str, auditor: AuditLogger) -> dict:
    """
    Full Step 2 pipeline: download response → parse → build XML → POST to Stibo.
    Called from the router in lambda_function.py.
    """
    local_response = TMP / response_key.replace("/", "_")
    records = []
    mapping_rules = []
    stem = Path(response_key).stem

    try:
        auditor.begin_file("s3://%s/%s" % (src_bucket, response_key))
        log.info("[Step 2] File received: s3://%s/%s", src_bucket, response_key)
        log.info("[Step 2] Downloading response file.")
        s3.download_file(src_bucket, response_key, str(local_response))

        mapping_rules = load_by_to_stibo_mapping(MAPPING_BUCKET, MAPPING_KEY_BY_TO_STIBO)
        records = parse_by_response(str(local_response), mapping_rules)
        if not records:
            auditor.end_file("skipped")
            return {"file": response_key, "status": "Skipped", "reason": "Empty response file"}

        lookup = {}
        if ARTICLE_ID_MODE == "variant":
            archived_xml = load_latest_archived_xml()
            if archived_xml:
                lookup = build_variant_lookup(archived_xml)
        else:
            log.info("[Step 2] ARTICLE_ID_MODE=direct - BY identifiers used directly.")

        for record in records:
            if not _by_identifier_value(record, mapping_rules):
                _record_by_response_issue(auditor, record, "missing_identifier", mapping_rules)
            elif ARTICLE_ID_MODE == "variant" and _by_identifier_value(record, mapping_rules).upper() not in lookup:
                _record_by_response_issue(auditor, record, "unresolved_identifier", mapping_rules)

        xml_bytes, counts = build_by_stibo_xml(records, lookup, mapping_rules)
        timestamp = datetime.datetime.now(_TZ_GMT7).strftime("%Y%m%d_%H%M%S")
        xml_filename = "by_stibo_response_%s_%s.xml" % (stem, timestamp)

        archived_key = archive_by_stibo_xml_to_s3(xml_bytes, xml_filename)
        log.info("[Step 2] XML archived: %s", archived_key or "(not archived)")

        post_result = post_by_xml_to_stibo(xml_bytes, xml_filename)

        response_id = post_result.get("id", "")
        auditor.set_stibo_response_id(response_id)
        if response_id:
            log.info("[Step 2] BGP ID received and stored: %s", response_id)

        auditor.end_file(
            "ok" if post_result.get("posted") else "failed",
            rows_produced=len(records),
            rows_upserted=counts["resolved"],
            rows_failed=counts["unresolved"] + counts["skipped"],
        )

        return {
            "file": response_key,
            "status": "Success" if post_result.get("posted") else "PostFailed",
            "records_total": len(records),
            "records_resolved": counts["resolved"],
            "records_unresolved": counts["unresolved"],
            "records_skipped": counts["skipped"],
            "archived_xml": archived_key,
            "stibo_posted": post_result.get("posted"),
            "stibo_status": post_result.get("status"),
            "stibo_response_id": post_result.get("id", ""),
        }

    except Exception as exc:
        log.error("[Step 2] Exception while processing %s: %s", response_key, exc)
        for record in records:
            _record_by_response_issue(auditor, record, "exception: %s" % exc, mapping_rules)
        auditor.end_file("failed")
        return {"file": response_key, "status": "Failed", "error": str(exc)}
    finally:
        local_response.unlink(missing_ok=True)
        (TMP / "mapping_by_to_stibo.xlsx").unlink(missing_ok=True)