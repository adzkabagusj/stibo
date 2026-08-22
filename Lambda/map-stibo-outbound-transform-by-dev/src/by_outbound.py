"""
by_outbound.py
==============
Step 1 — Stibo STEP XML → Blue Yonder (Outbound)

Reads XML from S3 (step-export/blueyonder/), transforms every product
into rows using the "BY Assortment Tempelate" mapping Excel, upserts
to SQL Server, and writes a pipe-delimited CSV to S3.
"""

import boto3
import base64
import csv
import datetime
import json
import logging
import os
import re
import sys
import unicodedata
import pyodbc
from io import StringIO
from pathlib import Path

from lxml import etree
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────
log = logging.getLogger("lambda_by_etl")

# ─────────────────────────────────────────────────────────────────────────────
# AWS CLIENTS
# ─────────────────────────────────────────────────────────────────────────────
s3             = boto3.client("s3")
secretsmanager = boto3.client("secretsmanager")

# ─────────────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ─────────────────────────────────────────────────────────────────────────────
SECRET_NAME      = os.environ.get("MAA_STIBO_SECRET_MANAGER", "")
MAPPING_BUCKET   = os.environ.get("MAPPING_BUCKET", "")
MAPPING_KEY      = os.environ.get("MAPPING_KEY", "")
MAPPING_SHEET    = os.environ.get("MAPPING_SHEET", "BY Assortment Tempelate")
PREFIX           = os.environ.get("PREFIX", "step-export/blueyonder/")
OUTPUT_BUCKET    = os.environ.get("OUTPUT_BUCKET") or MAPPING_BUCKET
OUTPUT_PREFIX    = os.environ.get("OUTPUT_PREFIX", "transformed/blueyonder/")
AUDIT_LOG_BUCKET = os.environ.get("AUDIT_LOG_BUCKET") or OUTPUT_BUCKET
AUDIT_LOG_PREFIX = os.environ.get("AUDIT_LOG_PREFIX", "etl-audit-logs/blueyonder")
XML_ARCHIVE_BUCKET = os.environ.get("XML_ARCHIVE_BUCKET") or OUTPUT_BUCKET
XML_ARCHIVE_PREFIX = os.environ.get("XML_ARCHIVE_PREFIX", "xml-archive/")

TMP = Path("/tmp")

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
HEADER_MARKER      = "S. No."
BY_HIERARCHY_TYPE  = "CPL_BYHierarchy"
SAP_HIERARCHY_TYPE = "CPL_SAPHierarchy"
LOV_ID             = "ID"
LOV_TEXT           = "TEXT"

# DB table name/schema are configurable via ENV so the same code can point at
# different environments (dev/test/prod) without a code change.
DB_TABLE_NAME = os.environ.get("DB_TABLE_NAME")
DB_TABLE_SCHEMA = os.environ.get("DB_TABLE_SCHEMA")
DB_TABLE     = os.environ.get(
    "DB_TABLE",
    f"[{DB_TABLE_SCHEMA}].[{DB_TABLE_NAME}]",
)
UPSERT_KEYS  = ["Brand_Principal_Gen_Article", "BY_Sub_Category_ID", "SBUCode"]
_DB_SKIP_COLS = {"InsertedDate", "created_img_at", "updated_img_at"}

_DEFAULT_TO_RE   = re.compile(r"^default\s+to\s+(.+)$", re.IGNORECASE)
_UPTO_RE         = re.compile(r"upto\s+(\d+)", re.IGNORECASE)
_BY_HIER_ID_RE   = re.compile(r"by\s+hierarchy.*\bID\b", re.IGNORECASE)
_BY_HIER_NAME_RE = re.compile(r"by\s+hierarchy.*\bname\b", re.IGNORECASE)
_BY_HIER_OBJ_RE  = re.compile(
    r"object\s+type\s*[:\-]?\s*(SBU|Brand|Division|Sub\s*Division|Category|Sub\s*Category)",
    re.IGNORECASE,
)
_SBU_JOIN_SEP = "/"

_SQL_FIELD_RENAMES = {
    "Price_Range": "PriceRange",
    "Compcode":    "compcode",
}
_SBU_CANONICAL_NAMES = {
    "sbucode": "SBUCode",
    "sbuname": "SBUName",
}

# Candidate strptime formats tried in order when normalising date fields.
_DATE_PARSE_FORMATS = (
    "%Y-%m-%d",
    "%m/%d/%Y",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y/%m/%d",
    "%m-%d-%Y",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%dT%H:%M:%S.%f",
    "%d%m%Y",   
    "%Y%m%d",     
    "%d %m %Y",    
    "%m %d %Y",     
    "%d/%m/%Y %H:%M:%S",  
)

# Fallback attr mappings for P_Hierarchy_L1-L5.
# Used in _get_rules() only when these sql_fields are absent from the Excel.
_PRINCIPAL_HIER_FALLBACK_RULES = (
    ("P_Hierarchy_L1", "AT_PrincipalMerchandiseHierarchyL1"),
    ("P_Hierarchy_L2", "AT_PrincipalMerchandiseHierarchyL2"),
    ("P_Hierarchy_L3", "AT_PrincipalMerchandiseHierarchyL3"),
    ("P_Hierarchy_L4", "AT_PrincipalMerchandiseHierarchyL4"),
    ("P_Hierarchy_L5", "AT_PrincipalMerchandiseHierarchyL5"),
)

# Fallback attr mappings for SBUCode / SBUName (BY Hierarchy — SBU level).
# Used in _get_rules() only when these sql_fields are absent from the Excel.
# Each entry: (sql_field, stibo_attr, stibo_type, comment)
#   stibo_type = "List of Value" + comment "Value ID" / "Value" drives the
#   same lov_extract derivation as normal Excel-sourced rules
#   (see _derive_lov_extract).
_SBU_FALLBACK_RULES = (
    ("SBUCode", "AT_SBU", "List of Value", "Value ID"),
    ("SBUName", "AT_SBU", "List of Value", "Value"),
)

# ─────────────────────────────────────────────────────────────────────────────
# TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────
_TZ_GMT7 = datetime.timezone(datetime.timedelta(hours=7))


# ─────────────────────────────────────────────────────────────────────────────
# SECRETS / DB
# ─────────────────────────────────────────────────────────────────────────────
_cached_secret = None


def _get_db_credentials() -> dict:
    global _cached_secret
    if _cached_secret is not None:
        log.info("Using cached DB credentials")
        return _cached_secret

    log.info("Fetching credentials from Secrets Manager: %s", SECRET_NAME)
    response = secretsmanager.get_secret_value(SecretId=SECRET_NAME)
    secret   = json.loads(response["SecretString"])

    username = secret.get("username")
    password = secret.get("password")
    if (not username or not password) and secret.get("app.stibo"):
        username = "app.stibo"
        password = secret["app.stibo"]

    if not username or not password:
        raise ValueError(
            f"Secret '{SECRET_NAME}' must contain 'username'/'password' "
            f"or an 'app.stibo' credential entry. "
        )
    _cached_secret = {"username": username, "password": password}
    return _cached_secret


def _get_connection():
    creds    = _get_db_credentials()
    server   = os.environ.get("DB_SERVER")
    database = os.environ.get("DB_NAME")
    if not server or not database:
        raise ValueError("DB_SERVER or DB_NAME missing from environment variables")

    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={server},1433;"
        f"DATABASE={database};"
        f"UID={creds['username']};"
        f"PWD={creds['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
        "Connection Timeout=30;"
    )
    try:
        log.info("Connecting to DB via SQL auth ...")
        conn = pyodbc.connect(conn_str, timeout=10)
        log.info("✅ DB connection successful (SQL auth)")
        verify_cursor = conn.cursor()
        verify_cursor.execute("SELECT DB_NAME(), USER_NAME()")
        db_name, user_name = verify_cursor.fetchone()
        verify_cursor.close()
        return conn
    except Exception as e:
        err_str = str(e)
        err_str = re.sub(r"(PWD=)[^;]+", r"\1***", err_str)
        err_str = re.sub(r"(UID=)[^;]+", r"\1***", err_str)
        log.error("❌ DB connection FAILED: %s", err_str)
        raise


# ─────────────────────────────────────────────────────────────────────────────
# AUDIT LOGGER
# ─────────────────────────────────────────────────────────────────────────────
def _utcnow_iso() -> str:
    return datetime.datetime.now(_TZ_GMT7).isoformat()


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
# MAPPING RULE
# ─────────────────────────────────────────────────────────────────────────────
class MappingRule:
    __slots__ = (
        "seq", "stibo_name", "stibo_attr", "stibo_type",
        "comment", "sql_field", "field_desc", "field_type", "nullable",
        "lov_extract",
    )

    def __init__(self, seq, stibo_name, stibo_attr, stibo_type,
                 comment, sql_field, field_desc, field_type, nullable):
        self.seq         = seq
        self.stibo_name  = stibo_name
        self.stibo_attr  = stibo_attr
        self.stibo_type  = stibo_type
        self.comment     = comment
        self.sql_field   = sql_field
        self.field_desc  = field_desc
        self.field_type  = field_type
        self.nullable    = nullable
        self.lov_extract = _derive_lov_extract(stibo_type, comment)

    def __repr__(self):
        return "<Rule #%s  %r [%s] -> %r  lov_extract=%r>" % (
            self.seq, self.stibo_attr, self.stibo_type, self.sql_field, self.lov_extract)


def _derive_lov_extract(stibo_type: str, comment: str) -> str:
    type_is_lov = bool(stibo_type and "list" in stibo_type.lower())
    if not type_is_lov:
        return None
    c = (comment or "").strip().lower()
    if c in ("value id", "value id ") or c.startswith("value id"):
        return LOV_ID
    return LOV_TEXT


# ─────────────────────────────────────────────────────────────────────────────
# WARM LAMBDA CACHING
# ─────────────────────────────────────────────────────────────────────────────
_cached_rules: list = None


# ─────────────────────────────────────────────────────────────────────────────
# FIELD NORMALISATION
# ─────────────────────────────────────────────────────────────────────────────
def _normalize_sql_field(name: str) -> str:
    if not name:
        return name
    name = unicodedata.normalize("NFKC", name).strip()
    name = re.sub(r"[\s\-]+", "_", name)
    name = re.sub(r"[^\w]", "", name)
    name = re.sub(r"_+", "_", name)
    name = name.strip("_")
    name = _SQL_FIELD_RENAMES.get(name, name)

    # Force any SBU-code/name spelling variant to the canonical name so the
    # per-row differentiation logic in process_xml_file() reliably matches.
    canon_key = name.lower().replace("_", "")
    name = _SBU_CANONICAL_NAMES.get(canon_key, name)

    return name


# ─────────────────────────────────────────────────────────────────────────────
# MAPPING LOADER
# ─────────────────────────────────────────────────────────────────────────────
def load_mapping(excel_path: str) -> list:
    wb = load_workbook(excel_path, read_only=True)
    if MAPPING_SHEET not in wb.sheetnames:
        raise ValueError(
            "Sheet %r not found in workbook. Available sheets: %s"
            % (MAPPING_SHEET, wb.sheetnames)
        )

    ws           = wb[MAPPING_SHEET]
    rules        = []
    data_started = False

    for row in ws.iter_rows(values_only=True):
        if not data_started:
            if str(row[0]).strip() == HEADER_MARKER:
                data_started = True
            continue

        seq = row[0]
        if seq is None:
            continue

        sql_field = _normalize_sql_field(str(row[5] or "").strip()) or None
        if not sql_field:
            continue

        raw_comment    = str(row[4] or "").strip()
        raw_stibo_name = str(row[1] or "").strip()

        if not raw_comment and raw_stibo_name.lower().startswith("default"):
            effective_comment = raw_stibo_name
        else:
            effective_comment = raw_comment

        rules.append(MappingRule(
            seq        = seq,
            stibo_name = raw_stibo_name or None,
            stibo_attr = str(row[2] or "").strip() or None,
            stibo_type = str(row[3] or "").strip() or None,
            comment    = effective_comment or None,
            sql_field  = sql_field,
            field_desc = str(row[6] or "").strip() or None,
            field_type = str(row[7] or "").strip() or None,
            nullable   = str(row[8] or "").strip() or None,
        ))

    log.info("Loaded %d mapping rules from sheet %r", len(rules), MAPPING_SHEET)
    for r in rules:
        log.info("  RULE: sql_field=%-35s  attr=%-25s  stibo_type=%-20s  comment=%-20r  lov_extract=%r",
                 r.sql_field, r.stibo_attr or "(none)", r.stibo_type or "(none)",
                 r.comment, r.lov_extract)
    return rules


def _get_rules() -> list:
    global _cached_rules
    if _cached_rules is not None:
        log.info("Using cached mapping rules (%d rules)", len(_cached_rules))
        return _cached_rules

    local_path = TMP / "by_mapping.xlsx"
    log.info("Downloading mapping Excel: s3://%s/%s", MAPPING_BUCKET, MAPPING_KEY)
    s3.download_file(MAPPING_BUCKET, MAPPING_KEY, str(local_path))
    _cached_rules = load_mapping(str(local_path))

    # ── P_Hierarchy_L1-L5 fallback ────────────────────────────────────────
    # If the mapping Excel already contains these sql_fields they are skipped.
    # If not (Excel not yet updated), they are appended so no data is lost.
    existing_fields = {r.sql_field for r in _cached_rules}
    next_seq = max(
        (r.seq for r in _cached_rules if isinstance(r.seq, (int, float))),
        default=0,
    ) + 1

    for sql_field, attr_id in _PRINCIPAL_HIER_FALLBACK_RULES:
        if sql_field in existing_fields:
            log.info(
                "  [P_Hier fallback] %-20s already in Excel — no injection needed",
                sql_field,
            )
            continue
        log.info(
            "  [P_Hier fallback] Injecting %-20s → %s", sql_field, attr_id
        )
        _cached_rules.append(MappingRule(
            seq        = next_seq,
            stibo_name = None,
            stibo_attr = attr_id,
            stibo_type = "Text",
            comment    = None,
            sql_field  = sql_field,
            field_desc = None,
            field_type = None,
            nullable   = None,
        ))
        next_seq += 1

    # ── SBUCode / SBUName fallback ────────────────────────────────────────
    # If the mapping Excel already contains these sql_fields they are skipped.
    # If not (Excel not yet updated), they are appended so no data is lost.
    existing_fields = {r.sql_field for r in _cached_rules}
    for sql_field, attr_id, stibo_type, comment in _SBU_FALLBACK_RULES:
        if sql_field in existing_fields:
            log.info(
                "  [SBU fallback] %-20s already in Excel — no injection needed",
                sql_field,
            )
            continue
        log.info(
            "  [SBU fallback] Injecting %-20s → %s (%s / %s)",
            sql_field, attr_id, stibo_type, comment,
        )
        _cached_rules.append(MappingRule(
            seq        = next_seq,
            stibo_name = "SBU",
            stibo_attr = attr_id,
            stibo_type = stibo_type,
            comment    = comment,
            sql_field  = sql_field,
            field_desc = None,
            field_type = None,
            nullable   = None,
        ))
        next_seq += 1

    return _cached_rules


# ─────────────────────────────────────────────────────────────────────────────
# XML VALUE READING HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _read_attr(product_el, attr_id: str, lov_extract: str):
    if not attr_id:
        return None
    ve = product_el.find("Values")
    if ve is None:
        return None
    for v in ve.findall("Value"):
        if v.get("AttributeID") == attr_id:
            return _extract_from_value_el(v, lov_extract)
    for mv in ve.findall("MultiValue"):
        if mv.get("AttributeID") == attr_id:
            first = mv.find("Value")
            if first is not None:
                return _extract_from_value_el(first, lov_extract)
    return None


def _extract_from_value_el(v_el, lov_extract: str):
    if lov_extract == LOV_ID:
        return (v_el.get("ID") or "").strip() or None
    text = (v_el.text or "").strip()
    if lov_extract == LOV_TEXT:
        return text or None
    if text:
        return text
    return (v_el.get("ID") or "").strip() or None


def _read_asset_cross_reference(product_el, ref_type: str = "Primary Image"):
    for acr in product_el.findall("AssetCrossReference"):
        if acr.get("Type") == ref_type:
            asset_id = acr.get("AssetID", "").strip()
            return asset_id or None
    return None


def _extract_all_by_hierarchies(product_el, hierarchy_lookup: dict):
    results = []
    seen_hierarchies = set()

    for cr in product_el.findall("ClassificationReference"):
        if cr.get("Type") != BY_HIERARCHY_TYPE:
            continue
        leaf_id = cr.get("ClassificationID", "")
        if not leaf_id or leaf_id not in hierarchy_lookup:
            continue

        hierarchy = {}
        current_id = leaf_id

        while current_id and current_id in hierarchy_lookup:
            node = hierarchy_lookup[current_id]
            node_type = node.get("type", "")
            node_id = node.get("id", "")
            node_name = node.get("name")

            if node_type == "CLS_SBU":
                hierarchy["BY_SBU_ID"] = node_id
            elif node_type == "CLS_Brand":
                hierarchy["BY_Brand_ID"] = node_id
            elif node_type == "CLS_BYDivision":
                hierarchy["BY_Division_ID"] = node_id
                hierarchy["BY_Division_Name"] = node_name
            elif node_type == "CLS_BYSubDivision":
                hierarchy["BY_Sub_Division_ID"] = node_id
                hierarchy["BY_Sub_Division_Name"] = node_name
            elif node_type == "CLS_BYCategory":
                hierarchy["BY_Category_ID"] = node_id
                hierarchy["BY_Category_Name"] = node_name
            elif node_type == "CLS_BYSubCategory":
                hierarchy["BY_Sub_Category_ID"] = node_id
                hierarchy["BY_Sub_Category_Name"] = node_name

            current_id = node.get("parent_id")

        if hierarchy:
            unique_key = hierarchy.get("BY_Sub_Category_ID")
            if unique_key and unique_key not in seen_hierarchies:
                seen_hierarchies.add(unique_key)
                results.append(hierarchy)
                log.debug("  ✓ Added hierarchy: BY_Sub_Category_ID=%s", unique_key)
            elif unique_key:
                log.debug("  ⊗ Skipped duplicate hierarchy: BY_Sub_Category_ID=%s", unique_key)

    log.info("  → Extracted %d unique hierarchy(ies)", len(results))
    return results if results else []


def _build_hierarchy_lookup(root_el) -> dict:
    lookup = {}
    classifications_el = root_el.find("Classifications")
    if classifications_el is None:
        log.info("No <Classifications> block — hierarchy lookup will be empty")
        return lookup

    for cls_el in classifications_el.iter("Classification"):
        cid     = cls_el.get("ID", "")
        name_el = cls_el.find("Name")
        parent    = cls_el.getparent()
        parent_id = (
            parent.get("ID")
            if (parent is not None and parent.tag == "Classification")
            else None
        )
        lookup[cid] = {
            "id":        cid,
            "name":      (name_el.text or "").strip() if name_el is not None else None,
            "parent_id": parent_id,
            "type":      cls_el.get("UserTypeID", ""),
        }

    log.info("Built hierarchy lookup: %d entries", len(lookup))
    return lookup


def _get_sap_division_name(product_el, hierarchy_lookup: dict) -> str:
    for cr in product_el.findall("ClassificationReference"):
        if cr.get("Type") != SAP_HIERARCHY_TYPE:
            continue
        current_id = cr.get("ClassificationID", "")
        while current_id and current_id in hierarchy_lookup:
            node = hierarchy_lookup[current_id]
            if node.get("type") == "CLS_SAPDivision":
                return node.get("name") or current_id
            current_id = node.get("parent_id")
        return None
    return None


def _build_season(product_el) -> str:
    season_id   = _read_attr(product_el, "AT_Season",     LOV_ID)
    season_year = _read_attr(product_el, "AT_SeasonYear", None)
    season_id   = season_id   or ""
    season_year = season_year or ""
    year_suffix = season_year[-2:] if len(season_year) >= 2 else season_year
    return (season_id + year_suffix) or None


def _parse_default_to(comment: str):
    if not comment:
        return None
    m = _DEFAULT_TO_RE.match(comment.strip())
    if not m:
        return None
    val = m.group(1).strip()
    if "if not available" in val.lower():
        return None
    return "" if val.lower().startswith("blank") else val

def _read_attr_all(product_el, attr_id: str, lov_extract: str) -> list:
    """Return every value for *attr_id*, including all MultiValue entries."""
    if not attr_id:
        return []
    ve = product_el.find("Values")
    if ve is None:
        return []
    results = []
    for v in ve.findall("Value"):
        if v.get("AttributeID") == attr_id:
            val = _extract_from_value_el(v, lov_extract)
            if val:
                results.append(val)
    for mv in ve.findall("MultiValue"):
        if mv.get("AttributeID") == attr_id:
            for v in mv.findall("Value"):
                val = _extract_from_value_el(v, lov_extract)
                if val:
                    results.append(val)
    return results


def _which_by_hierarchy_object(comment: str) -> str:
    if not comment:
        return None
    m = _BY_HIER_OBJ_RE.search(comment)
    if not m:
        return None
    raw = m.group(1).strip().upper().replace(" ", "")
    return {
        "SBU":         "SBU",
        "BRAND":       "Brand",
        "DIVISION":    "Division",
        "SUBDIVISION": "Sub_Division",
        "CATEGORY":    "Category",
        "SUBCATEGORY": "Sub_Category",
    }.get(raw, raw)


def _normalize_date_dd_mm_yyyy(raw: str, field: str) -> str:
    """Try each format in *_DATE_PARSE_FORMATS* and reformat to dd-MM-yyyy.

    Returns the original string unchanged (and logs a warning) when no
    candidate format matches — this prevents silently dropping the value.
    """
    stripped = raw.strip()
    for fmt in _DATE_PARSE_FORMATS:
        try:
            return datetime.datetime.strptime(stripped, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    log.warning(
        "⚠️  Could not parse date value %r for field %r "
        "— none of the %d candidate formats matched; keeping original value",
        stripped, field, len(_DATE_PARSE_FORMATS),
    )
    return stripped

def _split_sbu_string(s: str) -> list:
    """Split a possibly slash-joined SBU string into individual trimmed parts.
    Handles both 'SP/K2/TF' and 'MAA Sport / Foot Locker Malaysia / Footlocker Thailand'.
    """
    if not s:
        return []
    parts = [p.strip() for p in s.split(_SBU_JOIN_SEP)]
    return [p for p in parts if p]


def _get_sbu_pairs(product_el, attr_id: str = "AT_SBU") -> list:
    """Return a list of (sbu_code, sbu_name) tuples for all SBU values.

    Handles two XML shapes transparently:
      1. True MultiValue with one Value element per SBU (already distinct).
      2. A single Value/MultiValue entry containing multiple SBUs joined by
         `_SBU_JOIN_SEP` (e.g. 'SP/K2/TF'), which must be split further.
    """
    codes = _read_attr_all(product_el, attr_id, LOV_ID)
    names = _read_attr_all(product_el, attr_id, LOV_TEXT)

    flat_codes = []
    for c in codes:
        flat_codes.extend(_split_sbu_string(c))

    flat_names = []
    for n in names:
        flat_names.extend(_split_sbu_string(n))

    if not flat_codes:
        return [("", "")]

    return [
        (flat_codes[i], flat_names[i] if i < len(flat_names) else "")
        for i in range(len(flat_codes))
    ]

# ─────────────────────────────────────────────────────────────────────────────
# FIELD RESOLVER
# ─────────────────────────────────────────────────────────────────────────────
def _resolve_field(rule: MappingRule, product_el, by_hierarchy: dict,
                   hierarchy_lookup: dict, root):
    comment     = rule.comment or ""
    sql_field   = rule.sql_field or ""
    attr_id     = rule.stibo_attr
    lov_extract = rule.lov_extract

    default_val = _parse_default_to(comment)
    if default_val is not None:
        return default_val

    if _BY_HIER_ID_RE.search(comment) or _BY_HIER_NAME_RE.search(comment):
        obj_type  = _which_by_hierarchy_object(comment)
        want_name = bool(_BY_HIER_NAME_RE.search(comment))
        if obj_type:
            key = "BY_%s_%s" % (obj_type, "Name" if want_name else "ID")
            return by_hierarchy.get(key) or ""
        return by_hierarchy.get("BY_Sub_Category_ID") or ""

    if sql_field == "Div":
        return _get_sap_division_name(product_el, hierarchy_lookup) or ""

    if sql_field == "Season":
        return _build_season(product_el) or ""

    if sql_field in ("Image", "IMAGE_BINARY"):
        asset_id = None
        for acr in product_el.findall("AssetCrossReference"):
            asset_id = acr.get("AssetID")
            if asset_id:
                break
        if not asset_id:
            return None if sql_field == "IMAGE_BINARY" else ""

        asset = root.find(f".//Asset[@ID='{asset_id}']")
        log.info("AssetID: %s", asset_id)
        log.info("Asset FOUND: %s", asset is not None)

        if asset is None:
            return None if sql_field == "IMAGE_BINARY" else ""

        abc = asset.find("AssetBinaryContent")
        if abc is None:
            return None if sql_field == "IMAGE_BINARY" else ""

        if sql_field == "Image":
            return abc.get("Filename", "") or ""

        if sql_field == "IMAGE_BINARY":
            binary = abc.find("BinaryContent")
            if binary is not None and binary.text:
                try:
                    return base64.b64decode(binary.text.strip())
                except Exception as e:
                    log.error("Binary decode failed: %s", e)
                    return None
            return None

    # ── Date normalisation ────────────────────────────────────────────────
    # AT_IncomingMonth may arrive in many formats; always emit dd-MM-yyyy.
    if attr_id == "AT_IncomingMonth":
        raw = _read_attr(product_el, attr_id, lov_extract)
        if not raw:
            return ""
        return _normalize_date_dd_mm_yyyy(raw, sql_field)
        
    if sql_field == "Original_In_Store_Date":
        raw = _read_attr(product_el, attr_id or "AT_IncomingMonth", lov_extract)
        if not raw:
            return ""
        return _normalize_date_dd_mm_yyyy(raw, sql_field)

    if attr_id:
        raw = _read_attr(product_el, attr_id, lov_extract)
        if raw is None:
            return ""
        m = _UPTO_RE.search(comment)
        if m:
            raw = raw[:int(m.group(1))]
        return raw

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# XML FILE PROCESSING
# ─────────────────────────────────────────────────────────────────────────────
def process_xml_file(xml_path: str, rules: list) -> list:
    log.info("")
    log.info("=" * 70)
    log.info("  FILE: %s", xml_path)
    log.info("=" * 70)

    try:
        tree = etree.parse(xml_path)
    except etree.XMLSyntaxError as exc:
        log.error("XML parse error: %s", exc)
        return []

    root        = tree.getroot()
    products_el = root.find("Products")

    if products_el is None:
        log.warning("No <Products> element found — skipping.")
        return []

    hierarchy_lookup = _build_hierarchy_lookup(root)
    rows = []
    
    for r in rules:
        if "sbu" in (r.sql_field or "").lower():
            log.info("SBU rule: sql_field=%r  stibo_attr=%r  stibo_type=%r  comment=%r",
                     r.sql_field, r.stibo_attr, r.stibo_type, r.comment)

    for product in products_el.findall("Product"):
        pid = product.get("ID", "?")
        log.info("  > Product ID=%s  Type=%s", pid, product.get("UserTypeID", ""))
    
        all_hierarchies = _extract_all_by_hierarchies(product, hierarchy_lookup)
    
        if not all_hierarchies:
            log.info("  ℹ️ Product %s has no BY hierarchies — producing row with blank hierarchy fields", pid)
            all_hierarchies = [{}]   # single empty hierarchy dict so the row still gets built
    
        sbu_attr_id = next(
            (r.stibo_attr for r in rules if r.sql_field == "SBUCode"),
            "AT_SBU",
        )
        sbu_pairs = _get_sbu_pairs(product, sbu_attr_id or "AT_SBU")
        log.info("  > Product %s has %d SBU pair(s): %s", pid, len(sbu_pairs), sbu_pairs)
    
        for hierarchy in all_hierarchies:
            for sbu_code, sbu_name in sbu_pairs:
                row = {}
                for rule in rules:
                    if rule.sql_field == "SBUCode":
                        row[rule.sql_field] = sbu_code
                    elif rule.sql_field == "SBUName":
                        row[rule.sql_field] = sbu_name
                    else:
                        row[rule.sql_field] = _resolve_field(
                            rule, product, hierarchy, hierarchy_lookup, root
                        )
                rows.append(row)

    log.info("  done  %s  ->  %d row(s) produced", Path(xml_path).name, len(rows))
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# CSV SERIALISATION  (pipe-delimited)
# ─────────────────────────────────────────────────────────────────────────────
def rows_to_pipe_csv(rows: list, rules: list) -> str:
    columns = [r.sql_field for r in rules if r.sql_field]
    buf = StringIO()
    writer = csv.writer(buf, delimiter="|", lineterminator="\n",
                        quoting=csv.QUOTE_MINIMAL)
    writer.writerow(columns)
    for row in rows:
        writer.writerow([str(row.get(col, "") or "") for col in columns])
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE UPSERT
# ─────────────────────────────────────────────────────────────────────────────
def _upsert_rows(rows: list, rules: list, auditor: AuditLogger) -> None:
    if not rows:
        log.info("No rows to upsert — skipping DB write.")
        return

    conn   = _get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ?
          AND TABLE_NAME   = ?
    """, (DB_TABLE_SCHEMA, DB_TABLE_NAME))
    db_columns_info = {row[0]: (row[1].lower(), row[2] == "YES") for row in cursor.fetchall()}
    db_columns      = set(db_columns_info.keys())
    db_columns_ci   = {col.lower(): col for col in db_columns}
    log.info("DB table has %d columns", len(db_columns))

    def _match_db_col(field: str):
        if field in db_columns:
            return field
        return db_columns_ci.get(field.lower())

    columns = [
        _match_db_col(r.sql_field) for r in rules
        if r.sql_field
        and r.sql_field not in _DB_SKIP_COLS
        and _match_db_col(r.sql_field) is not None
    ]

    if not columns:
        log.error("No rule fields matched any DB column — skipping upsert entirely.")
        return

    col_to_rule_field = {}
    for r in rules:
        if r.sql_field and r.sql_field not in _DB_SKIP_COLS:
            db_col = _match_db_col(r.sql_field)
            if db_col:
                col_to_rule_field[db_col] = r.sql_field

    upsert_keys = [k for k in UPSERT_KEYS if _match_db_col(k) is not None]
    update_cols = [c for c in columns if c not in upsert_keys]

    # Bracketed column list used in INSERT and as the src alias list
    col_list = ", ".join(f"[{c}]" for c in columns)

    # USING VALUES params — keep CONVERT wrapper for IMAGE_BINARY so SQL Server
    # infers VARBINARY(MAX) rather than nvarchar for the column type.
    using_params = ", ".join(
        "CONVERT(VARBINARY(MAX), ?)" if c == "IMAGE_BINARY" else "?"
        for c in columns
    )

    # INSERT references src alias columns — no second set of ? needed.
    insert_src = ", ".join(f"src.[{c}]" for c in columns)

    on_clause = " AND ".join(f"tgt.[{k}] = src.[{k}]" for k in upsert_keys)

    # Guard: if every column is a key there are no update_cols; avoid "UPDATE SET ,"
    if update_cols:
        set_clause     = ",\n              ".join(f"[{c}] = src.[{c}]" for c in update_cols)
        matched_action = (
            f"UPDATE SET\n"
            f"              {set_clause},\n"
            f"              [updated_img_at]  = SYSUTCDATETIME(),\n"
            f"              [created_img_at]  = COALESCE(tgt.[created_img_at], SYSUTCDATETIME())"
        )
    else:
        matched_action = (
            "UPDATE SET [updated_img_at] = SYSUTCDATETIME(), "
            "[created_img_at] = COALESCE(tgt.[created_img_at], SYSUTCDATETIME())"
        )

    # ── FIX: use USING (VALUES (...)) AS src(...) directly ────────────────
    # The previous pattern — SELECT col FROM (VALUES (?)) AS v(col) — triggers
    # SQL Server error 156 "Incorrect syntax near the keyword 'FROM'" when
    # executed via pyodbc MERGE.  The VALUES-as-source form is the correct
    # SQL Server idiom and requires params only once (not doubled).
    merge_sql = f"""
    MERGE {DB_TABLE} AS tgt
    USING (VALUES ({using_params})) AS src({col_list})
      ON  {on_clause}
    WHEN MATCHED THEN
        {matched_action}
    WHEN NOT MATCHED THEN
        INSERT ({col_list}, [InsertedDate], [created_img_at])
        VALUES ({insert_src}, SYSUTCDATETIME(), SYSUTCDATETIME());
    """

    log.info("Upserting %d row(s) into %s", len(rows), DB_TABLE)
    log.debug("MERGE SQL:\n%s", merge_sql)

    def _prepare_value(val, db_col):
        if isinstance(val, bytes):
            return val
        if val == "" and db_col and db_col in db_columns_info:
            data_type, is_nullable = db_columns_info[db_col]
            if is_nullable and any(
                t in data_type for t in ["int", "numeric", "decimal", "float", "bit"]
            ):
                return None
        if val is None:
            return None
        return str(val)

    rows_upserted = 0
    rows_failed   = 0

    for row in rows:
        params = tuple(
            _prepare_value(row.get(col_to_rule_field.get(c, c)), c)
            for c in columns
        )
        try:
            cursor.execute(merge_sql, params)   # ← single params tuple, not doubled
            rows_upserted += 1
            auditor.record_row(row=row, warnings=[], db_status="upserted")
        except Exception as row_exc:
            err_msg = str(row_exc)
            log.error("Row upsert failed: %s", err_msg)
            rows_failed += 1
            auditor.record_row(row=row, warnings=[], db_status="failed", db_error=err_msg)
            conn.rollback()
            cursor = conn.cursor()

    conn.commit()
    cursor.close()
    conn.close()
    log.info("Upsert complete — succeeded: %d  failed: %d", rows_upserted, rows_failed)

# ─────────────────────────────────────────────────────────────────────────────
# S3 UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
def upload_output(content: str, s3_key: str) -> None:
    log.info("Uploading -> s3://%s/%s  (%d bytes)", OUTPUT_BUCKET, s3_key, len(content))
    s3.put_object(
        Bucket      = OUTPUT_BUCKET,
        Key         = s3_key,
        Body        = content.encode("utf-8"),
        ContentType = "text/csv",
    )
    log.info("Upload complete: s3://%s/%s", OUTPUT_BUCKET, s3_key)


def archive_original_xml_to_s3(local_xml_path: str, stem: str, timestamp: str):
    if not XML_ARCHIVE_BUCKET:
        log.warning("XML_ARCHIVE_BUCKET not set - skipping source XML archive.")
        return None

    archive_key = "%s/%s_%s.xml" % (
        XML_ARCHIVE_PREFIX.rstrip("/"),
        stem,
        timestamp,
    )
    log.info("Archiving source XML -> s3://%s/%s", XML_ARCHIVE_BUCKET, archive_key)
    try:
        s3.upload_file(local_xml_path, XML_ARCHIVE_BUCKET, archive_key)
        return archive_key
    except Exception as exc:
        log.error("Source XML archive upload failed: %s", exc)
        return None


# ─────────────────────────────────────────────────────────────────────────────
# HANDLER  (called by router)
# ─────────────────────────────────────────────────────────────────────────────
def handle_outbound(src_bucket: str, xml_key: str, auditor: AuditLogger) -> dict:
    """
    Full Step 1 pipeline: download XML → parse → upsert → CSV → S3.
    Called from the router in lambda_function.py.
    """
    stem = Path(xml_key).stem

    rules = _get_rules()

    xml_local = TMP / xml_key.replace("/", "_")
    s3.download_file(src_bucket, xml_key, str(xml_local))
    log.info("Downloaded XML -> %s", xml_local)

    timestamp = datetime.datetime.now(_TZ_GMT7).strftime("%Y%m%d_%H%M%S")

    auditor.begin_file(f"s3://{src_bucket}/{xml_key}")

    archive_original_xml_to_s3(str(xml_local), stem, timestamp)

    rows = process_xml_file(str(xml_local), rules)
    xml_local.unlink(missing_ok=True)

    if not rows:
        log.warning("No rows produced for %s — skipping upload", xml_key)
        auditor.end_file("no_rows")
        return {"statusCode": 200, "processed": 0, "rows": 0}

    try:
        _upsert_rows(rows, rules, auditor)
        upsert_status = "ok"
        upsert_failed = 0
    except Exception as exc:
        log.error("Upsert failed: %s", exc)
        upsert_status = "db_error"
        upsert_failed = len(rows)

    filename = "by_asst_%s_%s.csv" % (stem, timestamp)
    s3_key   = "%s/%s" % (OUTPUT_PREFIX.rstrip("/"), filename)
    content  = rows_to_pipe_csv(rows, rules)
    upload_output(content, s3_key)

    auditor.end_file(
        upsert_status,
        rows_produced = len(rows),
        rows_upserted = len(rows) - upsert_failed,
        rows_failed   = upsert_failed,
    )

    result = {
        "statusCode": 200,
        "processed":  1,
        "rows":       len(rows),
        "s3_output":  "s3://%s/%s" % (OUTPUT_BUCKET, s3_key),
    }
    log.info("Outbound run complete: %s", json.dumps(result))
    return result
