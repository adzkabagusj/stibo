"""
ref_grouping_sql.py
===================
Brand Category from SQL Server table DB_TOOLS.MST.REF_GROUPING.

This table is the ground truth for Brand Category; the "2. Source Mapping
related RNA" sheet of the Attributes List v6 workbook is no longer trusted.
There is deliberately no fallback to v6: when the query fails or no row
matches, the caller sends nothing and the gap shows up in the log.

Connection follows map-stibo-outbound-transform-by-dev/src/by_outbound.py
(Secrets Manager credentials + pyodbc / ODBC Driver 18 from the shared layer).

Environment variables:
    DB_SERVER                 SQL Server host
    MAA_STIBO_SECRET_MANAGER  secret holding the SQL login
    REF_GROUPING_DB           database to connect to (default DB_TOOLS)

Used by every module that writes AT_BrandCategory (rolled out brand by brand
since 2026-10-05; Diadora recap was the first, on 2026-09-30).
"""

import json
import logging
import os
import re

import boto3

log = logging.getLogger(__name__)

REF_GROUPING_TABLE = "DB_TOOLS.MST.REF_GROUPING"

# Reason of the last failed query (credentials masked), for the audit log in S3
LAST_ERROR = ""


def _norm_comp_code(value) -> str:
    """'0888', '888' and '888.0' all become '888' (same rule as the RNA loaders)."""
    s = str(value or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    if not s:
        return ""
    return s.lstrip("0") or "0"


def _key(country, comp_code, sbu, brand_code) -> tuple:
    return (
        str(country or "").strip().upper(),
        _norm_comp_code(comp_code),
        str(sbu or "").strip().upper(),
        str(brand_code or "").strip().upper(),
    )


def _get_credentials() -> dict:
    secret_name = os.environ.get("MAA_STIBO_SECRET_MANAGER", "")
    response = boto3.client("secretsmanager").get_secret_value(SecretId=secret_name)
    secret = json.loads(response["SecretString"])

    username = secret.get("username")
    password = secret.get("password")
    if (not username or not password) and secret.get("app.stibo"):
        username = "app.stibo"
        password = secret["app.stibo"]
    if not username or not password:
        raise ValueError(
            f"Secret '{secret_name}' must contain 'username'/'password' "
            f"or an 'app.stibo' credential entry."
        )
    return {"username": username, "password": password}


def _connect():
    import pyodbc  # from the shared Lambda layer; imported here so local runs without it still load

    server = os.environ.get("DB_SERVER")
    if not server:
        raise ValueError("DB_SERVER missing from environment variables")
    creds = _get_credentials()
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={server},1433;"
        f"DATABASE={os.environ.get('REF_GROUPING_DB', 'DB_TOOLS')};"
        f"UID={creds['username']};"
        f"PWD={creds['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
        "Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str, timeout=10)


def load_brand_categories(brand_code: str) -> dict | None:
    """
    Return {(country, comp_code, sbu, brand_code): brand_category} for one brand,
    read once per run.  None means the query itself failed (already logged).
    """
    global LAST_ERROR
    LAST_ERROR = ""
    brand_code = str(brand_code or "").strip().upper()
    try:
        conn = _connect()
        try:
            cursor = conn.cursor()
            cursor.execute(
                f"SELECT COUNTRY, COMPCODE, SBU, BRANDCODE, BRANDCATEGORY "
                f"FROM {REF_GROUPING_TABLE} WHERE UPPER(BRANDCODE) = ?",
                brand_code,
            )
            rows = cursor.fetchall()
        finally:
            conn.close()
    except Exception as exc:
        err = re.sub(r"(PWD=|UID=)[^;]+", r"\1***", f"{type(exc).__name__}: {exc}")
        LAST_ERROR = err
        log.error("[REF_GROUPING] Query failed for brand %s: %s", brand_code, err)
        return None

    lookup = {}
    for country, comp_code, sbu, bcode, category in rows:
        category = str(category or "").strip()
        if category:
            lookup[_key(country, comp_code, sbu, bcode)] = category
    log.info("[REF_GROUPING] %d Brand Category row(s) loaded for brand %s", len(lookup), brand_code)
    return lookup


def get_brand_category(lookup: dict | None, country, comp_code, sbu, brand_code) -> str:
    """Exact match on country + company code + SBU + brand code; '' when not found."""
    if not lookup:
        return ""
    return lookup.get(_key(country, comp_code, sbu, brand_code), "")


# Country codes used in file names → COUNTRY values in REF_GROUPING
COUNTRY_NAMES = {
    "ID": "INDONESIA", "MY": "MALAYSIA", "PH": "PHILIPPINES", "SG": "SINGAPORE",
    "TH": "THAILAND", "VN": "VIETNAM", "KH": "CAMBODIA",
}


def country_name(country) -> str:
    """'ID' → 'INDONESIA'; a full country name is returned upper-cased as is."""
    value = str(country or "").strip().upper()
    return COUNTRY_NAMES.get(value, value)


def load_for_run(brand_code, auditor=None) -> dict | None:
    """load_brand_categories() + the 'ref_grouping_sql' entry in the S3 audit log."""
    lookup = load_brand_categories(brand_code)
    if auditor is not None and hasattr(auditor, "record_loader"):
        auditor.record_loader(
            "ref_grouping_sql",
            "error" if lookup is None else "ok",
            rows=len(lookup or {}),
            error=LAST_ERROR,
        )
    return lookup


def brand_category_for_key(lookup: dict | None, brand_code, country, comp_code, sbu,
                           warned: set | None = None) -> str:
    """
    Exact-match Brand Category for one key ('' when not found or when the
    query failed).  A miss is logged once per key when the caller passes a
    `warned` set (one set per run), otherwise on every call.
    """
    if lookup is None:
        return ""
    category = get_brand_category(lookup, country_name(country), comp_code, sbu, brand_code)
    if not category:
        key = _key(country_name(country), comp_code, sbu, brand_code)
        if warned is None or key not in warned:
            if warned is not None:
                warned.add(key)
            log.warning("[REF_GROUPING] %s/%s/%s/%s has no row in REF_GROUPING — "
                        "AT_BrandCategory not sent",
                        country_name(country), comp_code, sbu, str(brand_code or "").upper())
    return category


def brand_category_for_file(brand_code, country, comp_code, sbu, auditor=None) -> str:
    """
    Brand Category for one input file whose rows all share the same country /
    company code / SBU / brand (taken from the file name).  Reads SQL once,
    records the loader in the audit log and logs a warning when nothing
    matches.  Returns '' when the query failed or no row matches exactly.
    """
    lookup = load_for_run(brand_code, auditor)
    return brand_category_for_key(lookup, brand_code, country, comp_code, sbu)
