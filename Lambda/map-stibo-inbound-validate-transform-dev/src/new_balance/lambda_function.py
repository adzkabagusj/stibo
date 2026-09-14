"""
lambda_function.py — Unified Lambda Handler (New Balance)
==========================================================
Function name : map-stibo-inbound-validate-transform-nb-dev

Handles ALL New Balance file types from a single S3 folder:
    raw/metadata/new-balance/

File-type → ETL dispatch:
    ┌─────────────────────────────────────────────────────────────────────────┐
    │ Filename contains                        → ETL module                   │
    ├─────────────────────────────────────────────────────────────────────────┤
    │ "Line List (Apparel & Accessories)"      → new_balance.main             │
    │ "Line List (Footwear)"                   → new_balance.main_footwear    │
    │ "Line List (Footwear GTM)" / "...GTM..."  → new_balance.inline_gtm_...   │
    │ "Line List"  (no Apparel/Footwear qual.) → new_balance.linelist_main    │
    │ "Ecommerce File"                         → new_balance.ecommerce_main   │
    │ "SMS Election" (Footwear)                → new_balance.sample_footwear  │
    │ "Apparel Sample Breakout"                → new_balance.sample_apparel   │
    │ "Accessories Sample Breakout"            → new_balance.sample_acc       │
    └─────────────────────────────────────────────────────────────────────────┘

    Detection order matters:
        1. linelist_footwear_gtm (GTM line sheet — most specific)
        2. linelist_footwear  (footwear price list)
        2. linelist_apparel   (explicit apparel/accessories check)
        3. ecommerce          (ecommerce / dtc / map_s keywords)
        4. linelist_licensed  (plain "line list" with no apparel/footwear qual.)

Linelist sub-type folders under input/:
    /tmp/stibo_workdir/input/linelist_apparel/    ← Apparel & Accessories
    /tmp/stibo_workdir/input/linelist_footwear/   ← Footwear (price list)
    /tmp/stibo_workdir/input/linelist_footwear_gtm/ ← Footwear GTM line sheet
    /tmp/stibo_workdir/input/linelist_licensed/   ← Licensed plain line list
    /tmp/stibo_workdir/input/ecommerce/           ← Ecommerce export
    /tmp/stibo_workdir/input/mdd/
    /tmp/stibo_workdir/input/attributes/

Global metadata files are fetched from raw/metadata/ root and bucket root.

────────────────────────────────────────────────────────────────────
ENVIRONMENT VARIABLES
────────────────────────────────────────────────────────────────────
    RAW_BUCKET        Source bucket  e.g. map-stibo-inbound-raw-dev
    PROCESSED_BUCKET  Destination bucket  e.g. map-stibo-inbound-processed-dev
    COMP_CODE         Optional fallback company code  (default: "0000")
"""

import json
import logging
import os
import re
import shutil
import sys
import types
import openpyxl
import pyxlsb
from pathlib import Path
from urllib.parse import unquote_plus

import boto3
import pandas as pd

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
# Set LAMBDA_TMP_DIR BEFORE importing ETL modules so BASE_DIR resolves correctly
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

import new_balance.inline_main_accessories           as etl_apparel        # NB inline apparel & accessories
import new_balance.inline_main_footwear  as etl_footwear        # NB inline footwear
import new_balance.inline_gtm_footwear_main as etl_footwear_gtm # NB inline footwear GTM line sheet
import new_balance.inline_appacc_preline_main as etl_appacc_preline  # NB inline App/Acc Preline line list
import new_balance.licensed_linelist_main  as etl_licensed_list   # NB licensed line list
import new_balance.licensed_ecommerce_main as etl_ecommerce       # NB ecommerce export
import new_balance.inline_ecommerce_main as etl_inline_ecommerce  # NB inline ecommerce export
import new_balance.inline_ean_source as etl_ean_source   #inline EAN SOURCE
import new_balance.licensed_ordersheet_main as etl_licensed_ordersheet  # NB licensed order sheet
import new_balance.inline_sample_footwear_main as etl_sample_footwear   # NB sample footwear (SMS Election)
import new_balance.inline_sample_apparel_main as etl_sample_apparel     # NB sample apparel breakout
import new_balance.inline_sample_accessories_main as etl_sample_accessories  # NB sample accessories breakout

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_nb_unified")

# ─────────────────────────────────────────────────────────────────────────────
# AWS client
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# Environment variables
# ─────────────────────────────────────────────────────────────────────────────
RAW_BUCKET        = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET  = os.environ["PROCESSED_BUCKET"]
DEFAULT_COMP_CODE = os.environ.get("COMP_CODE", "0000")

# ─────────────────────────────────────────────────────────────────────────────
# Brand name → brand code
# ─────────────────────────────────────────────────────────────────────────────
BRAND_NAME_TO_CODE: dict[str, str] = {
    "NEW BALANCE": "NEW",
    "ADIDAS":      "ADI",
    "NIKE":        "NIK",
    "SMIGGLE":     "SMI",
    "ALDO":        "ALD",
    "CROCS":       "CRO",
    "LOTTO":       "LOT",
    "BIRKENSTOCK": "BIR",
}

# ─────────────────────────────────────────────────────────────────────────────
# Category keyword → SBU code  (used by header-row fallback parser)
# ─────────────────────────────────────────────────────────────────────────────
CATEGORY_TO_SBU: dict[str, str] = {
    "footwear":    "FW",
    "apparel":     "AP",
    "accessory":   "AC",
    "accessories": "AC",
    "equipment":   "EQ",
}

# ─────────────────────────────────────────────────────────────────────────────
# File type detection
#
# Detection ORDER matters — more specific patterns must come first:
#
#   1. linelist_footwear  — "Line List (Footwear)" / "footwear" / "shoes"
#   2. linelist_apparel   — "Line List (Apparel & Accessories)" / "apparel" etc.
#   3. ecommerce          — ecomm / dtc / map_s / item tab keywords
#   4. linelist_licensed  — plain "Line List" with no apparel/footwear qualifier
#                           (catches "0888-SP-NEW BALANCE-Line List-Multi-SS26-1")
#   5. mdd / attributes  — shared support files
#
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # ── 0z. NB Sample templates (most specific — must precede the generic
    #     footwear/apparel/accessories keywords so breakout files are not
    #     swallowed by the plain line-list types) ─────────────────────────
    "sample_footwear": [
        "sms election",
        "footwear initial sms",
    ],
    "sample_apparel": [
        "apparel sample breakout",
        "inline apparel sample",
    ],
    "sample_accessories": [
        "accessories sample breakout",
        "licensed accessories sample",
    ],
    # ── 0 EAN ───────────────────────
    "ean_source": [
    "ean source", "ean_source",
    "upc list", "upc_list",
    "master upc", "master_upc",
    "barcode",
    ],
    # ── 1. Footwear GTM line sheet (more specific than plain footwear) ────
    "linelist_footwear_gtm": [
        "line list (footwear gtm)",
        "linelist (footwear gtm)",
        "line sheet (footwear gtm)",
        "footwear gtm",
        "gtm footwear",
    ],
    # ── 2. Footwear price list ────────────────────────────────────────────
    "linelist_footwear": [
        "line list (footwear)",
        "linelist (footwear)",
        "footwear",
        "shoes",
    ],
    # ── 1b. App/Acc Preline line list (more specific than plain apparel) ───
    "linelist_appacc_preline": [
        # Production names use "Line List App Acc - Updated ...", the frontend
        # label says "Line List (App Acc Preline - Updated)". Cover both.
        "line list app acc",
        "linelist app acc",
        "line list (app acc",
        "linelist (app acc",
        "app acc preline",
        "appacc preline",
        "preline linelist",
        "preline line list",
        "preline",
    ],
    # ── 2. Apparel & Accessories ──────────────────────────────────────────
    "linelist_apparel": [
        "line list (apparel",        # covers "Line List (Apparel & Accessories)"
        "linelist (apparel",
        "line list (accessories",
        "apparel",
        "accessories",
    ],
    # ── 3. Ecommerce ──────────────────────────────────────────────────────
    "ecommerce_licensed": [
        "ecommerce file licensed",       # 0888-...-Ecommerce File Licensed-...
        "ecommerce file-licensed",
        "ecommerce file - licensed",
        "ecommerce-licensed",
        "ecomm-licensed",
        "eommerce file licensed",        # typo variant (Eommerce)
        "eommerce file-licensed",
    ],
    "ecommerce_inline": [
        "ecommerce file inline",
        "ecommerce file-inline",
        "ecommerce file - inline",
        "ecommerce-inline",
        "ecomm-inline",
        "eommerce file inline",
        "eommerce file-inline",
        "inline",                   # ← ADD THIS — catches "...Inline..." anywhere in name
    ],
    "ecommerce": [
        "ecommerce file",
        "eommerce file",                 # typo variant (Eommerce)
        "ecomm", "ecommerce", "e-comm", "e_comm",
        "map_s", "map-s",
        "item tab", "item_tab",
        "dtc", "dtcecomm",
    ],
     
        # ── 4a. Licensed Order Sheet ──────────────────────────────────────────
    "ordersheet_licensed": [
        "order sheet licensed",
        "ordersheet licensed",
        "order sheet-licensed",
        "order-sheet-licensed",
        "order sheet (licensed)",
        "order sheet maa",          # ← ADD THIS
        "order sheet",              # ← ADD THIS (broad fallback)
    ],
    # ── 4. Licensed plain line list (no apparel/footwear qualifier) ───────
    "linelist_licensed": [
        "line list",
        "linelist",
        "price list", "pricelist", "price_list",
        "inline", "licensed", "apac",
    ],
   
    # ── 5. Shared / global support files ─────────────────────────────────
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": ["attributes", "attributes_list", "attribute list"],
}

# ─────────────────────────────────────────────────────────────────────────────
# ETL dispatcher  —  file_type key → ETL module
# ─────────────────────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "linelist_apparel":   etl_apparel,
    "linelist_footwear":  etl_footwear,
    "linelist_footwear_gtm": etl_footwear_gtm,
    "linelist_appacc_preline": etl_appacc_preline,
    "linelist_licensed":  etl_licensed_list,
    "ecommerce_licensed": etl_ecommerce,
    "ecommerce_inline":   etl_inline_ecommerce,
    "ean_source": etl_ean_source,
    "ordersheet_licensed":  etl_licensed_ordersheet,
    "sample_footwear":      etl_sample_footwear,
    "sample_apparel":       etl_sample_apparel,
    "sample_accessories":   etl_sample_accessories,
}

# SBU per ETL type — keeps output filenames unique
ETL_TYPE_SBU: dict[str, str] = {
    "linelist_apparel":   "AP",
    "linelist_footwear":  "FW",
    "linelist_footwear_gtm": "FW",
    "linelist_appacc_preline": "AP",
    "linelist_licensed":  "SP",
    "ecommerce_licensed": "SP",
    "ecommerce_inline":   "SP",
    "ean_source": "FW",
    "ordersheet_licensed":  "SP",
    "sample_footwear":     "FW",
    "sample_apparel":      "AP",
    "sample_accessories":  "AC",
}

# ─────────────────────────────────────────────────────────────────────────────
# Mandatory file rules
#
#   Non-linelist types always required: mdd, attributes
#   Plus at least ONE linelist/ecommerce type.
# ─────────────────────────────────────────────────────────────────────────────
REQUIRED_NON_LINELIST_TYPES: set[str] = {"mdd", "attributes"}

# Required types per triggered file type (mirrors NB Licensed logic)
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist_apparel":   {"linelist_apparel",   "mdd", "attributes"},
    "linelist_footwear":  {"linelist_footwear",  "mdd", "attributes"},
    "linelist_footwear_gtm": {"linelist_footwear_gtm", "mdd", "attributes"},
    "linelist_appacc_preline": {"linelist_appacc_preline", "mdd", "attributes"},
    "linelist_licensed":  {"linelist_licensed",  "mdd", "attributes"},
    "ecommerce_licensed": {"ecommerce_licensed", "mdd", "attributes"},
    "ecommerce_inline":   {"ecommerce_inline",   "mdd", "attributes"},
    "ean_source": {"ean_source", "mdd"},
    "ordersheet_licensed":  {"ordersheet_licensed", "mdd", "attributes"},
    "sample_footwear":     {"sample_footwear",     "mdd", "attributes"},
    "sample_apparel":      {"sample_apparel",      "mdd", "attributes"},
    "sample_accessories":  {"sample_accessories",  "mdd", "attributes"},
}

# Global types fetched from raw/metadata/ root (not brand subfolder)
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}


# =============================================================================
# HELPERS  (unchanged from originals — merged here)
# =============================================================================

def _clean_cell(v):
    """Strip illegal XML/Excel characters from string cell values."""
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    """Convert .xlsb → .xlsx streaming row-by-row (from NB Licensed handler)."""
    xlsx_path      = xlsb_path.with_suffix(".xlsx")
    wb_out         = openpyxl.Workbook(write_only=True)
    sheets_written = 0

    with pyxlsb.open_workbook(str(xlsb_path)) as wb:
        for sheet_name in wb.sheets:
            safe_name = "".join(
                c for c in sheet_name
                if 0x20 <= ord(c) <= 0xFFFD and c not in r"\/?*[]:"
            ).strip()[:31] or f"Sheet{sheets_written + 1}"

            try:
                ws_out = wb_out.create_sheet(title=safe_name)
            except Exception as e:
                log.warning("  Skipping sheet '%s' (bad name): %s", sheet_name, e)
                continue

            try:
                with wb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        try:
                            ws_out.append([
                                _clean_cell(item.v) if item is not None else None
                                for item in row
                            ])
                        except Exception as e:
                            log.warning("  Skipping bad row in '%s': %s", safe_name, e)
                sheets_written += 1
            except Exception as e:
                log.warning("  Error reading sheet '%s': %s", sheet_name, e)

    if sheets_written == 0:
        raise ValueError(f"No sheets converted from {xlsb_path.name}")

    wb_out.save(str(xlsx_path))
    log.info("  Converted %s → %s (%d sheets)", xlsb_path.name, xlsx_path.name, sheets_written)
    xlsb_path.unlink()
    return xlsx_path


def _convert_csv_to_xlsx(csv_path: Path) -> Path:
    """Convert a .csv to .xlsx so openpyxl can read it."""
    xlsx_path = csv_path.with_suffix(".xlsx")
    try:
        df = pd.read_csv(str(csv_path), dtype=str, keep_default_na=False)
    except Exception as e:
        raise ValueError(f"Cannot read CSV {csv_path.name}: {e}")

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append([_clean_cell(v) for v in row.tolist()])

    wb.save(str(xlsx_path))
    log.info("  Converted %s → %s", csv_path.name, xlsx_path.name)
    csv_path.unlink()
    return xlsx_path


def _detect_file_type(filename: str) -> str | None:
    name_lower = filename.lower()
    name_norm  = re.sub(r'\s*-\s*', ' ', name_lower)
    name_norm  = re.sub(r'\s+', ' ', name_norm).strip()
    name_norm  = name_norm.replace("eommerce", "ecommerce")

    # ✅ ADD THIS BLOCK — smart combo check before keyword loop
    has_ecomm_file = "ecommerce file" in name_norm
    has_licensed   = "licensed" in name_norm
    has_inline     = "inline" in name_norm
    has_ean_source = "ean source" in name_norm

    # GTM footwear line sheet — the token "GTM" can sit anywhere in the name
    # ("Copy of S227 GTM 2 APAC Footwear Line List July 2026"), so match on the
    # combination rather than on "gtm" alone, which would also swallow an
    # apparel line list that happens to mention GTM.
    has_gtm     = re.search(r'\bgtm\b', name_norm) is not None
    has_apparel = any(k in name_norm for k in ("apparel", "accessor"))
    has_fw_list = any(
        k in name_norm for k in ("footwear", "shoes", "line list", "linelist", "line sheet")
    )
    if has_gtm and has_fw_list and not has_apparel:
        return "linelist_footwear_gtm"

    # NB Sample templates — three distinct files:
    #   "S1 27 Footwear Initial SMS Election_…"   (footwear, no 'sample' word!)
    #   "S1'27 Global Inline Apparel Sample Breakout file_…"
    #   "S127 Licensed Accessories Sample Breakout file _ …"
    # Checked BEFORE the keyword loop so the generic apparel/accessories/
    # footwear entries cannot swallow the breakout files.
    has_sample_breakout = "sample breakout" in name_norm
    has_sms_election    = "sms election" in name_norm
    if has_sms_election and has_fw_list and not has_apparel:
        return "sample_footwear"
    if has_sample_breakout and "apparel" in name_norm:
        return "sample_apparel"
    if has_sample_breakout and "accessor" in name_norm:
        return "sample_accessories"

    # App/Acc line list — "App Acc" or "Preline" anywhere alongside a line-list
    # token. Guarded so it cannot swallow the plain
    # "Line List (Apparel & Accessories)" file, whose text has no "app acc".
    has_app_acc = re.search(r'\bapp\s*acc\b', name_norm) is not None
    has_preline = "preline" in name_norm
    has_list    = any(
        k in name_norm for k in ("line list", "linelist", "line sheet")
    )
    if (has_app_acc or has_preline) and has_list and not has_gtm:
        return "linelist_appacc_preline"

    if has_ean_source:
        return "ean_source"
    if has_ecomm_file and has_licensed and not has_inline:
        return "ecommerce_licensed"
    if has_ecomm_file and has_inline:
        return "ecommerce_inline"

    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        for kw in keywords:
            kw_norm = re.sub(r'\s*-\s*', ' ', kw.lower())
            kw_norm = re.sub(r'\s+', ' ', kw_norm).strip()
            kw_norm = kw_norm.replace("eommerce", "ecommerce")
            if kw_norm in name_norm:
                # ── NEW GUARD ──────────────────────────────────────────
                # Prevent generic "ecommerce" from matching when a more
                # specific sub-type keyword (inline / licensed) is present
                if ftype == "ecommerce" and (
                    "inline"   in name_norm or
                    "licensed" in name_norm
                ):
                    continue
                # The bare "inline" keyword under ecommerce_inline matches ANY
                # NB filename containing "Inline" — which is most line lists.
                # Only let it win when the name really is an ecommerce export.
                if (
                    ftype == "ecommerce_inline"
                    and kw_norm == "inline"
                    and not any(e in name_norm for e in ("ecomm", "ecommerce"))
                ):
                    continue
                # ───────────────────────────────────────────────────────
                return ftype
    return None


def _detect_triggered_file_type(key: str) -> str | None:
    """
    Derive the file type from the S3 key that triggered this invocation.
    Returns None for global fan-out placeholder keys (starting with '.__').
    """
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)


def _parse_event(event: dict) -> tuple[str, str]:
    """Extract bucket and key from an EventBridge S3 Object Created event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key: str) -> str | None:
    """
    Derive the principal (brand folder) from the S3 key.
    Key format: raw/metadata/{principal}/{filename}
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not match raw/metadata/{principal}/{file}: %s", key)
    return None


# These NB trigger types use the new Brand Mapping file instead of the v6 Attributes List.
_NB_NEW_MAPPING_TRIGGER_TYPES: set[str] = {
    "linelist_footwear_gtm", "linelist_appacc_preline",
    "sample_footwear", "sample_apparel", "sample_accessories",
}


def _list_principal_files(bucket: str, principal: str, triggered_ftype: str | None = None) -> dict[str, dict]:
    """
    Scan S3 for files belonging to this principal.
    - Brand-specific files: raw/metadata/{principal}/
    - Global shared files (mdd): raw/metadata/  (root level only)
    Brand-specific version always wins over global.
    """
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    # ── Brand-specific folder ─────────────────────────────────────────────
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue

            ftype = _detect_file_type(filename)
            if ftype is None:
                log.warning("  Unrecognised file (skipping): %s", filename)
                continue

            last_modified = obj["LastModified"]
            if ftype not in found or last_modified > found[ftype]["last_modified"]:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified,
                }
                log.info("  Classified  %-22s ← %s  [brand-specific]", ftype, filename)

    # ── Global folder (mdd only) ────────────────────────────────
    # GTM and App/Acc Preline modules use the new Brand Mapping file;
    # all other NB modules keep using the v6 Attributes List.
    _attr_principal = (
        None
        if triggered_ftype in _NB_NEW_MAPPING_TRIGGER_TYPES
        else "new_balance"
    )
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        principal=_attr_principal,
        log=log,
    )

    for page in paginator.paginate(
        Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"
    ):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue

            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES:
                continue

            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified,
                }
                log.info("  Classified  %-22s ← %s  [global]", ftype, filename)
            else:
                log.info(
                    "  Skipping global %s — brand-specific already found: %s",
                    ftype, found[ftype]["filename"],
                )

    return found


def _check_mandatory_files(found: dict[str, dict], triggered_ftype: str | None) -> list[str]:
    """
    Returns list of missing mandatory types.

    If triggered_ftype is known, use its specific required set.
    Otherwise fall back to: mdd + attributes + at least one linelist/ecommerce.
    """
    if triggered_ftype and triggered_ftype in REQUIRED_TYPES_BY_TRIGGER:
        required = REQUIRED_TYPES_BY_TRIGGER[triggered_ftype]
        return sorted(required - set(found.keys()))

    # Generic fallback
    missing = list(REQUIRED_NON_LINELIST_TYPES - set(found.keys()))
    has_any_brand_file = any(
        k in found for k in ("sample_apparel", "sample_accessories", "sample_footwear",
                            "linelist_apparel", "linelist_footwear",
                            "linelist_footwear_gtm", "linelist_appacc_preline",
                            "linelist_licensed", "ecommerce_licensed",
                            "ecommerce_inline", "linelist","ean_source","ordersheet_licensed")  # ← replace "ecommerce"
    )
    if not has_any_brand_file:
        missing.append("linelist or ecommerce")
    return sorted(missing)


def _prepare_tmp_dirs() -> dict[str, Path]:
    """
    Recreate /tmp/stibo_workdir/input/{type}/ from scratch each invocation.
    """
    base = Path(TMP_WORKDIR)

    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs: dict[str, Path] = {
        "linelist_apparel":   base / "input" / "linelist_apparel",
        "linelist_footwear":  base / "input" / "linelist_footwear",
        "linelist_footwear_gtm": base / "input" / "linelist_footwear_gtm",
        "linelist_appacc_preline": base / "input" / "linelist_appacc_preline",
        "linelist_licensed":  base / "input" / "linelist_licensed",
        "ecommerce_licensed": base / "input" / "ecommerce_licensed",
        "ecommerce_inline":   base / "input" / "ecommerce_inline",
        "ecommerce":          base / "input" / "ecommerce",
        "linelist":           base / "input" / "linelist",
        "mdd":                base / "input" / "mdd",
        "attributes":         base / "input" / "attributes",
        "ean_source": base / "input" / "ean_source",
        "ordersheet_licensed":  base / "input" / "ordersheet_licensed",
        "sample_footwear":      base / "input" / "sample_footwear",
        "sample_apparel":       base / "input" / "sample_apparel",
        "sample_accessories":   base / "input" / "sample_accessories",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    return dirs


def _download_files(
    bucket: str,
    found: dict[str, dict],
    dirs: dict[str, Path],
    auditor,
) -> None:
    """
    Download each identified S3 file to its corresponding local sub-folder.
    """
    for ftype, info in found.items():
        if ftype not in dirs:
            log.warning("  No local dir for type '%s' — skipping %s", ftype, info["filename"])
            continue

        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-22s ← s3://%s/%s", ftype, bucket, info["key"])
        try:
            s3.download_file(bucket, info["key"], str(local_path))
            log.info("  Saved → %s", local_path)
            if auditor:
                auditor.record_download(ftype, info["filename"], status="ok")
        except Exception as exc:
            log.error("  Download FAILED for %s: %s", info["key"], exc)
            if auditor:
                auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
            raise


def _upload_xml_outputs(principal: str) -> list[str]:
    """
    Upload all .xml files produced under output/xml/ to the processed bucket.
    Returns list of uploaded S3 keys.
    """
    xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
    uploaded = []

    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading → s3://%s/%s", PROCESSED_BUCKET, s3_key)
        try:
            s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
            uploaded.append(s3_key)
            log.info("  Uploaded: %s", xml_file.name)
        except Exception as exc:
            log.error("  Upload FAILED for %s: %s", xml_file.name, exc)

    return uploaded


def _parse_season_from_header(header: str) -> str | None:
    """
    Parse season code from a workbook header string or filename stem.
    Unchanged from original new_balance/lambda_function.py.
    """
    m = re.search(r'\b(FW|SS)(\d{2})\b', header, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}{m.group(2)}"

    m = re.search(r'\b(FW|SS)(20\d{2})\b', header, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}{m.group(2)[-2:]}"

    m = re.search(r'\b([SF])2\s+(\d{4})\b', header, re.IGNORECASE)
    if m:
        prefix = "SS" if m.group(1).upper() == "S" else "FW"
        return f"{prefix}{m.group(2)[-2:]}"

    m = re.search(r'\bFall\s*(\d{2,4})\b', header, re.IGNORECASE)
    if m:
        return f"FW{m.group(1)[-2:]}"

    m = re.search(r'\bSpring\s*(\d{2,4})\b', header, re.IGNORECASE)
    if m:
        return f"SS{m.group(1)[-2:]}"

    m = re.search(r'\bS2\s*(\d{2})\b', header, re.IGNORECASE)
    if m:
        return f"SS{m.group(1)}"

    m = re.search(r'\bF2\s*(\d{2})\b', header, re.IGNORECASE)
    if m:
        return f"FW{m.group(1)}"

    # GTM line-sheet title style: "S2'27 APAC Footwear GTM 2 Line Sheet"
    m = re.search(r"\b([SF])2['\u2019](\d{2})\b", header, re.IGNORECASE)
    if m:
        prefix = "SS" if m.group(1).upper() == "S" else "FW"
        return f"{prefix}{m.group(2)}"

    # NB sample-breakout title style: "S1'27", "S1 27", "S127"
    # (S + drop number + 2-digit year; S→SS, F→FW — same convention the
    # S2'27 GTM rule above uses). Covers the workbook title row AND the
    # filename stem, which is what the SMS Election file falls back to
    # (its first sheet has no usable title text).
    m = re.search(r"\b([SF])\d['\u2019\s]?(\d{2})\b", header, re.IGNORECASE)
    if m:
        prefix = "SS" if m.group(1).upper() == "S" else "FW"
        return f"{prefix}{m.group(2)}"

    # ── ADD: Carry Over / Collection season codes ─────────────────
    m = re.search(r'\b(CO|HO|AL|SM|SP|SU|FA|FL|WN)(20\d{2}|\d{2})\b', header, re.IGNORECASE)
    if m:
        prefix = m.group(1).upper()
        year   = m.group(2)
        return f"{prefix}{year[-2:]}"

    return None


def _parse_sbu_from_header(header: str) -> str:
    header_lower = header.lower()
    for keyword, sbu in CATEGORY_TO_SBU.items():
        if keyword in header_lower:
            return sbu
    return "GN"


def _read_workbook_header(xlsx_path: Path) -> str | None:
    """
    Scan first 5 rows, all columns, for a non-empty string cell that looks
    like a title. Unchanged from original new_balance/lambda_function.py.
    """
    try:
        wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(max_row=5, values_only=True):
            for val in row:
                if val and isinstance(val, str) and len(val.strip()) > 5:
                    wb.close()
                    return val.strip()
        wb.close()
        return None
    except Exception as exc:
        log.warning("Could not read workbook header from %s: %s", xlsx_path.name, exc)
        return None


def _parse_metadata_from_linelist_filename(
    linelist_filename: str,
    linelist_ftype: str = "linelist_apparel",
) -> dict:
    """
    Parse pipeline config from the linelist filename.

    Expected format:
        {comp_code}-{sbu}-{brand}-{file_type}-{multi_mono}-{season}-{seq}.xlsx
    e.g.:
        0888-SP-NEW BALANCE-Line List (Apparel & Accessories)-Multi-SS26-1.xlsx
        0888-SP-NEW BALANCE-Line List (Footwear)-Multi-SS26-1.xlsx
        0888-SP-NEW BALANCE-Line List-Multi-SS26-1.xlsx
        0888-SP-NEW BALANCE-Ecommerce File-Multi-SS26-1.xlsx

    Also handles NB Licensed fallback filenames (e.g. "Fall26_Inline_...xlsx")
    and NB Ecommerce filenames (e.g. "MAP_S226_APP.xlsx") using the same
    regex-based season fallback from the original NB Licensed handler.

    Unchanged logic from both originals — unified here.
    """
    stem = linelist_filename
    for ext in (".xlsx", ".xlsm", ".csv", ".xlsb"):
        if stem.lower().endswith(ext):
            stem = stem[: -len(ext)]
            break

    # ── Strategy 1: standard regex pattern ───────────────────────────────
    m = re.match(
        r'^(?P<comp>[^-]+)'
        r'-(?P<sbu>[^-]+)'
        r'-(?P<brand>.+?)'
        r'-(?P<file_type>.+?)'
        r'-(?P<multi_mono>[^-]+)'
        # Prefixes cover the MDD "Season LOV" codes (SP SM FL WN CO SS FW AL)
        # plus legacy aliases already in use (AW HO SU FA).
        r'-(?P<season>(SS|FW|FL|AW|HO|AL|CO|WN|SM|SP|SU|FA)\d{2,4})'  # e.g. SP2028, FW26, FL9080
        r'(?:-(?P<country>[A-Z]{2,3}))?'     # optional country token e.g. PH, MY, ID
        r'-(?P<seq>\d+)$',
        stem,
        re.IGNORECASE,
    )

    if m:
        comp_code  = m.group("comp").strip()
        sbu        = m.group("sbu").strip()
        brand      = m.group("brand").strip()
        file_type  = m.group("file_type").strip()
        season     = m.group("season").upper()
        seq_str    = m.group("seq").strip()
        seq        = int(seq_str) if seq_str.isdigit() else 1

       

        brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())
        log.info(
            "[Strategy 1] comp_code=%s  sbu=%s  brand=%s  "
            "brand_code=%s  season=%s  seq=%s  file_type=%s",
            comp_code, sbu, brand, brand_code, season, seq, file_type,
        )
        country_code = (m.group("country") or "").upper()

        return {
            "comp_code":    comp_code,
            "sbu":          sbu,
            "brand":        brand,
            "brand_code":   brand_code,
            "season":       season,
            "seq":          seq,
            "file_type":    file_type,
            "multi_mono":   m.group("multi_mono").strip(),
            "country_code": country_code,
        }

    log.warning(
        "[Strategy 1] Filename '%s' does not match standard pattern — "
        "falling back.", linelist_filename,
    )

    # ── Strategy 2: read workbook header cell ─────────────────────────────
    xlsx_path = Path(TMP_WORKDIR) / "input" / linelist_ftype / linelist_filename
    if not xlsx_path.exists():
        xlsx_path = xlsx_path.with_suffix(".xlsx")

    header = _read_workbook_header(xlsx_path) if xlsx_path.exists() else None

    if header:
        log.info("[Strategy 2] Workbook header: %s", header)
    else:
        log.warning(
            "[Strategy 2] Could not read header from workbook '%s' — "
            "will try filename-based parsing.", linelist_filename,
        )

    season = _parse_season_from_header(header) if header else None

    # ── Strategy 3: parse season directly from filename stem ─────────────
    if not season:
        log.info("[Strategy 3] Trying season from filename stem: '%s'", stem)
        season = _parse_season_from_header(stem)

    # ── Strategy 4: NB Licensed-style fallback (Fall26, MAP_S226 etc.) ───
    if not season:
        season_match = re.search(
            r"(FW|SS|AW|HO|Fall|Spring|Summer|Winter)[\s_]?(\d{2,4})",
            stem, re.IGNORECASE,
        )
        if season_match:
            prefix_raw = season_match.group(1).upper()
            year_raw   = season_match.group(2)
            prefix_map = {"FALL": "FW", "WINTER": "FW", "SPRING": "SS", "SUMMER": "SS"}
            prefix     = prefix_map.get(prefix_raw, prefix_raw[:2])
            year_short = year_raw[-2:]
            season     = f"{prefix}{year_short}"
            log.info("[Strategy 4] Parsed season from fallback pattern: %s", season)

    # ── Strategy 5: NB Ecommerce filename like MAP_S226_APP ──────────────
    if not season:
        nb_ecomm_match = re.search(r"MAP[_\-]S?(\d{2,4})[_\-]", stem, re.IGNORECASE)
        if nb_ecomm_match:
            code_str = nb_ecomm_match.group(1)
            if len(code_str) == 3:
                sea_num  = code_str[0]
                year_2d  = code_str[1:]
                prefix   = "FW" if sea_num == "2" else "SS"
                season   = f"{prefix}{year_2d}"
                log.info("[Strategy 5] Parsed season from NB ecomm filename: %s", season)

    if not season:
        log.error(
            "[Strategy 2-5] Cannot extract season from header '%s' or filename '%s'",
            header, stem,
        )
        return {}

    # ── Extract fields directly from filename parts ───────────────────────
    parts     = stem.split("-")
    comp_code = parts[0].strip() if len(parts) > 0 and parts[0].strip().isdigit() else DEFAULT_COMP_CODE
    seq_str   = parts[-1].strip() if parts else "1"
    seq       = int(seq_str) if seq_str.isdigit() else 1

    # sbu — read directly from filename part[1]
    sbu_from_filename = parts[1].strip() if len(parts) > 1 else ""
    # Only trust slot [1] when the filename actually has the convention shape
    # ({comp}-{sbu}-{brand}-{type}-{multi}-{season}[-{country}]-{seq} = 7+ parts).
    # A non-conventional name can put something else there entirely
    # (e.g. "S227 APP ACC Preline Linelist - SEA" → "SEA"), which must not be
    # mistaken for an SBU. Deliberately not a fixed code list — real SBUs
    # include CH, SB, GO and others that a whitelist would keep rejecting.
    if len(parts) >= 7 and re.fullmatch(r'[A-Za-z0-9]{2,4}', sbu_from_filename):
        sbu = sbu_from_filename.upper()
    elif linelist_ftype == "ean_source":
        sbu = ETL_TYPE_SBU.get("ean_source", "FW")
    else:
        sbu = _parse_sbu_from_header(header) if header else _parse_sbu_from_header(stem)

    # multi_mono — read directly from filename part[-3]
    multi_mono = parts[-3].strip() if len(parts) >= 3 else "Multi"
    if multi_mono.lower() not in ("multi", "mono"):
        # fallback if the slot doesn't look right
        stem_lower = stem.lower()
        multi_mono = "Mono" if "mono" in stem_lower else "Multi"

    # file_type — everything between brand (parts[2]) and multi_mono (parts[-3])
    if linelist_ftype == "ean_source":
        file_type_str = "EAN Source"
    elif len(parts) >= 7:
        file_type_str = "-".join(p.strip() for p in parts[3:-3])
    else:
        file_type_str = "Line List"

    brand      = "New Balance"
    brand_code = BRAND_NAME_TO_CODE["NEW BALANCE"]

    log.info(
        "[Strategy 2-5] comp_code=%s  sbu=%s  brand=%s  "
        "brand_code=%s  season=%s  seq=%s  multi_mono=%s  file_type=%s",
        comp_code, sbu, brand, brand_code, season, seq, multi_mono, file_type_str,
    )
    country_code_fallback = ""
    for p in reversed(parts):
        tok = p.strip()
        if re.match(r'^[A-Z]{2,3}$', tok, re.IGNORECASE) and not tok.isdigit():
            # Exclude known non-country tokens
            if tok.upper() not in ("SP", "AP", "FW", "SS", "GN", "AC", "NEW", "NIK",
                                    "ADI", "SMI", "ALD", "CRO", "LOT", "BIR", "NB",
                                    "MULTI", "MONO"):
                country_code_fallback = tok.upper()
                break

    return {
        "comp_code":   comp_code,
        "sbu":         sbu,
        "brand":       brand,
        "brand_code":  brand_code,
        "season":      season,
        "seq":         seq,
        "file_type":   file_type_str,
        "multi_mono":  multi_mono,
        "country_code": country_code_fallback,
    }
 
# =============================================================================
# LAMBDA HANDLER
# =============================================================================

def lambda_handler(event, context, auditor=None):
    """
    Unified Lambda entry point for all New Balance file types.

    Routing:
        linelist_apparel   → new_balance.main             (inline apparel)
        linelist_footwear  → new_balance.main_footwear    (inline footwear)
        linelist_licensed  → new_balance.linelist_main    (licensed line list)
        ecommerce          → new_balance.ecommerce_main   (ecommerce export)
    """
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "new-balance",
            routing_decision = "new_balance unified lambda_handler invoked directly",
        )

    log.info("Event received: %s", json.dumps(event))

    # ── 1. Parse EventBridge event ────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Guard: only process keys under raw/metadata/ ──────────────────
    if not key.startswith("raw/metadata/"):
        log.warning("Key outside raw/metadata/ — ignoring: %s", key)
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}

    # ── 3. Extract principal ───────────────────────────────────────────────
    principal = _extract_principal(key)
    if not principal:
        log.error("Could not extract principal from key: %s", key)
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    log.info("Principal (brand folder): %s", principal)

    # ── 3b. Detect triggered file type ───────────────────────────────────
    triggered_ftype = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_ftype or "unknown/global-fanout")

    etl_module = ETL_DISPATCHER.get(triggered_ftype)

    # ── 4. Discover + classify all files in S3 ───────────────────────────
    log.info("Listing files under raw/metadata/%s/ ...", principal)
    found = _list_principal_files(bucket, principal, triggered_ftype=triggered_ftype)
    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})

    if auditor:
        auditor.set_files_classified(found)

    # ── 4b. Global fan-out: resolve ETL from whatever brand file is present
    #        Priority: linelist_apparel > linelist_footwear >
    #                  linelist_licensed > ecommerce
    if etl_module is None:
        for candidate in ("sample_apparel", "sample_accessories", "sample_footwear",
                  "linelist_apparel", "linelist_footwear",
                  "linelist_footwear_gtm", "linelist_appacc_preline",
                  "linelist_licensed", "ecommerce_licensed",
                  "ecommerce_inline","ean_source","ordersheet_licensed"):
            if candidate in found:
                etl_module      = ETL_DISPATCHER[candidate]
                triggered_ftype = candidate
                log.info("Global fan-out: resolved ETL module → '%s'", candidate)
                break

    if etl_module is None:
        log.info(
            "No brand file present yet for '%s' during global fan-out — skipping.",
            principal,
        )
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "skipped_no_brand_file",
        }

    # ── 5. Check mandatory files ──────────────────────────────────────────
    missing = _check_mandatory_files(found, triggered_ftype)
    if missing:
        log.info(
            "Mandatory files not yet present for '%s' — waiting for: %s",
            principal, missing,
        )
        if auditor:
            auditor.set_files_missing(missing)
            auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    missing,
            "status":     "waiting_for_mandatory_files",
        }

    log.info(
        "Mandatory files present for '%s' — triggered type: %s",
        principal, triggered_ftype,
    )

    # ── 6. Prepare clean /tmp/ directories ───────────────────────────────
    dirs = _prepare_tmp_dirs()

    # ── 7. Download files ──────────────────────────────────────────────────
    log.info("Downloading %d files (%s) ...", len(found), sorted(found.keys()))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        log.error("Download failed: %s", exc)
        if auditor:
            auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": str(exc)}

    # ── 7b. Convert .xlsb / .csv → .xlsx ──────────────────────────────────
    log.info("Converting .xlsb / .csv files to .xlsx ...")
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                if auditor:
                    auditor.record_conversion(
                        xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
                    )
            except Exception as e:
                log.error("xlsb conversion failed for %s: %s", xlsb_file.name, e)
                if auditor:
                    auditor.record_conversion(
                        xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx",
                        status="failed", error=str(e),
                    )
                    auditor.set_lambda_status(
                        "error", error=f"xlsb conversion failed: {xlsb_file.name}: {e}"
                    )
                if _direct_invocation:
                    auditor.flush(principal)
                return {
                    "statusCode": 500, "principal": principal,
                    "error": f"xlsb conversion failed for {xlsb_file.name}: {e}",
                }

        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                if auditor:
                    auditor.record_conversion(
                        csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
                    )
            except Exception as e:
                log.error("CSV conversion failed for %s: %s", csv_file.name, e)
                if auditor:
                    auditor.record_conversion(
                        csv_file.name, from_ext=".csv", to_ext=".xlsx",
                        status="failed", error=str(e),
                    )
                    auditor.set_lambda_status(
                        "error", error=f"csv conversion failed: {csv_file.name}: {e}"
                    )
                if _direct_invocation:
                    auditor.flush(principal)
                return {
                    "statusCode": 500, "principal": principal,
                    "error": f"csv conversion failed for {csv_file.name}: {e}",
                }

    # ── 8. Parse metadata from the triggered brand file ───────────────────
    # Pick the best filename to parse metadata from: triggered type first,
    # then fall through in priority order.
    meta_ftype    = triggered_ftype
    meta_filename = found.get(meta_ftype, {}).get("filename")

    if not meta_filename:
        for candidate in ("sample_apparel", "sample_accessories", "sample_footwear",
                  "linelist_apparel", "linelist_footwear",
                  "linelist_footwear_gtm", "linelist_appacc_preline",
                  "linelist_licensed", "ecommerce_licensed",
                  "ecommerce_inline",  "linelist","ean_source","ordersheet_licensed"):
            if candidate in found:
                meta_ftype    = candidate
                meta_filename = found[candidate]["filename"]
                break

    log.info("Parsing metadata from: %s (type: %s)", meta_filename, meta_ftype)
    meta = _parse_metadata_from_linelist_filename(meta_filename, meta_ftype)

    if not meta:
        log.error("Cannot parse metadata from linelist '%s'.", meta_filename)
        if auditor:
            auditor.set_lambda_status(
                "error", error="Unparseable or missing metadata in linelist",
            )
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 400,
            "principal":  principal,
            "error":      "Unparseable or missing metadata in linelist",
        }

    log.info(
        "Metadata: brand=%s  brand_code=%s  comp_code=%s  "
        "sbu=%s  season=%s  seq=%s",
        meta["brand"], meta["brand_code"], meta["comp_code"],
        meta["sbu"], meta["season"], meta["seq"],
    )
    if auditor:
        auditor.set_metadata_parsed(meta)

    # ── 9. Run ETL for the triggered file type only ───────────────────────
    log.info(
        "── Running %s ETL → %s ──",
        triggered_ftype.upper(), etl_module.__name__,
    )

    args = types.SimpleNamespace(
        brand      = meta["brand"].title(),
        brand_code = meta["brand_code"],
        comp_code  = meta["comp_code"],
        sbu        = meta["sbu"],
        season     = meta["season"],
        seq        = meta["seq"],
        file_type  = meta.get("file_type", "Line List"),
        multi_mono = meta.get("multi_mono", "Multi"),
        country_code = meta.get("country_code", "") or principal.upper(),  # ← FIXED
    )

    try:
        etl_module.run(args, auditor=auditor)
        log.info("── %s ETL complete ──", triggered_ftype.upper())
    except Exception as exc:
        log.error("%s ETL failed: %s", triggered_ftype.upper(), exc, exc_info=True)
        if auditor:
            auditor.set_lambda_status("error", error=f"{triggered_ftype} ETL failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": str(exc)}

    # ── 10. Upload generated XML(s) ───────────────────────────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("ETL completed but no XML files were produced under output/xml/")
        if auditor:
            auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "no_xml_generated",
        }

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    if auditor:
        auditor.set_xml_uploads(uploaded)
        auditor.set_lambda_status("ok")
    if _direct_invocation:
        auditor.flush(principal)

    brand_file_info = found.get(triggered_ftype, {})
    return {
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_ftype,
        "brand":               meta["brand"],
        "brand_code":          meta["brand_code"],
        "season":              meta["season"],
        "sbu":                 meta["sbu"],
        "brand_file_used":     brand_file_info.get("filename"),
        "brand_file_s3_key":   brand_file_info.get("key"),
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
    }
