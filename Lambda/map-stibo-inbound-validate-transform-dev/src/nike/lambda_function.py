"""
lambda_function.py — Lambda Handler for Nike
=============================================
Function name : map-stibo-inbound-validate-transform-nike-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/nike/*

Nike-specific flow:
    • ONE input file from the brand  : linelist (SP27 ROOKIE)
    • No TDD, no Backlog
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types: {"linelist", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        linelist          → nike.linelist_main.run()
        orderconfirmation → nike.orderConfirmation_main.run()

S3 layout:
    raw/metadata/nike/{linelist_file}.xlsx         ← Nike uploads here
    raw/metadata/nike/{orderconfirm_file}.xlsx     ← Nike uploads here
    raw/metadata/mdd_file.xlsx                     ← shared global MDD
    raw/metadata/attributes_list.xlsx              ← shared global attr list

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/nike/{filename}.xml
"""

import json
import logging
import os
import shutil
import sys
import types
import pyxlsb
import openpyxl
from pathlib import Path
from urllib.parse import unquote_plus

import boto3
import pandas as pd

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_nike"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import nike.linelist_rookie_main as linelist_etl                # noqa: E402
import nike.sp27_maa_linelist_main as sp27_maa_etl       # noqa: E402
import nike.orderConfirmation_main as order_confirm_etl  # noqa: E402
import nike.nike_360_linelist_main as nike_360_etl       # noqa: E402
import nike.nike_360_ean_main as nike_360_ean_etl        # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_nike")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# ─────────────────────────────────────────────────────────────────────────────
BRAND_NAME_TO_CODE: dict[str, str] = {
    "ADIDAS":      "ADI",
    "NIKE":        "NIKE",
    "NEW BALANCE": "NEW",
    "SMIGGLE":     "SMI",
    "ALDO":        "ALD",
    "CROCS":       "CRO",
    "LOTTO":       "LOT",
    "BIRKENSTOCK": "BIR",
}

# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "linelist": [
        "line list", "linelist", "rookie",
        "catalogue", "catalog", "handover",
        "360", "order form", "orderform",
        # order confirmation variants — routed via LINELIST_VARIANTS
        "order confirm", "orderconfirm", "order_confirm",
        "so confirm", "soconfirm", "order conf",
        "confirmed order", "booking confirm",
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "attributes", "attributes_list", "attribute list",
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template", "mapping files template",
    ],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist": {"linelist", "mdd", "attributes"},
}

REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes"}
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES:   set[str] = {"mdd", "attributes"}

ETL_DISPATCHER: dict[str, object] = {
    "linelist": linelist_etl,
}

# Sub-variants — picked by filename pattern after file-type dispatch.
# First match wins; order matters.
LINELIST_VARIANTS: list[tuple[tuple[str, ...], object]] = [
    (("360", "ean"),           nike_360_ean_etl),  # Nike 360 EAN files → nike_360_ean_main
    (("360", "order", "form"), nike_360_etl),    # Nike 360 Order Form → nike_360_linelist_main
    (("nike", "360"),          nike_360_etl),    # Nike 360 files
    (("360", "maa"),           nike_360_etl),    # 360 MAA files
    (("sp27", "maa"),          sp27_maa_etl),
    (("rookie",),              linelist_etl),    # e.g. SP27 Rookie Linelist → linelist_rookie_main
    (("order", "confirm"),     order_confirm_etl),
    (("orderconfirm",),        order_confirm_etl),
    (("order_confirm",),       order_confirm_etl),
    (("soconfirm",),           order_confirm_etl),
    (("so", "confirm"),      order_confirm_etl),
    (("booking", "confirm"), order_confirm_etl),
    (("confirmed", "order"), order_confirm_etl),
]

# ─────────────────────────────────────────────────────────────────────────────
# Directories needed per file-type (used in _prepare_tmp_dirs)
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_DIRS: list[str] = ["linelist", "mdd", "attributes", "mapping"]


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
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
                log.warning("  Skipping sheet '%s': %s", sheet_name, e)
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
    xlsx_path = csv_path.with_suffix(".xlsx")
    df        = pd.read_csv(str(csv_path), dtype=str, keep_default_na=False)
    wb        = openpyxl.Workbook(write_only=True)
    ws        = wb.create_sheet()
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append([_clean_cell(v) for v in row.tolist()])
    wb.save(str(xlsx_path))
    log.info("  Converted %s → %s", csv_path.name, xlsx_path.name)
    csv_path.unlink()
    return xlsx_path


def _safe_ts_to_epoch(ts) -> int:
    try:
        if hasattr(ts, "timestamp"):
            return int(ts.timestamp())
        elif isinstance(ts, (int, float)):
            return int(float(ts))
        elif isinstance(ts, str):
            from datetime import datetime as _dt
            d = _dt.fromisoformat(ts.replace("Z", "+00:00"))
            return int(d.timestamp())
    except Exception:
        pass
    return 0


def _detect_file_type(filename: str) -> str | None:
    name = filename.lower()
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw in name for kw in keywords):
            return ftype
    return None


def _detect_triggered_file_type(key: str) -> str | None:
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)


def _pick_linelist_variant(filename: str, default_module):
    """
    Inspect the linelist filename and return the matching ETL sub-module.
    Falls back to default_module (linelist_etl) if no variant matches.
    """
    name = filename.lower()
    for tokens, module in LINELIST_VARIANTS:
        if all(tok in name for tok in tokens):
            log.info("  Variant matched: %s → %s", tokens, module.__name__)
            return module
    return default_module


def _is_order_confirm_module(module) -> bool:
    return module is order_confirm_etl


def _parse_event(event: dict) -> tuple[str, str]:
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key: str) -> str | None:
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not follow raw/metadata/{principal}/{file}: %s", key)
    return None


def _list_principal_files(bucket: str, principal: str,
                          exclude_types: set = None,
                          triggered_filename: str = None) -> dict[str, dict]:
    exclude_types = exclude_types or set()
    prefix        = f"raw/metadata/{principal}/"
    paginator     = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None or ftype in exclude_types:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}
                log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

    # Nike 360 files keep using the NEW Brand Mapping file (NIK 360 tab);
    # everything else (rookie, SP27 MAA, order confirmation) uses the V6
    # Attributes List — same "some triggers stay on NEW mapping" carve-out
    # pattern as new_balance/lambda_function.py's _NB_NEW_MAPPING_TRIGGER_TYPES.
    _is_nike_360   = bool(triggered_filename) and "360" in triggered_filename.lower()
    _attr_principal = None if _is_nike_360 else "nike"
    add_root_metadata_files(
        s3_client=s3, bucket=bucket, found=found,
        include_types=ROOT_TYPES, exclude_types=exclude_types,
        principal=_attr_principal, log=log,
    )

    for page in paginator.paginate(Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES or ftype in exclude_types:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}
                log.info("  Classified  %-12s ← %s  [global]", ftype, filename)

    return found


def _parse_metadata_from_linelist_filename(
    dirs: dict,
    triggered_file_type: str = None,
    etl_module=None,
) -> dict:
    import re

    # Order confirmation is data-driven — return a minimal stub;
    # orderConfirmation_main.run() resolves season/brand from the data rows.
    if _is_order_confirm_module(etl_module):
        linelist_dir = dirs.get("linelist")
        if linelist_dir and linelist_dir.exists():
            files = list(linelist_dir.glob("*.xlsx"))
            if files:
                log.info(
                    "Order confirmation variant detected — using data-driven metadata stub "
                    "(file: %s)", files[0].name
                )
                return {
                    "comp_code":                   "0888",
                    "sbu":                         "SP",
                    "brand":                       "Nike",
                    "brand_code":                  "NIK",
                    "season":                      "",   # resolved from data rows
                    "seq":                         1,
                    "multi_mono":                  "Multi",
                    "country_code":                "",
                    "article_type_from_filename":  "",
                }
        log.error("No file found in linelist dir for order confirmation: %s", linelist_dir)
        return {}

    # Standard linelist filename parse
    candidate_dirs = [dirs.get("linelist")]
    seen = set()
    candidate_dirs = [
        d for d in candidate_dirs
        if d is not None and str(d) not in seen and not seen.add(str(d))
    ]

    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        for xlsx_file in folder.glob("*.xlsx"):
            stem = xlsx_file.stem
            log.info("Attempting metadata parse from filename: '%s'", stem)

            parts = re.split(r"\s*-\s*", stem)
            if len(parts) < 7:
                log.warning(
                    "  Filename '%s' has only %d parts (need ≥7) — skipping", stem, len(parts)
                )
                continue

            comp_code = parts[0].strip()
            sbu       = parts[1].strip()

            # ── Extract brand name + brand code from parts[2] ────
            # Supports both:
            #   "NIKE"         → brand="Nike",  brand_code="NIK"  (default)
            #   "NIKE NIK"     → brand="Nike",  brand_code="NIK"  (embedded)
            brand      = "Nike"   # default
            brand_code = "NIK"    # default
            raw_brand_token = parts[2].strip()
            brand_token_match = re.match(
                r"^([A-Za-z][A-Za-z\s]*?)\s+([A-Z]{3})$",
                raw_brand_token,
                re.IGNORECASE,
            )
            if brand_token_match:
                # e.g. "NIKE NIK" → brand="Nike", brand_code="NIK" (from filename)
                brand      = brand_token_match.group(1).strip().title()
                brand_code = brand_token_match.group(2).upper()
                log.info(
                    "  Brand token '%s' → brand=%r  brand_code=%r (embedded in filename)",
                    raw_brand_token, brand, brand_code,
                )
            else:
                # No embedded code — use defaults: brand="Nike", brand_code="NIK"
                brand = raw_brand_token.title()
                log.info(
                    "  Brand token '%s' → brand=%r  brand_code=%r (default fallback)",
                    raw_brand_token, brand, brand_code,
                )

            if not comp_code:
                continue

            season     = None
            season_idx = None
            for i, p in enumerate(parts):
                if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
                    season     = p.strip().upper()
                    season_idx = i
                    break

            if not season or season_idx is None or season_idx < 3:
                continue

            trailing = parts[season_idx + 1:]
            seq      = 1
            for p in reversed(trailing):
                if p.strip().isdigit():
                    seq = int(p.strip())
                    break

            country_code = ""
            for p in trailing:
                p = p.strip()
                if re.match(r"^[A-Z]{2,3}$", p, re.IGNORECASE) and not p.isdigit():
                    country_code = p.upper()
                    break

            multi_mono = parts[season_idx - 1].strip() if season_idx >= 1 else "Multi"

            article_type_from_filename = ""
            for token in parts[3:season_idx]:
                t = token.strip().lower()
                if "inline" in t:
                    article_type_from_filename = "Inline"
                    break
                if "license" in t or "licence" in t:
                    article_type_from_filename = "License"
                    break

            log.info(
                "Parsed: comp_code=%s  sbu=%s  brand=%s  brand_code=%s  season=%s  seq=%s",
                comp_code, sbu, brand, brand_code, season, seq,
            )
            return {
                "comp_code":                  comp_code,
                "sbu":                        sbu,
                "brand":                      brand,
                "brand_code":                 brand_code,
                "season":                     season,
                "seq":                        seq,
                "multi_mono":                 multi_mono,
                "country_code":               country_code,
                "article_type_from_filename": article_type_from_filename,
            }

    log.error("No valid brand file found in candidate dirs.")
    return {}


def _prepare_tmp_dirs() -> dict[str, Path]:
    base       = Path(TMP_WORKDIR)
    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs: dict[str, Path] = {ftype: base / "input" / ftype for ftype in FILE_TYPE_DIRS}
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


def _download_files(bucket, found, dirs, auditor):
    for ftype, info in found.items():
        if ftype not in dirs:
            continue
        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-16s ← s3://%s/%s", ftype, bucket, info["key"])
        try:
            s3.download_file(bucket, info["key"], str(local_path))
            auditor.record_download(ftype, info["filename"], status="ok")
        except Exception as exc:
            log.error("  Download FAILED for %s: %s", info["key"], exc)
            auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
            raise


def _upload_xml_outputs(principal: str) -> list[str]:
    xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
    uploaded = []
    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading %s → s3://%s/%s", xml_file.name, PROCESSED_BUCKET, s3_key)
        s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
        uploaded.append(s3_key)
    return uploaded


def _resolve_brand_code_from_mdd(brand_name: str, mdd_dir: Path) -> str:
    """
    Look up the brand LOV ID directly from the MDD Brand LOV sheet.

    MDD Brand LOV layout:
        Col A = LOV ID      (e.g. "NIK")
        Col B = Brand Name  (e.g. "NIKE")

    Searches Col B for a case-insensitive match against brand_name,
    returns the corresponding Col A value as the brand code.
    Falls back to BRAND_NAME_TO_CODE, then brand_name[:3].upper().
    """
    import re as _re
    brand_upper = (brand_name or "").strip().upper()

    mdd_files = list(mdd_dir.glob("*.xlsx"))
    if not mdd_files:
        log.warning("[brand_code] No MDD file found in %s — falling back to BRAND_NAME_TO_CODE", mdd_dir)
        return BRAND_NAME_TO_CODE.get(brand_upper, brand_upper[:3])

    try:
        wb = openpyxl.load_workbook(str(mdd_files[0]), read_only=True, data_only=True)
        # Find the Brand LOV sheet
        brand_sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()), None
        )
        if not brand_sheet:
            log.warning("[brand_code] No 'Brand LOV' sheet found in MDD — falling back to BRAND_NAME_TO_CODE")
            wb.close()
            return BRAND_NAME_TO_CODE.get(brand_upper, brand_upper[:3])

        rows = list(wb[brand_sheet].iter_rows(values_only=True))
        wb.close()

        for row in rows[1:]:   # skip header
            if not row or len(row) < 2:
                continue
            col_a = str(row[0]).strip() if row[0] else ""   # LOV ID
            col_b = str(row[1]).strip() if row[1] else ""   # Brand Name
            if col_b.upper() == brand_upper:
                log.info(
                    "[brand_code] MDD Brand LOV match → brand_name='%s'  col_B='%s'  col_A(lov_id)='%s'",
                    brand_name, col_b, col_a,
                )
                return col_a

        log.warning(
            "[brand_code] No match for '%s' in MDD Brand LOV col B — falling back to BRAND_NAME_TO_CODE",
            brand_name,
        )
    except Exception as exc:
        log.warning("[brand_code] Error reading MDD Brand LOV: %s — falling back to BRAND_NAME_TO_CODE", exc)

    return BRAND_NAME_TO_CODE.get(brand_upper, brand_upper[:3])


# ═════════════════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event, context, auditor: AuditLogger = None):
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type="direct_invocation", original_key="",
            detected_brand="nike", routing_decision="nike lambda_handler invoked directly",
        )

    log.info("Event received: %s", json.dumps(event))

    bucket, key = _parse_event(event)
    if not bucket or not key:
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    if not key.startswith("raw/metadata/"):
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True}

    principal = _extract_principal(key)
    if not principal:
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    triggered_file_type = _detect_triggered_file_type(key)
    etl_module     = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(triggered_file_type, REQUIRED_TYPES_DEFAULT)

    found = _list_principal_files(
        bucket, principal,
        exclude_types={triggered_file_type} if triggered_file_type else set(),
        triggered_filename=Path(key).name,
    )

    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {
            "key": key, "filename": trigger_filename, "last_modified": None,
        }

    auditor.set_files_classified(found)

    if etl_module is None:
        if "linelist" in found:
            etl_module          = ETL_DISPATCHER["linelist"]
            triggered_file_type = "linelist"
            required_types      = REQUIRED_TYPES_BY_TRIGGER.get("linelist", REQUIRED_TYPES_DEFAULT)

    if etl_module is None:
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 200, "status": "skipped_no_brand_file"}

    missing = required_types - set(found.keys())
    if missing:
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 200, "missing": sorted(missing), "status": "waiting"}

    dirs = _prepare_tmp_dirs()

    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "error": f"Download failed: {exc}"}

    # Convert .xlsb / .csv
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                auditor.record_conversion(xlsb_file.name, ".xlsb", ".xlsx", "ok")
            except Exception as e:
                auditor.set_lambda_status("error", error=f"xlsb conv failed: {e}")
                if _direct_invocation:
                    auditor.flush(principal)
                return {"statusCode": 500, "error": str(e)}
        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                auditor.record_conversion(csv_file.name, ".csv", ".xlsx", "ok")
            except Exception as e:
                auditor.set_lambda_status("error", error=f"csv conv failed: {e}")
                if _direct_invocation:
                    auditor.flush(principal)
                return {"statusCode": 500, "error": str(e)}

    # ── Pick linelist sub-variant (sp27_maa, order_confirm, etc.) ────────────
    if "linelist" in found:
        chosen = _pick_linelist_variant(found["linelist"]["filename"], etl_module)
        if chosen is not etl_module:
            etl_module = chosen

    meta = _parse_metadata_from_linelist_filename(
        dirs,
        triggered_file_type=triggered_file_type,
        etl_module=etl_module,
    )
    if not meta:
        auditor.set_lambda_status("error", error="Unparseable metadata in filename")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 400, "error": "Unparseable metadata"}

    auditor.set_metadata_parsed(meta)

    # ── Build args namespace ─────────────────────────────────────────────────
    args = types.SimpleNamespace(
        brand                      = meta["brand"],
        brand_code                 = meta["brand_code"],
        comp_code                  = meta["comp_code"],
        sbu                        = meta["sbu"],
        season                     = meta.get("season", ""),
        seq                        = meta.get("seq", 1),
        multi_mono                 = meta.get("multi_mono", "Multi"),
        country_code               = meta.get("country_code", ""),
        article_type_from_filename = meta.get("article_type_from_filename", ""),
    )

    # For order confirmation the input file lives in the linelist dir;
    # pass its path explicitly so orderConfirmation_main.run() can find it.
    if _is_order_confirm_module(etl_module):
        linelist_dir = dirs["linelist"]
        order_file   = next(linelist_dir.glob("*.xlsx"), None)
        args.order_confirm_path = str(order_file) if order_file else None
        log.info("Order confirmation input: %s", args.order_confirm_path)

    log.info("Running Nike ETL → %s ...", etl_module.__name__)
    etl_module.run(args, auditor=auditor)

    uploaded = _upload_xml_outputs(principal)
    if not uploaded:
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 200, "status": "no_xml_generated"}

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct_invocation:
        auditor.flush(principal)

    # ── Resolve brand LOV ID for inclusion in return payload ────────────────
    _brand_lov_id   = meta.get("brand_code", "")
    _brand_label    = meta.get("brand", "")
    try:
        from nike.linelist_rookie_main import resolve_brand_lov_id, MDDLoader as _MDDLoader
        _mdd_dir   = Path(TMP_WORKDIR) / "input" / "mdd"
        _mdd_files = list(_mdd_dir.glob("*.xlsx"))
        if _mdd_files:
            _mdd = _MDDLoader(_mdd_files[0])
            _brand_lov_id, _brand_label = resolve_brand_lov_id(
                _mdd, brand_name=meta.get("brand", ""), fallback_code=meta.get("brand_code", ""),
            )
    except Exception as _e:
        log.warning("[lambda return] Could not resolve brand LOV: %s", _e)

    return {
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_file_type,
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
        "brand_code_input":    meta.get("brand_code", ""),
        "brand_lov_id":        _brand_lov_id,
        "brand_label":         _brand_label,
    }