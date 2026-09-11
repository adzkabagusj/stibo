"""
lambda_function.py — Lambda Handler for 2XU
============================================
Function name : map-stibo-inbound-validate-transform-2xu-{env}

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-{env}/raw/metadata/2xu/*

2XU-specific flow vs Smiggle:
    • Only ONE input file from the brand : Product Bible (.xlsx)
    • No TDD, no Backlog, no Linelist diff
    • mdd + attributes are shared/global files
    • Required types: {"productbible", "mdd", "attributes"}
    • Product Bible filename keyword detection looks for:
        "product bible", "productbible", "product_bible", "bible",
        "sku list", "skulist"

File-type dispatch:
    productbible → twoxu.product_bible_main.run()

S3 layout:
    raw/metadata/2xu/{product_bible_file}.xlsx
    raw/metadata/mdd_file.xlsx                    ← shared global MDD
    raw/metadata/attributes_list.xlsx             ← shared global attr list

Output:
    s3://map-stibo-inbound-processed-{env}/
        processed/stepxml/2xu/{filename}.xml

NOTE on package name: Python identifiers cannot start with a digit, so the
brand package is `twoxu/` but the S3 principal/folder remains `2xu`.
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
# Set LAMBDA_TMP_DIR BEFORE importing 2XU modules so that their module-level
# BASE_DIR resolves to /tmp/stibo_workdir_2xu (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_2xu"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules — one per 2XU file type ──────────────────────────────────────
import twoxu.product_bible_main      as productbible_etl    # noqa: E402
import twoxu.order_confirmation_main      as orderconfirmation_etl  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_2xu")

# ─────────────────────────────────────────────────────────────────────────────
# AWS client
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# Environment variables
# ─────────────────────────────────────────────────────────────────────────────
RAW_BUCKET       = os.environ.get("RAW_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# ─────────────────────────────────────────────────────────────────────────────
# Brand lookup
# ─────────────────────────────────────────────────────────────────────────────
BRAND_NAME_TO_CODE: dict[str, str] = {
    "ADIDAS":      "ADI",
    "NIKE":        "NIK",
    "NEW BALANCE": "NEW",
    "SMIGGLE":     "SMI",
    "ALDO":        "ALD",
    "CROCS":       "CRO",
    "LOTTO":       "LOT",
    "BIRKENSTOCK": "BIR",
    "2XU":         "2XU",
}

# ─────────────────────────────────────────────────────────────────────────────
# File type detection
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # 2XU primary input (Product Bible workbook)
    "productbible": [
        "product bible", "productbible", "product_bible",
        "bible",
        "sku list", "skulist", "sku_list",
    ],
    # 2XU secondary input (Order Confirmation workbook)
    "orderconfirmation": [
        "order confirmation", "orderconfirmation", "order_confirmation",
        "order conf", "orderconf", "order_conf",
    ],
    # Shared / global
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": ["attributes", "attributes_list", "attribute list"],
}

# ── Required types per triggered file type ────────────────────────────────────
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "productbible":      {"productbible", "mdd", "attributes"},
    "orderconfirmation": {"orderconfirmation", "mdd", "attributes"},
}
REQUIRED_TYPES_DEFAULT: set[str] = {"productbible", "mdd", "attributes"}

# Global types fetched from raw/metadata/ root (not inside 2xu/)
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}

# ── ETL dispatcher ────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "productbible":      productbible_etl,
    "orderconfirmation": orderconfirmation_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    xlsx_path = xlsb_path.with_suffix(".xlsx")
    wb_out    = openpyxl.Workbook(write_only=True)
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
    xlsx_path = csv_path.with_suffix(".xlsx")
    df = pd.read_csv(str(csv_path), dtype=str, keep_default_na=False)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet()
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append([_clean_cell(v) for v in row.tolist()])
    wb.save(str(xlsx_path))
    csv_path.unlink()
    return xlsx_path


def _safe_ts_to_epoch(ts) -> int:
    try:
        if hasattr(ts, "timestamp"):
            return int(ts.timestamp())
        if isinstance(ts, (int, float)):
            return int(float(ts))
        if isinstance(ts, str):
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
                          exclude_types: set = None) -> dict[str, dict]:
    exclude_types = exclude_types or set()
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None:
                log.warning("  Unrecognised file — skipping: %s", filename)
                continue
            if ftype in exclude_types:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified
                }
                log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

        add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=GLOBAL_TYPES,
        exclude_types=exclude_types,
        principal="2xu",
        log=log,
    )

    return found


def _parse_metadata_from_filename(dirs: dict, triggered_file_type: str = None) -> dict:
    """
    Parse comp_code / sbu / brand / season / seq / multi_mono / country / article_type
    from the triggered brand file's filename.

    Convention:
        {COMP}-{SBU}-{BRAND}-{FileType}-{Multi/Mono}-{SEASON}-{SEQ}.xlsx
    Example:
        0888-SP-2XU-Product Bible-Multi-H226-1.xlsx
    """
    import re

    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs = [
            dirs.get(triggered_file_type),
            dirs.get("productbible"),
            dirs.get("orderconfirmation"),
        ]
    else:
        candidate_dirs = [
            dirs.get("productbible"),
            dirs.get("orderconfirmation"),
        ]

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
                    "  Filename '%s' has only %d '-'-separated parts (need ≥7) — "
                    "falling back to defaults.", stem, len(parts),
                )
                continue

            comp_code = parts[0].strip()
            sbu       = parts[1].strip()
            brand     = parts[2].strip()   # keep raw — 2XU should stay as "2XU"
            if not comp_code:
                continue


            # ── Locate season token by regex (e.g. H226, SS26, WN2002, FW2026) ──
            # First char must be a letter so numeric comp codes like "0888" are excluded.
            season     = None
            season_idx = None
            for i, p in enumerate(parts):
                if re.match(r'^[A-Z][A-Z0-9]\d{2,4}$', p.strip(), re.IGNORECASE):
                    season     = p.strip().upper()
                    season_idx = i
                    break

            if not season or season_idx is None or season_idx < 3:
                log.warning("  No valid season token in '%s' — skipping", stem)
                continue

            trailing = parts[season_idx + 1:]

            seq = 1
            for p in reversed(trailing):
                if p.strip().isdigit():
                    seq = int(p.strip())
                    break

            country_code = ""
            for p in trailing:
                pp = p.strip()
                if re.match(r'^[A-Z]{2,3}$', pp, re.IGNORECASE) and not pp.isdigit():
                    country_code = pp.upper()
                    break

            multi_mono = parts[season_idx - 1].strip() if season_idx >= 1 else "Multi"

            brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand.upper()[:3])

            article_type_from_filename = ""
            for token in parts[3:season_idx]:
                t = token.strip().lower()
                if "inline" in t:
                    article_type_from_filename = "Inline"; break
                if "license" in t or "licence" in t:
                    article_type_from_filename = "License"; break

            log.info(
                "Parsed from filename: comp_code=%s  sbu=%s  brand=%s  brand_code=%s  "
                "season=%s  seq=%s  multi_mono=%s  country=%s  article_type=%s",
                comp_code, sbu, brand, brand_code, season, seq, multi_mono,
                country_code, article_type_from_filename,
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

    # ── Fallback defaults — 2XU Product Bible may arrive without the
    # MAP naming convention (it comes directly from the brand).  Use
    # the file's stem to extract season if possible.
    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        for xlsx_file in folder.glob("*.xlsx"):
            stem = xlsx_file.stem
            m = re.search(r'(?<![A-Z0-9])([A-Z][A-Z0-9])[\s_-]?(\d{2,4})(?!\d)', stem, re.IGNORECASE)
            if m:
                hp, yr = m.group(1).upper(), m.group(2)
                season = f"{hp}{yr}"
            else:
                season = "H226"
            log.warning(
                "Falling back to defaults for '%s' — season=%s", stem, season,
            )
            return {
                "comp_code":  "0888",
                "sbu":        "SP",
                "brand":      "2XU",
                "brand_code": "2XU",
                "season":     season,
                "seq":        1,
                "multi_mono": "Multi",
                "country_code":              "AU",
                "article_type_from_filename": "Inline",
            }

    return {}


def _prepare_tmp_dirs() -> dict[str, Path]:
    base = Path(TMP_WORKDIR)
    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))
    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs: dict[str, Path] = {
        "productbible":      base / "input" / "productbible",
        "orderconfirmation": base / "input" / "orderconfirmation",
        "mdd":               base / "input" / "mdd",
        "attributes":        base / "input" / "attributes",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


def _download_files(bucket, found, dirs, auditor):
    for ftype, info in found.items():
        if ftype not in dirs:
            continue
        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-12s ← s3://%s/%s", ftype, bucket, info["key"])
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


# ═════════════════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event, context, auditor: AuditLogger = None):
    """
    Main Lambda entry point for 2XU.
    """
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "2xu",
            routing_decision = "2xu lambda_handler invoked directly",
        )


    # ── 1. Parse event ────────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Guard: only raw/metadata/ keys ─────────────────────────
    if not key.startswith("raw/metadata/"):
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}

    # ── 3. Extract principal ──────────────────────────────────────
    principal = _extract_principal(key)
    if not principal:
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    log.info("Principal: %s", principal)

    # ── 3b. Detect which file type triggered this invocation ──────
    triggered_file_type = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_file_type or "unknown/global-fanout")

    etl_module     = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(
        triggered_file_type, REQUIRED_TYPES_DEFAULT,
    )

    # ── 4. List + classify files ──────────────────────────────────
    found = _list_principal_files(
        bucket, principal,
        exclude_types={triggered_file_type} if triggered_file_type else set(),
    )

    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {
            "key":           key,
            "filename":      trigger_filename,
            "last_modified": None,
        }
        log.info("Pinned trigger file → type='%s'  file='%s'",
                 triggered_file_type, trigger_filename)

    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    # ── 4b. If triggered by global fan-out, pick ETL module from
    #         whatever brand file is present in S3 ────────────────
    if etl_module is None:
        for ftype in ("productbible", "orderconfirmation"):
            if ftype in found:
                etl_module          = ETL_DISPATCHER[ftype]
                triggered_file_type = ftype
                required_types      = REQUIRED_TYPES_BY_TRIGGER.get(
                    ftype, REQUIRED_TYPES_DEFAULT,
                )
                log.info("Global fan-out: resolved ETL module from present files → '%s'",
                         ftype)
                break

    if etl_module is None:
        log.info(
            "No brand file present yet for '%s' during global fan-out — skipping.",
            principal,
        )
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "principal": principal, "status": "skipped_no_brand_file"}

    # ── 5. Check mandatory files ──────────────────────────────────
    missing = required_types - set(found.keys())
    if missing:
        log.info("Mandatory files not yet present for '%s' — waiting for: %s",
                 principal, sorted(missing))
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation: auditor.flush(principal)
        return {
            "statusCode": 200, "principal": principal,
            "present": sorted(found.keys()), "missing": sorted(missing),
            "status":  "waiting_for_mandatory_files",
        }

    # ── 6. Prepare /tmp/ directories ──────────────────────────────
    dirs = _prepare_tmp_dirs()

    # ── 7. Download files ─────────────────────────────────────────
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "principal": principal,
                "error": f"Download failed: {exc}"}

    # ── 7b. Convert .xlsb / .csv → .xlsx ──────────────────────────
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                auditor.record_conversion(xlsb_file.name, from_ext=".xlsb",
                                          to_ext=".xlsx", status="ok")
            except Exception as e:
                auditor.record_conversion(xlsb_file.name, from_ext=".xlsb",
                                          to_ext=".xlsx", status="failed", error=str(e))
                auditor.set_lambda_status("error",
                                          error=f"xlsb conversion failed: {xlsb_file.name}: {e}")
                if _direct_invocation: auditor.flush(principal)
                return {"statusCode": 500, "principal": principal,
                        "error": f"xlsb conversion failed: {xlsb_file.name}: {e}"}
        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                auditor.record_conversion(csv_file.name, from_ext=".csv",
                                          to_ext=".xlsx", status="ok")
            except Exception as e:
                auditor.record_conversion(csv_file.name, from_ext=".csv",
                                          to_ext=".xlsx", status="failed", error=str(e))
                auditor.set_lambda_status("error",
                                          error=f"csv conversion failed: {csv_file.name}: {e}")
                if _direct_invocation: auditor.flush(principal)
                return {"statusCode": 500, "principal": principal,
                        "error": f"csv conversion failed: {csv_file.name}: {e}"}

    # ── 8. Parse pipeline metadata from filename (with fallback) ──
    meta = _parse_metadata_from_filename(dirs, triggered_file_type=triggered_file_type)
    if not meta:
        auditor.set_lambda_status("error",
                                  error="Unparseable metadata in Product Bible filename")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 400, "principal": principal,
                "error": "Unparseable metadata in Product Bible filename"}

    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand                      = meta["brand"],
        brand_code                 = meta["brand_code"],
        comp_code                  = meta["comp_code"],
        sbu                        = meta["sbu"],
        season                     = meta["season"],
        seq                        = meta["seq"],
        multi_mono                 = meta.get("multi_mono", "Multi"),
        country_code               = meta.get("country_code", ""),
        article_type_from_filename = meta.get("article_type_from_filename", ""),
    )

    # ── 9. Run ETL ────────────────────────────────────────────────
    log.info("Running 2XU ETL → %s (triggered_file_type=%s) ...",
             etl_module.__name__, triggered_file_type)
    etl_module.run(args, auditor=auditor)

    # ── 10. Upload generated XMLs ─────────────────────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct_invocation: auditor.flush(principal)

    pb_info = found.get("productbible", {})
    return {
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_file_type,
        "input_used":          pb_info.get("filename"),
        "input_s3_key":        pb_info.get("key"),
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
    }