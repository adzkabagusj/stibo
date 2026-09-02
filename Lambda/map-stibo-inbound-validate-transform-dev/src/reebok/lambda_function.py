"""
lambda_function.py — Lambda Handler for Reebok
================================================
Function name : map-stibo-inbound-validate-transform-reebok-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/reebok/*

Reebok-specific flow:
    • Input files from the brand:
        - apparel   → Article Master data with image (APPAREL)
                      e.g. "Reebok article master data with image - APPAREL (SS27).xlsx"
        - footwear  → Article Master data with image (FOOTWEAR)
                      e.g. "Reebok article master data with image - FOOTWEAR (SS27).xlsx"
    • mdd + attributes are shared/global files stored at bucket root
    • Required types per trigger:
        apparel   → {"apparel", "mdd", "attributes"}
        footwear  → {"footwear", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        apparel   → reebok.article_master_apparel.run()
        footwear  → reebok.article_master_footwear.run()

S3 layout:
    raw/metadata/reebok/{input_file}.xlsx              ← Reebok uploads here
    raw/metadata/mdd_file.xlsx                          ← shared global MDD
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global brand mapping

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/reebok/{filename}.xml
"""

import logging
import os
import re
import sys
import types
import pyxlsb
import openpyxl
from pathlib import Path

import boto3
import pandas as pd

from audit_logger import AuditLogger

# ─────────────────────────────────────────────────────────────────────────────
# Set LAMBDA_TMP_DIR BEFORE importing reebok modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_reebok
# (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_reebok"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import reebok.article_master_apparel as apparel_etl  # noqa: E402  # Article Master (Apparel)
import reebok.article_master_footwear as footwear_etl  # noqa: E402  # Article Master (Footwear)
import reebok.ean_main as ean_etl
import reebok.recap_main as recap_etl  # noqa: E402  # EAN Source (barcode data — apparel + footwear)

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_reebok")

# ─────────────────────────────────────────────────────────────────────────────
# AWS client
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# Environment variables
# ─────────────────────────────────────────────────────────────────────────────
RAW_BUCKET       = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# ─────────────────────────────────────────────────────────────────────────────
# File type detection
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # Reebok EAN Source (barcode data — covers both apparel and footwear).
    # Checked BEFORE "footwear"/"apparel" below: the real filename embeds
    # "Footwear" (e.g. "... EAN Source (MAA Fashion Footwear) Inline..."),
    # which would otherwise false-match the "footwear" keyword first.
    # Example: "0888-FF-REEBOK-REEBOK_Article Attributes EAN Source
    #           (MAA Fashion Footwear) Inline-Multi-SM2027-ID-1.xlsx"
    "ean": ["ean source", "article attributes ean source"],
    # Reebok Apparel Article Master
    # Example: "Reebok article master data with image - APPAREL (SS27).xlsx"
    "apparel": ["apparel"],
    # Reebok Footwear Article Master
    # Example: "Reebok article master data with image - FOOTWEAR (SS27).xlsx"
    "footwear": ["footwear"],
    "recap": ["recap", "recap sample"],
    # Shared / global
    # NOTE: kept specific ("master data dictionary" / "mdd") on purpose —
    # Reebok's own principal filename contains the substring "article
    # master data", which would otherwise collide with a looser "master
    # data" keyword and mis-classify the brand file as the global MDD.
    "mdd": ["master data dictionary", "mdd"],
    "attributes": [
        "brand mapping", "brand_mapping", "new - brand mapping files template",
        "mapping template", "brand template",
    ],
}

# ── Required types per triggered file type ────────────────────────────────────
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "apparel":  {"apparel", "mdd", "attributes"},
    "footwear": {"footwear", "mdd", "attributes"},
    "recap":    {"recap", "mdd", "attributes"},
    # ean_main.py only needs the MDD (for the Size LOV lookup) — it emits
    # barcode data only, no RNA/brand-mapping attributes.
    "ean":      {"ean", "mdd"},
}

REQUIRED_TYPES_DEFAULT: set[str] = {"apparel", "mdd", "attributes"}

# Global types fetched from raw/metadata/ root (not inside reebok/)
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}

# ── ETL dispatcher ────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "apparel":  apparel_etl,
    "footwear": footwear_etl,
    "recap":    recap_etl,
    "ean":      ean_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    """Strip illegal XML / Excel characters from string cell values."""
    if not isinstance(v, str):
        return v
    return "".join(
        c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r"
    )


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    """Convert .xlsb → .xlsx streaming row-by-row."""
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
    log.info(
        "  Converted %s → %s (%d sheets)",
        xlsb_path.name, xlsx_path.name, sheets_written,
    )
    xlsb_path.unlink()
    return xlsx_path


def _convert_csv_to_xlsx(csv_path: Path) -> Path:
    """Convert .csv → .xlsx."""
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


def _safe_ts_to_epoch(ts) -> int:
    """Convert ANY timestamp type to Unix epoch seconds. Returns 0 on failure."""
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
    """Return the file type key for a filename, or None if unrecognised."""
    name = filename.lower()
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw in name for kw in keywords):
            return ftype
    return None


def _detect_triggered_file_type(key: str) -> str | None:
    """Derive the file type from the S3 key that triggered this invocation."""
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)


def _parse_event(event: dict) -> tuple[str, str]:
    """Extract (bucket, key) from an EventBridge S3 Object Created event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key: str) -> str | None:
    """Derive principal from key: raw/metadata/{principal}/{filename}."""
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not follow raw/metadata/{principal}/{file}: %s", key)
    return None


def _list_principal_files(bucket: str, principal: str, exclude_types: set = None) -> dict[str, dict]:
    """
    Scan S3 for all recognised files:
      1. Brand-specific files  ← raw/metadata/{principal}/
      2. Bucket root           ← s3://bucket/ root  (mdd + attributes)

    Brand-specific takes priority if both exist for the same type.
    Bucket root always wins for attributes and mdd.
    """
    exclude_types = exclude_types or set()
    prefix        = f"raw/metadata/{principal}/"
    paginator     = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    # ── Step 1: brand-specific folder ───────────────────────────
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
                log.warning("  Unrecognised file — skipping: %s", filename)
                continue
            if ftype in exclude_types:
                log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified)
                > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified,
                }
                log.info("  Classified  %-18s ← %s  [brand-specific]", ftype, filename)

    # ── Step 2: bucket root — MDD + attributes ─────────────────
    try:
        resp = s3.list_objects_v2(Bucket=bucket, Delimiter="/")
        for obj in resp.get("Contents", []):
            obj_key  = obj["Key"]
            filename = Path(obj_key).name
            if not filename.lower().endswith((".xlsx", ".xlsm")):
                continue

            ftype = _detect_file_type(filename)

            if ftype == "attributes" and "attributes" not in exclude_types:
                found["attributes"] = {
                    "key":           obj_key,
                    "filename":      filename,
                    "last_modified": obj["LastModified"],
                }
                log.info("  Classified  attributes         ← %s  [bucket root]", filename)

            elif ftype == "mdd" and "mdd" not in exclude_types:
                found["mdd"] = {
                    "key":           obj_key,
                    "filename":      filename,
                    "last_modified": obj["LastModified"],
                }
                log.info("  Classified  mdd                ← %s  [bucket root]", filename)
    except Exception as e:
        log.warning("[bucket root] Scan failed: %s", e)

    return found


def _prepare_tmp_dirs() -> dict:
    """Create /tmp/ subdirectories for each file type."""
    workdir = Path(TMP_WORKDIR)
    workdir.mkdir(parents=True, exist_ok=True)

    dirs = {
        "apparel":    workdir / "input" / "apparel",
        "footwear":   workdir / "input" / "footwear",
        "ean":        workdir / "input" / "ean",
        "recap":      workdir / "input" / "recap",
        "mdd":        workdir / "input" / "mdd",
        "attributes": workdir / "input" / "attributes",
        "xml":        workdir / "output" / "xml",
        "logs":       workdir / "output" / "logs",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    return dirs


def _download_files(bucket: str, found: dict, dirs: dict, auditor: AuditLogger):
    """Download all found files to their corresponding /tmp/ directories."""
    for ftype, meta in found.items():
        key      = meta["key"]
        filename = meta["filename"]

        if ftype in dirs:
            target_dir = dirs[ftype]
        elif ftype in GLOBAL_TYPES:
            target_dir = dirs.get("mdd") if ftype == "mdd" else dirs.get("attributes")
        else:
            log.warning("  No target directory for type '%s' — skipping", ftype)
            continue

        local_path = target_dir / filename
        log.info("  Downloading %s → %s", key, local_path.name)

        try:
            s3.download_file(bucket, key, str(local_path))
            auditor.record_download(key, str(local_path), status="ok")
        except Exception as e:
            log.error("  Download failed: %s — %s", key, e)
            auditor.record_download(key, str(local_path), status="failed", error=str(e))
            raise


def _upload_xml_outputs(principal: str) -> list[str]:
    """Upload all generated XML files to S3 processed bucket."""
    xml_dir = Path(TMP_WORKDIR) / "output" / "xml"
    if not xml_dir.exists():
        return []

    uploaded = []
    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading %s → s3://%s/%s", xml_file.name, PROCESSED_BUCKET, s3_key)
        try:
            s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
            uploaded.append(s3_key)
        except Exception as e:
            log.error("  Upload failed: %s — %s", xml_file.name, e)
            raise

    return uploaded


# ═════════════════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event, context, auditor=None):
    """
    Main Lambda entry point for Reebok ETL.

    Flow:
    1. Parse S3 event → extract bucket + key
    2. Detect triggered file type
    3. Determine required files
    4. List + download all required files
    5. Convert xlsb/csv → xlsx if needed
    6. Parse metadata from filename
    7. Run appropriate ETL module
    8. Upload generated XML
    """
    log.info("=" * 80)
    log.info("REEBOK ETL INVOKED")
    log.info("=" * 80)
    log.info("Event: %s", event)

    _direct_invocation = event.get("direct_invocation", False)

    # ── 1. Parse event ────────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Invalid event — missing bucket or key")
        return {"statusCode": 400, "error": "Invalid event structure"}

    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Extract principal ──────────────────────────────────────
    principal = _extract_principal(key)
    if not principal or principal != "reebok":
        log.error("Key does not belong to reebok principal: %s", key)
        return {"statusCode": 400, "error": "Not a reebok file"}

    if auditor is None:
        auditor = AuditLogger(
            principal=principal,
            s3_bucket=bucket,
            s3_key=key,
            lambda_request_id=context.request_id if context else "local",
        )
    auditor.set_lambda_status("started")

    # ── 3. Detect triggered file type ────────────────────────────
    triggered_file_type = _detect_triggered_file_type(key)
    if not triggered_file_type:
        log.warning("Cannot classify triggered file: %s", key)
        auditor.set_lambda_status("unrecognised_file_type")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "unrecognised_file_type",
        }

    log.info("Triggered file type: %s", triggered_file_type)
    auditor.set_router_context(
        trigger_type=triggered_file_type,
        original_key=key,
        detected_brand=principal,
        routing_decision=triggered_file_type,
    )

    # ── 4. Determine required file types ──────────────────────────
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(
        triggered_file_type, REQUIRED_TYPES_DEFAULT
    )
    log.info("Required file types: %s", sorted(required_types))

    # ── 5. Determine ETL module ───────────────────────────────────
    etl_module = ETL_DISPATCHER.get(triggered_file_type)
    if not etl_module:
        log.error("No ETL module mapped for file type: %s", triggered_file_type)
        auditor.set_lambda_status("error", error=f"No ETL module for {triggered_file_type}")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 500,
            "principal": principal,
            "error": f"No ETL module for {triggered_file_type}",
        }

    # ── 6. List all principal files ───────────────────────────────
    log.info("Scanning S3 for required files ...")
    found = _list_principal_files(bucket, principal)

    if triggered_file_type not in found:
        log.warning(
            "Triggered file type '%s' not found in scan — may have been deleted",
            triggered_file_type,
        )
        auditor.set_lambda_status("triggered_file_missing")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "skipped_no_brand_file",
        }

    # ── 7. Check mandatory files ──────────────────────────────────
    missing = required_types - set(found.keys())
    if missing:
        log.info(
            "Mandatory files not yet present for '%s' — waiting for: %s",
            principal, sorted(missing),
        )
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    sorted(missing),
            "status":     "waiting_for_mandatory_files",
        }

    # ── 8. Prepare /tmp/ directories ─────────────────────────────
    dirs = _prepare_tmp_dirs()

    # ── 9. Download files ─────────────────────────────────────────
    log.info("Downloading %d files ...", len(found))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 500, "principal": principal,
            "error": f"Download failed: {exc}",
        }

    # ── 10. Convert .xlsb / .csv → .xlsx ─────────────────────────
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                auditor.record_conversion(
                    xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
                )
            except Exception as e:
                log.error("xlsb conversion failed: %s — %s", xlsb_file.name, e)
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
                    "error": f"xlsb conversion failed: {xlsb_file.name}: {e}",
                }

        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                auditor.record_conversion(
                    csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
                )
            except Exception as e:
                log.error("csv conversion failed: %s — %s", csv_file.name, e)
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
                    "error": f"csv conversion failed: {csv_file.name}: {e}",
                }

    # ── 11. Parse pipeline metadata from filename ──────────────────
    # Use the resolved module's own copy — apparel/footwear currently have
    # identical logic, but this keeps the two from silently diverging.
    meta = etl_module._parse_metadata_from_filename(
        dirs, triggered_file_type=triggered_file_type,
    )
    if not meta:
        log.error(
            "Cannot parse metadata from brand filename. "
            "Ensure file is named like: "
            "{CompanyCode}-{SBU}-Reebok-{FileType}-{Season}-{Seq}.xlsx  "
            "or 'Reebok article master data with image - APPAREL ({Season}).xlsx'"
        )
        auditor.set_lambda_status(
            "error", error="Unparseable metadata in brand filename"
        )
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 400, "principal": principal,
            "error": "Unparseable metadata in brand filename",
        }

    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand        = meta["brand"],
        brand_code   = meta["brand_code"],
        comp_code    = meta["comp_code"],
        sbu          = meta["sbu"],
        season       = meta["season"],
        seq          = meta["seq"],
        country_code = meta.get("country_code", ""),
    )

    # ── 12. Run ETL — dispatched to the correct module ─────────────
    log.info(
        "Running Reebok ETL → %s (triggered_file_type=%s) ...",
        etl_module.__name__, triggered_file_type,
    )
    try:
        etl_module.run(args, auditor=auditor)
    except Exception as exc:
        log.error("ETL execution failed: %s", exc, exc_info=True)
        auditor.set_lambda_status("error", error=f"ETL failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 500, "principal": principal,
            "error": f"ETL failed: {exc}",
        }

    # ── 13. Upload generated XMLs ─────────────────────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("ETL completed but no XML files were produced.")
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "no_xml_generated",
        }

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct_invocation:
        auditor.flush(principal)

    trigger_info = found.get(triggered_file_type, {})
    return {
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_file_type,
        "source_file":         trigger_info.get("filename"),
        "source_s3_key":       trigger_info.get("key"),
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
    }
