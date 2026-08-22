"""
lambda_function.py — Lambda Handler for Clarks
================================================
Function name : map-stibo-inbound-validate-transform-clarks-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/clarks/*

Clarks-specific flow:
    • Input files from the brand:
        - order_form_accs     → Accessories Order Form (e.g. RAW Order Form Clarks Accs SS27.xlsx)
        - order_form_footwear → Footwear Order Form
        - ean                 → EAN Source (e.g. RAW Delivery Performance Clarks Footwear & Accs...)
    • mdd + attributes are shared/global files stored at bucket root
    • Required types per trigger:
        order_form_accs     → {"order_form_accs", "mdd", "attributes"}
        order_form_footwear → {"order_form_footwear", "mdd", "attributes"}
        ean                 → {"ean", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        order_form_accs     → clarks.order_form_accs_main.run()
        order_form_footwear → clarks.order_form_footwear_main.run()
        ean                 → clarks.ean_source_main.run()

S3 layout:
    raw/metadata/clarks/{input_file}.xlsx             ← Clarks uploads here
    raw/metadata/mdd_file.xlsx                        ← shared global MDD
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global brand mapping

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/clarks/{filename}.xml
"""

import logging
import os
import re
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
# Set LAMBDA_TMP_DIR BEFORE importing clarks modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_clarks
# (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_clarks"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import clarks.order_form_accs_main as order_form_accs_etl  # noqa: E402  # Accessories Order Form
import clarks.order_form_footwear_main as order_form_footwear_etl  # noqa: E402  # Footwear Order Form
import clarks.ean_source_main as ean_etl  # noqa: E402  # EAN Source

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_clarks")

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
    # Clarks EAN Source (check FIRST - most specific)
    # Example: "0888-FF-CLARKS-RAW Delivery Performance Clarks Footwear & Accs (MAA Fashion Footwear) Inline-Multi-SP2027-ID-1"
    "ean": [
        "delivery performance clarks footwear & accs",
        "delivery performance clarks footwear and accs",
        "delivery performance clarks",
        "raw delivery performance",
        "delivery performance",
    ],
    # Clarks Accessories Order Form
    # Example: "RAW Order Form Clarks Accs SS27.xlsx"
    "order_form_accs": [
        "order form clarks accs",
        "order form accs",
        "clarks accs",
    ],
    # Clarks Footwear Order Form
    # Example: "0888-FF-CLARKS-RAW Order Form Clarks Footwear SS27 (MAA Fashion Footwear) Inline-Multi-SP2029-ID-1"
    "order_form_footwear": [
        "raw order form clarks footwear",
        "order form clarks footwear",
    ],
    # Shared / global
    "mdd": ["mdd", "master data", "master data dictionary", "master_data"],
    "attributes": [
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template",
    ],
}

# ── Required types per triggered file type ────────────────────────────────────
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "order_form_accs": {"order_form_accs", "mdd", "attributes"},
    "order_form_footwear": {"order_form_footwear", "mdd", "attributes"},
    "ean": {"ean", "mdd", "attributes"},
}

REQUIRED_TYPES_DEFAULT: set[str] = {"order_form_accs", "order_form_footwear", "ean", "mdd", "attributes"}

# Global types fetched from raw/metadata/ root (not inside clarks/)
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}

# ── ETL dispatcher ────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "order_form_accs": order_form_accs_etl,
    "order_form_footwear": order_form_footwear_etl,
    "ean": ean_etl,
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
    """
    Derive the file type from the S3 key that triggered this invocation.
    Returns a key from FILE_TYPE_KEYWORDS or None.
    """
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
    """
    Derive principal from key:  raw/metadata/{principal}/{filename}
    Returns the principal (e.g. "clarks") or None.
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not follow raw/metadata/{principal}/{file}: %s", key)
    return None


def _list_principal_files(
    bucket:        str,
    principal:     str,
    exclude_types: set = None,
) -> dict[str, dict]:
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


def _parse_metadata_from_filename(
    dirs: dict,
    triggered_file_type: str = None,
) -> dict:
    """
    Parse comp_code, sbu, brand, brand_code, season, seq from the
    triggered brand file's filename.

    Expected format (flexible):
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{Season}-{Seq}.xlsx
        RAW Order Form Clarks Accs SS27.xlsx   ← also accepted (no dashes)

    Falls back to safe defaults when the filename does not match the strict
    dash-separated pattern (which is common for Clarks RAW files).
    """
    candidate_dirs = []
    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs.append(dirs[triggered_file_type])

    seen: set = set()
    candidate_dirs = [
        d for d in candidate_dirs
        if d is not None and str(d) not in seen and not seen.add(str(d))
    ]

    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        xls_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
        for xlsx_file in xls_files:
            stem = xlsx_file.stem
            log.info("Attempting metadata parse from filename: '%s'", stem)

            # ── Try strict dash-separated format first ────────────
            parts = re.split(r"\s*-\s*", stem)
            if len(parts) >= 6:
                comp_code = parts[0].strip()
                sbu       = parts[1].strip()
                brand     = parts[2].strip().title()
                brand_code = "CKS"

                # Locate season token
                season = None
                season_idx = None
                for i, p in enumerate(parts):
                    if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
                        season     = p.strip().upper()
                        season_idx = i
                        break

                if season and season_idx:
                    trailing     = parts[season_idx + 1:]
                    seq          = next((int(p) for p in reversed(trailing) if p.strip().isdigit()), 1)
                    country_code = next(
                        (p.strip().upper() for p in trailing
                         if re.match(r"^[A-Z]{2,3}$", p.strip(), re.IGNORECASE) and not p.strip().isdigit()),
                        "",
                    )
                    log.info(
                        "Parsed (strict): comp=%s sbu=%s brand=%s code=%s season=%s seq=%s country=%s",
                        comp_code, sbu, brand, brand_code, season, seq, country_code,
                    )
                    return {
                        "comp_code":   comp_code,
                        "sbu":         sbu,
                        "brand":       brand,
                        "brand_code":  brand_code,
                        "season":      season,
                        "seq":         seq,
                        "multi_mono":  "Multi",
                        "country_code": country_code,
                    }

            # ── Fallback: extract season from anywhere in filename ─
            season_match = re.search(r"([A-Z]{2}\d{2,4})", stem, re.IGNORECASE)
            season = season_match.group(1).upper() if season_match else "SS27"

            log.info(
                "Parsed (fallback): comp=0888 sbu=SP brand=Clarks code=CLK season=%s", season,
            )
            return {
                "comp_code":   "0888",
                "sbu":         "SP",
                "brand":       "Clarks",
                "brand_code":  "CKS",
                "season":      season,
                "seq":         1,
                "multi_mono":  "Multi",
                "country_code": "",
            }

    log.error("No valid brand file found in candidate dirs.")
    return {}


def _prepare_tmp_dirs() -> dict[str, Path]:
    """Create clean /tmp/stibo_workdir_clarks/input/{type}/ dirs."""
    base       = Path(TMP_WORKDIR)
    input_base = base / "input"

    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs: dict[str, Path] = {
        "order_form_accs": base / "input" / "order_form_accs",
        "order_form_footwear": base / "input" / "order_form_footwear",
        "ean":             base / "input" / "ean",
        "mdd":             base / "input" / "mdd",
        "attributes":      base / "input" / "attributes",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    return dirs


def _download_files(
    bucket:  str,
    found:   dict[str, dict],
    dirs:    dict[str, Path],
    auditor: AuditLogger,
) -> None:
    """Download each classified S3 file to its local sub-folder."""
    for ftype, info in found.items():
        if ftype not in dirs:
            log.warning("  No local dir for type '%s' — skipping download", ftype)
            continue
        local_path = dirs[ftype] / info["filename"]
        try:
            s3.download_file(bucket, info["key"], str(local_path))
            log.info("  Saved → %s", local_path)
            auditor.record_download(ftype, info["filename"], status="ok")
        except Exception as exc:
            log.error("  Download FAILED for %s: %s", info["key"], exc)
            auditor.record_download(
                ftype, info["filename"], status="failed", error=str(exc)
            )
            raise


def _upload_xml_outputs(principal: str) -> list[str]:
    """
    Upload every .xml under TMP_WORKDIR/output/xml/ to:
        s3://{PROCESSED_BUCKET}/processed/stepxml/{principal}/{filename}
    Returns list of uploaded S3 keys.
    """
    xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
    uploaded = []
    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
        uploaded.append(s3_key)
    return uploaded


# ═════════════════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event, context, auditor: AuditLogger = None):
    """
    Main Lambda entry point.

    auditor is injected by router.py in the shared platform;
    if invoked directly (local test / manual trigger), we create our own.
    """
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "clarks",
            routing_decision = "clarks lambda_handler invoked directly",
        )

    # ── 1. Parse event ────────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Guard: only raw/metadata/ keys ────────────────────────
    if not key.startswith("raw/metadata/"):
        log.warning("Key outside raw/metadata/ — ignoring: %s", key)
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}

    # ── 3. Extract principal ──────────────────────────────────────
    principal = _extract_principal(key)
    if not principal:
        log.error("Could not extract principal from key: %s", key)
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    log.info("Principal: %s", principal)

    # ── 3b. Detect which file type triggered this invocation ──────
    triggered_file_type = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_file_type or "unknown/global-fanout")

    etl_module     = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(
        triggered_file_type, REQUIRED_TYPES_DEFAULT
    )

    # ── 4. List + classify files ──────────────────────────────────
    found = _list_principal_files(
        bucket, principal,
        exclude_types={triggered_file_type} if triggered_file_type else set(),
    )

    # Always pin the trigger file as the authoritative source for its type
    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {
            "key":           key,
            "filename":      trigger_filename,
            "last_modified": None,
        }
        log.info(
            "Pinned trigger file → type='%s'  file='%s'",
            triggered_file_type, trigger_filename,
        )

    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    # ── 4b. If triggered by global fan-out, resolve ETL module ───
    if etl_module is None:
        for file_type in ["order_form_accs", "order_form_footwear"]:
            if file_type in found:
                etl_module          = ETL_DISPATCHER.get(file_type)
                triggered_file_type = file_type
                log.info(
                    "Global fan-out: resolved ETL module from present files → '%s'",
                    file_type,
                )
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

    # ── 5. Check mandatory files ──────────────────────────────────
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

    # ── 6. Prepare /tmp/ directories ─────────────────────────────
    dirs = _prepare_tmp_dirs()

    # ── 7. Download files ─────────────────────────────────────────
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

    # ── 7b. Convert .xlsb / .csv → .xlsx ─────────────────────────
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

    # ── 8. Parse pipeline metadata from filename ──────────────────
    meta = _parse_metadata_from_filename(
        dirs, triggered_file_type=triggered_file_type,
    )
    if not meta:
        log.error(
            "Cannot parse metadata from brand filename. "
            "Ensure file is named like: "
            "{CompanyCode}-{SBU}-Clarks-{FileType}-{Season}-{Seq}.xlsx  "
            "or RAW Order Form Clarks Accs {Season}.xlsx"
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
        multi_mono   = meta.get("multi_mono", "Multi"),
        country_code = meta.get("country_code", ""),
    )

    # ── 9. Run ETL — dispatched to the correct module ─────────────
    log.info(
        "Running Clarks ETL → %s (triggered_file_type=%s) ...",
        etl_module.__name__, triggered_file_type,
    )
    etl_module.run(args, auditor=auditor)

    # ── 10. Upload generated XMLs ─────────────────────────────────
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
