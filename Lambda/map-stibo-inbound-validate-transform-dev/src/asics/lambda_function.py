"""
lambda_function.py — Lambda Handler for Asics
==============================================
Function name : map-stibo-inbound-validate-transform-asics-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/asics/*

Asics-specific flow:
    • Input files from the brand:
        - footwear  → Footwear Order Form (e.g. Asics Footwear SS27.xlsx)
    • mdd + attributes are shared/global files stored at bucket root
    • Required types per trigger:
        footwear  → {"footwear", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        footwear  → asics.footwear_main.run()

S3 layout:
    raw/metadata/asics/{input_file}.xlsx             ← Asics uploads here
    raw/metadata/mdd_file.xlsx                        ← shared global MDD
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global brand mapping

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/asics/{filename}.xml
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
# Set LAMBDA_TMP_DIR BEFORE importing asics modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_asics
# (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_asics"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import asics.footwear_main as footwear_etl  # noqa: E402  # Footwear Order Form
import asics.app_acc_main as app_acc_etl    # noqa: E402  # Apparel & Accessories Order Form
import asics.ean_main as ean_etl            # noqa: E402  # EAN Source File

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_asics")

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
    # Asics Footwear Order Form
    # Example: "Asics Footwear SS27.xlsx"
    # Example: "0999-SP-ASICS-Footwear Order Form-SP2027-ID-1.xlsx"
    # Example: "0888-SP-ASICS-OTB AW26 P.RUN MAA (MAA Sport) Inline-Multi-SP2029-ID-1"
    "footwear": [
        "p.run",
        "p run",
        "footwear",
    ],
    # Asics Apparel & Accessories Order Form
    # Example: "0888-SP-ASICS-OTB AW26 APEQ MAA (MAA Sport) Inline-Multi-SP2027-ID-1"
    "apparel_accessories": [
        "apeq",
        "apparel",
        "accessories",
    ],
    # Asics EAN Source File
    # Example: "0999-SP-ASICS-EAN Source-SP2027-ID-1.xlsx"
    # Example: "0888-SP-ASICS-MAA_AW26 EAN CODE UPDATE_272026 (MAA Sport) Inline-Multi-SP2029-ID-1"
    "ean": [
        "ean source",
        "ean_source",
        "ean-source",
        "ean code update",
        "ean code",
        "barcode source",
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
    "footwear": {"footwear", "mdd", "attributes"},
    "apparel_accessories": {"apparel_accessories", "mdd", "attributes"},
    "ean": {"ean"},
}

REQUIRED_TYPES_DEFAULT: set[str] = {"footwear", "mdd", "attributes"}

# Global types fetched from raw/metadata/ root (not inside asics/)
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}

# ── ETL dispatcher ────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "footwear": footwear_etl,
    "apparel_accessories": app_acc_etl,
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
    Returns the principal (e.g. "asics") or None.
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


# Static fallback map: brand name (upper) → 3-letter SAP brand code
BRAND_NAME_TO_CODE: dict[str, str] = {
    "ASICS": "ASC",
    "ASIC":  "ASC",
}


def _resolve_brand_code_from_mdd(brand_name: str, mdd_dir: Path) -> str:
    """
    Look up the brand LOV ID from the MDD 'Brand LOV' sheet.

    MDD Brand LOV layout:
        Col A = LOV ID      (e.g. "ASC")
        Col B = Brand Name  (e.g. "ASICS")

    Searches Col B for a case-insensitive match against brand_name,
    returns the corresponding Col A value as the brand code.
    Falls back to BRAND_NAME_TO_CODE, then brand_name[:3].upper().
    """
    brand_upper = (brand_name or "").strip().upper()

    mdd_files = list(mdd_dir.glob("*.xlsx"))
    if not mdd_files:
        log.warning("[brand_code] No MDD file found in %s — falling back to BRAND_NAME_TO_CODE", mdd_dir)
        return BRAND_NAME_TO_CODE.get(brand_upper, brand_upper[:3])

    try:
        wb = openpyxl.load_workbook(str(mdd_files[0]), read_only=True, data_only=True)
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
            "[brand_code] No match for '%s' in MDD Brand LOV — falling back to BRAND_NAME_TO_CODE",
            brand_name,
        )
    except Exception as exc:
        log.warning("[brand_code] Error reading MDD Brand LOV: %s — falling back to BRAND_NAME_TO_CODE", exc)

    return BRAND_NAME_TO_CODE.get(brand_upper, brand_upper[:3])


def _parse_metadata_from_filename(
    dirs: dict,
    triggered_file_type: str = None,
) -> dict:
    """
    Parse comp_code, sbu, brand, brand_code, season, seq from the
    triggered brand file's filename.

    Expected format (flexible):
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{Season}-{Seq}.xlsx
        Asics Footwear SS27.xlsx   ← also accepted (no dashes)

    Falls back to safe defaults when the filename does not match the strict
    dash-separated pattern.
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

                # ── Dynamic brand name + brand code from parts[2] ──
                # Supports:
                #   "ASICS"         → brand="Asics",  brand_code from MDD / fallback
                #   "ASICS ASC"     → brand="Asics",  brand_code="ASC" (embedded)
                brand      = "Asics"  # default
                brand_code = "ASC"    # default
                raw_brand_token = parts[2].strip()
                brand_token_match = re.match(
                    r"^([A-Za-z][A-Za-z\s]*?)\s+([A-Z]{3})$",
                    raw_brand_token,
                    re.IGNORECASE,
                )
                if brand_token_match:
                    # e.g. "ASICS ASC" → brand="Asics", brand_code="ASC" (from filename)
                    brand      = brand_token_match.group(1).strip().title()
                    brand_code = brand_token_match.group(2).upper()
                    log.info(
                        "  Brand token '%s' → brand=%r  brand_code=%r (embedded in filename)",
                        raw_brand_token, brand, brand_code,
                    )
                else:
                    # No embedded code — read brand name, resolve code from MDD
                    brand      = raw_brand_token.title()
                    mdd_dir    = dirs.get("mdd", Path(TMP_WORKDIR) / "input" / "mdd")
                    brand_code = _resolve_brand_code_from_mdd(brand, mdd_dir)
                    log.info(
                        "  Brand token '%s' → brand=%r  brand_code=%r (resolved from MDD/fallback)",
                        raw_brand_token, brand, brand_code,
                    )

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
                        "comp_code":    comp_code,
                        "sbu":          sbu,
                        "brand":        brand,
                        "brand_code":   brand_code,
                        "season":       season,
                        "seq":          seq,
                        "country_code": country_code,
                    }

            # ── Fallback: extract season from anywhere in filename ────
            season_match = re.search(r"\b([A-Z]{2})(\d{2,4})\b", stem, re.IGNORECASE)
            if season_match:
                season_prefix = season_match.group(1).upper()
                season_year   = season_match.group(2)
                if len(season_year) == 2:
                    season_year = f"20{season_year}"
                season = f"{season_prefix}{season_year}"
                
                log.info(
                    "Parsed (fallback): brand=Asics code=ASC season=%s (defaults applied)",
                    season,
                )
                return {
                    "comp_code":    "0999",
                    "sbu":          "SP",
                    "brand":        "Asics",
                    "brand_code":   "ASC",
                    "season":       season,
                    "seq":          1,
                    "country_code": "",
                }

    # Final fallback
    log.warning("Could not parse metadata from filename — using defaults")
    return {
        "comp_code":    "0999",
        "sbu":          "SP",
        "brand":        "Asics",
        "brand_code":   "ASC",
        "season":       "SS27",
        "seq":          1,
        "country_code": "",
    }


def _prepare_tmp_dirs() -> dict:
    """Create clean /tmp/stibo_workdir_asics/input/{type}/ dirs."""
    base       = Path(TMP_WORKDIR)
    input_base = base / "input"

    # Wipe stale files from previous Lambda container reuse
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    logs_out = base / "output" / "logs"
    logs_out.mkdir(parents=True, exist_ok=True)

    dirs = {
        "footwear":            base / "input" / "footwear",
        "apparel_accessories": base / "input" / "apparel_accessories",
        "ean":                 base / "input" / "ean",
        "mdd":                 base / "input" / "mdd",
        "attributes":          base / "input" / "attributes",
        "xml":                 xml_out,
        "logs":                logs_out,
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    return dirs


def _download_files(bucket: str, found: dict, dirs: dict, auditor: AuditLogger):
    """Download all found files to their corresponding /tmp/ directories."""
    for ftype, meta in found.items():
        key      = meta["key"]
        filename = meta["filename"]
        
        # Determine target directory
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
    Main Lambda entry point for Asics ETL.
    
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
    log.info("ASICS ETL INVOKED")
    log.info("=" * 80)
    log.info("Event: %s", event)

    # Check if direct invocation (for testing)
    _direct_invocation = event.get("direct_invocation", False)

    # ── 1. Parse event ────────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Invalid event — missing bucket or key")
        return {"statusCode": 400, "error": "Invalid event structure"}

    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Extract principal ──────────────────────────────────────
    principal = _extract_principal(key)
    if not principal or principal != "asics":
        log.error("Key does not belong to asics principal: %s", key)
        return {"statusCode": 400, "error": "Not an asics file"}

    # Use passed auditor (from router) or create a new one
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
    meta = _parse_metadata_from_filename(
        dirs, triggered_file_type=triggered_file_type,
    )
    if not meta:
        log.error(
            "Cannot parse metadata from brand filename. "
            "Ensure file is named like: "
            "{CompanyCode}-{SBU}-Asics-{FileType}-{Season}-{Seq}.xlsx  "
            "or Asics Footwear {Season}.xlsx"
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
        "Running Asics ETL → %s (triggered_file_type=%s) ...",
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
