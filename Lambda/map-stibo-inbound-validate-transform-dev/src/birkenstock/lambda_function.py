"""
lambda_function.py — Lambda Handler for Birkenstock
====================================================
Function name : map-stibo-inbound-validate-transform-birkenstock-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/birkenstock/*

Birkenstock-specific flow:
    • linelist file : OFS Mitra (Order File Source) → OFS_FC_Mitra_ss27.run()
    • oor file      : Birkenstock EAN / OOR report  → oor.run()
    • No TDD, no Backlog
    • mdd + attributes are shared/global files (raw/metadata/ root)

File-type dispatch (by triggered S3 key):
    ofs_mitra → birkenstock.OFS_FC_Mitra_ss27.run()
    oor / ean → birkenstock.oor.run()

S3 layout:
    raw/metadata/birkenstock/OFS_FC_Mitra_ss27.xlsx   ← Birkenstock uploads here
    raw/metadata/birkenstock/OOR_EAN_Report.xlsx      ← Birkenstock uploads here
    raw/metadata/mdd_file.xlsx                         ← shared global MDD
    raw/metadata/attributes_list.xlsx                  ← shared global attr list
    raw/metadata/brand_mapping_template.xlsx           ← shared mapping template

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/birkenstock/{filename}.xml
"""

import json
import logging
import os
import shutil
import sys
from argparse import Namespace
from pathlib import Path
from urllib.parse import unquote_plus

import boto3

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_birkenstock"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ───────────────────────────────────────────────────────────────
import birkenstock.OFS_FC_Mitra_ss27 as ofs_mitra_etl
import birkenstock.oor as oor_etl

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_birkenstock")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ.get("RAW_BUCKET",       "map-stibo-inbound-raw-dev")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "map-stibo-inbound-processed-dev")

# ─────────────────────────────────────────────────────────────────────────────
BRAND_CODE   = "BCK"
BRAND_NAME   = "BIRKENSTOCK"
BRAND_FOLDER = "birkenstock"

# ─────────────────────────────────────────────────────────────────────────────
# FILE TYPE DETECTION
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "oor": [
        "oor",
        "oor report",
        "oor_report",
        "oor-report",
        "ean report",
        "ean_report",
        "order form fw",
        "order_form_fw",
    ],
    "linelist": [
        # OFS filename tokens
        "ofs",
        "ofs_fc",
        "ofs_sp",
        "mitra",
        "order file",
        "order_file",
        "fc_mitra",
        # Fallback generic price list keywords
        "price master",
        "price list",
        "linelist",
        "line list",
    ],
    "mdd": [
        "mdd",
        "master_data",
        "master data",
        "master data dictionary",
    ],
    "attributes": [
        "attributes",
        "attributes_list",
        "attribute list",
        "brand mapping",
        "brand_mapping",
        "NEW - Brand mapping files Template",
        "mapping template",
        "brand template",
        "mapping files template",
    ],
}

REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes"}
ROOT_TYPES: set[str]              = {"mdd", "attributes"}

# Mandatory files scoped by the triggered file type
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist": {"linelist", "mdd", "attributes"},
    # oor now also pulls the global MDD (Size Code LOV) — see _gather_files
    "oor":      {"oor", "mdd"},
}

# File types this Lambda will actually process (others are skipped)
PROCESSABLE_TYPES: set[str] = {"linelist", "oor"}

# ─────────────────────────────────────────────────────────────────────────────
# LINELIST VARIANT → ETL MODULE DISPATCH
# Each entry: (tuple-of-tokens-ALL-must-match, etl_module)
# ─────────────────────────────────────────────────────────────────────────────
LINELIST_VARIANTS: list[tuple[tuple[str, ...], object]] = [
    (("ofs",   "mitra"),  ofs_mitra_etl),
    (("ofs",   "fc"),     ofs_mitra_etl),
    (("ofs",),            ofs_mitra_etl),
    (("mitra",),          ofs_mitra_etl),
]

# Sub-directories needed in /tmp workdir
FILE_TYPE_DIRS: list[str] = ["linelist", "oor", "mdd", "attributes", "mapping"]


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _detect_file_type(filename: str) -> str | None:
    name = filename.lower()
    # Order matters: check 'oor' before 'linelist' to avoid overlap
    for ftype in ["oor", "linelist", "mdd", "attributes"]:
        keywords = FILE_TYPE_KEYWORDS.get(ftype, [])
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
    Falls back to default_module if no variant matches.

    Example:
        OFS_FC_Mitra_ss27.xlsx → all tokens ("ofs", "fc") present → ofs_mitra_etl
    """
    name = filename.lower()
    for tokens, module in LINELIST_VARIANTS:
        if all(tok in name for tok in tokens):
            log.info("  Variant matched: %s → %s", tokens, module.__name__)
            return module
    log.info("  No variant matched, using default: %s", default_module.__name__)
    return default_module


def _parse_event(event: dict) -> tuple[str, str]:
    """Parse S3 bucket and key from EventBridge event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object", {}).get("key",   "")
    return bucket, unquote_plus(key)


def _prepare_tmp_dirs() -> Path:
    """(Re-)create the temp working directory tree."""
    base = Path(TMP_WORKDIR)
    if base.exists():
        shutil.rmtree(base)

    input_dir  = base / "input"
    output_dir = base / "output"

    for subdir in FILE_TYPE_DIRS:
        (input_dir / subdir).mkdir(parents=True, exist_ok=True)

    (output_dir / "xml").mkdir(parents=True, exist_ok=True)
    (output_dir / "logs").mkdir(parents=True, exist_ok=True)

    log.info("[Setup] Temp directories created at %s", base)
    return base


def _download_file(bucket: str, key: str, local_path: Path) -> Path:
    """Download a single S3 object to a local path."""
    local_path.parent.mkdir(parents=True, exist_ok=True)
    s3.download_file(bucket, key, str(local_path))
    log.info("[S3] Downloaded: %s → %s", key, local_path.name)
    return local_path


def _upload_output_files(base_dir: Path, bucket: str) -> list[str]:
    """Upload all generated XML files to S3 processed bucket."""
    xml_dir  = base_dir / "output" / "xml"
    uploaded = []

    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{BRAND_FOLDER}/{xml_file.name}"
        s3.upload_file(str(xml_file), bucket, s3_key)
        log.info("[S3] Uploaded: %s → s3://%s/%s", xml_file.name, bucket, s3_key)
        uploaded.append(s3_key)

    return uploaded


def _gather_files(
    bucket: str,
    triggered_key: str,
    triggered_type: str,
    base_dir: Path,
) -> dict[str, Path]:
    """
    Download the triggered file into the correct sub-dir and all required
    global support files.

    - linelist flow needs mdd + attributes (+ brand mapping)
    - oor flow needs mdd only (for the Size Code LOV lookup)

    Returns a dict mapping file-type → local Path.
    """
    files: dict[str, Path] = {}
    input_dir = base_dir / "input"

    # ── 1. Download triggered file into its sub-dir ──────────────────────────
    filename = Path(triggered_key).name
    subdir   = "oor" if triggered_type == "oor" else "linelist"
    local_path = input_dir / subdir / filename
    _download_file(bucket, triggered_key, local_path)
    files[triggered_type] = local_path

    # ── 2. Global files needed per flow ───────────────────────────────────────
    # oor only needs mdd (Size Code LOV); linelist needs mdd + attributes
    include_types = {"mdd"} if triggered_type == "oor" else ROOT_TYPES

    found: dict = {}
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=include_types,
        log=log,
    )

    for ftype, file_info in found.items():
        if not file_info:
            continue
        s3_key = file_info.get("key",      "")
        fname  = file_info.get("filename", "")
        if not s3_key or not fname:
            continue

        # Choose sub-directory
        if ftype == "mdd":
            target_dir = input_dir / "mdd"
        elif ftype == "attributes":
            target_dir = input_dir / "attributes"
        else:
            target_dir = input_dir / "mapping"

        local_file = target_dir / fname
        try:
            _download_file(bucket, s3_key, local_file)
            files[ftype] = local_file

            # Brand mapping doubles as the mapping input
            if "brand mapping" in fname.lower() or "mapping" in fname.lower():
                files["mapping"] = local_file
        except Exception as e:
            log.warning("[Global] Failed to download %s (%s): %s", ftype, fname, e)

    return files


def _validate_required_files(files: dict[str, Path], triggered_type: str) -> list[str]:
    """Return a list of missing required file types, scoped by triggered type."""
    required = REQUIRED_TYPES_BY_TRIGGER.get(triggered_type, REQUIRED_TYPES_DEFAULT)
    missing = []
    for req in required:
        path = files.get(req)
        if not path or not Path(path).exists():
            missing.append(req)
    return missing


# ═════════════════════════════════════════════════════════════════════════════
# MAIN HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event: dict, context, auditor: AuditLogger = None) -> dict:
    """
    Main Lambda handler for Birkenstock file processing (OFS linelist + OOR/EAN).
    """
    log.info("=" * 60)
    log.info("  Birkenstock Lambda Handler")
    log.info("=" * 60)
    log.info("Event: %s", json.dumps(event, default=str)[:500])

    if auditor is None:
        auditor = AuditLogger(brand=BRAND_FOLDER)

    try:
        # ── Parse event ───────────────────────────────────────────────────────
        bucket, key = _parse_event(event)
        if not bucket or not key:
            log.error("Missing bucket or key in event")
            return {"statusCode": 400, "error": "Missing bucket or key"}

        log.info("[Event] Bucket: %s | Key: %s", bucket, key)

        # ── Detect file type ──────────────────────────────────────────────────
        file_type = _detect_triggered_file_type(key)
        if not file_type:
            log.warning("Unknown file type for key: %s", key)
            return {"statusCode": 200, "status": "skipped", "reason": "unknown file type"}

        log.info("[FileType] Detected: %s", file_type)

        # Only process linelist / oor files
        if file_type not in PROCESSABLE_TYPES:
            log.info("Skipping non-processable file type: %s", file_type)
            return {"statusCode": 200, "status": "skipped", "reason": f"{file_type} file"}

        # ── Prepare workspace ─────────────────────────────────────────────────
        base_dir = _prepare_tmp_dirs()

        # ── Gather all required files ─────────────────────────────────────────
        files = _gather_files(bucket, key, file_type, base_dir)
        log.info("[Files] Gathered: %s", list(files.keys()))

        # ── Validate required files ───────────────────────────────────────────
        missing = _validate_required_files(files, file_type)
        if missing:
            msg = f"Missing required files: {missing}"
            log.error("[Validate] %s", msg)
            auditor.set_lambda_status("error", error=msg)
            return {"statusCode": 400, "status": "error", "error": msg}

        filename = Path(key).name

        # ═════════════════════════════════════════════════════════════════════
        # DISPATCH
        # ═════════════════════════════════════════════════════════════════════
        if file_type == "oor":
            # ── OOR / EAN flow ────────────────────────────────────────────────
            log.info("[ETL] Dispatching → birkenstock.oor")
            oor_args = Namespace(
                brand=BRAND_NAME,
                brand_code=BRAND_CODE,
                oor_path=str(files.get("oor", "")),
                season="",
            )
            oor_etl.run(oor_args, auditor)
            result = {"status": "success"}

        else:
            # ── OFS linelist flow ─────────────────────────────────────────────
            etl_module = _pick_linelist_variant(filename, ofs_mitra_etl)
            log.info("[ETL] Using module: %s", etl_module.__name__)

            # NOTE: country_code, comp_code, sbu, season are parsed dynamically
            #       from the filename inside the ETL module. Values below are
            #       safe fallbacks only.
            etl_args = {
                "linelist_path":   str(files.get("linelist",   "")),
                "mdd_path":        str(files.get("mdd",        "")),
                "attributes_path": str(files.get("attributes", "")),
                "mapping_path":    str(files.get("mapping",    "")),
                "country_code": "ID",
                "comp_code":    "0888",
                "sbu":          "SP",
                "season":       "",
            }

            log.info(
                "[ETL Args] linelist=%s | mdd=%s | attrs=%s | mapping=%s",
                Path(etl_args["linelist_path"]).name,
                Path(etl_args["mdd_path"]).name        if etl_args["mdd_path"]        else "—",
                Path(etl_args["attributes_path"]).name if etl_args["attributes_path"] else "—",
                Path(etl_args["mapping_path"]).name    if etl_args["mapping_path"]    else "—",
            )

            result = etl_module.run(etl_args, auditor)

        log.info("[ETL] Result: %s", result)

        if isinstance(result, dict) and result.get("status") != "success":
            error_msg = result.get("message", "ETL run failed")
            log.error("[ETL] Failed: %s", error_msg)
            auditor.set_lambda_status("error", error=error_msg)
            return {"statusCode": 500, "status": "error", "error": error_msg}

        # ── Upload output XML files ───────────────────────────────────────────
        uploaded = _upload_output_files(base_dir, PROCESSED_BUCKET)
        log.info("[Upload] %d file(s) uploaded", len(uploaded))

        # ── Finalise audit ────────────────────────────────────────────────────
        auditor.set_xml_uploads(uploaded)
        auditor.set_lambda_status("success")

        return {
            "statusCode":    200,
            "status":        "success",
            "brand":         BRAND_FOLDER,
            "file_type":     file_type,
            "article_count": result.get("article_count", 0) if isinstance(result, dict) else 0,
            "output_files":  uploaded,
        }

    except Exception as e:
        log.exception("Lambda handler error: %s", e)
        auditor.set_lambda_status("error", error=str(e))
        return {"statusCode": 500, "error": str(e)}


# ═════════════════════════════════════════════════════════════════════════════
# LOCAL TESTING
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # ── Test 1: OFS linelist flow ─────────────────────────────────────────────
    test_event_linelist = {
        "detail": {
            "bucket": {"name": "map-stibo-inbound-raw-dev"},
            "object": {"key": "raw/metadata/birkenstock/OFS_FC_Mitra_ss27.xlsx"},
        }
    }

    # ── Test 2: OOR / EAN flow ────────────────────────────────────────────────
    test_event_oor = {
        "detail": {
            "bucket": {"name": "map-stibo-inbound-raw-dev"},
            "object": {"key": "raw/metadata/birkenstock/OOR_EAN_Report.xlsx"},
        }
    }

    # Toggle which event to run locally
    result = lambda_handler(test_event_linelist, None)
    print(f"Result: {result}")