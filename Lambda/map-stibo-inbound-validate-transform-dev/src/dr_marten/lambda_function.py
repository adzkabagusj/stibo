"""
lambda_function.py — Lambda Handler for DR Marten
==================================================
Function name : map-stibo-inbound-validate-transform-dr-marten-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/dr-marten/*

DR Marten-specific flow:
    • ONE input file from the brand: Retail Price Master
    • No TDD, no Backlog
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types: {"linelist", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        retail_price_master → dr_marten.retail_price_master_main.run()

S3 layout:
    raw/metadata/dr-marten/{retail_price_master}.xlsx  ← DR Marten uploads here
    raw/metadata/mdd_file.xlsx                         ← shared global MDD
    raw/metadata/attributes_list.xlsx                  ← shared global attr list
    raw/metadata/brand_mapping_template.xlsx           ← shared mapping template

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/dr-marten/{filename}.xml
"""

import json
import logging
import os
import shutil
import sys
from pathlib import Path
from urllib.parse import unquote_plus

import boto3
import openpyxl

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_dr_marten"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import dr_marten.retail_price_master_main as retail_price_master_etl  # noqa: E402
import dr_marten.inline_ean_source as inline_ean_source_etl  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_dr_marten")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET = os.environ.get("RAW_BUCKET", "map-stibo-inbound-raw-dev")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "map-stibo-inbound-processed-dev")

# ─────────────────────────────────────────────────────────────────────────────
BRAND_CODE = "DRM"
BRAND_NAME = "DR MARTEN"
BRAND_FOLDER = "dr-marten"

# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "ean": [
        "ean source", "ean_source", "ean-source",
        "ean", "barcode", "bar code",
    ],
    "linelist": [
        "retail price", "price master", "rrp", "msrp",
        "sea_ss", "b2b suggested", "price list",
        "linelist", "line list",
    ],
    "mdd": ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "attributes", "attributes_list", "attribute list",
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template", "mapping files template",
    ],
}

REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes"}
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}

# Sub-variants
LINELIST_VARIANTS: list[tuple[tuple[str, ...], object]] = [
    (("retail", "price"),      retail_price_master_etl),
    (("price", "master"),      retail_price_master_etl),
    (("b2b", "suggested"),     retail_price_master_etl),
    (("sea_ss",),              retail_price_master_etl),
    (("rrp",),                 retail_price_master_etl),
]

# Directories needed per file-type
FILE_TYPE_DIRS: list[str] = ["linelist", "ean", "mdd", "attributes", "mapping"]


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

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
    Falls back to default_module if no variant matches.
    """
    name = filename.lower()
    for tokens, module in LINELIST_VARIANTS:
        if all(tok in name for tok in tokens):
            log.info("  Variant matched: %s → %s", tokens, module.__name__)
            return module
    return default_module


def _parse_event(event: dict) -> tuple[str, str]:
    """Parse S3 bucket and key from EventBridge event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key = detail.get("object", {}).get("key", "")
    return bucket, unquote_plus(key)


def _prepare_tmp_dirs():
    """Create temporary directories for file processing."""
    base = Path(TMP_WORKDIR)
    if base.exists():
        shutil.rmtree(base)

    input_dir = base / "input"
    output_dir = base / "output"

    for subdir in FILE_TYPE_DIRS:
        (input_dir / subdir).mkdir(parents=True, exist_ok=True)

    (output_dir / "xml").mkdir(parents=True, exist_ok=True)
    (output_dir / "logs").mkdir(parents=True, exist_ok=True)

    log.info("[Setup] Temp directories created at %s", base)
    return base


def _download_file(bucket: str, key: str, local_path: Path) -> Path:
    """Download a file from S3."""
    local_path.parent.mkdir(parents=True, exist_ok=True)
    s3.download_file(bucket, key, str(local_path))
    log.info("[S3] Downloaded: %s → %s", key, local_path.name)
    return local_path


def _upload_output_files(base_dir: Path, bucket: str) -> list[str]:
    """Upload generated XML files to S3."""
    xml_dir = base_dir / "output" / "xml"
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
    Download the triggered file and any required supporting files.
    Returns a dict mapping file types to local paths.
    """
    files: dict[str, Path] = {}
    input_dir = base_dir / "input"

    # Download triggered file
    filename = Path(triggered_key).name
    local_path = input_dir / triggered_type / filename
    _download_file(bucket, triggered_key, local_path)
    files[triggered_type] = local_path

    # Build found dict for add_root_metadata_files
    found: dict = {}
    
    # Get global files (mdd, attributes, mapping) from bucket root
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        log=log,
    )
    
    # Download found global files
    for ftype, file_info in found.items():
        if not file_info:
            continue
        s3_key = file_info.get("key", "")
        fname = file_info.get("filename", "")
        if not s3_key or not fname:
            continue
        
        # Determine target subdirectory
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
            
            # Also set mapping if this is a brand mapping file
            if "brand mapping" in fname.lower() or "mapping" in fname.lower():
                files["mapping"] = local_file
        except Exception as e:
            log.warning("[Global] Failed to download %s: %s", ftype, e)

    return files


# ═════════════════════════════════════════════════════════════════════════════
# MAIN HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event: dict, context, auditor: AuditLogger = None) -> dict:
    """
    Main Lambda handler for DR Marten file processing.
    """
    log.info("=" * 60)
    log.info("  DR Marten Lambda Handler")
    log.info("=" * 60)
    log.info("Event: %s", json.dumps(event, default=str)[:500])

    # Use provided auditor or create a new one
    if auditor is None:
        auditor = AuditLogger(brand=BRAND_FOLDER)

    try:
        bucket, key = _parse_event(event)
        if not bucket or not key:
            log.error("Missing bucket or key in event")
            return {"statusCode": 400, "error": "Missing bucket or key"}

        log.info("[Event] Bucket: %s | Key: %s", bucket, key)

        # Detect file type
        file_type = _detect_triggered_file_type(key)
        if not file_type:
            log.warning("Unknown file type for key: %s", key)
            return {"statusCode": 200, "status": "skipped", "reason": "unknown file type"}

        log.info("[FileType] Detected: %s", file_type)

        # Only process linelist- or ean-type files
        if file_type not in ("linelist", "ean"):
            log.info("Skipping non-processable file type: %s", file_type)
            return {"statusCode": 200, "status": "skipped", "reason": f"{file_type} file"}

        # Prepare temp directories
        base_dir = _prepare_tmp_dirs()

        # Gather files
        files = _gather_files(bucket, key, file_type, base_dir)
        log.info("[Files] Gathered: %s", list(files.keys()))

        filename = Path(key).name

        # ── EAN Source flow ─────────────────────────────────────────────
        if file_type == "ean":
            log.info("[ETL] Using module: %s", inline_ean_source_etl.__name__)
            etl_args = {
                "linelist_path": "",
                "ean_path": str(files.get("ean", "")),
                "mdd_path": str(files.get("mdd", "")),
                "attributes_path": str(files.get("attributes", "")),
                "mapping_path": str(files.get("mapping", "")),
                "brand": BRAND_NAME,
                "brand_code": BRAND_CODE,
            }
            inline_ean_source_etl.run(etl_args, auditor)

            uploaded = _upload_output_files(base_dir, PROCESSED_BUCKET)
            log.info("[Upload] %d files uploaded", len(uploaded))

            auditor.set_xml_uploads(uploaded)
            auditor.set_lambda_status("success")
            return {
                "statusCode": 200,
                "status": "success",
                "brand": BRAND_FOLDER,
                "file_type": file_type,
                "output_files": uploaded,
            }

        # ── Linelist / Retail Price Master flow ─────────────────────────
        # Pick ETL module variant
        etl_module = _pick_linelist_variant(filename, retail_price_master_etl)
        log.info("[ETL] Using module: %s", etl_module.__name__)

        # Build args for ETL module
        etl_args = {
            "linelist_path": str(files.get("linelist", "")),
            "mdd_path": str(files.get("mdd", "")),
            "attributes_path": str(files.get("attributes", "")),
            "mapping_path": str(files.get("mapping", "")),
            "country_code": "ID",  # Default, can be overridden
            "comp_code": "0888",
            "sbu": "SP",
        }

        # Run ETL
        result = etl_module.run(etl_args, auditor)
        log.info("[ETL] Result: %s", result)

        if result.get("status") != "success":
            error_msg = result.get("message", "ETL failed")
            auditor.set_lambda_status("error", error=error_msg)
            return {"statusCode": 500, "status": "error", "error": error_msg}

        # Upload output files
        uploaded = _upload_output_files(base_dir, PROCESSED_BUCKET)
        log.info("[Upload] %d files uploaded", len(uploaded))

        # Record success
        auditor.set_xml_uploads(uploaded)
        auditor.set_lambda_status("success")

        return {
            "statusCode": 200,
            "status": "success",
            "brand": BRAND_FOLDER,
            "file_type": file_type,
            "article_count": result.get("article_count", 0),
            "output_files": uploaded,
        }

    except Exception as e:
        log.exception("Lambda handler error: %s", e)
        auditor.set_lambda_status("error", error=str(e))
        return {"statusCode": 500, "error": str(e)}


# ═════════════════════════════════════════════════════════════════════════════
# LOCAL TESTING
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # For local testing
    test_event = {
        "detail": {
            "bucket": {"name": "map-stibo-inbound-raw-dev"},
            "object": {"key": "raw/metadata/dr-marten/SEA_SS27_B2B_Suggested_Retail_Price.xlsx"},
        }
    }

    result = lambda_handler(test_event, None)
    print(f"Result: {result}")
