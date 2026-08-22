"""
lambda_function.py — Lambda Handler for Pazzion
================================================
Function name : map-stibo-inbound-validate-transform-pazzion-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/pazzion/*

Pazzion-specific flow:
    • ONE primary input file from the brand : order form (Spring Summer Order Form)
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types: {"orderform", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        orderform    → pazzion.order_form.run()

S3 layout:
    raw/metadata/pazzion/{orderform_file}.xlsx           ← Pazzion uploads here
    raw/metadata/mdd_file.xlsx                           ← shared global MDD
    raw/metadata/NEW - Brand mapping files Template.xlsx ← shared global brand mapping

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/pazzion/{filename}.xml
"""

import json
import logging
import os
import re
import shutil
import sys
import types
from pathlib import Path
from urllib.parse import unquote_plus

import boto3

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
# Set LAMBDA_TMP_DIR BEFORE importing pazzion modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_pazzion
# (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_pazzion"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules — Pazzion order form and barcode ────────────────────────────
import pazzion.order_form as orderform_etl     # noqa: E402  # Dynamic mapping from brand file
import pazzion.barcode as barcode_etl          # noqa: E402  # Barcode/EAN processing

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_pazzion")

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
    # Pazzion barcode/EAN file (checked first for priority)
    # Example: "0888-SP-PAZZION (PZZ)-Barcode (MAA Sport) Inline-Multi-SP2029-ID-1.xlsx"
    "barcode": [
        "barcode", "ean", "pos code", "ci1167",
    ],
    
    # Pazzion primary input (Spring Summer Order Form / Fall Winter Order Form)
    # Example: "IDM - Spring Summer'26 Order Form.xlsx"
    "orderform": [
        "order form", "orderform", "spring summer", "fall winter",
        "ss", "fw", "idm", "pazzion",
    ],
    
    # MDD file (shared global)
    "mdd": [
        "mdd", "master data", "lov", "attribute",
    ],
    
    # Brand mapping file (shared global)
    "attributes": [
        "brand mapping", "mapping files template", "template",
        "pazzion(inline)", "inline",
    ],
}


def _detect_file_type(s3_key: str, local_path: Path) -> str | None:
    """
    Determine file type from S3 key or filename.
    Returns: "orderform", "mdd", "attributes", or None if unknown.
    """
    key_lower = s3_key.lower()
    file_lower = local_path.name.lower()
    
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw in key_lower or kw in file_lower:
                return ftype
    
    return None


# ─────────────────────────────────────────────────────────────────────────────
# S3 helpers
# ─────────────────────────────────────────────────────────────────────────────
def _download_s3_file(bucket: str, key: str, dest: Path) -> None:
    """Download a single S3 object to dest."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    log.info("[S3] Downloading  s3://%s/%s  →  %s", bucket, key, dest)
    s3.download_file(bucket, key, str(dest))
    size_kb = dest.stat().st_size // 1024
    log.info("[S3]   ✓ %s  (%dKB)", dest.name, size_kb)


def _upload_s3_file(local_path: Path, bucket: str, key: str) -> None:
    """Upload local file to S3 bucket + key."""
    log.info("[S3] Uploading  %s  →  s3://%s/%s", local_path, bucket, key)
    s3.upload_file(str(local_path), bucket, key)
    log.info("[S3]   ✓ uploaded")


def _parse_s3_event(event: dict) -> list[tuple[str, str]]:
    """
    Parse S3 event and return list of (bucket, key) tuples.
    Supports both EventBridge and direct S3 event formats.
    """
    results = []
    
    # EventBridge format
    if "detail" in event and "bucket" in event["detail"]:
        bucket = event["detail"]["bucket"]["name"]
        key = unquote_plus(event["detail"]["object"]["key"])
        results.append((bucket, key))
    
    # Direct S3 event format
    elif "Records" in event:
        for record in event["Records"]:
            if "s3" in record:
                bucket = record["s3"]["bucket"]["name"]
                key = unquote_plus(record["s3"]["object"]["key"])
                results.append((bucket, key))
    
    return results


# ─────────────────────────────────────────────────────────────────────────────
# Main Lambda handler
# ─────────────────────────────────────────────────────────────────────────────
def lambda_handler(event, context, auditor=None):
    """
    Lambda entry point for Pazzion order form processing.
    """
    log.info("═══════════════════════════════════════════════════════════")
    log.info("   PAZZION LAMBDA HANDLER — DYNAMIC ORDER FORM PROCESSING  ")
    log.info("═══════════════════════════════════════════════════════════")
    log.info("Event: %s", json.dumps(event, indent=2, default=str))
    
    # Parse S3 event
    s3_files = _parse_s3_event(event)
    if not s3_files:
        log.error("No S3 files found in event")
        return {"statusCode": 400, "body": "No S3 files in event"}
    
    bucket, key = s3_files[0]  # Process first file
    log.info("Processing: s3://%s/%s", bucket, key)
    
    # Initialize audit logger if not provided by router
    if auditor is None:
        auditor = AuditLogger(
            s3_key=key,
            lambda_name=context.function_name if context else "local",
            request_id=context.request_id if context else "local-request",
        )
    
    try:
        # Clean workdir
        if Path(TMP_WORKDIR).exists():
            shutil.rmtree(TMP_WORKDIR)
        Path(TMP_WORKDIR).mkdir(parents=True, exist_ok=True)
        
        # Create directory structure
        (Path(TMP_WORKDIR) / "input" / "orderform").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "input" / "barcode").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "input" / "mdd").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "input" / "attributes").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "output" / "orderforms").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "output" / "xml").mkdir(parents=True, exist_ok=True)
        
        # Download triggered file
        filename = Path(key).name
        local_path = Path(TMP_WORKDIR) / "input" / filename
        _download_s3_file(bucket, key, local_path)
        
        # Detect file type
        file_type = _detect_file_type(key, local_path)
        if not file_type:
            log.error("Unknown file type: %s", filename)
            auditor.set_lambda_status("failed", error=f"Unknown file type: {filename}")
            return {"statusCode": 400, "body": f"Unknown file type: {filename}"}
        
        log.info("Detected file type: %s", file_type)
        
        # Move file to appropriate subdirectory
        if file_type == "orderform":
            dest_path = Path(TMP_WORKDIR) / "input" / "orderform" / filename
        elif file_type == "barcode":
            dest_path = Path(TMP_WORKDIR) / "input" / "barcode" / filename
        elif file_type == "mdd":
            dest_path = Path(TMP_WORKDIR) / "input" / "mdd" / filename
        elif file_type == "attributes":
            dest_path = Path(TMP_WORKDIR) / "input" / "attributes" / filename
        else:
            dest_path = local_path
        
        if dest_path != local_path:
            shutil.move(str(local_path), str(dest_path))
            log.info("Moved to: %s", dest_path)
        
        # Build found dict for global files
        found = {}
        add_root_metadata_files(
            s3_client=s3,
            bucket=bucket,
            found=found,
            include_types={"mdd", "attributes"},
            log=log,
        )
        
        # Download global files (MDD + Brand Mapping)
        dirs = {
            "mdd": Path(TMP_WORKDIR) / "input" / "mdd",
            "attributes": Path(TMP_WORKDIR) / "input" / "attributes",
        }
        for ftype, info in found.items():
            if ftype in dirs:
                local_file = dirs[ftype] / info["filename"]
                s3.download_file(bucket, info["key"], str(local_file))
                log.info("Downloaded %s: %s", ftype, info["filename"])
        
        # Verify all required files are present
        mdd_files = list((Path(TMP_WORKDIR) / "input" / "mdd").glob("*.xlsx"))
        attr_files = list((Path(TMP_WORKDIR) / "input" / "attributes").glob("*.xlsx"))
        orderform_files = list((Path(TMP_WORKDIR) / "input" / "orderform").glob("*.xlsx"))
        barcode_files = list((Path(TMP_WORKDIR) / "input" / "barcode").glob("*.xlsx"))
        
        log.info("Files found:")
        log.info("  MDD files: %s", [f.name for f in mdd_files])
        log.info("  Attribute files: %s", [f.name for f in attr_files])
        log.info("  Order form files: %s", [f.name for f in orderform_files])
        log.info("  Barcode files: %s", [f.name for f in barcode_files])
        
        if not mdd_files:
            raise FileNotFoundError("No MDD file found")
        if not attr_files:
            raise FileNotFoundError("No brand mapping file found")
        if file_type == "orderform" and not orderform_files:
            raise FileNotFoundError("No order form file found")
        if file_type == "barcode" and not barcode_files:
            raise FileNotFoundError("No barcode file found")
        
        # Extract metadata from S3 key or filename
        # Expected format: raw/metadata/pazzion/{filename}.xlsx
        # Filename examples: 
        #   "0888-SP-PAZZION PZZ-Order Form MAA Sport Inline-Multi-SS2029-KH-1.xlsx"
        #   "IDM - Spring Summer'26 Order Form.xlsx"
        
        # Default values
        brand = "Pazzion"
        brand_code = "PZZ"
        comp_code = "0888"
        sbu = "SP"
        country_code = "SG"  # Default to Singapore for Pazzion
        season = "SS26"  # Default season
        
        # Try to extract all metadata from filename
        # Pattern: {comp}-{sbu}-{BRAND_NAME} {BRAND_CODE}-Order Form...Multi-{SEASON}-{COUNTRY}...
        # Example: 0888-SP-PAZZION PZZ-Order Form MAA Sport Inline-Multi-SS2029-KH-1.xlsx
        full_match = re.search(
            r'(\d+)-([A-Z]+)-([A-Z\s]+)\s+([A-Z]{3})-Order\s+Form.*?-([A-Z]{2}\d+)-([A-Z]{2})',
            filename,
            re.IGNORECASE
        )
        if full_match:
            comp_code = full_match.group(1)
            sbu = full_match.group(2).upper()
            brand = full_match.group(3).strip().title()
            brand_code = full_match.group(4).upper()
            season = full_match.group(5).upper()
            country_code = full_match.group(6).upper()
            log.info("Extracted all metadata from filename:")
            log.info("  comp_code=%s  sbu=%s  brand=%s (%s)  season=%s  country=%s",
                     comp_code, sbu, brand, brand_code, season, country_code)
        else:
            # Fallback: try partial extraction
            # Try to extract brand name and brand code from filename
            # Pattern: {comp}-{sbu}-{BRAND_NAME} {BRAND_CODE}-Order Form...
            # Example: 0888-SP-PAZZION PZZ-Order Form...
            brand_match = re.search(r'(\d+)-([A-Z]+)-([A-Z\s]+)\s+([A-Z]{3})-Order\s+Form', filename, re.IGNORECASE)
            if brand_match:
                comp_code = brand_match.group(1)
                sbu = brand_match.group(2).upper()
                brand = brand_match.group(3).strip().title()
                brand_code = brand_match.group(4).strip().upper()
                log.info("Extracted brand and codes from filename: comp=%s sbu=%s brand=%s (%s)",
                         comp_code, sbu, brand, brand_code)
            
                # Try to extract season from filename
            filename_lower = filename.lower()
            
            # Extract season code (SS/FW + year)
            season_match = re.search(r"(spring summer|ss|fall winter|fw|autumn winter|aw)\s*['\"]?(\d{2})", filename_lower)
            if season_match:
                season_type = season_match.group(1)
                year = season_match.group(2)
                
                # Map season type to code
                if "spring" in season_type or season_type == "ss":
                    sea_code = "SS"
                elif "fall" in season_type or season_type == "fw":
                    sea_code = "FW"
                elif "autumn" in season_type or season_type == "aw":
                    sea_code = "AW"
                else:
                    sea_code = "SS"  # default
                
                season = f"{sea_code}{year}"
        
        log.info("Extracted metadata:")
        log.info("  Brand: %s (%s)", brand, brand_code)
        log.info("  Company: %s", comp_code)
        log.info("  SBU: %s", sbu)
        log.info("  Country: %s", country_code)
        log.info("  Season: %s", season)
        
        # Build args namespace
        args = types.SimpleNamespace(
            brand=brand,
            brand_code=brand_code,
            comp_code=comp_code,
            sbu=sbu,
            season=season,
            country_code=country_code,
        )
        
        # Dispatch to ETL module based on file type
        if file_type == "orderform":
            log.info("Calling pazzion.order_form.run() ...")
            orderform_etl.run(args, auditor=auditor)
        elif file_type == "barcode":
            log.info("Calling pazzion.barcode.run() ...")
            barcode_etl.run(args, auditor=auditor)
        else:
            log.warning("File type '%s' does not trigger ETL (metadata only)", file_type)
            auditor.set_lambda_status("skipped")
            return {"statusCode": 200, "body": f"File type '{file_type}' skipped (metadata only)"}
        
        # Upload XML outputs to S3
        xml_dir = Path(TMP_WORKDIR) / "output" / "xml"
        xml_files = list(xml_dir.glob("*.xml"))
        
        if not xml_files:
            log.warning("No XML files generated")
            auditor.set_lambda_status("completed")
            return {"statusCode": 200, "body": "No XML files generated"}
        
        for xml_file in xml_files:
            s3_key = f"processed/stepxml/pazzion/{xml_file.name}"
            _upload_s3_file(xml_file, PROCESSED_BUCKET, s3_key)
        
        log.info("✓ Processing complete — %d XML file(s) uploaded", len(xml_files))
        auditor.set_lambda_status("completed")
        
        return {
            "statusCode": 200,
            "body": json.dumps({
                "message": "Success",
                "input_file": filename,
                "output_files": [f.name for f in xml_files],
            }),
        }
    
    except Exception as e:
        log.error("Processing failed: %s", e, exc_info=True)
        auditor.set_lambda_status("error", error=str(e))
        
        return {
            "statusCode": 500,
            "body": json.dumps({
                "error": str(e),
                "input_file": filename if 'filename' in locals() else "unknown",
            }),
        }


# ─────────────────────────────────────────────────────────────────────────────
# Local testing
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # For local testing, simulate an S3 event
    import argparse
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--bucket", default="map-stibo-inbound-raw-dev")
    parser.add_argument("--key", default="raw/metadata/pazzion/test_orderform.xlsx")
    
    args = parser.parse_args()
    
    # Simulate event
    test_event = {
        "detail": {
            "bucket": {"name": args.bucket},
            "object": {"key": args.key},
        }
    }
    
    # Mock context
    test_context = types.SimpleNamespace(
        function_name="map-stibo-inbound-validate-transform-pazzion-dev",
        request_id="local-test-request-id",
    )
    
    # Set environment variables for local testing
    os.environ["RAW_BUCKET"] = args.bucket
    os.environ["PROCESSED_BUCKET"] = "map-stibo-inbound-processed-dev"
    
    result = lambda_handler(test_event, test_context)
    print("\n" + "="*60)
    print("RESULT:", json.dumps(result, indent=2))
    print("="*60)
