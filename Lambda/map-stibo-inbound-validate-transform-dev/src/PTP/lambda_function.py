"""
lambda_function.py — Lambda Handler for PTP
================================================
Function name : map-stibo-inbound-validate-transform-ptp-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/ptp/*

PTP-specific flow:
    • TWO distinct input files from the brand, each its own trigger
      (mirrors k_swiss's orderform / fob_order_info split — see
      k_swiss/lambda_function.py):
        - orderform   : Asia Order Form USD FOB
                        (e.g. "V2 ASIA ORDER FORM USD FOB 20260207.xlsx")
                        → Generic-article XML (FOB, hierarchies, pricing…)
        - ean_source  : PTP_EAN Source
                        (e.g. "...PTP_EAN Source MAA Sport Inline-...xlsx")
                        → Variant-level barcode-only XML (DC_Barcode)
      These are column-for-column near-duplicates of each other (the EAN
      source file is the order form minus its SIZE SHORT/SIZE LONG
      columns), but each upload triggers ONLY its own ETL — an EAN-source
      upload must NOT also regenerate a full duplicate order-form XML.
      That coupling used to exist here (both ETLs ran off every
      "orderform"-classified upload, because "ptp" was a keyword for
      "orderform" and matched the EAN file's name too) and produced a
      bogus full Generic-article XML under the EAN file's name with no
      barcode data in it. Fixed by giving ean_source its own, more
      specific keyword match (checked first) and its own dispatch branch.
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types:
        orderform    → {"orderform", "mdd", "attributes"}
        ean_source   → {"ean_source", "mdd", "attributes"}

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        orderform    → PTP.order_form_usd_fob.run()
        ean_source   → PTP.ean_source.run()

S3 layout:
    raw/metadata/ptp/{orderform_file}.xlsx                 ← order form upload
    raw/metadata/ptp/{ean_source_file}.xlsx                ← EAN source upload
    raw/metadata/mdd_file.xlsx                             ← shared global MDD
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global brand mapping

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/ptp/{filename}.xml

NOTE on metadata extraction:
    Unlike other brands' order forms, the PTP input file
    ("V2 ASIA ORDER FORM USD FOB <date>.xlsx") has no comp_code / sbu /
    season / country segments embedded in its filename — it's a single
    combined "Asia" order form. The regex-based extraction below mirrors
    the convention used by other brands (in case a differently-named file
    is later uploaded), but for the current filename it will fall through
    to the PTP-specific defaults. FLAG: confirm with the team how
    comp_code/sbu/season/country should really be supplied per upload
    (e.g. separate uploads per country, or a fixed default per run).
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
# Set LAMBDA_TMP_DIR BEFORE importing PTP modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_ptp
# (isolated from other brands).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_ptp"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules — PTP order form (USD FOB) + EAN/barcode source ────────────
# Each is its own distinct upload/trigger — see FILE_TYPE_KEYWORDS below.
import PTP.order_form_usd_fob as orderform_etl     # noqa: E402  # Dynamic mapping from brand file
import PTP.ean_source as ean_source_etl            # noqa: E402  # Barcode-only XML (own upload)

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
# force=True is required in Lambda: the runtime pre-attaches its own handler
# to the root logger (default level WARNING) before this module runs, so a
# plain basicConfig() call is a silent no-op and INFO-level logs from this
# file AND every module it imports (MDDLoader, BrandMappingLoader, RNALoader,
# PTPOrderFormLoader, ean_source.py) never reach CloudWatch — only WARNING+
# gets through. Confirmed via a real invocation where every "[MDD]" /
# "[BrandMapping]" / "[RNA]" info log was missing from CloudWatch even
# though the run succeeded.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
    force=True,
)
log = logging.getLogger("lambda_ptp")

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
    # "ean_source" MUST appear before "orderform" — more-specific keywords
    # first (same convention as k_swiss's "fob_order_info" vs "orderform").
    # The EAN source filename ("...PTP_EAN Source MAA Sport Inline-...")
    # also contains "PTP", so if the generic "orderform" match ran first
    # it would misclassify this upload as the order form — which is
    # exactly the bug this split fixes (see module docstring).
    #
    # PTP EAN/barcode source
    # Example: "0888-SP-PTP-PTP_EAN Source MAA Sport Inline-Multi-SM28-ID-1.xlsx"
    "ean_source": [
        "ean source", "ean_source", "ean-source", "ptp_ean", "ptp ean",
    ],

    # PTP primary input (Asia Order Form USD FOB)
    # Example: "V2 ASIA ORDER FORM USD FOB 20260207.xlsx"
    # NOTE: bare "ptp" is deliberately NOT a keyword here — it matched
    # every PTP filename including the EAN source upload above.
    "orderform": [
        "order form", "orderform", "usd fob", "asia order form",
    ],

    # MDD file (shared global)
    "mdd": [
        "mdd", "master data", "lov", "attribute",
    ],

    # Brand mapping file (shared global)
    "attributes": [
        "brand mapping", "mapping files template", "template",
    ],
}


def _detect_file_type(s3_key: str, local_path: Path) -> str | None:
    """
    Determine file type from S3 key or filename.
    Returns: "ean_source", "orderform", "mdd", "attributes", or None if unknown.
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
    Lambda entry point for PTP order form (USD FOB) processing.
    """
    log.info("═══════════════════════════════════════════════════════════")
    log.info("   PTP LAMBDA HANDLER — DYNAMIC ORDER FORM PROCESSING      ")
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
        (Path(TMP_WORKDIR) / "input" / "ean_source").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "input" / "mdd").mkdir(parents=True, exist_ok=True)
        (Path(TMP_WORKDIR) / "input" / "attributes").mkdir(parents=True, exist_ok=True)
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
        elif file_type == "ean_source":
            dest_path = Path(TMP_WORKDIR) / "input" / "ean_source" / filename
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
        ean_source_files = list((Path(TMP_WORKDIR) / "input" / "ean_source").glob("*.xlsx"))

        log.info("Files found:")
        log.info("  MDD files: %s", [f.name for f in mdd_files])
        log.info("  Attribute files: %s", [f.name for f in attr_files])
        log.info("  Order form files: %s", [f.name for f in orderform_files])
        log.info("  EAN source files: %s", [f.name for f in ean_source_files])

        if not mdd_files:
            raise FileNotFoundError("No MDD file found")
        if not attr_files:
            raise FileNotFoundError("No brand mapping file found")
        if file_type == "orderform" and not orderform_files:
            raise FileNotFoundError("No order form file found")
        if file_type == "ean_source" and not ean_source_files:
            raise FileNotFoundError("No EAN source file found")

        # Extract metadata from S3 key or filename
        # Expected format: raw/metadata/ptp/{filename}.xlsx
        # Filename example (bare convention — no embedded metadata):
        #   "V2 ASIA ORDER FORM USD FOB 20260207.xlsx"
        # Filename example (real PTP upload convention — metadata as a
        # hyphen-delimited prefix/suffix around the free-text order form
        # name):
        #   "0888-SP-PTP-V2 ASIA ORDER FORM USD FOB 20260207 MAA Sport
        #    Inline-Multi-SM28-ID-1.xlsx"
        #    → comp_code=0888 sbu=SP brand_code=PTP season=SM28 country=ID
        # Filename example (other brands' convention, supported defensively
        # in case a differently-shaped file is uploaded):
        #   "0888-SP-PTP PTP-Order Form MAA Sport Inline-Multi-SS2029-KH-1.xlsx"

        # PTP-specific defaults (order form ships to Asia via PT MAP Aktif
        # Adiperkasa / Indonesia by default — see "Source Mapping related
        # RNA" sheet). Used only when no metadata can be parsed from the
        # filename. FLAG: confirm real per-upload convention with the team.
        brand = "PTP"
        brand_code = "PTP"
        comp_code = "0888"
        sbu = "SP"
        country_code = "ID"  # Default to Indonesia for PTP
        season = "SS26"  # Default season

        # 1) Real PTP convention: "{comp}-{sbu}-{brand}-<free text>-{season}-{country}-{seq}.xlsx"
        #    e.g. "0888-SP-PTP-V2 ASIA ORDER FORM USD FOB 20260207 MAA Sport
        #    Inline-Multi-SM28-ID-1.xlsx". Matched as a prefix + suffix pair
        #    (rather than one full-string regex) so the free-text order-form
        #    description in the middle — which varies and contains no fixed
        #    "-Order Form" marker — doesn't have to be matched literally.
        prefix_match = re.match(r'^(\d+)-([A-Z]+)-([A-Z]+)-', filename, re.IGNORECASE)
        suffix_match = re.search(r'-([A-Z]{2}\d{2,4})-([A-Z]{2})-\d+\.xlsx$', filename, re.IGNORECASE)

        # 2) Other brands' convention (defensive fallback): brand name +
        #    3-letter brand code immediately followed by literal "-Order Form".
        full_match = None if (prefix_match and suffix_match) else re.search(
            r'(\d+)-([A-Z]+)-([A-Z\s]+)\s+([A-Z]{3})-Order\s+Form.*?-([A-Z]{2}\d+)-([A-Z]{2})',
            filename,
            re.IGNORECASE
        )

        if prefix_match and suffix_match:
            comp_code = prefix_match.group(1)
            sbu = prefix_match.group(2).upper()
            brand = brand_code = prefix_match.group(3).upper()
            season = suffix_match.group(1).upper()
            country_code = suffix_match.group(2).upper()
            log.info("Extracted metadata from filename (PTP convention):")
            log.info("  comp_code=%s  sbu=%s  brand=%s (%s)  season=%s  country=%s",
                     comp_code, sbu, brand, brand_code, season, country_code)
        elif full_match:
            comp_code = full_match.group(1)
            sbu = full_match.group(2).upper()
            brand = full_match.group(3).strip().upper()
            brand_code = full_match.group(4).upper()
            season = full_match.group(5).upper()
            country_code = full_match.group(6).upper()
            log.info("Extracted metadata from filename (other-brand convention):")
            log.info("  comp_code=%s  sbu=%s  brand=%s (%s)  season=%s  country=%s",
                     comp_code, sbu, brand, brand_code, season, country_code)
        else:
            log.info("Filename does not match any known upload convention — "
                     "using PTP defaults: comp_code=%s sbu=%s brand=%s (%s) "
                     "season=%s country=%s",
                     comp_code, sbu, brand, brand_code, season, country_code)

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

        # Dispatch to ETL module based on file type — each trigger runs
        # ONLY its own ETL (see module docstring: an ean_source upload
        # must not also regenerate a full duplicate order-form XML).
        if file_type == "orderform":
            log.info("Calling PTP.order_form_usd_fob.run() ...")
            orderform_etl.run(args, auditor=auditor)
        elif file_type == "ean_source":
            log.info("Calling PTP.ean_source.run() (barcode-only XML) ...")
            ean_source_etl.run(args, auditor=auditor)
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
            s3_key = f"processed/stepxml/ptp/{xml_file.name}"
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
    parser.add_argument("--key", default="raw/metadata/ptp/test_orderform.xlsx")

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
        function_name="map-stibo-inbound-validate-transform-ptp-dev",
        request_id="local-test-request-id",
    )

    # Set environment variables for local testing
    os.environ["RAW_BUCKET"] = args.bucket
    os.environ["PROCESSED_BUCKET"] = "map-stibo-inbound-processed-dev"

    result = lambda_handler(test_event, test_context)
    print("\n" + "="*60)
    print("RESULT:", json.dumps(result, indent=2))
    print("="*60)
