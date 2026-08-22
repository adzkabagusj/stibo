"""
lambda_function.py — Lambda Handler for New Era
================================================
Function name : map-stibo-inbound-validate-transform-new_era-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/new_era/*

New Era-specific flow:
    • THREE possible input files from the brand — three separate physical
      workbooks, disambiguated by the "ACC" / "HW" / "APP" token in the
      filename (see _detect_orderform_subtype() below; the token is a
      bare word, not parenthesized, e.g. "FW26 ACC_Revised...",
      "FW26 HW_Revised...", "FW26 APP 1..." all occur):
        - Order Form Acc (Bags and Other Accessories), e.g.
          "0888-SP-NEW ERA-Copy of Order form - APAC FW26 ACC_Revised
           Size MAA Sport Inline-Multi-SP2029-ID-1.xlsx"
        - Order Form HW (Headwear), e.g.
          "0888-SP-NEW ERA-Copy of Order form - APAC FW26 HW_Revised
           Pricing MAA Sport Inline-Multi-SM28-ID-1.xlsx"
        - Order Form App (Apparel), e.g.
          "0888-SP-NEW ERA-Copy of Order form - APAC FW26 APP 1
           MAA Sport Inline-Multi-SP2028-ID-1.xlsx"
    • "attributes" (brand mapping workbook) is MANDATORY — the
      "New Era(Inline)" tab of that workbook drives every per-row
      attribute all three ETLs send, and its "Source Mapping related RNA"
      tab resolves AT_BrandType/Status/Group (see
      new_era/order_form_acc_main.py, new_era/order_form_app_main.py,
      new_era/order_form_HW_main.py). "mdd" is MANDATORY too — it
      resolves AT_Brand and the LOV IDs for BrandType/Status/Group.
    • Required types: {"orderform_acc"|"orderform_app"|"orderform_hw", "attributes", "mdd"}

File-type dispatch:
    orderform_acc → new_era.order_form_acc_main.run()
    orderform_app → new_era.order_form_app_main.run()
    orderform_hw  → new_era.order_form_HW_main.run()
    delivery_schedule_ean → new_era.delivery_schedule_ean_main.run()

S3 layout:
    raw/metadata/new_era/{order_form_file}.xlsx         ← New Era uploads here
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global mapping (has "New Era(Inline)" tab)

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/new_era/{filename}.xml
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
TMP_WORKDIR = "/tmp/stibo_workdir_new_era"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import new_era.order_form_acc_main as orderform_acc_etl  # noqa: E402
import new_era.order_form_app_main as orderform_app_etl  # noqa: E402
import new_era.order_form_HW_main as orderform_hw_etl  # noqa: E402
import new_era.delivery_schedule_ean_main as delivery_schedule_ean_etl  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_new_era")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ.get("RAW_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# ─────────────────────────────────────────────────────────────────────────────
# File type detection — keyword order matters (first match wins)
# ─────────────────────────────────────────────────────────────────────────────
# "order form"/"orderform" is handled specially in _detect_file_type() below
# (not listed here) — all three New Era Order Form physical files (Acc/App/
# HW) share that generic keyword, so the split is resolved by
# _detect_orderform_subtype()'s ACC/APP/HW token search instead.
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "delivery_schedule_ean": [
        "delivery schedule ean", "delivery schedule", "delivery_schedule",
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template",
    ],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "orderform_acc": {"orderform_acc", "attributes", "mdd"},
    "orderform_app": {"orderform_app", "attributes", "mdd"},
    "orderform_hw":  {"orderform_hw", "attributes", "mdd"},
    "delivery_schedule_ean": {"delivery_schedule_ean", "mdd"},
}
REQUIRED_TYPES_DEFAULT: set[str] = {"orderform_acc", "attributes", "mdd"}
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}
ROOT_TYPES:   set[str] = {"mdd", "attributes"}
ORDERFORM_TYPES: tuple[str, ...] = (
    "orderform_acc", "orderform_app", "orderform_hw", "delivery_schedule_ean",
)

ETL_DISPATCHER: dict[str, object] = {
    "orderform_acc": orderform_acc_etl,
    "orderform_app": orderform_app_etl,
    "orderform_hw":  orderform_hw_etl,
    "delivery_schedule_ean": delivery_schedule_ean_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

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


# Matches a bare "acc"/"app"/"hw" token bounded by anything other than a
# letter or digit — so "(ACC)", " ACC_", "-HW-", " APP " all match, but
# "Accessories" and "PATHWAY" don't. Underscore/parenthesis/dash/space are
# all treated as separators here (unlike \b, which treats "_" as a word
# character and would miss "FW26 HW_Revised...", the real HW filename's
# shape). None of ACC/APP/HW are substrings of each other, so match order
# doesn't matter for correctness.
_ACC_TOKEN_RE = re.compile(r"(?<![a-z0-9])acc(?![a-z0-9])")
_APP_TOKEN_RE = re.compile(r"(?<![a-z0-9])app(?![a-z0-9])")
_HW_TOKEN_RE  = re.compile(r"(?<![a-z0-9])hw(?![a-z0-9])")


def _detect_orderform_subtype(name_lower: str) -> str | None:
    """
    Disambiguate a New Era "Order Form" filename between the three
    physical workbooks — Bags and Other Accessories (orderform_acc),
    Apparel (orderform_app), and Headwear (orderform_hw) — all of which
    share the generic "order form" keyword.

    Returns None (unrecognized) if none of the three tokens are found —
    deliberately NOT defaulting to any one subtype: an earlier version of
    this function defaulted unmatched filenames to "orderform_acc", which
    silently misrouted an Apparel upload into the Acc ETL (wrong sheet
    name expected -> crash) instead of surfacing it as an unrecognized
    file. Callers should treat None the same as any other unrecognized
    file type.
    """
    if _ACC_TOKEN_RE.search(name_lower):
        return "orderform_acc"
    if _APP_TOKEN_RE.search(name_lower):
        return "orderform_app"
    if _HW_TOKEN_RE.search(name_lower):
        return "orderform_hw"
    log.warning(
        "Could not find an 'ACC'/'APP'/'HW' token in Order Form filename '%s' — "
        "unrecognized subtype", name_lower,
    )
    return None


def _detect_file_type(filename: str) -> str | None:
    name = filename.lower()
    # Priority: delivery schedule files must route to delivery_schedule_ean.
    if "delivery schedule" in name or "delivery_schedule" in name:
        return "delivery_schedule_ean"
    if "order form" in name or "orderform" in name:
        return _detect_orderform_subtype(name)
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw.lower() in name for kw in keywords):
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
    brand_exclude = exclude_types | ROOT_TYPES
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsm")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None or ftype in brand_exclude:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}
                log.info("  Classified  %-14s <- %s  [brand-specific]", ftype, filename)

    # ── Scan bucket root exclusively for MDD + Brand Mapping ────
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        exclude_types=exclude_types,
        log=log,
    )

    return found


def _prepare_tmp_dirs() -> dict[str, Path]:
    base = Path(TMP_WORKDIR)
    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs = {
        "orderform_acc": base / "input" / "orderform_acc",
        "orderform_app": base / "input" / "orderform_app",
        "orderform_hw":  base / "input" / "orderform_hw",
        "delivery_schedule_ean": base / "input" / "delivery_schedule_ean",
        "mdd":           base / "input" / "mdd",
        "attributes":    base / "input" / "attributes",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


def _download_files(bucket, found, dirs, auditor):
    for ftype, info in found.items():
        if ftype not in dirs:
            continue
        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-14s <- s3://%s/%s", ftype, bucket, info["key"])
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
        log.info("  Uploading %s -> s3://%s/%s", xml_file.name, PROCESSED_BUCKET, s3_key)
        s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
        uploaded.append(s3_key)
    return uploaded


# ═════════════════════════════════════════════════════════════════════════════
# LAMBDA HANDLER
# ═════════════════════════════════════════════════════════════════════════════

def lambda_handler(event, context, auditor: AuditLogger = None):
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type="direct_invocation", original_key="",
            detected_brand="new_era", routing_decision="new_era lambda_handler invoked directly",
        )

    log.info("Event received: %s", json.dumps(event))

    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    if not key.startswith("raw/metadata/"):
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True}

    principal = _extract_principal(key)
    if not principal:
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation: auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    triggered_file_type = _detect_triggered_file_type(key)
    etl_module     = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(triggered_file_type, REQUIRED_TYPES_DEFAULT)

    found = _list_principal_files(bucket, principal,
                                  exclude_types={triggered_file_type} if triggered_file_type else set())

    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {"key": key, "filename": trigger_filename, "last_modified": None}
        log.info("Pinned trigger file -> type='%s' file='%s'", triggered_file_type, trigger_filename)

    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    if etl_module is None:
        # Trigger was mdd/attributes (global fan-out) rather than an order
        # form itself — check whether either order form subtype is already
        # present for this principal.
        for ftype in ORDERFORM_TYPES:
            if ftype in found:
                etl_module = ETL_DISPATCHER[ftype]
                triggered_file_type = ftype
                required_types = REQUIRED_TYPES_BY_TRIGGER.get(ftype, REQUIRED_TYPES_DEFAULT)
                break

    if etl_module is None:
        log.info("No brand file present yet for '%s' during global fan-out — skipping.", principal)
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "principal": principal, "status": "skipped_no_brand_file"}

    missing = required_types - set(found.keys())
    if missing:
        log.info("Mandatory files not yet present for '%s' — waiting for: %s", principal, sorted(missing))
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation: auditor.flush(principal)
        return {
            "statusCode": 200, "principal": principal,
            "present": sorted(found.keys()), "missing": sorted(missing),
            "status": "waiting_for_mandatory_files",
        }

    dirs = _prepare_tmp_dirs()

    log.info("Downloading %d files ...", len(found))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": f"Download failed: {exc}"}

    meta = etl_module._parse_metadata_from_filename(dirs, triggered_file_type=triggered_file_type)
    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand        = meta["brand"],
        brand_code   = meta["brand_code"],
        comp_code    = meta["comp_code"],
        sbu          = meta["sbu"],
        season       = meta["season"],
        country_code = meta.get("country_code", ""),
    )

    log.info("Running New Era ETL -> %s (triggered_file_type=%s) ...", etl_module.__name__, triggered_file_type)
    # Re-assert LAMBDA_TMP_DIR: brand lambda modules set this env var at
    # import time, so the last-imported brand module can overwrite it in a
    # shared router process. Reset right before the ETL call.
    os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
    try:
        etl_module.run(args, auditor=auditor)
    except Exception as exc:
        log.error("[NewEra] ETL run failed: %s", exc)
        auditor.set_lambda_status("error", error=str(exc))
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": str(exc)}

    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("ETL completed but no XML files were produced.")
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct_invocation: auditor.flush(principal)

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
