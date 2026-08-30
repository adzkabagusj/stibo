"""
lambda_function.py — Lambda Handler for Anta
================================================
Function name : map-stibo-inbound-validate-transform-anta-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/anta/*

Anta-specific flow:
    • ONE input file from the brand: "Stibo Line List Template - Anta.xlsx"
    • "attributes" (brand mapping workbook) is MANDATORY — the ANTA tab of
      that workbook drives every per-row attribute this ETL sends, and its
      "Source Mapping related RNA" tab resolves AT_BrandType/Status/Group
      (see anta/linelist_main.py). "mdd" is MANDATORY too — it resolves
      AT_Brand and the LOV IDs for BrandType/Status/Group.
    • Required types: {"linelist", "attributes", "mdd"}
    • EAN Source file: "...ANTA_EAN Source (MAA Sport) Inline-Multi-...xlsx"
      needs "mdd" too — its Size Code LOV (Col F/G) resolves the Size LOV ID
      used in KEY_InboundVariant, same mechanism as
      crocs/shipment_confirmation_main.py (see anta/ean_main.py).
      Required types: {"ean", "mdd"}

File-type dispatch:
    linelist → anta.linelist_main.run()
    ean      → anta.ean_main.run()

S3 layout:
    raw/metadata/anta/{linelist_file}.xlsx             ← Anta uploads here
    raw/metadata/anta/{ean_file}.xlsx                  ← Anta uploads here
    raw/metadata/NEW - Brand mapping files Template.xlsx  ← shared global mapping (has "ANTA" tab)

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/anta/{filename}.xml
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
TMP_WORKDIR = "/tmp/stibo_workdir_anta_1"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import anta_1.linelist_main as linelist_etl  # noqa: E402
import anta_1.ean_main as ean_etl            # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_anta_1")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ.get("RAW_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# ─────────────────────────────────────────────────────────────────────────────
# File type detection — keyword order matters (first match wins)
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # Anta EAN Source File
    # Example: "0888-SP-ANTA-ANTA_EAN Source (MAA Sport) Inline-Multi-SP2027-ID-1"
    "ean": [
        "ean source",
        "ean_source",
        "ean-source",
        "ean code update",
        "ean code",
        "barcode source",
    ],
    "linelist": [
        "stibo line list template - anta", "stibo line list template",
        "line list", "linelist", "line_list", "line sheet", "anta",
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template",
    ],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist": {"linelist", "attributes", "mdd"},
    "ean": {"ean", "mdd"},
}
REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "attributes", "mdd"}
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}
ROOT_TYPES:   set[str] = {"mdd", "attributes"}

ETL_DISPATCHER: dict[str, object] = {
    "linelist": linelist_etl,
    "ean": ean_etl,
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


def _detect_file_type(filename: str) -> str | None:
    name = filename.lower()
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
                log.info("  Classified  %-12s <- %s  [brand-specific]", ftype, filename)

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


def _parse_metadata_from_filename(dirs: dict) -> dict:
    """
    Parse comp_code, sbu, brand, season from the triggering file's name
    (linelist or EAN — whichever one was actually downloaded this run).

    Expected portal format (dash-separated), same convention as other brands:
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{SeasonCode}{Year}-{Country}-{Seq}

    Falls back to safe defaults when the filename does not match — this is
    common for a plain "Stibo Line List Template - Anta.xlsx" upload.
    """
    defaults = {
        "comp_code":    "0888",
        "sbu":          "SP",
        "brand":        "Anta_1",
        "brand_code":   "ATA",
        "season":       "SS27",
        "country_code": "",
    }

    # Prefer the linelist file when both are present; otherwise fall back to
    # whichever brand file was actually downloaded this run (e.g. "ean").
    xls_files: list[Path] = []
    for ftype in ("linelist", "ean"):
        folder = dirs.get(ftype)
        if folder and folder.exists():
            xls_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
            if xls_files:
                break
    if not xls_files:
        return defaults

    stem = xls_files[0].stem
    log.info("[Anta] Parsing metadata from filename: '%s'", stem)
    parts = re.split(r"\s*-\s*", stem)

    meta = dict(defaults)

    if len(parts) >= 6:
        comp_code = parts[0].strip()
        sbu       = parts[1].strip()
        if comp_code:
            meta["comp_code"] = comp_code
        if sbu:
            meta["sbu"] = sbu.upper()

        # ── Brand name + brand code from parts[2] (same convention as Nike) ──
        # Supports both:
        #   "ANTA"      → brand="Anta", brand_code="ATA"  (default fallback)
        #   "ANTA ATA"  → brand="Anta", brand_code="ATA"  (embedded in filename)
        raw_brand_token = parts[2].strip()
        brand_token_match = re.match(
            r"^([A-Za-z][A-Za-z\s]*?)\s+([A-Z]{3})$",
            raw_brand_token,
            re.IGNORECASE,
        )
        if brand_token_match:
            meta["brand"]      = brand_token_match.group(1).strip().title()
            meta["brand_code"] = brand_token_match.group(2).upper()
            log.info(
                "[Anta] Brand token '%s' -> brand=%r brand_code=%r (embedded in filename)",
                raw_brand_token, meta["brand"], meta["brand_code"],
            )
        elif raw_brand_token:
            meta["brand"] = raw_brand_token.title()
            log.info(
                "[Anta] Brand token '%s' -> brand=%r brand_code=%r (default fallback)",
                raw_brand_token, meta["brand"], meta["brand_code"],
            )

        for i, p in enumerate(parts):
            m = re.match(r"^([A-Z]{2})(20\d{2}|\d{2})$", p.strip(), re.IGNORECASE)
            if m:
                prefix   = m.group(1).upper()
                year_raw = m.group(2)
                meta["season"] = f"{prefix}{year_raw[-2:]}"
                for p2 in parts[i + 1:]:
                    p2 = p2.strip()
                    if re.match(r"^[A-Z]{2,3}$", p2, re.IGNORECASE) and not p2.isdigit():
                        meta["country_code"] = p2.upper()
                        break
                break
    else:
        season_match = re.search(r"([A-Z]{2}\d{2,4})", stem, re.IGNORECASE)
        if season_match:
            meta["season"] = season_match.group(1).upper()
        log.info("[Anta] Filename has < 6 dash-parts — using defaults with parsed season: %s", meta["season"])

    log.info(
        "[Anta] Metadata: comp_code=%s sbu=%s season=%s country=%s",
        meta["comp_code"], meta["sbu"], meta["season"], meta["country_code"],
    )
    return meta


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
        "linelist":   base / "input" / "linelist",
        "ean":        base / "input" / "ean",
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs


def _download_files(bucket, found, dirs, auditor):
    for ftype, info in found.items():
        if ftype not in dirs:
            continue
        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-12s <- s3://%s/%s", ftype, bucket, info["key"])
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
            detected_brand="anta_1", routing_decision="anta lambda_handler invoked directly",
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
        if "linelist" in found:
            etl_module = ETL_DISPATCHER["linelist"]
            triggered_file_type = "linelist"
            required_types = REQUIRED_TYPES_BY_TRIGGER.get("linelist", REQUIRED_TYPES_DEFAULT)

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

    meta = _parse_metadata_from_filename(dirs)
    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand        = meta["brand"],
        brand_code   = meta["brand_code"],
        comp_code    = meta["comp_code"],
        sbu          = meta["sbu"],
        season       = meta["season"],
        country_code = meta.get("country_code", ""),
    )

    log.info("Running Anta ETL -> %s (triggered_file_type=%s) ...", etl_module.__name__, triggered_file_type)
    # Re-assert LAMBDA_TMP_DIR: brand lambda modules set this env var at
    # import time, so the last-imported brand module can overwrite it in a
    # shared router process. Reset right before the ETL call.
    os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
    try:
        etl_module.run(args, auditor=auditor)
    except Exception as exc:
        log.error("[Anta] ETL run failed: %s", exc)
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
