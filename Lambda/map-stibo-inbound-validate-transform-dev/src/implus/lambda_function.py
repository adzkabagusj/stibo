"""
lambda_function.py -- Lambda Handler for implus (implus)
====================================================================
Function name : map-stibo-inbound-validate-transform-implus-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/implus/*

implus-specific flow:
    * ONE input file from the brand: linelist / linesheet / article planning
      OR an EAN Update file (same Excel, different keyword in filename)
    * mdd is optional (linelist is self-contained -- sizes & styles embedded)
    * Required types: {"linelist"} or {"ean_update"}
    * Keyword detection:
        linelist   -- "linelist", "linesheet", "implus", "selection", "article planning"
        ean_update -- "ean update", "ean_update", "ean-update", "barcode update",
                      " ean " (bare EAN token, e.g. "...Source_Balega EAN MAA Sport...")

File-type dispatch:
    linelist   -> implus.linelist_main_source_balega.run()        (default)
    linelist   -> implus.linelist_main_source_harbinger.run()     (filename contains "harbinger")
    ean_update -> implus.ean_update_source_balega.run()

S3 layout:
    raw/metadata/implus/{linelist_file}.xlsx   <- implus uploads here
    raw/metadata/mdd_file.xlsx                   <- shared global MDD (optional)

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/implus/{filename}.xml
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

# -----------------------------------------------------------------------------
TMP_WORKDIR = "/tmp/stibo_workdir_implus"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# -- ETL modules --------------------------------------------------------------
import implus.linelist_main_source_balega as linelist_etl  # noqa: E402  (default sub-brand parser)
import implus.linelist_main_source_harbinger as harbinger_etl  # noqa: E402
import implus.linelist_main_source_sofsole as sofsole_etl  # noqa: E402
import implus.ean_update_source_balega as ean_update_etl  # noqa: E402
import implus.ean_update_source_harbinger as ean_update_harbinger_etl  # noqa: E402
import implus.ean_update_source_sofsole as ean_update_sofsole_etl  # noqa: E402

# -----------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_implus")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ.get("RAW_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# ─────────────────────────────────────────────────────────────────────────────
# File type detection — keyword order matters (first match wins)
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "ean_update": [
        "ean update", "ean_update", "ean-update",
        "barcode update", "barcode_update",
        " ean ", "ean source", "ean_source", "ean-source",
    ],
    "linelist": [
        "linesheet", "linelist", "line_list", "line sheet",
        "implus", "selection consolidation", "article planning",
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template",
    ],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist":   {"linelist"},
    "ean_update": {"ean_update"},
}
REQUIRED_TYPES_DEFAULT: set[str] = {"linelist"}
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}
ROOT_TYPES:   set[str] = {"mdd", "attributes"}

ETL_DISPATCHER: dict[str, object] = {
    "linelist":   linelist_etl,
    "ean_update": ean_update_etl,
}

# Sub-brand source files share the "linelist" file-type but need different
# parsers (different sheet layouts). Picked by filename right before the ETL
# call — keeps the file-type detection above (required-files / missing-files
# checks) generic while routing to the correct parser.
def _select_linelist_module(filename: str):
    name = (filename or "").lower()
    if "harbinger" in name:
        return harbinger_etl
    if "sofsole" in name:
        return sofsole_etl
    return linelist_etl  # default: Balega parser


# EAN Update files follow the same sub-brand-by-filename pattern as linelist
# files (different sheet layout -> different parser).
def _select_ean_update_module(filename: str):
    name = (filename or "").lower()
    if "harbinger" in name:
        return ean_update_harbinger_etl
    if "sofsole" in name:
        return ean_update_sofsole_etl
    return ean_update_etl  # default: Balega parser


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
    """
    Scan S3 for all recognised files:
      Step 1: Brand-specific files  ← raw/metadata/{principal}/
              (mdd + attributes excluded here — they live at bucket root)
      Step 2: Bucket root           ← s3://{bucket}/  (mdd + attributes only)
              via add_root_metadata_files() which always picks the latest file.

    Mirrors the Clarks pattern exactly:
      - Brand folder  → linelist / ean_update
      - Bucket root   → mdd, attributes (shared across all brands)
    """
    exclude_types = exclude_types or set()
    # Exclude ROOT_TYPES from the brand folder scan — they only come from bucket root
    brand_exclude = exclude_types | ROOT_TYPES
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    # ── Step 1: brand-specific folder ────────────────────────────────────────
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsm")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None or ftype in brand_exclude:
                if ftype is None:
                    log.warning("  Unrecognised file — skipping: %s", filename)
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}
                log.info("  Classified  %-18s ← %s  [brand-specific]", ftype, filename)

    # ── Step 2: bucket root — MDD + attributes (shared global files) ─────────
    # add_root_metadata_files scans the true bucket root (no prefix / Delimiter="/")
    # and always overwrites with the latest version — same behaviour as Clarks.
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
    Parse sbu, country_code, season_prefix, season_year from the downloaded
    linelist filename.

    Expected portal format (dash-separated):
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{SeasonCode}{Year}-{Country}-{Seq}
    e.g. 0888-SP-STACCATO IMP-Linesheet MAA Sport Inline-Multi-SM2027-VN-1.xlsx
    """
    defaults = {
        "comp_code":     "0888",
        "sbu":           "FF",
        "brand":         "IMPLUS",
        "brand_code":    "IPL",
        "country_code":  "ID",
        "season_prefix": None,
        "season_year":   None,
        "article_type":  "Inline",
    }

    folder = dirs.get("linelist")
    if not folder or not folder.exists():
        return defaults

    xls_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
    if not xls_files:
        return defaults

    stem = xls_files[0].stem
    log.info("[implus] Parsing metadata from filename: '%s'", stem)
    parts = re.split(r"\s*-\s*", stem)

    if len(parts) < 6:
        log.warning("[implus] Filename has < 6 dash-parts — using defaults: '%s'", stem)
        return defaults

    meta = dict(defaults)

    comp_code = parts[0].strip()
    if comp_code:
        meta["comp_code"] = comp_code

    sbu = parts[1].strip()
    if sbu:
        meta["sbu"] = sbu

    # Brand name from parts[2] (e.g. "IMPLUS" in "0888-SB-IMPLUS-...").
    # brand_code stays hardcoded "IPL" — the ETL will resolve the LOV ID from MDD.
    brand_from_filename = parts[2].strip()
    if brand_from_filename:
        meta["brand"] = brand_from_filename.title()  # e.g. "Implus"
        log.info("[implus] Brand extracted from filename: '%s'", meta["brand"])

    # article_type: scan parts[3..] up to season token for "inline" / "license"
    upper_stem = stem.upper()
    if "INLINE" in upper_stem:
        meta["article_type"] = "Inline"
    elif "LICENSE" in upper_stem or "LICENCE" in upper_stem or "LICENSED" in upper_stem:
        meta["article_type"] = "License"

    # Find season token (e.g. SM2027, FW26, SS2028) — must appear at index ≥ 3
    for i, p in enumerate(parts):
        m = re.match(r"^([A-Z]{2})(20\d{2}|\d{2})$", p.strip(), re.IGNORECASE)
        if m and i >= 3:
            prefix   = m.group(1).upper()
            year_raw = m.group(2)
            year_str = f"20{year_raw}" if len(year_raw) == 2 else year_raw
            meta["season_prefix"] = prefix
            meta["season_year"]   = year_str

            # country = first 2–3 uppercase-letter token after the season token
            for p2 in parts[i + 1:]:
                p2 = p2.strip()
                if re.match(r"^[A-Z]{2,3}$", p2, re.IGNORECASE) and not p2.isdigit():
                    meta["country_code"] = p2.upper()
                    break
            break

    log.info(
        "[implus] Metadata: comp_code=%s  sbu=%s  season=%s%s  country=%s  article_type=%s",
        meta["comp_code"], meta["sbu"],
        meta["season_prefix"] or "?", meta["season_year"] or "?",
        meta["country_code"], meta["article_type"],
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

    linelist_dir = base / "input" / "linelist"
    dirs = {
        "linelist":   linelist_dir,
        "ean_update": linelist_dir,   # EAN Update uses the same Excel as linelist
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
    }
    for d in set(dirs.values()):
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
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type="direct_invocation", original_key="",
            detected_brand="implus", routing_decision="implus lambda_handler invoked directly",
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

    auditor.set_files_classified(found)

    if etl_module is None:
        # Trigger file itself wasn't recognized (e.g. mdd/attributes landed
        # first) — fall back to whichever brand file-type is already present.
        for ftype in ("linelist", "ean_update"):
            if ftype in found:
                etl_module = ETL_DISPATCHER[ftype]
                triggered_file_type = ftype
                required_types = REQUIRED_TYPES_BY_TRIGGER.get(ftype, REQUIRED_TYPES_DEFAULT)
                break

    if etl_module is None:
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "status": "skipped_no_brand_file"}

    missing = required_types - set(found.keys())
    if missing:
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "missing": sorted(missing), "status": "waiting_for_mandatory_files"}

    dirs = _prepare_tmp_dirs()

    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "error": f"Download failed: {exc}"}

    meta = _parse_metadata_from_filename(dirs)
    log.info("[implus] Using metadata: %s", meta)
    auditor.set_metadata_parsed(meta)

    # Resolve the downloaded linelist file path explicitly so run() doesn't
    # fall back to _find_input_files() (which relies on LAMBDA_TMP_DIR and
    # breaks when multiple brand lambdas overwrite that env var at import time).
    linelist_dir = dirs.get("linelist", Path(TMP_WORKDIR) / "input" / "linelist")
    linelist_files = sorted(
        [f for f in linelist_dir.glob("*") if f.suffix.lower() in (".xlsx", ".xlsm")],
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    input_file = str(linelist_files[0]) if linelist_files else None
    log.info("[implus] Input file resolved: %s", input_file)

    args = types.SimpleNamespace(
        brand=meta["brand"],
        brand_code=meta["brand_code"],
        comp_code=meta["comp_code"],
        sbu=meta["sbu"],
        country_code=meta["country_code"],
        season_prefix=meta.get("season_prefix"),
        season_year=meta.get("season_year"),
        article_type=meta.get("article_type", "Inline"),
        input_file=input_file,
    )

    # Both EAN Update and Linelist files share the "sub-brand parser picked
    # off the actual resolved input file" pattern (falls back to the
    # triggering S3 key's filename if nothing was downloaded locally).
    resolved_filename = Path(input_file).name if input_file else trigger_filename
    if triggered_file_type == "ean_update":
        run_module = _select_ean_update_module(resolved_filename)
    else:
        run_module = _select_linelist_module(resolved_filename)

    log.info("Running implus ETL → %s ...", run_module.__name__)
    # Re-assert LAMBDA_TMP_DIR here: all brand lambda_function modules set this
    # env var at import time, so the last imported module (e.g. twoxu) overwrites
    # it. Resetting it right before the ETL call guarantees the XML is written to
    # TMP_WORKDIR/output/xml, which is exactly where _upload_xml_outputs scans.
    os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
    try:
        run_module.run(args, auditor=auditor)
    except Exception as exc:
        log.error("[implus] ETL run failed: %s", exc)
        auditor.set_lambda_status("error", error=str(exc))
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "error": str(exc)}

    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 200, "status": "no_xml_generated"}

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct_invocation: auditor.flush(principal)

    return {
        "statusCode": 200, "principal": principal,
        "triggered_file_type": triggered_file_type,
        "uploaded": uploaded, "count": len(uploaded), "status": "ok",
    }
