"""
lambda_function.py — Lambda Handler for ON Running (ONR)
========================================================
Function name : map-stibo-inbound-validate-transform-onr-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/onr/*

ON Running-specific flow:
    • ONE input file from the brand: linesheet (Planet Sports)
    • No TDD, no Backlog
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types: {"linelist", "mdd", "attributes"}
    • Linesheet filename keyword detection recognises:
        "linesheet", "linelist", "line list", "planet"

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        linelist         → onr.linelist_main.run()
        packaging_list_taf / taf_ss26
                         → onr.taf_ss26_main.run()
                           Trigger  : SAP Article Creation
                           Endpoint : EAN Update
                           (File type formerly known as TAF; renamed by
                            MAA to "Packaging List TAF")


S3 layout:
    raw/metadata/onr/{linesheet_file}.xlsx   ← ON Running uploads here
    raw/metadata/mdd_file.xlsx               ← shared global MDD
    raw/metadata/attributes_list.xlsx        ← shared global attr list

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/onr/{filename}.xml
"""

import json
import logging
import os
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
TMP_WORKDIR = "/tmp/stibo_workdir_onr"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import onr.linelist_planet_sport_main as linelist_planet_sport_etl  # noqa: E402
import onr.linelist_footwear_taf_main as linelist_footwear_taf_etl   # noqa: E402
import onr.packaging_list_planet_sport_main as packaging_list_planet_sport_etl  # noqa: E402  (Renamed by MAA from planet_sport)
import onr.taf_ss26_main as taf_ss26_etl   # noqa: E402  (SS26 layout — EAN + US sizes)

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_onr")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# ─────────────────────────────────────────────────────────────────────────────
# Brand lookup
# ─────────────────────────────────────────────────────────────────────────────
BRAND_NAME_TO_CODE: dict[str, str] = {
    "ADIDAS":      "ADI",
    "NIKE":        "NIK",
    "NEW BALANCE": "NEW",
    "SMIGGLE":     "SMI",
    "ALDO":        "ALD",
    "CROCS":       "CRO",
    "LOTTO":       "LOT",
    "BIRKENSTOCK": "BIR",
    "ON RUNNING":  "ONR",
}

# ─────────────────────────────────────────────────────────────────────────────
# File type detection
# ─────────────────────────────────────────────────────────────────────────────
# IMPORTANT: keyword order matters — first match wins.
#   "taf_ss26" must precede "taf" (SS26 files also contain "taf").
#   "taf"      must precede "linelist" (TAF files contain "linesheet").
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "taf_ss26": [
        # New MAA name variants — "Packaging List TAF"
        "packaging list taf", "packaging_list_taf", "packaginglisttaf",
        "packaging list - taf", "packing list taf",
        "ss26 (taf)", "taf - on ss26", "taf-on-ss26", "ss26 taf", "taf ss26",
        "ss26.xlsx", "ss26.xlsb", "ss26.csv",
    ],
    "taf": [
        # New MAA name variants — "Packaging List TAF" (non-SS26)
        "packaging list taf", "packaging_list_taf", "packaginglisttaf",
        "packaging list - taf", "packing list taf",
        "(taf)", " taf ", "_taf_", "-taf-", " taf-", "-taf ", " taf.", "_taf.",
        "taf.xlsx", "taf.xlsb", "taf.csv",
        "footwear taf", "line list footwear taf", "linelist footwear taf", "taf-multi", "taf_multi",
        "athlete's foot", "athletes foot", "the athletes foot", "the athlete's foot",
    ],
    "linelist_planet_sport": [
        # Line List variants for Planet Sport (original naming)
        "line list footwear planet sport", "linelist footwear planet sport",
        "line list planet sport", "linelist planet sport",
        "linesheet footwear planet sport", "linesheet planet sport",
    ],
    "planet_sport": [
        # New MAA name variants — "Packaging List Planet Sport" / "Packing List Planet Sport"
        "packaging list planet sport", "packaging_list_planet_sport", "packaginglistplanetsport",
        "packing list planet sport", "packing_list_planet_sport", "packinglistplanetsport",
        "packaging list - planet sport", "packing list - planet sport",
        "planet sport", "planet_sport", "planetsport",
    ],
    "linelist": [
        "linesheet", "line sheet", "linelist", "line list",
        "planet", "catalogue", "catalog",
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": ["attributes", "attributes_list", "attribute list"],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist":              {"linelist", "mdd", "attributes"},
    "linelist_planet_sport": {"linelist_planet_sport", "mdd", "attributes"},
    "planet_sport":          {"planet_sport", "mdd", "attributes"},
    "taf":                   {"taf", "mdd", "attributes"},  # FW26 pipeline (requires attributes for RNA lookup)
    "taf_ss26":              {"taf_ss26", "mdd"},           # SS26 pipeline (EAN + US sizes)
}
REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes"}
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}

ETL_DISPATCHER: dict[str, object] = {
    "linelist":              linelist_planet_sport_etl,
    "linelist_planet_sport": linelist_planet_sport_etl,  # Line List Footwear Planet Sport
    "planet_sport":          packaging_list_planet_sport_etl,  # Renamed by MAA: "Packaging List Planet Sport"
    "taf":                   linelist_footwear_taf_etl,
    "taf_ss26":              taf_ss26_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  (same as smiggle — shared utility functions)
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    xlsx_path  = xlsb_path.with_suffix(".xlsx")
    wb_out     = openpyxl.Workbook(write_only=True)
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
                            ws_out.append([_clean_cell(item.v) if item is not None else None for item in row])
                        except Exception as e:
                            log.warning("  Skipping bad row in '%s': %s", safe_name, e)
                sheets_written += 1
            except Exception as e:
                log.warning("  Error reading sheet '%s': %s", sheet_name, e)
    if sheets_written == 0:
        raise ValueError(f"No sheets converted from {xlsb_path.name}")
    wb_out.save(str(xlsx_path))
    log.info("  Converted %s → %s (%d sheets)", xlsb_path.name, xlsx_path.name, sheets_written)
    xlsb_path.unlink()
    return xlsx_path


def _convert_csv_to_xlsx(csv_path: Path) -> Path:
    xlsx_path = csv_path.with_suffix(".xlsx")
    df = pd.read_csv(str(csv_path), dtype=str, keep_default_na=False)
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
    try:
        if hasattr(ts, 'timestamp'):
            return int(ts.timestamp())
        elif isinstance(ts, (int, float)):
            return int(float(ts))
        elif isinstance(ts, str):
            from datetime import datetime as _dt
            d = _dt.fromisoformat(ts.replace('Z', '+00:00'))
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
    exclude_types = exclude_types or set()
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None or ftype in exclude_types:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}

    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        exclude_types=exclude_types,
        principal="onr",
        log=log,
    )

    for page in paginator.paginate(Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES or ftype in exclude_types:
                continue
            if ftype not in found:
                found[ftype] = {"key": key, "filename": filename, "last_modified": obj["LastModified"]}

    return found


def _parse_metadata_from_linelist_filename(dirs: dict, triggered_file_type: str = None) -> dict:
    import re

    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs = [dirs.get(triggered_file_type), dirs.get("linelist")]
    else:
        candidate_dirs = [dirs.get("linelist")]

    seen = set()
    candidate_dirs = [d for d in candidate_dirs if d is not None and str(d) not in seen and not seen.add(str(d))]

    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        for xlsx_file in folder.glob("*.xlsx"):
            stem = xlsx_file.stem
            log.info("Attempting metadata parse from filename: '%s'", stem)

            parts = re.split(r"\s*-\s*", stem)
            if len(parts) < 7:
                log.warning("  Filename '%s' has only %d parts (need ≥7) — skipping", stem, len(parts))
                continue

            comp_code = parts[0].strip()
            sbu       = parts[1].strip()
            brand     = parts[2].strip().title()

            if not comp_code:
                continue

            season = None
            season_idx = None
            for i, p in enumerate(parts):
                if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
                    season = p.strip().upper()
                    season_idx = i
                    break

            if not season or season_idx is None or season_idx < 3:
                continue

            trailing = parts[season_idx + 1:]
            seq = 1
            for p in reversed(trailing):
                if p.strip().isdigit():
                    seq = int(p.strip())
                    break

            country_code = ""
            for p in trailing:
                p = p.strip()
                if re.match(r'^[A-Z]{2,3}$', p, re.IGNORECASE) and not p.isdigit():
                    country_code = p.upper()
                    break

            multi_mono = parts[season_idx - 1].strip() if season_idx >= 1 else "Multi"
            brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())

            article_type_from_filename = ""
            file_type_tokens = parts[3:season_idx]
            for token in file_type_tokens:
                t = token.strip().lower()
                if "inline" in t:
                    article_type_from_filename = "Inline"
                    break
                if "license" in t or "licence" in t:
                    article_type_from_filename = "License"
                    break

            log.info(
                "Parsed: comp_code=%s sbu=%s brand=%s brand_code=%s season=%s seq=%s",
                comp_code, sbu, brand, brand_code, season, seq,
            )
            return {
                "comp_code": comp_code, "sbu": sbu, "brand": brand,
                "brand_code": brand_code, "season": season, "seq": seq,
                "multi_mono": multi_mono, "country_code": country_code,
                "article_type_from_filename": article_type_from_filename,
            }

    log.error("No valid brand file found in candidate dirs.")
    return {}


def _taf_default_metadata(dirs: dict) -> dict:
    """
    Fallback metadata for TAF linesheets (which don't follow the strict
    'comp-sbu-brand-...-season-...-seq' dash filename format).

    Extracts season from filename if present (e.g. 'FW26', 'SS25'),
    otherwise uses ON Running defaults.
    """
    import re as _re
    season = "SS26"
    folder = dirs.get("taf_ss26") or dirs.get("taf") or dirs.get("linelist")
    if folder and folder.exists():
        for xlsx_file in folder.glob("*.xlsx"):
            m = _re.search(r"\b([A-Z]{2}\d{2,4})\b", xlsx_file.stem, _re.IGNORECASE)
            if m:
                season = m.group(1).upper()
                break
    return {
        "comp_code":   "0888",
        "sbu":         "SP",
        "brand":       "On Running",
        "brand_code":  "ONR",
        "season":      season,
        "seq":         1,
        "multi_mono":  "Multi",
        "country_code": "ID",
        "article_type_from_filename": "Inline",
    }


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
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
    }
    # Planet Sport, TAF and TAF_SS26 linesheets all share the same input folder
    # as 'linelist' so their respective ETL modules find them via LINELIST_DIR.
    dirs["linelist_planet_sport"] = dirs["linelist"]
    dirs["planet_sport"]          = dirs["linelist"]
    dirs["taf"]                   = dirs["linelist"]
    dirs["taf_ss26"]              = dirs["linelist"]
    for d in dirs.values():
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
            detected_brand="onr", routing_decision="onr lambda_handler invoked directly",
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
    etl_module    = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(triggered_file_type, REQUIRED_TYPES_DEFAULT)

    found = _list_principal_files(bucket, principal,
                                  exclude_types={triggered_file_type} if triggered_file_type else set())

    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {"key": key, "filename": trigger_filename, "last_modified": None}

    auditor.set_files_classified(found)

    if etl_module is None:
        for ftype in ("taf_ss26", "taf", "linelist_planet_sport", "planet_sport", "linelist"):
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
        # Only download files that are actually required for this pipeline.
        # This prevents stale brand files of other types (e.g. a TAF file
        # sitting in S3 when a Planet Sport file was the trigger) from being
        # downloaded into the shared linelist folder and corrupting the run.
        found_to_download = {k: v for k, v in found.items() if k in required_types}
        _download_files(bucket, found_to_download, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation: auditor.flush(principal)
        return {"statusCode": 500, "error": f"Download failed: {exc}"}

    # Convert formats
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            _convert_xlsb_to_xlsx(xlsb_file)
        for csv_file in type_dir.glob("*.csv"):
            _convert_csv_to_xlsx(csv_file)

    meta = _parse_metadata_from_linelist_filename(dirs, triggered_file_type=triggered_file_type)
    if not meta:
        # TAF linesheets don't follow the strict 7-part dash filename format.
        # Fall back to defaults + season extraction from filename.
        if triggered_file_type in ("taf", "taf_ss26"):
            meta = _taf_default_metadata(dirs)
            log.info("[Packaging List TAF / %s] Using fallback metadata: %s", triggered_file_type.upper(), meta)
        else:
            auditor.set_lambda_status("error", error="Unparseable metadata in linelist filename")
            if _direct_invocation: auditor.flush(principal)
            return {"statusCode": 400, "error": "Unparseable metadata in linelist filename"}

    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand=meta["brand"], brand_code=meta["brand_code"],
        comp_code=meta["comp_code"], sbu=meta["sbu"],
        season=meta["season"], seq=meta["seq"],
        multi_mono=meta.get("multi_mono", "Multi"),
        country_code=meta.get("country_code", ""),
        article_type_from_filename=meta.get("article_type_from_filename", ""),
    )

    log.info("Running ON Running ETL → %s ...", etl_module.__name__)
    etl_module.run(args, auditor=auditor)

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
