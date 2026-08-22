"""
lambda_function.py — Lambda Handler for Airwalk (LOT)
====================================================
Function name : map-stibo-inbound-validate-transform-airwalk-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/airwalk/*

Airwalk-specific flow:
    • ONE input file from the brand: orderform
      (e.g. "FW26 LIC FOB ORDERFORM (1).xlsx")
    • mdd is optional (orderform is self-contained — sizes & prices embedded)
    • Required types: {"orderform"}
    • Keyword detection: "orderform", "order form", "fob"

File-type dispatch:
    orderform → airwalk.orderform_main.run()

S3 layout:
    raw/metadata/airwalk/{orderform_file}.xlsx   ← Airwalk uploads here
    raw/metadata/mdd_file.xlsx                 ← shared global MDD (optional)

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/airwalk/{filename}.xml
"""

import json
import logging
import os
import shutil
import sys
import types
from pathlib import Path
from urllib.parse import unquote_plus

import boto3
import openpyxl
import pandas as pd
import pyxlsb

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_airwalk"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules ──────────────────────────────────────────────────────────────
import airwalk.recap_main           as recap_etl            # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_airwalk")

# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

RAW_BUCKET       = os.environ.get("RAW_BUCKET", "")
PROCESSED_BUCKET = os.environ.get("PROCESSED_BUCKET", "")

# ─────────────────────────────────────────────────────────────────────────────
# File type detection — keyword order matters (first match wins)
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # fob_order_info MUST appear before orderform — more-specific keywords first.
    "fob_order_info": [
        "fob order info", "order info update", "fob_order_info", "ean update",
        "fob order",      # portal-renamed files (e.g. "0888-SP-LOTTO-FOB Order…")
    ],
    "pricelist": [
        "price list", "pricelist", "price-list", "price_list",
        "licensee price", "usd licensee",
    ],
    "recap": [
        "recap", "recap sample", "licensed recap", "sample development",
    ],
    "orderform": [
        "orderform", "order form", "order-form",
        "lic fob", "airwalk",
        # "fob order" intentionally omitted — caught by fob_order_info above
    ],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": [
        "brand mapping", "brand_mapping", "NEW - Brand mapping files Template",
        "mapping template", "brand template",
    ],
}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "orderform":      {"orderform"},
    "fob_order_info": {"fob_order_info"},
    "recap":          {"recap", "mdd", "attributes"},
    "pricelist":      {"pricelist"},
}
REQUIRED_TYPES_DEFAULT: set[str] = {"orderform"}
GLOBAL_TYPES: set[str] = {"mdd", "attributes"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}

ETL_DISPATCHER: dict[str, object] = {
    "recap":          recap_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    xlsx_path = xlsb_path.with_suffix(".xlsx")
    wb_out = openpyxl.Workbook(write_only=True)
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
    exclude_types = exclude_types or set()
    # ROOT_TYPES (mdd, attributes) are fetched exclusively from the bucket root;
    # exclude them from the brand-folder and raw/metadata/ scans.
    brand_exclude = exclude_types | ROOT_TYPES
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None or ftype in brand_exclude:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {"key": key, "filename": filename, "last_modified": last_modified}

    for page in paginator.paginate(Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename or not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES or ftype in brand_exclude:
                continue
            if ftype not in found:
                found[ftype] = {"key": key, "filename": filename, "last_modified": obj["LastModified"]}

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


def _airwalk_default_metadata(dirs: dict) -> dict:
    """
    Build default metadata for Lotto files (no strict filename format).
    Extracts season & country_code from filename if present (e.g. 'FW26', 'PH', 'KH').
    Searches recap, orderform, then fob_order_info dirs.
    """
    import re as _re
    season = "FW26"
    country_code = "ID"
    country_map = {
        "INDONESIA": "ID", "ID": "ID",
        "PHILIPPINES": "PH", "PHILIPPINE": "PH", "PH": "PH",
        "THAILAND": "TH", "TH": "TH",
        "SINGAPORE": "SG", "SG": "SG",
        "MALAYSIA": "MY", "MY": "MY",
        "VIETNAM": "VN", "VN": "VN",
        "CAMBODIA": "KH", "CAMBODIAN": "KH", "KH": "KH",
    }
    search_dirs = [dirs.get("recap"), dirs.get("orderform"), dirs.get("fob_order_info")]
    for folder in search_dirs:
        if not folder or not folder.exists():
            continue
        for f in list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm")):
            m = _re.search(r"\b([A-Z]{2}\d{2,4})\b", f.stem, _re.IGNORECASE)
            if m:
                season = m.group(1).upper()
            stem_tokens = f.stem.upper().replace("_", " ").replace("-", " ").split()
            for token in stem_tokens:
                if token in country_map and token != "SP":
                    country_code = country_map[token]
                    break
            if m:
                break
        else:
            continue
        break
    return {
        "comp_code":   "0888",
        "sbu":         "SP",
        "brand":       "airwalk",
        "brand_code":  "AIW",
        "season":      season,
        "seq":         1,
        "multi_mono":  "Multi",
        "country_code": country_code,
        "article_type_from_filename": "License",
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
        "orderform":      base / "input" / "orderform",
        "fob_order_info": base / "input" / "fob_order_info",
        "recap":          base / "input" / "recap",
        "pricelist":      base / "input" / "pricelist",
        "mdd":            base / "input" / "mdd",
        "attributes":     base / "input" / "attributes",
    }
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
            detected_brand="airwalk", routing_decision="airwalk lambda_handler invoked directly",
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
        for ftype in ("recap", "pricelist", "fob_order_info", "orderform"):
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

    # Convert formats
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            _convert_xlsb_to_xlsx(xlsb_file)
        for csv_file in type_dir.glob("*.csv"):
            _convert_csv_to_xlsx(csv_file)

    meta = _airwalk_default_metadata(dirs)
    log.info("[Airwalk] Using metadata: %s", meta)
    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand=meta["brand"], brand_code=meta["brand_code"],
        comp_code=meta["comp_code"], sbu=meta["sbu"],
        season=meta["season"], seq=meta["seq"],
        multi_mono=meta.get("multi_mono", "Multi"),
        country_code=meta.get("country_code", ""),
        article_type_from_filename=meta.get("article_type_from_filename", ""),
    )

    log.info("Running Airwalk ETL → %s ...", etl_module.__name__)
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
