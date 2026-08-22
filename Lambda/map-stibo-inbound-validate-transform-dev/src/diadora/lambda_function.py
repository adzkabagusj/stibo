# """
# lambda_function.py — Unified Lambda Handler (Diadora)
# ======================================================
# Function name : map-stibo-inbound-validate-transform-diadora-dev
 
# Handles Diadora file types from a single S3 folder:
#     raw/metadata/diadora/
 
# File-type → ETL dispatch:
#     ┌──────────────────────────────────────────────────────────────┐
#     │ Filename contains           → ETL module                     │
#     ├──────────────────────────────────────────────────────────────┤
#     │ "Pricelist" / "Price List"  → diadora.main  (inline)        │
#     └──────────────────────────────────────────────────────────────┘
 
#     Detection order:
#         1. pricelist  — inline pricelist (the only Diadora ETL for now)
#         2. mdd / attributes / naming — shared global support files
 
# Global files (mdd, naming) fetched from raw/metadata/ root.
 
# ────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ────────────────────────────────────────────────────────────────────
#     RAW_BUCKET        Source bucket       e.g. map-stibo-inbound-raw-dev
#     PROCESSED_BUCKET  Destination bucket  e.g. map-stibo-inbound-processed-dev
#     COMP_CODE         Optional fallback company code  (default: "0888")
# """
 
# import json
# import logging
# import os
# import re
# import shutil
# import sys
# import types
# import openpyxl
# import pyxlsb
# from pathlib import Path
# from urllib.parse import unquote_plus
 
# import boto3
# import pandas as pd
 
# from audit_logger import AuditLogger
# from metadata_s3 import add_root_metadata_files
 
# # ─────────────────────────────────────────────────────────────────
# # Set LAMBDA_TMP_DIR BEFORE importing ETL modules
# # ─────────────────────────────────────────────────────────────────
# TMP_WORKDIR = "/tmp/stibo_workdir"
# os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
 
# import diadora.main as etl_pricelist   # Diadora inline pricelist
# import diadora.recap as etl_recap      # Diadora licensed recap
 
# # ─────────────────────────────────────────────────────────────────
# # Logging
# # ─────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s  %(levelname)-8s  %(message)s",
#     handlers=[logging.StreamHandler(sys.stdout)],
# )
# log = logging.getLogger("lambda_diadora")
 
# # ─────────────────────────────────────────────────────────────────
# # AWS client
# # ─────────────────────────────────────────────────────────────────
# s3 = boto3.client("s3")
 
# # ─────────────────────────────────────────────────────────────────
# # Environment variables
# # ─────────────────────────────────────────────────────────────────
# RAW_BUCKET        = os.environ["RAW_BUCKET"]
# PROCESSED_BUCKET  = os.environ["PROCESSED_BUCKET"]
# # ─────────────────────────────────────────────────────────────────
# # Brand name → brand code
# # ─────────────────────────────────────────────────────────────────
# BRAND_NAME_TO_CODE: dict[str, str] = {
#     "DIADORA":     "DIA",
#     "NEW BALANCE": "NEW",
#     "ADIDAS":      "ADI",
#     "NIKE":        "NIK",
#     "LOTTO":       "LOT",
# }
 
# # ─────────────────────────────────────────────────────────────────
# # File type detection
# # ─────────────────────────────────────────────────────────────────
# FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
#     "pricelist":  ["pricelist", "price list", "price_list"],
#     "recap":      ["recap", "recap sample"],
#     "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
#     "attributes": ["attributes", "attributes_list", "attribute list"],
#     "naming":     ["naming", "naming_convention", "naming convention"],
# }
 
# # ─────────────────────────────────────────────────────────────────
# # ETL dispatcher  —  file_type key → ETL module
# # ─────────────────────────────────────────────────────────────────
# ETL_DISPATCHER: dict[str, object] = {
#     "pricelist": etl_pricelist,
#     "recap":     etl_recap,
# }
 
# # ─────────────────────────────────────────────────────────────────
# # Mandatory file rules
# # ─────────────────────────────────────────────────────────────────
# REQUIRED_NON_BRAND_TYPES: set[str] = {"mdd", "attributes", "naming"}
 
# REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
#     "pricelist": {"pricelist", "mdd", "attributes", "naming"},
#     "recap":     {"recap", "mdd", "attributes"},
# }
 
# # Global types fetched from raw/metadata/ root (not brand subfolder)
# GLOBAL_TYPES: set[str] = {"mdd", "naming"}
# ROOT_TYPES: set[str] = {"mdd", "attributes"}
 
 
# # =============================================================================
# # HELPERS
# # =============================================================================
 
# def _clean_cell(v):
#     """Strip illegal XML/Excel characters from string cell values."""
#     if not isinstance(v, str):
#         return v
#     return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")
 
 
# def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
#     """Convert .xlsb → .xlsx streaming row-by-row."""
#     xlsx_path      = xlsb_path.with_suffix(".xlsx")
#     wb_out         = openpyxl.Workbook(write_only=True)
#     sheets_written = 0
 
#     with pyxlsb.open_workbook(str(xlsb_path)) as wb:
#         for sheet_name in wb.sheets:
#             safe_name = "".join(
#                 c for c in sheet_name
#                 if 0x20 <= ord(c) <= 0xFFFD and c not in r"\/?*[]:"
#             ).strip()[:31] or f"Sheet{sheets_written + 1}"
 
#             try:
#                 ws_out = wb_out.create_sheet(title=safe_name)
#             except Exception as e:
#                 log.warning("  Skipping sheet '%s' (bad name): %s", sheet_name, e)
#                 continue
 
#             try:
#                 with wb.get_sheet(sheet_name) as sheet:
#                     for row in sheet.rows():
#                         try:
#                             ws_out.append([
#                                 _clean_cell(item.v) if item is not None else None
#                                 for item in row
#                             ])
#                         except Exception as e:
#                             log.warning("  Skipping bad row in '%s': %s", safe_name, e)
#                 sheets_written += 1
#             except Exception as e:
#                 log.warning("  Error reading sheet '%s': %s", sheet_name, e)
 
#     if sheets_written == 0:
#         raise ValueError(f"No sheets converted from {xlsb_path.name}")
 
#     wb_out.save(str(xlsx_path))
#     log.info("  Converted %s → %s (%d sheets)", xlsb_path.name, xlsx_path.name, sheets_written)
#     xlsb_path.unlink()
#     return xlsx_path
 
 
# def _convert_csv_to_xlsx(csv_path: Path) -> Path:
#     """Convert a .csv to .xlsx so openpyxl can read it."""
#     xlsx_path = csv_path.with_suffix(".xlsx")
#     try:
#         df = pd.read_csv(str(csv_path), dtype=str, keep_default_na=False)
#     except Exception as e:
#         raise ValueError(f"Cannot read CSV {csv_path.name}: {e}")
 
#     wb = openpyxl.Workbook(write_only=True)
#     ws = wb.create_sheet()
#     ws.append(list(df.columns))
#     for _, row in df.iterrows():
#         ws.append([_clean_cell(v) for v in row.tolist()])
 
#     wb.save(str(xlsx_path))
#     log.info("  Converted %s → %s", csv_path.name, xlsx_path.name)
#     csv_path.unlink()
#     return xlsx_path
 
 
# def _detect_file_type(filename: str) -> str | None:
#     """
#     Return the file type key for a given filename, or None if unrecognised.
#     Case-insensitive keyword substring match.
#     """
#     name_lower = filename.lower()
#     for ftype, keywords in FILE_TYPE_KEYWORDS.items():
#         if any(kw.lower() in name_lower for kw in keywords):
#             return ftype
#     return None
 
 
# def _detect_triggered_file_type(key: str) -> str | None:
#     """
#     Derive the file type from the S3 key that triggered this invocation.
#     Returns None for global fan-out placeholder keys (starting with '.__').
#     """
#     filename = Path(key).name
#     if filename.startswith(".__"):
#         return None
#     return _detect_file_type(filename)
 
 
# def _parse_event(event: dict) -> tuple[str, str]:
#     """Extract bucket and key from an EventBridge S3 Object Created event."""
#     detail = event.get("detail", {})
#     bucket = detail.get("bucket", {}).get("name", "")
#     key    = detail.get("object",  {}).get("key",  "")
#     return bucket, str(key)
 
 
# def _extract_principal(key: str) -> str | None:
#     """
#     Derive the principal (brand folder) from the S3 key.
#     Key format: raw/metadata/{principal}/{filename}
#     """
#     parts = key.strip("/").split("/")
#     if len(parts) >= 4:
#         return parts[2]
#     log.warning("Key does not match raw/metadata/{principal}/{file}: %s", key)
#     return None
 
 
# def _list_principal_files(bucket: str, principal: str) -> dict[str, dict]:
#     """
#     Scan S3 for files belonging to this principal.
#     Brand-specific files: raw/metadata/{principal}/
#     Global shared files (mdd, naming): raw/metadata/ (root level only)
#     Brand-specific version always wins over global.
#     """
#     prefix    = f"raw/metadata/{principal}/"
#     paginator = s3.get_paginator("list_objects_v2")
#     found: dict[str, dict] = {}
 
#     # ── Brand-specific folder ─────────────────────────────────
#     for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
#         for obj in page.get("Contents", []):
#             key      = obj["Key"]
#             filename = Path(key).name
#             if not filename:
#                 continue
#             if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
#                 continue
 
#             ftype = _detect_file_type(filename)
#             if ftype is None:
#                 log.warning("  Unrecognised file (skipping): %s", filename)
#                 continue
 
#             last_modified = obj["LastModified"]
#             if ftype not in found or last_modified > found[ftype]["last_modified"]:
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified,
#                 }
#                 log.info("  Classified  %-22s ← %s  [brand-specific]", ftype, filename)
 
#     # ── Global folder (mdd + naming only) ────────────────────
#     add_root_metadata_files(
#         s3_client=s3,
#         bucket=bucket,
#         found=found,
#         include_types=ROOT_TYPES,
#         log=log,
#     )

#     for page in paginator.paginate(
#         Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"
#     ):
#         for obj in page.get("Contents", []):
#             key      = obj["Key"]
#             filename = Path(key).name
#             if not filename:
#                 continue
#             if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
#                 continue
 
#             ftype = _detect_file_type(filename)
#             if ftype not in GLOBAL_TYPES:
#                 continue
 
#             last_modified = obj["LastModified"]
#             if ftype not in found:
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified,
#                 }
#                 log.info("  Classified  %-22s ← %s  [global]", ftype, filename)
#             else:
#                 log.info(
#                     "  Skipping global %s — brand-specific already found: %s",
#                     ftype, found[ftype]["filename"],
#                 )
 
#     return found
 
 
# def _check_mandatory_files(
#     found: dict[str, dict], triggered_ftype: str | None
# ) -> list[str]:
#     """Returns list of missing mandatory types."""
#     if triggered_ftype and triggered_ftype in REQUIRED_TYPES_BY_TRIGGER:
#         required = REQUIRED_TYPES_BY_TRIGGER[triggered_ftype]
#         return sorted(required - set(found.keys()))
 
#     # Generic fallback
#     missing = list(REQUIRED_NON_BRAND_TYPES - set(found.keys()))
#     has_brand_file = any(ftype in found for ftype in ETL_DISPATCHER)
#     if not has_brand_file:
#         missing.append("pricelist")
#     return sorted(missing)
 
 
# def _prepare_tmp_dirs() -> dict[str, Path]:
#     """Recreate /tmp/stibo_workdir/input/ from scratch each invocation."""
#     base = Path(TMP_WORKDIR)
 
#     input_base = base / "input"
#     if input_base.exists():
#         shutil.rmtree(str(input_base))
 
#     xml_out = base / "output" / "xml"
#     if xml_out.exists():
#         shutil.rmtree(str(xml_out))
#     xml_out.mkdir(parents=True, exist_ok=True)
 
#     dirs: dict[str, Path] = {
#         "pricelist":  base / "input" / "pricelist",
#         "recap":      base / "input" / "recap",
#         "mdd":        base / "input" / "mdd",
#         "attributes": base / "input" / "attributes",
#         "naming":     base / "input" / "naming",
#     }
#     for d in dirs.values():
#         d.mkdir(parents=True, exist_ok=True)
 
#     return dirs
 
 
# def _download_files(
#     bucket: str,
#     found:  dict[str, dict],
#     dirs:   dict[str, Path],
#     auditor,
# ) -> None:
#     """Download each identified S3 file to its corresponding local sub-folder."""
#     for ftype, info in found.items():
#         if ftype not in dirs:
#             log.warning("  No local dir for type '%s' — skipping %s", ftype, info["filename"])
#             continue
 
#         local_path = dirs[ftype] / info["filename"]
#         log.info("  Downloading %-22s ← s3://%s/%s", ftype, bucket, info["key"])
#         try:
#             s3.download_file(bucket, info["key"], str(local_path))
#             log.info("  Saved → %s", local_path)
#             if auditor:
#                 auditor.record_download(ftype, info["filename"], status="ok")
#         except Exception as exc:
#             log.error("  Download FAILED for %s: %s", info["key"], exc)
#             if auditor:
#                 auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
#             raise
 
 
# def _upload_xml_outputs(principal: str) -> list[str]:
#     """Upload all .xml files under output/xml/ to the processed bucket."""
#     xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
#     uploaded = []
 
#     for xml_file in xml_dir.glob("*.xml"):
#         s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
#         log.info("  Uploading → s3://%s/%s", PROCESSED_BUCKET, s3_key)
#         try:
#             s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
#             uploaded.append(s3_key)
#             log.info("  Uploaded: %s", xml_file.name)
#         except Exception as exc:
#             log.error("  Upload FAILED for %s: %s", xml_file.name, exc)
 
#     return uploaded
 
 
# def _parse_metadata_from_brand_filename(brand_filename: str) -> dict:
#     """
#     Parse pipeline config from a Diadora brand input filename.

#     Required format (same standard as all other brands):
#         {CompanyCode}-{SBU}-{Brand}-{FileType}-{MultiMono}-{Season}-{Seq}.xlsx
#     Example:
#         0888-SP-Diadora-Pricelist-Multi-FW26-1.xlsx
#         0888-SP-DIADORA-Recap Sample (MAA Sport) Licensed-Multi-FL2029-KH-1.xlsx
#     """
#     stem = brand_filename
#     for ext in (".xlsx", ".csv", ".xlsb"):
#         if stem.lower().endswith(ext):
#             stem = stem[: -len(ext)]
#             break

#     parts = re.split(r"\s*-\s*", stem)

#     # Season: 2 letters + 2 OR 4 digits (FW26, FW2026)
#     season = ""
#     season_idx = None
#     for i, p in enumerate(parts):
#         tok = p.strip()
#         if re.match(r"^[A-Z]{2}\d{2,4}$", tok, re.IGNORECASE):
#             season = tok.upper()
#             season_idx = i
#             break

#     if len(parts) >= 4 and season_idx is not None and season_idx >= 3:
#         comp_code  = parts[0].strip()
#         sbu        = parts[1].strip() if len(parts) > 1 else ""
#         brand      = parts[2].strip() if len(parts) > 2 else ""
#         file_type  = parts[3].strip() if len(parts) > 3 else ""
#         multi_mono = parts[4].strip() if season_idx > 4 else "Multi"

#         trailing = parts[season_idx + 1:]

#         # seq = last all-digit token after season
#         seq = 1
#         for p in reversed(trailing):
#             tok = p.strip()
#             if tok.isdigit():
#                 seq = int(tok)
#                 break

#         # Normalise FW2026 → FW26
#         season = re.sub(
#             r'^(FW|SS|AL)(20)(\d{2})$', lambda x: f"{x.group(1)}{x.group(3)}",
#             season, flags=re.IGNORECASE,
#         ).upper()

#         article_type_from_filename = ""
#         for at in ("Inline", "License", "SSE"):
#             if at.lower() in file_type.lower():
#                 article_type_from_filename = at
#                 break

#         country_code = ""
#         for p in trailing:
#             tok = p.strip()
#             if re.match(r"^[A-Z]{2,3}$", tok, re.IGNORECASE) and not tok.isdigit():
#                 country_code = tok.upper()
#                 break

#         if comp_code and season:
#             brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())
#             log.info(
#                 "Parsed metadata: comp=%s  sbu=%s  brand=%s  brand_code=%s  "
#                 "season=%s  seq=%s  multi_mono=%s",
#                 comp_code, sbu, brand, brand_code, season, seq, multi_mono,
#             )
#             return {
#                 "comp_code":  comp_code,
#                 "sbu":        sbu,
#                 "brand":      brand,
#                 "brand_code": brand_code,
#                 "season":     season,
#                 "seq":        seq,
#                 "file_type":  file_type,
#                 "multi_mono": multi_mono,
#                 "country_code": country_code,
#                 "article_type_from_filename": article_type_from_filename,
#             }

#     log.error(
#         "Brand filename '%s' does not follow the required naming convention. "
#         "Expected: {CompanyCode}-{SBU}-{Brand}-{FileType}-{MultiMono}-{Season}-{Seq}.xlsx "
#         "e.g. 0888-SP-Diadora-Pricelist-Multi-FW26-1.xlsx",
#         brand_filename,
#     )
#     return {}
 
 
# # =============================================================================
# # LAMBDA HANDLER
# # =============================================================================
 
# def lambda_handler(event, context, auditor=None):
#     """
#     Lambda entry point for Diadora file processing.
 
#     Routing:
#         pricelist → diadora.main  (inline pricelist → STEP XML)
#     """
#     _direct_invocation = auditor is None
#     if _direct_invocation:
#         auditor = AuditLogger(event=event, context=context)
#         auditor.set_router_context(
#             trigger_type     = "direct_invocation",
#             original_key     = "",
#             detected_brand   = "diadora",
#             routing_decision = "diadora lambda_handler invoked directly",
#         )
 
#     log.info("Event received: %s", json.dumps(event))
 
#     # ── 1. Parse EventBridge event ────────────────────────────────
#     bucket, key = _parse_event(event)
#     if not bucket or not key:
#         log.error("Cannot parse bucket/key from event")
#         auditor.set_lambda_status("error", error="Unparseable event")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 400, "error": "Unparseable event"}
 
#     key = unquote_plus(key)
#     log.info("Triggered by: s3://%s/%s", bucket, key)
 
#     # ── 2. Guard: only process keys under raw/metadata/ ──────────
#     if not key.startswith("raw/metadata/"):
#         log.warning("Key outside raw/metadata/ — ignoring: %s", key)
#         auditor.set_lambda_status("skipped_outside_prefix")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}
 
#     # ── 3. Extract principal ──────────────────────────────────────
#     principal = _extract_principal(key)
#     if not principal:
#         log.error("Could not extract principal from key: %s", key)
#         auditor.set_lambda_status("error", error="Cannot determine principal")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 400, "error": "Cannot determine principal"}
 
#     log.info("Principal (brand folder): %s", principal)
 
#     # ── 3b. Detect triggered file type ───────────────────────────
#     triggered_ftype = _detect_triggered_file_type(key)
#     log.info("Triggered file type: %s", triggered_ftype or "unknown/global-fanout")
 
#     etl_module = ETL_DISPATCHER.get(triggered_ftype)
 
#     # ── 4. Discover + classify all files in S3 ───────────────────
#     log.info("Listing files under raw/metadata/%s/ ...", principal)
#     found = _list_principal_files(bucket, principal)
#     log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
 
#     if auditor:
#         auditor.set_files_classified(found)
 
#     # ── 4b. Global fan-out: resolve ETL from files present ───────
#     if etl_module is None:
#         for candidate in ("recap", "pricelist"):
#             if candidate in found:
#                 etl_module      = ETL_DISPATCHER[candidate]
#                 triggered_ftype = candidate
#                 log.info("Global fan-out: resolved ETL module → '%s'", candidate)
#                 break
 
#     if etl_module is None:
#         log.info(
#             "No brand file present yet for '%s' during global fan-out — skipping.",
#             principal,
#         )
#         auditor.set_lambda_status("skipped_no_brand_file")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "status":     "skipped_no_brand_file",
#         }
 
#     # ── 5. Check mandatory files ──────────────────────────────────
#     missing = _check_mandatory_files(found, triggered_ftype)
#     if missing:
#         log.info(
#             "Mandatory files not yet present for '%s' — waiting for: %s",
#             principal, missing,
#         )
#         if auditor:
#             auditor.set_files_missing(missing)
#             auditor.set_lambda_status("waiting_for_mandatory_files")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "present":    sorted(found.keys()),
#             "missing":    missing,
#             "status":     "waiting_for_mandatory_files",
#         }
 
#     log.info(
#         "Mandatory files present for '%s' — triggered type: %s",
#         principal, triggered_ftype,
#     )
 
#     # ── 6. Prepare clean /tmp/ directories ───────────────────────
#     dirs = _prepare_tmp_dirs()
 
#     # ── 7. Download files ─────────────────────────────────────────
#     log.info("Downloading %d files (%s) ...", len(found), sorted(found.keys()))
#     try:
#         _download_files(bucket, found, dirs, auditor)
#     except Exception as exc:
#         log.error("Download failed: %s", exc)
#         if auditor:
#             auditor.set_lambda_status("error", error=f"Download failed: {exc}")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {"statusCode": 500, "principal": principal, "error": str(exc)}
 
#     # ── 7b. Convert .xlsb / .csv → .xlsx ─────────────────────────
#     log.info("Converting .xlsb / .csv files to .xlsx ...")
#     for type_dir in dirs.values():
#         for xlsb_file in type_dir.glob("*.xlsb"):
#             try:
#                 _convert_xlsb_to_xlsx(xlsb_file)
#                 if auditor:
#                     auditor.record_conversion(
#                         xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
#                     )
#             except Exception as e:
#                 log.error("xlsb conversion failed for %s: %s", xlsb_file.name, e)
#                 if auditor:
#                     auditor.record_conversion(
#                         xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx",
#                         status="failed", error=str(e),
#                     )
#                     auditor.set_lambda_status(
#                         "error", error=f"xlsb conversion failed: {xlsb_file.name}: {e}"
#                     )
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {
#                     "statusCode": 500, "principal": principal,
#                     "error": f"xlsb conversion failed for {xlsb_file.name}: {e}",
#                 }
 
#         for csv_file in type_dir.glob("*.csv"):
#             try:
#                 _convert_csv_to_xlsx(csv_file)
#                 if auditor:
#                     auditor.record_conversion(
#                         csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
#                     )
#             except Exception as e:
#                 log.error("CSV conversion failed for %s: %s", csv_file.name, e)
#                 if auditor:
#                     auditor.record_conversion(
#                         csv_file.name, from_ext=".csv", to_ext=".xlsx",
#                         status="failed", error=str(e),
#                     )
#                     auditor.set_lambda_status(
#                         "error", error=f"csv conversion failed: {csv_file.name}: {e}"
#                     )
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {
#                     "statusCode": 500, "principal": principal,
#                     "error": f"csv conversion failed for {csv_file.name}: {e}",
#                 }
 
#     # ── 8. Parse metadata from the triggering brand filename ──────
#     meta_filename = found.get(triggered_ftype, {}).get("filename")
#     if not meta_filename:
#         log.error("Brand filename not found in classified files for type '%s'.", triggered_ftype)
#         auditor.set_lambda_status("error", error=f"{triggered_ftype} file missing from found set")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {"statusCode": 400, "principal": principal, "error": f"No {triggered_ftype} file found"}
 
#     log.info("Parsing metadata from: %s", meta_filename)
#     meta = _parse_metadata_from_brand_filename(meta_filename)
 
#     if not meta:
#         log.error("Cannot parse metadata from brand file '%s'.", meta_filename)
#         if auditor:
#             auditor.set_lambda_status(
#                 "error", error=f"Unparseable metadata in {triggered_ftype}",
#             )
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 400,
#             "principal":  principal,
#             "error":      f"Unparseable metadata in {triggered_ftype}",
#         }
 
#     log.info(
#         "Metadata: brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
#         meta["brand"], meta["brand_code"], meta["comp_code"],
#         meta["sbu"], meta["season"], meta["seq"],
#     )
#     if auditor:
#         auditor.set_metadata_parsed(meta)
 
#     # ── 9. Run ETL ────────────────────────────────────────────────
#     log.info("── Running %s ETL → %s ──", triggered_ftype.upper(), etl_module.__name__)
 
#     args = types.SimpleNamespace(
#         brand      = meta["brand"].title(),
#         brand_code = meta["brand_code"],
#         comp_code  = meta["comp_code"],
#         sbu        = meta["sbu"],
#         season     = meta["season"],
#         seq        = meta["seq"],
#         file_type  = meta.get("file_type", "Pricelist"),
#         multi_mono = meta.get("multi_mono", "Multi"),
#         country_code = meta.get("country_code", ""),
#         article_type_from_filename = meta.get("article_type_from_filename", ""),
#     )
 
#     try:
#         etl_module.run(args, auditor=auditor)
#         log.info("── %s ETL complete ──", triggered_ftype.upper())
#     except Exception as exc:
#         log.error("%s ETL failed: %s", triggered_ftype.upper(), exc, exc_info=True)
#         if auditor:
#             auditor.set_lambda_status("error", error=f"{triggered_ftype} ETL failed: {exc}")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {"statusCode": 500, "principal": principal, "error": str(exc)}
 
#     # ── 10. Upload generated XML(s) ───────────────────────────────
#     uploaded = _upload_xml_outputs(principal)
 
#     if not uploaded:
#         log.warning("ETL completed but no XML files were produced under output/xml/")
#         if auditor:
#             auditor.set_lambda_status("no_xml_generated")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "status":     "no_xml_generated",
#         }
 
#     log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
#     if auditor:
#         auditor.set_xml_uploads(uploaded)
#         auditor.set_lambda_status("ok")
#     if _direct_invocation:
#         auditor.flush(principal)
 
#     brand_file_info = found.get(triggered_ftype, {})
#     return {
#         "statusCode":          200,
#         "principal":           principal,
#         "triggered_file_type": triggered_ftype,
#         "brand":               meta["brand"],
#         "brand_code":          meta["brand_code"],
#         "season":              meta["season"],
#         "sbu":                 meta["sbu"],
#         "brand_file_used":     brand_file_info.get("filename"),
#         "brand_file_s3_key":   brand_file_info.get("key"),
#         "uploaded":            uploaded,
#         "count":               len(uploaded),
#         "status":              "ok",
#     }


"""
lambda_function.py — Unified Lambda Handler (Diadora)
======================================================
Function name : map-stibo-inbound-validate-transform-diadora-dev
 
Handles Diadora file types from a single S3 folder:
    raw/metadata/diadora/
 
File-type → ETL dispatch:
    ┌──────────────────────────────────────────────────────────────┐
    │ Filename contains           → ETL module                     │
    ├──────────────────────────────────────────────────────────────┤
    │ "Pricelist" / "Price List"  → diadora.main  (inline)        │
    └──────────────────────────────────────────────────────────────┘
 
    Detection order:
        1. pricelist  — inline pricelist (the only Diadora ETL for now)
        2. mdd / attributes — shared global support files
 
Global files (mdd) fetched from raw/metadata/ root.
 
────────────────────────────────────────────────────────────────────
ENVIRONMENT VARIABLES
────────────────────────────────────────────────────────────────────
    RAW_BUCKET        Source bucket       e.g. map-stibo-inbound-raw-dev
    PROCESSED_BUCKET  Destination bucket  e.g. map-stibo-inbound-processed-dev
    COMP_CODE         Optional fallback company code  (default: "0888")
"""
 
import json
import logging
import os
import re
import shutil
import sys
import types
import openpyxl
import pyxlsb
from pathlib import Path
from urllib.parse import unquote_plus
 
import boto3
import pandas as pd
 
from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files
 
# ─────────────────────────────────────────────────────────────────
# Set LAMBDA_TMP_DIR BEFORE importing ETL modules
# ─────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
 
import diadora.main as etl_pricelist   # Diadora inline pricelist
import diadora.recap as etl_recap      # Diadora licensed recap
 
# ─────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_diadora")
 
# ─────────────────────────────────────────────────────────────────
# AWS client
# ─────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")
 
# ─────────────────────────────────────────────────────────────────
# Environment variables
# ─────────────────────────────────────────────────────────────────
RAW_BUCKET        = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET  = os.environ["PROCESSED_BUCKET"]
# ─────────────────────────────────────────────────────────────────
# Brand name → brand code
# ─────────────────────────────────────────────────────────────────
BRAND_NAME_TO_CODE: dict[str, str] = {
    "DIADORA":     "DIA",
    "NEW BALANCE": "NEW",
    "ADIDAS":      "ADI",
    "NIKE":        "NIK",
    "LOTTO":       "LOT",
}
 
# ─────────────────────────────────────────────────────────────────
# File type detection
# ─────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "pricelist":  ["pricelist", "price list", "price_list"],
    "recap":      ["recap", "recap sample"],
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": ["attributes", "attributes_list", "attribute list"],
}
 
# ─────────────────────────────────────────────────────────────────
# ETL dispatcher  —  file_type key → ETL module
# ─────────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "pricelist": etl_pricelist,
    "recap":     etl_recap,
}
 
# ─────────────────────────────────────────────────────────────────
# Mandatory file rules
# ─────────────────────────────────────────────────────────────────
REQUIRED_NON_BRAND_TYPES: set[str] = {"mdd", "attributes"}
 
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "pricelist": {"pricelist", "mdd", "attributes"},
    "recap":     {"recap", "mdd", "attributes"},
}
 
# Global types fetched from raw/metadata/ root (not brand subfolder)
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}
 
 
# =============================================================================
# HELPERS
# =============================================================================
 
def _clean_cell(v):
    """Strip illegal XML/Excel characters from string cell values."""
    if not isinstance(v, str):
        return v
    return "".join(c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r")
 
 
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
    log.info("  Converted %s → %s (%d sheets)", xlsb_path.name, xlsx_path.name, sheets_written)
    xlsb_path.unlink()
    return xlsx_path
 
 
def _convert_csv_to_xlsx(csv_path: Path) -> Path:
    """Convert a .csv to .xlsx so openpyxl can read it."""
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
 
 
def _detect_file_type(filename: str) -> str | None:
    """
    Return the file type key for a given filename, or None if unrecognised.
    Case-insensitive keyword substring match.
    """
    name_lower = filename.lower()
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw.lower() in name_lower for kw in keywords):
            return ftype
    return None
 
 
def _detect_triggered_file_type(key: str) -> str | None:
    """
    Derive the file type from the S3 key that triggered this invocation.
    Returns None for global fan-out placeholder keys (starting with '.__').
    """
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)
 
 
def _parse_event(event: dict) -> tuple[str, str]:
    """Extract bucket and key from an EventBridge S3 Object Created event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)
 
 
def _extract_principal(key: str) -> str | None:
    """
    Derive the principal (brand folder) from the S3 key.
    Key format: raw/metadata/{principal}/{filename}
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not match raw/metadata/{principal}/{file}: %s", key)
    return None
 
 
def _list_principal_files(bucket: str, principal: str) -> dict[str, dict]:
    """
    Scan S3 for files belonging to this principal.
    Brand-specific files: raw/metadata/{principal}/
    Global shared files (mdd): raw/metadata/ (root level only)
    Brand-specific version always wins over global.
    """
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}
 
    # ── Brand-specific folder ─────────────────────────────────
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
 
            ftype = _detect_file_type(filename)
            if ftype is None:
                log.warning("  Unrecognised file (skipping): %s", filename)
                continue
 
            last_modified = obj["LastModified"]
            if ftype not in found or last_modified > found[ftype]["last_modified"]:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified,
                }
                log.info("  Classified  %-22s ← %s  [brand-specific]", ftype, filename)
 
    # ── Global folder (mdd only) ────────────────────
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        log=log,
    )

    for page in paginator.paginate(
        Bucket=bucket, Prefix="raw/metadata/", Delimiter="/"
    ):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue
 
            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES:
                continue
 
            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified,
                }
                log.info("  Classified  %-22s ← %s  [global]", ftype, filename)
            else:
                log.info(
                    "  Skipping global %s — brand-specific already found: %s",
                    ftype, found[ftype]["filename"],
                )
 
    return found
 
 
def _check_mandatory_files(
    found: dict[str, dict], triggered_ftype: str | None
) -> list[str]:
    """Returns list of missing mandatory types."""
    if triggered_ftype and triggered_ftype in REQUIRED_TYPES_BY_TRIGGER:
        required = REQUIRED_TYPES_BY_TRIGGER[triggered_ftype]
        return sorted(required - set(found.keys()))
 
    # Generic fallback
    missing = list(REQUIRED_NON_BRAND_TYPES - set(found.keys()))
    has_brand_file = any(ftype in found for ftype in ETL_DISPATCHER)
    if not has_brand_file:
        missing.append("pricelist")
    return sorted(missing)
 
 
def _prepare_tmp_dirs() -> dict[str, Path]:
    """Recreate /tmp/stibo_workdir/input/ from scratch each invocation."""
    base = Path(TMP_WORKDIR)
 
    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))
 
    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)
 
    dirs: dict[str, Path] = {
        "pricelist":  base / "input" / "pricelist",
        "recap":      base / "input" / "recap",
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
 
    return dirs
 
 
def _download_files(
    bucket: str,
    found:  dict[str, dict],
    dirs:   dict[str, Path],
    auditor,
) -> None:
    """Download each identified S3 file to its corresponding local sub-folder."""
    for ftype, info in found.items():
        if ftype not in dirs:
            log.warning("  No local dir for type '%s' — skipping %s", ftype, info["filename"])
            continue
 
        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-22s ← s3://%s/%s", ftype, bucket, info["key"])
        try:
            s3.download_file(bucket, info["key"], str(local_path))
            log.info("  Saved → %s", local_path)
            if auditor:
                auditor.record_download(ftype, info["filename"], status="ok")
        except Exception as exc:
            log.error("  Download FAILED for %s: %s", info["key"], exc)
            if auditor:
                auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
            raise
 
 
def _upload_xml_outputs(principal: str) -> list[str]:
    """Upload all .xml files under output/xml/ to the processed bucket."""
    xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
    uploaded = []
 
    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading → s3://%s/%s", PROCESSED_BUCKET, s3_key)
        try:
            s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
            uploaded.append(s3_key)
            log.info("  Uploaded: %s", xml_file.name)
        except Exception as exc:
            log.error("  Upload FAILED for %s: %s", xml_file.name, exc)
 
    return uploaded
 
 
def _parse_metadata_from_brand_filename(brand_filename: str) -> dict:
    """
    Parse pipeline config from a Diadora brand input filename.

    Required format (same standard as all other brands):
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{MultiMono}-{Season}-{Seq}.xlsx
    Example:
        0888-SP-Diadora-Pricelist-Multi-FW26-1.xlsx
        0888-SP-DIADORA-Recap Sample (MAA Sport) Licensed-Multi-FL2029-KH-1.xlsx
    """
    stem = brand_filename
    for ext in (".xlsx", ".csv", ".xlsb"):
        if stem.lower().endswith(ext):
            stem = stem[: -len(ext)]
            break

    parts = re.split(r"\s*-\s*", stem)

    # Season: 2 letters + 2 OR 4 digits (FW26, FW2026)
    season = ""
    season_idx = None
    for i, p in enumerate(parts):
        tok = p.strip()
        if re.match(r"^[A-Z]{2}\d{2,4}$", tok, re.IGNORECASE):
            season = tok.upper()
            season_idx = i
            break

    if len(parts) >= 4 and season_idx is not None and season_idx >= 3:
        comp_code  = parts[0].strip()
        sbu        = parts[1].strip() if len(parts) > 1 else ""
        brand      = parts[2].strip() if len(parts) > 2 else ""
        file_type  = parts[3].strip() if len(parts) > 3 else ""
        multi_mono = parts[4].strip() if season_idx > 4 else "Multi"

        trailing = parts[season_idx + 1:]

        # seq = last all-digit token after season
        seq = 1
        for p in reversed(trailing):
            tok = p.strip()
            if tok.isdigit():
                seq = int(tok)
                break

        # Normalise FW2026 → FW26
        season = re.sub(
            r'^(FW|SS|AL)(20)(\d{2})$', lambda x: f"{x.group(1)}{x.group(3)}",
            season, flags=re.IGNORECASE,
        ).upper()

        article_type_from_filename = ""
        for at in ("Inline", "License", "SSE"):
            if at.lower() in file_type.lower():
                article_type_from_filename = at
                break

        country_code = ""
        for p in trailing:
            tok = p.strip()
            if re.match(r"^[A-Z]{2,3}$", tok, re.IGNORECASE) and not tok.isdigit():
                country_code = tok.upper()
                break

        if comp_code and season:
            brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())
            log.info(
                "Parsed metadata: comp=%s  sbu=%s  brand=%s  brand_code=%s  "
                "season=%s  seq=%s  multi_mono=%s",
                comp_code, sbu, brand, brand_code, season, seq, multi_mono,
            )
            return {
                "comp_code":  comp_code,
                "sbu":        sbu,
                "brand":      brand,
                "brand_code": brand_code,
                "season":     season,
                "seq":        seq,
                "file_type":  file_type,
                "multi_mono": multi_mono,
                "country_code": country_code,
                "article_type_from_filename": article_type_from_filename,
            }

    log.error(
        "Brand filename '%s' does not follow the required filename format. "
        "Expected: {CompanyCode}-{SBU}-{Brand}-{FileType}-{MultiMono}-{Season}-{Seq}.xlsx "
        "e.g. 0888-SP-Diadora-Pricelist-Multi-FW26-1.xlsx",
        brand_filename,
    )
    return {}
 
 
# =============================================================================
# LAMBDA HANDLER
# =============================================================================
 
def lambda_handler(event, context, auditor=None):
    """
    Lambda entry point for Diadora file processing.
 
    Routing:
        pricelist → diadora.main  (inline pricelist → STEP XML)
    """
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "diadora",
            routing_decision = "diadora lambda_handler invoked directly",
        )
 
    log.info("Event received: %s", json.dumps(event))
 
    # ── 1. Parse EventBridge event ────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}
 
    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)
 
    # ── 2. Guard: only process keys under raw/metadata/ ──────────
    if not key.startswith("raw/metadata/"):
        log.warning("Key outside raw/metadata/ — ignoring: %s", key)
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}
 
    # ── 3. Extract principal ──────────────────────────────────────
    principal = _extract_principal(key)
    if not principal:
        log.error("Could not extract principal from key: %s", key)
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}
 
    log.info("Principal (brand folder): %s", principal)
 
    # ── 3b. Detect triggered file type ───────────────────────────
    triggered_ftype = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_ftype or "unknown/global-fanout")
 
    etl_module = ETL_DISPATCHER.get(triggered_ftype)
 
    # ── 4. Discover + classify all files in S3 ───────────────────
    log.info("Listing files under raw/metadata/%s/ ...", principal)
    found = _list_principal_files(bucket, principal)
    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
 
    if auditor:
        auditor.set_files_classified(found)
 
    # ── 4b. Global fan-out: resolve ETL from files present ───────
    if etl_module is None:
        for candidate in ("recap", "pricelist"):
            if candidate in found:
                etl_module      = ETL_DISPATCHER[candidate]
                triggered_ftype = candidate
                log.info("Global fan-out: resolved ETL module → '%s'", candidate)
                break
 
    if etl_module is None:
        log.info(
            "No brand file present yet for '%s' during global fan-out — skipping.",
            principal,
        )
        auditor.set_lambda_status("skipped_no_brand_file")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "skipped_no_brand_file",
        }
 
    # ── 5. Check mandatory files ──────────────────────────────────
    missing = _check_mandatory_files(found, triggered_ftype)
    if missing:
        log.info(
            "Mandatory files not yet present for '%s' — waiting for: %s",
            principal, missing,
        )
        if auditor:
            auditor.set_files_missing(missing)
            auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    missing,
            "status":     "waiting_for_mandatory_files",
        }
 
    log.info(
        "Mandatory files present for '%s' — triggered type: %s",
        principal, triggered_ftype,
    )
 
    # ── 6. Prepare clean /tmp/ directories ───────────────────────
    dirs = _prepare_tmp_dirs()
 
    # ── 7. Download files ─────────────────────────────────────────
    log.info("Downloading %d files (%s) ...", len(found), sorted(found.keys()))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        log.error("Download failed: %s", exc)
        if auditor:
            auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": str(exc)}
 
    # ── 7b. Convert .xlsb / .csv → .xlsx ─────────────────────────
    log.info("Converting .xlsb / .csv files to .xlsx ...")
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                if auditor:
                    auditor.record_conversion(
                        xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
                    )
            except Exception as e:
                log.error("xlsb conversion failed for %s: %s", xlsb_file.name, e)
                if auditor:
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
                    "error": f"xlsb conversion failed for {xlsb_file.name}: {e}",
                }
 
        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                if auditor:
                    auditor.record_conversion(
                        csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
                    )
            except Exception as e:
                log.error("CSV conversion failed for %s: %s", csv_file.name, e)
                if auditor:
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
                    "error": f"csv conversion failed for {csv_file.name}: {e}",
                }
 
    # ── 8. Parse metadata from the triggering brand filename ──────
    meta_filename = found.get(triggered_ftype, {}).get("filename")
    if not meta_filename:
        log.error("Brand filename not found in classified files for type '%s'.", triggered_ftype)
        auditor.set_lambda_status("error", error=f"{triggered_ftype} file missing from found set")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 400, "principal": principal, "error": f"No {triggered_ftype} file found"}
 
    log.info("Parsing metadata from: %s", meta_filename)
    meta = _parse_metadata_from_brand_filename(meta_filename)
 
    if not meta:
        log.error("Cannot parse metadata from brand file '%s'.", meta_filename)
        if auditor:
            auditor.set_lambda_status(
                "error", error=f"Unparseable metadata in {triggered_ftype}",
            )
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 400,
            "principal":  principal,
            "error":      f"Unparseable metadata in {triggered_ftype}",
        }
 
    log.info(
        "Metadata: brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        meta["brand"], meta["brand_code"], meta["comp_code"],
        meta["sbu"], meta["season"], meta["seq"],
    )
    if auditor:
        auditor.set_metadata_parsed(meta)
 
    # ── 9. Run ETL ────────────────────────────────────────────────
    log.info("── Running %s ETL → %s ──", triggered_ftype.upper(), etl_module.__name__)
 
    args = types.SimpleNamespace(
        brand      = meta["brand"].title(),
        brand_code = meta["brand_code"],
        comp_code  = meta["comp_code"],
        sbu        = meta["sbu"],
        season     = meta["season"],
        seq        = meta["seq"],
        file_type  = meta.get("file_type", "Pricelist"),
        multi_mono = meta.get("multi_mono", "Multi"),
        country_code = meta.get("country_code", ""),
        article_type_from_filename = meta.get("article_type_from_filename", ""),
    )
 
    try:
        etl_module.run(args, auditor=auditor)
        log.info("── %s ETL complete ──", triggered_ftype.upper())
    except Exception as exc:
        log.error("%s ETL failed: %s", triggered_ftype.upper(), exc, exc_info=True)
        if auditor:
            auditor.set_lambda_status("error", error=f"{triggered_ftype} ETL failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": str(exc)}
 
    # ── 10. Upload generated XML(s) ───────────────────────────────
    uploaded = _upload_xml_outputs(principal)
 
    if not uploaded:
        log.warning("ETL completed but no XML files were produced under output/xml/")
        if auditor:
            auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "no_xml_generated",
        }
 
    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    if auditor:
        auditor.set_xml_uploads(uploaded)
        auditor.set_lambda_status("ok")
    if _direct_invocation:
        auditor.flush(principal)
 
    brand_file_info = found.get(triggered_ftype, {})
    return {
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_ftype,
        "brand":               meta["brand"],
        "brand_code":          meta["brand_code"],
        "season":              meta["season"],
        "sbu":                 meta["sbu"],
        "brand_file_used":     brand_file_info.get("filename"),
        "brand_file_s3_key":   brand_file_info.get("key"),
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
    }
