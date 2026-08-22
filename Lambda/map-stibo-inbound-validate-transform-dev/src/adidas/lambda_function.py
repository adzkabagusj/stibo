
# """
# lambda_function.py — Lambda 1 Handler
# ======================================
# Function name : map-stibo-inbound-validate-transform-dev

# Trigger:
#     EventBridge rule (S3 Object Created) scoped to:
#         s3://map-stibo-inbound-raw-dev/raw/metadata/*

# Flow:
#     1. Parse EventBridge event → extract principal (brand) from S3 key
#        Key format: raw/metadata/{principal}/{filename}.xlsx

#     2. List ALL .xlsx files under raw/metadata/{principal}/

#     3. Classify each file into one of 6 types by filename keyword:
#          linelist / backlog / tdd / mdd / attributes / naming

#     4. If any mandatory type is missing → log and return 200 (wait)

#     5. Download all files to /tmp/stibo_workdir/input/{type}/

#     6. Call main.run(args) ← ALL existing ETL logic — UNCHANGED

#     7. Upload every .xml produced under /tmp/stibo_workdir/output/xml/
#        to:
#            s3://map-stibo-inbound-processed-dev/
#                processed/stepxml/{principal}/{filename}

#     8. Return structured response

# ────────────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# ────────────────────────────────────────────────────────────────────
#     RAW_BUCKET        Source bucket
#                       e.g.  map-stibo-inbound-raw-dev

#     PROCESSED_BUCKET  Destination bucket (triggers Lambda 2)
#                       e.g.  map-stibo-inbound-processed-dev

# ────────────────────────────────────────────────────────────────────
# NOTE — Debug break in main.py
# ────────────────────────────────────────────────────────────────────
#     main.py currently has a `break` inside the article loop (~line 1371)
#     that stops after the first matched article. Remove before production.
# """

# import json
# import logging
# import os
# import re  
# import shutil
# import sys
# import types
# import pyxlsb
# import openpyxl
# from pathlib import Path
# from urllib.parse import unquote_plus


# # from adidas.linelist_diff import run_diff_and_upload

# import boto3
# import pandas as pd

# from audit_logger import AuditLogger
# from metadata_s3 import add_root_metadata_files

# # ─────────────────────────────────────────────────────────────────────────────
# # CRITICAL: Set LAMBDA_TMP_DIR BEFORE importing main so that main.py's
# # module-level BASE_DIR resolves to /tmp/stibo_workdir instead of
# # the Lambda read-only /var/task/ directory.
# # ─────────────────────────────────────────────────────────────────────────────
# # TMP_WORKDIR = "/tmp/stibo_workdir"
# # os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
# TMP_WORKDIR = "/tmp/stibo_workdir_adidas"
# os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
# from adidas.tdd_main import run_tdd
# from adidas.backlog_main import run_backlog
# # from adidas.dtb_main import run_dtb
# from adidas.dtb_main import run as run_dtb
# import adidas.main as etl  # noqa: E402  (import after env var intentional)

# # ─────────────────────────────────────────────────────────────────────────────
# # Logging
# # ─────────────────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s  %(levelname)-8s  %(message)s",
#     handlers=[logging.StreamHandler(sys.stdout)],
# )
# log = logging.getLogger("lambda1")

# # ─────────────────────────────────────────────────────────────────────────────
# # AWS client (reused across warm invocations)
# # ─────────────────────────────────────────────────────────────────────────────
# s3 = boto3.client("s3")

# # ─────────────────────────────────────────────────────────────────────────────
# # Environment variables — only 2 needed
# # ─────────────────────────────────────────────────────────────────────────────
# RAW_BUCKET       = os.environ["RAW_BUCKET"]
# PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]


# # ─────────────────────────────────────────────────────────────────────────────
# # Brand name → brand code lookup
# # ─────────────────────────────────────────────────────────────────────────────
# BRAND_NAME_TO_CODE: dict[str, str] = {
#     "ADIDAS":      "ADI",
#     "NIKE":        "NIK",
#     "NEW BALANCE": "NEW",
#     "SMIGGLE":     "SMI",
#     "ALDO":        "ALD",
#     "CROCS":       "CRO",
#     "LOTTO":       "LOT",
#     "BIRKENSTOCK": "BIR",
# }

# # ─────────────────────────────────────────────────────────────────────────────
# # File type identification keywords
# # ─────────────────────────────────────────────────────────────────────────────
# FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
#     "linelist":      ["line list", "linelist", "LINELIST"],
#     "backlog":       ["backlog", "Backlog"],
#     "tdd":           ["tdd", "TDD", "( ros )", "(ros)", "ros"],
#     "dtb":           ["dtb", "DTB"],
#     "mdd":           ["mdd", "master_data", "master data"],
#     "attributes":    ["attributes", "attributes_list", "attribute list"],
#     "naming":        ["naming", "naming_convention", "naming convention"],
#     # Licensed brand types
#     "recap":         ["recap", "recap sample"],
#     "order_sheet":   ["order sheet", "ordersheet", "order_sheet"],
#     "ean_source":    ["ean source", "ean_source", "upc source"],
#     "packing_list":  ["packing list", "packing_list", "packlist"],
#     "oor":           ["oor", "out of ratio", "out_of_ratio"],
#     "open_shipment": ["open shipment", "open_shipment", "openshipment"],
# }

# REQUIRED_TYPES_BY_BRAND = {
#     "ADI": {"linelist"},
# }
# _DEFAULT_REQUIRED = {"linelist"}


# # =============================================================================
# # HELPERS
# # =============================================================================

# def _clean_cell(v):
#     """Strip illegal XML/Excel characters from string cell values."""
#     if not isinstance(v, str):
#         return v
#     return ''.join(
#         c for c in v
#         if 0x20 <= ord(c) <= 0xFFFD or c in '\t\n\r'
#     )


# def _safe_ts_to_epoch(ts) -> int:
#     """
#     Convert ANY timestamp type (datetime, int, float, str) to Unix epoch seconds.
#     Returns 0 on failure — avoids TypeError on mixed-type comparisons.
#     """
#     try:
#         if hasattr(ts, 'timestamp'):            # datetime / aware datetime
#             return int(ts.timestamp())
#         elif isinstance(ts, (int, float)):
#             return int(float(ts))
#         elif isinstance(ts, str):
#             from datetime import datetime
#             dt = datetime.fromisoformat(ts.replace('Z', '+00:00'))
#             return int(dt.timestamp())
#     except Exception:
#         pass
#     return 0


# def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
#     """Convert a .xlsb to .xlsx — streaming row-by-row to keep memory low."""
#     xlsx_path = xlsb_path.with_suffix(".xlsx")
#     wb_out = openpyxl.Workbook(write_only=True)
#     sheets_written = 0

#     with pyxlsb.open_workbook(str(xlsb_path)) as wb:
#         for sheet_name in wb.sheets:
#             safe_name = ''.join(
#                 c for c in sheet_name
#                 if 0x20 <= ord(c) <= 0xFFFD and c not in '\\/?*[]:'
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
#                             log.warning("  Skipping bad row in sheet '%s': %s", safe_name, e)
#                             continue
#                 sheets_written += 1
#                 log.info("  Written sheet: '%s' → '%s'", sheet_name, safe_name)
#             except Exception as e:
#                 log.warning("  Error reading sheet '%s': %s", sheet_name, e)
#                 continue

#     if sheets_written == 0:
#         raise ValueError(f"No sheets could be converted from {xlsb_path.name}")

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


# def _parse_metadata_from_linelist_filename(dirs: dict, folder_type: str = "linelist") -> dict:
#     folder = dirs.get(folder_type)
#     if not folder or not folder.exists():
#         log.error("%s directory not found: %s", folder_type, folder)
#         return {}

#     for xlsx_file in folder.glob("*.xlsx"):
#         stem = xlsx_file.stem
#         log.info("Attempting metadata parse from filename: '%s'", stem)

#         parts = re.split(r"\s*-\s*", stem)

#         # Season: 2 letters + 2 OR 4 digits (SP26, FW2026, SP2029)
#         season     = None
#         season_idx = None
#         for i, p in enumerate(parts):
#             if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
#                 season     = p.strip().upper()
#                 season_idx = i
#                 break

#         if not season or season_idx is None or season_idx < 3:  # relaxed from 5 → 3
#             log.warning("Filename '%s': no valid season token found — skipping", stem)
#             continue

#         comp_code  = parts[0].strip()
#         sbu        = parts[1].strip()
#         brand      = parts[2].strip()
#         multi_mono = parts[4].strip() if season_idx > 4 else "Multi"

#         # Everything between season and the last all-digit token
#         # may include country code(s) like MY, ID, etc.
#         trailing = parts[season_idx + 1:]

#         # country = first 2-letter uppercase token after season (if not a digit)
#         country_code = ""
#         for p in trailing:
#             p = p.strip()
#             if re.match(r'^[A-Z]{2,3}$', p, re.IGNORECASE) and not p.isdigit():
#                 country_code = p.upper()
#                 break

#         # seq = last all-digit token in trailing
#         seq = 1
#         for p in reversed(trailing):
#             if p.strip().isdigit():
#                 seq = int(p.strip())
#                 break

#         if not comp_code:
#             log.warning("Filename '%s': empty comp_code — skipping", stem)
#             continue

#         brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())

#         # Extract article_type from file_type token (parts[3])
#         # e.g. "Line List (MAA Sport) Inline" → article_type = "Inline"
#         file_type_token = parts[3].strip() if len(parts) > 3 else ""
#         article_type_from_filename = ""
#         for at in ["Inline", "License", "SSE"]:
#             if at.lower() in file_type_token.lower():
#                 article_type_from_filename = at
#                 break

#         log.info(
#             "Parsed → comp=%s sbu=%s brand=%s brand_code=%s season=%s "
#             "seq=%s country=%s article_type=%s",
#             comp_code, sbu, brand, brand_code, season,
#             seq, country_code, article_type_from_filename,
#         )
#         return {
#             "comp_code":    comp_code,
#             "sbu":          sbu,
#             "brand":        brand,
#             "brand_code":   brand_code,
#             "season":       season,
#             "seq":          seq,
#             "multi_mono":   multi_mono,
#             "country_code": country_code,
#             "article_type_from_filename": article_type_from_filename,
#         }

#     log.error("Cannot parse metadata from filename.")
#     return {}


# def _detect_file_type(filename: str) -> str | None:
#     """
#     Return the file type for a given filename, or None if unclassifiable.
#     Case-insensitive keyword search.
#     """
#     name = filename.lower()
#     for ftype, keywords in FILE_TYPE_KEYWORDS.items():
#         if any(kw.lower() in name for kw in keywords):
#             return ftype
#     return None


# def _parse_event(event: dict) -> tuple[str, str]:
#     """
#     Extract bucket name and object key from an EventBridge S3 Object Created event.
#     """
#     detail = event.get("detail", {})
#     bucket = detail.get("bucket", {}).get("name", "")
#     key    = detail.get("object",  {}).get("key",  "")
#     return bucket, str(key)


# def _extract_principal(key: str) -> str | None:
#     """
#     Derive the principal (brand folder) from the S3 object key.
#     Key format : raw/metadata/{principal}/{filename}
#     """
#     parts = key.strip("/").split("/")
#     if len(parts) >= 4:
#         return parts[2]
#     log.warning("Key does not follow raw/metadata/{principal}/{file} pattern: %s", key)
#     return None


# def _list_principal_files(bucket: str, principal: str,
#                            exclude_types: set = None) -> dict[str, dict]:
#     """
#     Lists files for a brand principal:
#     - Brand-specific files from raw/metadata/{principal}/
#     - Global files (mdd, naming) from raw/metadata/ root
#     Brand-specific files take priority over global if duplicated.

#     exclude_types: file types to skip entirely (used to prevent the
#                    trigger file's type being overridden by a newer S3 object).
#     """
#     exclude_types = exclude_types or set()
#     GLOBAL_TYPES  = {"mdd", "naming"}
#     ROOT_TYPES    = {"mdd", "attributes"}
#     prefix        = f"raw/metadata/{principal}/"
#     paginator     = s3.get_paginator("list_objects_v2")
#     found: dict[str, dict] = {}

#     # ── Step 1: Scan brand-specific folder ───────────────────────────────
#     for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
#         for obj in page.get("Contents", []):
#             obj_key  = obj["Key"]                    # ← renamed from 'key'
#             filename = Path(obj_key).name
#             if not filename:
#                 continue
#             if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
#                 continue

#             ftype = _detect_file_type(filename)
#             if ftype is None:
#                 log.warning("  Unrecognised file (no type match): %s — skipping", filename)
#                 continue
#             if ftype in exclude_types:               # ← skip trigger type
#                 log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
#                 continue

#             last_modified = obj["LastModified"]
#             if ftype not in found or (
#                 _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
#             ):
#                 found[ftype] = {
#                     "key":           obj_key,
#                     "filename":      filename,
#                     "last_modified": last_modified,
#                 }
#                 log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

#     # ── Step 2: Scan global folder raw/metadata/ for mdd + naming ────────
#     add_root_metadata_files(
#         s3_client=s3,
#         bucket=bucket,
#         found=found,
#         include_types=ROOT_TYPES,
#         exclude_types=exclude_types,
#         log=log,
#     )

#     global_prefix = "raw/metadata/"
#     for page in paginator.paginate(Bucket=bucket, Prefix=global_prefix, Delimiter="/"):
#         for obj in page.get("Contents", []):
#             obj_key  = obj["Key"]                    # ← renamed from 'key'
#             filename = Path(obj_key).name
#             if not filename:
#                 continue
#             if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
#                 continue

#             ftype = _detect_file_type(filename)
#             if ftype not in GLOBAL_TYPES:
#                 continue
#             if ftype in exclude_types:               # ← skip trigger type
#                 continue

#             last_modified = obj["LastModified"]
#             if ftype not in found:
#                 found[ftype] = {
#                     "key":           obj_key,
#                     "filename":      filename,
#                     "last_modified": last_modified,
#                 }
#                 log.info("  Classified  %-12s ← %s  [global]", ftype, filename)
#             else:
#                 log.info(
#                     "  Skipping global %s — brand-specific version already found: %s",
#                     ftype, found[ftype]["filename"],
#                 )

#     return found
# def _prepare_tmp_dirs() -> dict[str, Path]:
#     """
#     Recreate /tmp/stibo_workdir/input/{type}/ sub-folders from scratch.
#     Also clears output/xml/ so only current run's XMLs are uploaded.
#     Returns {file_type: local_dir_path} for all input types.
#     """
#     base = Path(TMP_WORKDIR)

#     input_base = base / "input"
#     if input_base.exists():
#         shutil.rmtree(str(input_base))

#     xml_out = base / "output" / "xml"
#     if xml_out.exists():
#         shutil.rmtree(str(xml_out))
#     xml_out.mkdir(parents=True, exist_ok=True)

#     # All possible input type dirs (standard + licensed brand types)
#     dir_types = [
#         "linelist", "backlog", "tdd", "dtb", "mdd", "attributes", "naming",
#         "recap", "order_sheet", "ean_source", "packing_list", "oor", "open_shipment",
#     ]
#     dirs: dict[str, Path] = {}
#     for dtype in dir_types:
#         d = base / "input" / dtype
#         d.mkdir(parents=True, exist_ok=True)
#         dirs[dtype] = d

#     return dirs


# def _download_files(
#     bucket: str,
#     found: dict[str, dict],
#     dirs: dict[str, Path],
#     auditor: AuditLogger,
# ) -> None:
#     """
#     Download each identified S3 file to its corresponding local sub-folder.
#     Records each download outcome into the auditor.
#     Raises on download failure so the caller can handle / surface the error.
#     """
#     for ftype, info in found.items():
#         if ftype not in dirs:
#             log.warning("  No local dir mapped for type '%s' — skipping download", ftype)
#             continue

#         local_path = dirs[ftype] / info["filename"]
#         log.info("  Downloading %-12s ← s3://%s/%s", ftype, bucket, info["key"])
#         try:
#             s3.download_file(bucket, info["key"], str(local_path))
#             log.info("  Saved → %s", local_path)
#             auditor.record_download(ftype, info["filename"], status="ok")
#         except Exception as exc:
#             log.error("  Download FAILED for %s: %s", info["key"], exc)
#             auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
#             raise


# def _upload_xml_outputs(principal: str) -> list[str]:
#     """
#     Find all .xml files under output/xml/ and upload to processed bucket.
#     Returns list of S3 keys successfully uploaded.
#     """
#     xml_dir = Path(TMP_WORKDIR) / "output" / "xml"
#     if not xml_dir.exists():
#         return []

#     uploaded = []
#     for xml_file in xml_dir.glob("*.xml"):
#         key = f"processed/stepxml/{principal}/{xml_file.name}"
#         log.info("  Uploading XML → s3://%s/%s", PROCESSED_BUCKET, key)
#         try:
#             s3.upload_file(str(xml_file), PROCESSED_BUCKET, key)
#             uploaded.append(key)
#         except Exception as e:
#             log.error("  Upload failed for %s: %s", xml_file.name, e)
#     return uploaded

# # =============================================================================
# # LAMBDA HANDLER
# # =============================================================================

# def lambda_handler(event, context, auditor: AuditLogger = None):
#     """
#     Main Lambda entry point.
#     auditor is created and flushed by router.py.
#     If invoked directly (not via router), creates its own auditor as fallback.
#     """
#     # ── Fallback: direct invocation (not via router.py) ───────────────────
#     if auditor is None:
#         auditor = AuditLogger(event=event, context=context)
#         auditor.set_router_context(
#             trigger_type     = "direct_invocation",
#             original_key     = "",
#             detected_brand   = "",
#             routing_decision = "lambda_handler invoked directly without router",
#         )
#         _direct_invocation = True
#     else:
#         _direct_invocation = False

#     log.info("Event received: %s", json.dumps(event))

#     # ── 1. Parse EventBridge event ────────────────────────────────────────
#     bucket, key = _parse_event(event)

#     if not bucket or not key:
#         log.error("Cannot parse bucket/key from event")
#         auditor.set_lambda_status("error", error="Unparseable event")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 400, "error": "Unparseable event"}

#     key = unquote_plus(key)
#     log.info("Triggered by: s3://%s/%s", bucket, key)

#     # Capture trigger info NOW before any S3 listing corrupts 'key'
#     trigger_filename = Path(key).name
#     trigger_type     = _detect_file_type(trigger_filename)
#     log.info("Trigger file: '%s'  type: '%s'", trigger_filename, trigger_type)

#     # ── 2. Guard: only process keys under raw/metadata/ ──────────────────
#     if not key.startswith("raw/metadata/"):
#         log.warning("Key is outside raw/metadata/ — ignoring: %s", key)
#         auditor.set_lambda_status("skipped_outside_prefix")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}

#     # ── 3. Extract principal (brand folder name) from key ─────────────────
#     principal = _extract_principal(key)
#     if not principal:
#         log.error("Could not extract principal from key: %s", key)
#         auditor.set_lambda_status("error", error="Cannot determine principal")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 400, "error": "Cannot determine principal"}

#     log.info("Principal (brand folder): %s", principal)

#     # ── 4. List + classify all files for this principal ───────────────────
#     log.info("Listing files under raw/metadata/%s/ ...", principal)
#     # exclude_types: don't let the listing override the actual trigger file
#     found = _list_principal_files(bucket, principal,
#                                 exclude_types={trigger_type} if trigger_type else set())

#     # Always pin the trigger file — it is the authoritative source for its type
#     if trigger_type:
#         found[trigger_type] = {
#             "key":           key,
#             "filename":      trigger_filename,
#             "last_modified": None,
#         }
#         log.info("Pinned trigger file → type='%s'  file='%s'", trigger_type, trigger_filename)

#     log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
#     auditor.set_files_classified(found)

#     # ── 5a. DTB trigger: self-contained early exit (doesn't need linelist) ──
#     _trigger_fname = Path(key).name
#     _trigger_type_dtb = (
#         _detect_file_type(_trigger_fname)
#         if not _trigger_fname.startswith(".__")
#         else None
#     )

#     if _trigger_type_dtb == "dtb":
#         if "dtb" not in found:
#             log.info("DTB file not yet present for '%s' — waiting.", principal)
#             auditor.set_files_missing(["dtb"])
#             auditor.set_lambda_status("waiting_for_mandatory_files")
#             if _direct_invocation:
#                 auditor.flush(principal)
#             return {
#                 "statusCode": 200,
#                 "principal":  principal,
#                 "present":    sorted(found.keys()),
#                 "missing":    ["dtb"],
#                 "status":     "waiting_for_mandatory_files",
#             }

#         dirs = _prepare_tmp_dirs()
#         try:
#             _download_files(bucket, found, dirs, auditor)
#         except Exception as exc:
#             auditor.set_lambda_status("error", error=f"Download failed: {exc}")
#             if _direct_invocation:
#                 auditor.flush(principal)
#             return {"statusCode": 500, "principal": principal, "error": f"Download failed: {exc}"}

#         for type_dir in dirs.values():
#             for xlsb_file in type_dir.glob("*.xlsb"):
#                 try:
#                     _convert_xlsb_to_xlsx(xlsb_file)
#                     auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok")
#                 except Exception as e:
#                     auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="failed", error=str(e))
#             for csv_file in type_dir.glob("*.csv"):
#                 try:
#                     _convert_csv_to_xlsx(csv_file)
#                     auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok")
#                 except Exception as e:
#                     auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="failed", error=str(e))

#         meta = _parse_metadata_from_linelist_filename(dirs, folder_type="dtb")
#         if not meta:
#             auditor.set_lambda_status("error", error="Unparseable metadata in DTB filename")
#             if _direct_invocation:
#                 auditor.flush(principal)
#             return {
#                 "statusCode": 400,
#                 "principal":  principal,
#                 "error":      "Unparseable metadata in DTB filename. Ensure: {CompCode}-{SBU}-{Brand}-DTB-{Multi/Mono}-{Season}-{Seq}.xlsx",
#             }

#         auditor.set_metadata_parsed(meta)
#         dtb_args = types.SimpleNamespace(
#             brand=meta["brand"],
#             brand_code=meta["brand_code"],
#             comp_code=meta["comp_code"],
#             sbu=meta["sbu"],
#             season=meta["season"],
#             seq=meta["seq"],
#             multi_mono=meta.get("multi_mono", "Multi"),
#             country_code=meta.get("country_code", ""),  # ← ADD THIS LINE
#             article_type_from_filename=meta.get("article_type_from_filename", ""),
#         )

#         log.info("DTB trigger — running DTB ecommerce enrichment pipeline ...")
#         run_dtb(dtb_args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))

#         uploaded = _upload_xml_outputs(principal)
#         if not uploaded:
#             auditor.set_lambda_status("no_xml_generated")
#             if _direct_invocation:
#                 auditor.flush(principal)
#             return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

#         auditor.set_xml_uploads(uploaded)
#         auditor.set_lambda_status("ok")
#         if _direct_invocation:
#             auditor.flush(principal)
#         dtb_info = found.get("dtb", {})
#         return {
#             "statusCode":    200,
#             "principal":     principal,
#             "trigger":       "dtb",
#             "dtb_file_used": dtb_info.get("filename"),
#             "uploaded":      uploaded,
#             "count":         len(uploaded),
#             "status":        "ok",
#         }

#     # ── 5b. Early check: linelist must be present before we do anything ────
#     if "linelist" not in found:
#         log.info("Linelist not yet present for '%s' — waiting.", principal)
#         auditor.set_files_missing(["linelist"])
#         auditor.set_lambda_status("waiting_for_mandatory_files")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "present":    sorted(found.keys()),
#             "missing":    ["linelist"],
#             "status":     "waiting_for_mandatory_files",
#         }

#     # ── 6. Prepare clean /tmp/ input + output directories ─────────────────
#     dirs = _prepare_tmp_dirs()

#     # ── 7. Download all files into their sub-folders ──────────────────────
#     log.info("Downloading %d input files (%s) ...", len(found), sorted(found.keys()))
#     try:
#         _download_files(bucket, found, dirs, auditor)
#     except Exception as exc:
#         auditor.set_lambda_status("error", error=f"Download failed: {exc}")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {"statusCode": 500, "principal": principal, "error": f"Download failed: {exc}"}

#     # ── 7b. Convert .xlsb and .csv files to .xlsx ─────────────────────────
#     log.info("Converting .xlsb and .csv files to .xlsx ...")
#     for type_dir in dirs.values():
#         for xlsb_file in type_dir.glob("*.xlsb"):
#             try:
#                 _convert_xlsb_to_xlsx(xlsb_file)
#                 auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok")
#             except Exception as e:
#                 log.error("Failed to convert %s: %s", xlsb_file.name, e)
#                 auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx",
#                                           status="failed", error=str(e))
#                 auditor.set_lambda_status("error", error=f"xlsb conversion failed: {e}")
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {"statusCode": 500, "principal": principal,
#                         "error": f"xlsb conversion failed for {xlsb_file.name}: {e}"}

#         for csv_file in type_dir.glob("*.csv"):
#             try:
#                 _convert_csv_to_xlsx(csv_file)
#                 auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok")
#             except Exception as e:
#                 log.error("Failed to convert %s: %s", csv_file.name, e)
#                 auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx",
#                                           status="failed", error=str(e))
#                 auditor.set_lambda_status("error", error=f"csv conversion failed: {e}")
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {"statusCode": 500, "principal": principal,
#                         "error": f"csv conversion failed for {csv_file.name}: {e}"}

#     # ── 8. Parse pipeline config from linelist filename ───────────────────
#     meta = _parse_metadata_from_linelist_filename(dirs)

#     if not meta:
#         auditor.set_lambda_status("error", error="Unparseable metadata in linelist filename")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 400,
#             "principal":  principal,
#             "error": (
#                 "Unparseable metadata in linelist filename. "
#                 "Ensure filename follows: "
#                 "{CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Seq}.xlsx "
#                 "e.g. 0888-SP-Adidas-Line List-Multi-FW26-1.xlsx"
#             ),
#         }

#     auditor.set_metadata_parsed(meta)

#     # ── 8b. Brand-specific mandatory file check ───────────────────────────
#     required_types    = REQUIRED_TYPES_BY_BRAND.get(meta["brand_code"], _DEFAULT_REQUIRED)
#     missing_mandatory = required_types - set(found.keys())

#     optional_present = [t for t in ("backlog", "tdd") if t in found]
#     optional_missing = [t for t in ("backlog", "tdd") if t not in found]

#     if missing_mandatory:
#         log.info(
#             "Mandatory files not yet present for '%s' (%s) — waiting for: %s",
#             principal, meta["brand_code"], sorted(missing_mandatory),
#         )
#         auditor.set_files_missing(sorted(missing_mandatory))
#         auditor.set_lambda_status("waiting_for_mandatory_files")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "present":    sorted(found.keys()),
#             "missing":    sorted(missing_mandatory),
#             "status":     "waiting_for_mandatory_files",
#         }

#     log.info(
#         "Mandatory files confirmed for '%s' (%s) — starting ETL. "
#         "Optional present: %s | Optional missing (empty stubs): %s",
#         principal, meta["brand_code"], optional_present, optional_missing,
#     )

#     args = types.SimpleNamespace(
#         brand        = meta["brand"],
#         brand_code   = meta["brand_code"],
#         comp_code    = meta["comp_code"],
#         sbu          = meta["sbu"],
#         season       = meta["season"],
#         seq          = meta["seq"],
#         multi_mono   = meta.get("multi_mono", "Multi"),
#         country_code = meta.get("country_code", principal.upper()),  # ← was principal.upper()
#         article_type_from_filename = meta.get("article_type_from_filename", ""),  # ← NEW
#     )

#     # ── 9. Run ETL pipeline — ALL main.py logic UNCHANGED ─────────────────
#     # log.info("Running ETL pipeline (main.run) ...")
#         # ── 9. Run Pipeline Logic (Sahi tarika) ─────────────────
#     # ── 9. Run Pipeline Logic ─────────────────────────────────────────────
#     log.info("Routing to pipeline — trigger_type: '%s'", trigger_type)

#     # if trigger_type == "linelist":
#     #     log.info("Linelist trigger - running ETL pipeline ...")
#     #     etl.run(args, auditor=auditor)
#     # elif trigger_type == "tdd":
#     #     log.info("TDD trigger - running TDD pipeline ...")
#     #     run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
#     # elif trigger_type == "backlog":
#     #     log.info("Backlog trigger - running Backlog pipeline ...")
#     #     run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
#     # else:
#     #     log.info("Default trigger (%s) - running ETL pipeline ...", trigger_type)
#     #     etl.run(args, auditor=auditor)

#     if trigger_type == "tdd":
#         log.info("TDD trigger - running TDD pipeline ...")
#         run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
#     elif trigger_type == "backlog":
#         log.info("Backlog trigger - running Backlog pipeline ...")
#         run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
#     else:
        
#         log.info("ETL trigger (%s) - running ETL pipeline ...", trigger_type)
#         etl.run(args, auditor=auditor)

#     # backlog
#     # if trigger_type != "backlog" and "backlog" in found:
#     #     log.info("Backlog file present — running Backlog pipeline additionally ...")
#     #     run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))

#     # tdd
#     # if trigger_type != "tdd" and "tdd" in found:
#     #     log.info("TDD file present — running TDD pipeline additionally ...")
#     #     run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
       
#     # ── 10. Upload generated XML(s) to processed bucket ───────────────────
#     uploaded = _upload_xml_outputs(principal)

#     if not uploaded:
#         log.warning("ETL completed but no XML files were produced under output/xml/")
#         auditor.set_lambda_status("no_xml_generated")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "status":     "no_xml_generated",
#         }

#     log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
#     auditor.set_xml_uploads(uploaded)
#     auditor.set_lambda_status("ok")

#     if _direct_invocation:
#         auditor.flush(principal)

#     linelist_info = found.get("linelist", {})
#     return {
#         "statusCode":      200,
#         "principal":       principal,
#         "linelist_used":   linelist_info.get("filename"),
#         "linelist_s3_key": linelist_info.get("key"),
#         "uploaded":        uploaded,
#         "count":           len(uploaded),
#         "status":          "ok",
#     }



"""
lambda_function.py — Lambda 1 Handler
======================================
Function name : map-stibo-inbound-validate-transform-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/*

Flow:
    1. Parse EventBridge event → extract principal (brand) from S3 key
       Key format: raw/metadata/{principal}/{filename}.xlsx

    2. List ALL .xlsx files under raw/metadata/{principal}/

    3. Classify each file into one of 6 types by filename keyword:
         linelist / backlog / tdd / mdd / attributes

    4. If any mandatory type is missing → log and return 200 (wait)

    5. Download all files to /tmp/stibo_workdir/input/{type}/

    6. Call main.run(args) ← ALL existing ETL logic — UNCHANGED

    7. Upload every .xml produced under /tmp/stibo_workdir/output/xml/
       to:
           s3://map-stibo-inbound-processed-dev/
               processed/stepxml/{principal}/{filename}

    8. Return structured response

────────────────────────────────────────────────────────────────────
ENVIRONMENT VARIABLES
────────────────────────────────────────────────────────────────────
    RAW_BUCKET        Source bucket
                      e.g.  map-stibo-inbound-raw-dev

    PROCESSED_BUCKET  Destination bucket (triggers Lambda 2)
                      e.g.  map-stibo-inbound-processed-dev

────────────────────────────────────────────────────────────────────
NOTE — Debug break in main.py
────────────────────────────────────────────────────────────────────
    main.py currently has a `break` inside the article loop (~line 1371)
    that stops after the first matched article. Remove before production.
"""

import json
import logging
import os
import re  
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
# CRITICAL: Set LAMBDA_TMP_DIR BEFORE importing main so that main.py's
# module-level BASE_DIR resolves to /tmp/stibo_workdir instead of
# the Lambda read-only /var/task/ directory.
# ─────────────────────────────────────────────────────────────────────────────
# TMP_WORKDIR = "/tmp/stibo_workdir"
TMP_WORKDIR = "/tmp/stibo_workdir_adidas"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR
from adidas.tdd_main import run_tdd
from adidas.backlog_main import run_backlog
# from adidas.dtb_main import run_dtb
from adidas.dtb_main import run as run_dtb
import adidas.main as etl  

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda1")

# ─────────────────────────────────────────────────────────────────────────────
# AWS client (reused across warm invocations)
# ─────────────────────────────────────────────────────────────────────────────
s3 = boto3.client("s3")

# ─────────────────────────────────────────────────────────────────────────────
# Environment variables — only 2 needed
# ─────────────────────────────────────────────────────────────────────────────
RAW_BUCKET       = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]


# ─────────────────────────────────────────────────────────────────────────────
# Brand name → brand code lookup
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
}

# ─────────────────────────────────────────────────────────────────────────────
# File type identification keywords
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    "linelist":      ["line list", "linelist", "LINELIST"],
    "backlog":       ["backlog", "Backlog"],
    "tdd":           ["tdd", "TDD", "( ros )", "(ros)", "ros"],
    "dtb":           ["dtb", "DTB"],
    "mdd":           ["mdd", "master_data", "master data"],
    "attributes":    ["attributes", "attributes_list", "attribute list"],
    # Licensed brand types
    "recap":         ["recap", "recap sample"],
    "order_sheet":   ["order sheet", "ordersheet", "order_sheet"],
    "ean_source":    ["ean source", "ean_source", "upc source"],
    "packing_list":  ["packing list", "packing_list", "packlist"],
    "oor":           ["oor", "out of ratio", "out_of_ratio"],
    "open_shipment": ["open shipment", "open_shipment", "openshipment"],
}

REQUIRED_TYPES_BY_BRAND = {
    "ADI": {"linelist"},
}
_DEFAULT_REQUIRED = {"linelist"}


# =============================================================================
# HELPERS
# =============================================================================

def _clean_cell(v):
    """Strip illegal XML/Excel characters from string cell values."""
    if not isinstance(v, str):
        return v
    return ''.join(
        c for c in v
        if 0x20 <= ord(c) <= 0xFFFD or c in '\t\n\r'
    )


def _safe_ts_to_epoch(ts) -> int:
    """
    Convert ANY timestamp type (datetime, int, float, str) to Unix epoch seconds.
    Returns 0 on failure — avoids TypeError on mixed-type comparisons.
    """
    try:
        if hasattr(ts, 'timestamp'):            # datetime / aware datetime
            return int(ts.timestamp())
        elif isinstance(ts, (int, float)):
            return int(float(ts))
        elif isinstance(ts, str):
            from datetime import datetime
            dt = datetime.fromisoformat(ts.replace('Z', '+00:00'))
            return int(dt.timestamp())
    except Exception:
        pass
    return 0


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    """Convert a .xlsb to .xlsx — streaming row-by-row to keep memory low."""
    xlsx_path = xlsb_path.with_suffix(".xlsx")
    wb_out = openpyxl.Workbook(write_only=True)
    sheets_written = 0

    with pyxlsb.open_workbook(str(xlsb_path)) as wb:
        for sheet_name in wb.sheets:
            safe_name = ''.join(
                c for c in sheet_name
                if 0x20 <= ord(c) <= 0xFFFD and c not in '\\/?*[]:'
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
                            log.warning("  Skipping bad row in sheet '%s': %s", safe_name, e)
                            continue
                sheets_written += 1
                log.info("  Written sheet: '%s' → '%s'", sheet_name, safe_name)
            except Exception as e:
                log.warning("  Error reading sheet '%s': %s", sheet_name, e)
                continue

    if sheets_written == 0:
        raise ValueError(f"No sheets could be converted from {xlsb_path.name}")

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


def _parse_metadata_from_linelist_filename(dirs: dict, folder_type: str = "linelist") -> dict:
    folder = dirs.get(folder_type)
    if not folder or not folder.exists():
        log.error("%s directory not found: %s", folder_type, folder)
        return {}

    for xlsx_file in folder.glob("*.xlsx"):
        stem = xlsx_file.stem
        log.info("Attempting metadata parse from filename: '%s'", stem)

        parts = re.split(r"\s*-\s*", stem)

        # Season: 2 letters + 2 OR 4 digits (SP26, FW2026, SP2029)
        season     = None
        season_idx = None
        for i, p in enumerate(parts):
            if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
                season     = p.strip().upper()
                season_idx = i
                break

        if not season or season_idx is None or season_idx < 3:  # relaxed from 5 → 3
            log.warning("Filename '%s': no valid season token found — skipping", stem)
            continue

        comp_code  = parts[0].strip()
        sbu        = parts[1].strip()
        brand      = parts[2].strip()
        multi_mono = parts[4].strip() if season_idx > 4 else "Multi"

        # Everything between season and the last all-digit token
        # may include country code(s) like MY, ID, etc.
        trailing = parts[season_idx + 1:]

        # country = first 2-letter uppercase token after season (if not a digit)
        country_code = ""
        for p in trailing:
            p = p.strip()
            if re.match(r'^[A-Z]{2,3}$', p, re.IGNORECASE) and not p.isdigit():
                country_code = p.upper()
                break

        # seq = last all-digit token in trailing
        seq = 1
        for p in reversed(trailing):
            if p.strip().isdigit():
                seq = int(p.strip())
                break

        if not comp_code:
            log.warning("Filename '%s': empty comp_code — skipping", stem)
            continue

        brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())

        # Extract article_type from file_type token (parts[3])
        # e.g. "Line List (MAA Sport) Inline" → article_type = "Inline"
        file_type_token = parts[3].strip() if len(parts) > 3 else ""
        article_type_from_filename = ""
        for at in ["Inline", "License", "SSE"]:
            if at.lower() in file_type_token.lower():
                article_type_from_filename = at
                break

        log.info(
            "Parsed → comp=%s sbu=%s brand=%s brand_code=%s season=%s "
            "seq=%s country=%s article_type=%s",
            comp_code, sbu, brand, brand_code, season,
            seq, country_code, article_type_from_filename,
        )
        return {
            "comp_code":    comp_code,
            "sbu":          sbu,
            "brand":        brand,
            "brand_code":   brand_code,
            "season":       season,
            "seq":          seq,
            "multi_mono":   multi_mono,
            "country_code": country_code,
            "article_type_from_filename": article_type_from_filename,
        }

    log.error("Cannot parse metadata from filename.")
    return {}


def _detect_file_type(filename: str) -> str | None:
    """
    Return the file type for a given filename, or None if unclassifiable.
    Case-insensitive keyword search.
    """
    name = filename.lower()
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw.lower() in name for kw in keywords):
            return ftype
    return None


def _parse_event(event: dict) -> tuple[str, str]:
    """
    Extract bucket name and object key from an EventBridge S3 Object Created event.
    """
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key: str) -> str | None:
    """
    Derive the principal (brand folder) from the S3 object key.
    Key format : raw/metadata/{principal}/{filename}
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not follow raw/metadata/{principal}/{file} pattern: %s", key)
    return None


def _list_principal_files(bucket: str, principal: str,
                           exclude_types: set = None) -> dict[str, dict]:
    """
    Lists files for a brand principal:
    - Brand-specific files from raw/metadata/{principal}/
    - Global files (mdd) from raw/metadata/ root
    Brand-specific files take priority over global if duplicated.

    exclude_types: file types to skip entirely (used to prevent the
                   trigger file's type being overridden by a newer S3 object).
    """
    exclude_types = exclude_types or set()
    GLOBAL_TYPES  = {"mdd"}
    ROOT_TYPES    = {"mdd", "attributes"}
    prefix        = f"raw/metadata/{principal}/"
    paginator     = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    # ── Step 1: Scan brand-specific folder ───────────────────────────────
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            obj_key  = obj["Key"]                    # ← renamed from 'key'
            filename = Path(obj_key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue

            ftype = _detect_file_type(filename)
            if ftype is None:
                log.warning("  Unrecognised file (no type match): %s — skipping", filename)
                continue
            if ftype in exclude_types:               # ← skip trigger type
                log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
                continue

            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {
                    "key":           obj_key,
                    "filename":      filename,
                    "last_modified": last_modified,
                }
                log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

    # ── Step 2: Scan global folder raw/metadata/ for mdd ────────
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        exclude_types=exclude_types,
        log=log,
    )

    global_prefix = "raw/metadata/"
    for page in paginator.paginate(Bucket=bucket, Prefix=global_prefix, Delimiter="/"):
        for obj in page.get("Contents", []):
            obj_key  = obj["Key"]                    # ← renamed from 'key'
            filename = Path(obj_key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsb", ".csv")):
                continue

            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES:
                continue
            if ftype in exclude_types:               # ← skip trigger type
                continue

            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {
                    "key":           obj_key,
                    "filename":      filename,
                    "last_modified": last_modified,
                }
                log.info("  Classified  %-12s ← %s  [global]", ftype, filename)
            else:
                log.info(
                    "  Skipping global %s — brand-specific version already found: %s",
                    ftype, found[ftype]["filename"],
                )

    return found
def _prepare_tmp_dirs() -> dict[str, Path]:
    """
    Recreate /tmp/stibo_workdir/input/{type}/ sub-folders from scratch.
    Also clears output/xml/ so only current run's XMLs are uploaded.
    Returns {file_type: local_dir_path} for all input types.
    """
    base = Path(TMP_WORKDIR)

    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    # All possible input type dirs (standard + licensed brand types)
    dir_types = [
        "linelist", "backlog", "tdd", "dtb", "mdd", "attributes",
        "recap", "order_sheet", "ean_source", "packing_list", "oor", "open_shipment",
    ]
    dirs: dict[str, Path] = {}
    for dtype in dir_types:
        d = base / "input" / dtype
        d.mkdir(parents=True, exist_ok=True)
        dirs[dtype] = d

    return dirs


def _download_files(
    bucket: str,
    found: dict[str, dict],
    dirs: dict[str, Path],
    auditor: AuditLogger,
) -> None:
    """
    Download each identified S3 file to its corresponding local sub-folder.
    Records each download outcome into the auditor.
    Raises on download failure so the caller can handle / surface the error.
    """
    for ftype, info in found.items():
        if ftype not in dirs:
            log.warning("  No local dir mapped for type '%s' — skipping download", ftype)
            continue

        local_path = dirs[ftype] / info["filename"]
        log.info("  Downloading %-12s ← s3://%s/%s", ftype, bucket, info["key"])
        try:
            s3.download_file(bucket, info["key"], str(local_path))
            log.info("  Saved → %s", local_path)
            auditor.record_download(ftype, info["filename"], status="ok")
        except Exception as exc:
            log.error("  Download FAILED for %s: %s", info["key"], exc)
            auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
            raise


def _upload_xml_outputs(principal: str) -> list[str]:
    """
    Find all .xml files under output/xml/ and upload to processed bucket.
    Returns list of S3 keys successfully uploaded.
    """
    xml_dir = Path(TMP_WORKDIR) / "output" / "xml"
    if not xml_dir.exists():
        return []

    uploaded = []
    for xml_file in xml_dir.glob("*.xml"):
        key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading XML → s3://%s/%s", PROCESSED_BUCKET, key)
        try:
            s3.upload_file(str(xml_file), PROCESSED_BUCKET, key)
            uploaded.append(key)
        except Exception as e:
            log.error("  Upload failed for %s: %s", xml_file.name, e)
    return uploaded

# =============================================================================
# LAMBDA HANDLER
# =============================================================================

def lambda_handler(event, context, auditor: AuditLogger = None):
    """
    Main Lambda entry point.
    auditor is created and flushed by router.py.
    If invoked directly (not via router), creates its own auditor as fallback.
    """
    # ── Fallback: direct invocation (not via router.py) ───────────────────
    if auditor is None:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "",
            routing_decision = "lambda_handler invoked directly without router",
        )
        _direct_invocation = True
    else:
        _direct_invocation = False

    log.info("Event received: %s", json.dumps(event))

    # ── 1. Parse EventBridge event ────────────────────────────────────────
    bucket, key = _parse_event(event)

    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    # Capture trigger info NOW before any S3 listing corrupts 'key'
    trigger_filename = Path(key).name
    trigger_type     = _detect_file_type(trigger_filename)
    log.info("Trigger file: '%s'  type: '%s'", trigger_filename, trigger_type)

    # ── 2. Guard: only process keys under raw/metadata/ ──────────────────
    if not key.startswith("raw/metadata/"):
        log.warning("Key is outside raw/metadata/ — ignoring: %s", key)
        auditor.set_lambda_status("skipped_outside_prefix")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 200, "skipped": True, "reason": "outside_prefix"}

    # ── 3. Extract principal (brand folder name) from key ─────────────────
    principal = _extract_principal(key)
    if not principal:
        log.error("Could not extract principal from key: %s", key)
        auditor.set_lambda_status("error", error="Cannot determine principal")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    log.info("Principal (brand folder): %s", principal)

    # ── 4. List + classify all files for this principal ───────────────────
    log.info("Listing files under raw/metadata/%s/ ...", principal)
    # exclude_types: don't let the listing override the actual trigger file
    found = _list_principal_files(bucket, principal,
                                exclude_types={trigger_type} if trigger_type else set())

    # Always pin the trigger file — it is the authoritative source for its type
    if trigger_type:
        found[trigger_type] = {
            "key":           key,
            "filename":      trigger_filename,
            "last_modified": None,
        }
        log.info("Pinned trigger file → type='%s'  file='%s'", trigger_type, trigger_filename)

    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    # ── 5a. DTB trigger: self-contained early exit (doesn't need linelist) ──
    _trigger_fname = Path(key).name
    _trigger_type_dtb = (
        _detect_file_type(_trigger_fname)
        if not _trigger_fname.startswith(".__")
        else None
    )

    if _trigger_type_dtb == "dtb":
        if "dtb" not in found:
            log.info("DTB file not yet present for '%s' — waiting.", principal)
            auditor.set_files_missing(["dtb"])
            auditor.set_lambda_status("waiting_for_mandatory_files")
            if _direct_invocation:
                auditor.flush(principal)
            return {
                "statusCode": 200,
                "principal":  principal,
                "present":    sorted(found.keys()),
                "missing":    ["dtb"],
                "status":     "waiting_for_mandatory_files",
            }

        dirs = _prepare_tmp_dirs()
        try:
            _download_files(bucket, found, dirs, auditor)
        except Exception as exc:
            auditor.set_lambda_status("error", error=f"Download failed: {exc}")
            if _direct_invocation:
                auditor.flush(principal)
            return {"statusCode": 500, "principal": principal, "error": f"Download failed: {exc}"}

        for type_dir in dirs.values():
            for xlsb_file in type_dir.glob("*.xlsb"):
                try:
                    _convert_xlsb_to_xlsx(xlsb_file)
                    auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok")
                except Exception as e:
                    auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="failed", error=str(e))
            for csv_file in type_dir.glob("*.csv"):
                try:
                    _convert_csv_to_xlsx(csv_file)
                    auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok")
                except Exception as e:
                    auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="failed", error=str(e))

        meta = _parse_metadata_from_linelist_filename(dirs, folder_type="dtb")
        if not meta:
            auditor.set_lambda_status("error", error="Unparseable metadata in DTB filename")
            if _direct_invocation:
                auditor.flush(principal)
            return {
                "statusCode": 400,
                "principal":  principal,
                "error":      "Unparseable metadata in DTB filename. Ensure: {CompCode}-{SBU}-{Brand}-DTB-{Multi/Mono}-{Season}-{Seq}.xlsx",
            }

        auditor.set_metadata_parsed(meta)
        dtb_args = types.SimpleNamespace(
            brand=meta["brand"],
            brand_code=meta["brand_code"],
            comp_code=meta["comp_code"],
            sbu=meta["sbu"],
            season=meta["season"],
            seq=meta["seq"],
            multi_mono=meta.get("multi_mono", "Multi"),
            country_code=meta.get("country_code", ""),  # ← ADD THIS LINE
            article_type_from_filename=meta.get("article_type_from_filename", ""),
        )

        log.info("DTB trigger — running DTB ecommerce enrichment pipeline ...")
        run_dtb(dtb_args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))

        uploaded = _upload_xml_outputs(principal)
        if not uploaded:
            auditor.set_lambda_status("no_xml_generated")
            if _direct_invocation:
                auditor.flush(principal)
            return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

        auditor.set_xml_uploads(uploaded)
        auditor.set_lambda_status("ok")
        if _direct_invocation:
            auditor.flush(principal)
        dtb_info = found.get("dtb", {})
        return {
            "statusCode":    200,
            "principal":     principal,
            "trigger":       "dtb",
            "dtb_file_used": dtb_info.get("filename"),
            "uploaded":      uploaded,
            "count":         len(uploaded),
            "status":        "ok",
        }

    # ── 5b. Early check: linelist must be present before we do anything ────
    if "linelist" not in found:
        log.info("Linelist not yet present for '%s' — waiting.", principal)
        auditor.set_files_missing(["linelist"])
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    ["linelist"],
            "status":     "waiting_for_mandatory_files",
        }

    # ── 6. Prepare clean /tmp/ input + output directories ─────────────────
    dirs = _prepare_tmp_dirs()

    # ── 7. Download all files into their sub-folders ──────────────────────
    log.info("Downloading %d input files (%s) ...", len(found), sorted(found.keys()))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal, "error": f"Download failed: {exc}"}

    # ── 7b. Convert .xlsb and .csv files to .xlsx ─────────────────────────
    log.info("Converting .xlsb and .csv files to .xlsx ...")
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok")
            except Exception as e:
                log.error("Failed to convert %s: %s", xlsb_file.name, e)
                auditor.record_conversion(xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx",
                                          status="failed", error=str(e))
                auditor.set_lambda_status("error", error=f"xlsb conversion failed: {e}")
                if _direct_invocation:
                    auditor.flush(principal)
                return {"statusCode": 500, "principal": principal,
                        "error": f"xlsb conversion failed for {xlsb_file.name}: {e}"}

        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok")
            except Exception as e:
                log.error("Failed to convert %s: %s", csv_file.name, e)
                auditor.record_conversion(csv_file.name, from_ext=".csv", to_ext=".xlsx",
                                          status="failed", error=str(e))
                auditor.set_lambda_status("error", error=f"csv conversion failed: {e}")
                if _direct_invocation:
                    auditor.flush(principal)
                return {"statusCode": 500, "principal": principal,
                        "error": f"csv conversion failed for {csv_file.name}: {e}"}

    # ── 8. Parse pipeline config from linelist filename ───────────────────
    meta = _parse_metadata_from_linelist_filename(dirs)

    if not meta:
        auditor.set_lambda_status("error", error="Unparseable metadata in linelist filename")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 400,
            "principal":  principal,
            "error": (
                "Unparseable metadata in linelist filename. "
                "Ensure filename follows: "
                "{CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Seq}.xlsx "
                "e.g. 0888-SP-Adidas-Line List-Multi-FW26-1.xlsx"
            ),
        }

    auditor.set_metadata_parsed(meta)

    # ── 8b. Brand-specific mandatory file check ───────────────────────────
    required_types    = REQUIRED_TYPES_BY_BRAND.get(meta["brand_code"], _DEFAULT_REQUIRED)
    missing_mandatory = required_types - set(found.keys())

    optional_present = [t for t in ("backlog", "tdd") if t in found]
    optional_missing = [t for t in ("backlog", "tdd") if t not in found]

    if missing_mandatory:
        log.info(
            "Mandatory files not yet present for '%s' (%s) — waiting for: %s",
            principal, meta["brand_code"], sorted(missing_mandatory),
        )
        auditor.set_files_missing(sorted(missing_mandatory))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    sorted(missing_mandatory),
            "status":     "waiting_for_mandatory_files",
        }

    log.info(
        "Mandatory files confirmed for '%s' (%s) — starting ETL. "
        "Optional present: %s | Optional missing (empty stubs): %s",
        principal, meta["brand_code"], optional_present, optional_missing,
    )

    args = types.SimpleNamespace(
        brand        = meta["brand"],
        brand_code   = meta["brand_code"],
        comp_code    = meta["comp_code"],
        sbu          = meta["sbu"],
        season       = meta["season"],
        seq          = meta["seq"],
        multi_mono   = meta.get("multi_mono", "Multi"),
        country_code = meta.get("country_code", principal.upper()),  # ← was principal.upper()
        article_type_from_filename = meta.get("article_type_from_filename", ""),  # ← NEW
    )

    # ── 9. Run ETL pipeline — ALL main.py logic UNCHANGED ─────────────────
    # log.info("Running ETL pipeline (main.run) ...")
        
    # ── 9. Run Pipeline Logic ─────────────────────────────────────────────
    log.info("Routing to pipeline — trigger_type: '%s'", trigger_type)

    # if trigger_type == "linelist":
    #     log.info("Linelist trigger - running ETL pipeline ...")
    #     etl.run(args, auditor=auditor)
    # elif trigger_type == "tdd":
    #     log.info("TDD trigger - running TDD pipeline ...")
    #     run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
    # elif trigger_type == "backlog":
    #     log.info("Backlog trigger - running Backlog pipeline ...")
    #     run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
    # else:
    #     log.info("Default trigger (%s) - running ETL pipeline ...", trigger_type)
    #     etl.run(args, auditor=auditor)

    if trigger_type == "tdd":
        log.info("TDD trigger - running TDD pipeline ...")
        run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
    elif trigger_type == "backlog":
        log.info("Backlog trigger - running Backlog pipeline ...")
        run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
    else:
        
        log.info("ETL trigger (%s) - running ETL pipeline ...", trigger_type)
        etl.run(args, auditor=auditor)

    # backlog
    # if trigger_type != "backlog" and "backlog" in found:
    #     log.info("Backlog file present — running Backlog pipeline additionally ...")
    #     run_backlog(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))

    # tdd
    # if trigger_type != "tdd" and "tdd" in found:
    #     log.info("TDD file present — running TDD pipeline additionally ...")
    #     run_tdd(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))
       
    # ── 10. Upload generated XML(s) to processed bucket ───────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("ETL completed but no XML files were produced under output/xml/")
        auditor.set_lambda_status("no_xml_generated")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "status":     "no_xml_generated",
        }

    log.info("ETL complete. %d XML file(s) uploaded: %s", len(uploaded), uploaded)
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")

    if _direct_invocation:
        auditor.flush(principal)

    linelist_info = found.get("linelist", {})
    return {
        "statusCode":      200,
        "principal":       principal,
        "linelist_used":   linelist_info.get("filename"),
        "linelist_s3_key": linelist_info.get("key"),
        "uploaded":        uploaded,
        "count":           len(uploaded),
        "status":          "ok",
    }
