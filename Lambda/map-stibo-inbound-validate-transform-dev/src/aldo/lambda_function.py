# """
# aldo/lambda_function.py
# ========================
# Aldo brand handler — same structure as adidas/lambda_function.py
# Source: TrenzaShop (.xlsm) + OOR Report + OOR Size Run + MCR (.xlsx)

# File type detection for Aldo:
#   - linelist  : "trenzashop", "line list", "linelist"
#   - oor       : "wsl001", "oor report"  (but NOT "size run")
#   - oor_size  : "oor_size", "oor size", "size run"
#   - mcr       : "wsl004", "mcr", "material composition"
#   - mdd       : "mdd", "master data"
#   - attributes: "attributes"
#   - naming    : "naming"
#   - ecommerce : "ecommerce file", "ecommerce"

# PARAMS sheet (inside TrenzaShop .xlsm) layout:
#   Row 0: DB connection string
#   Row 1: catalog_code  e.g. "ALDLFTWSPRING1S26"
#   Row 2: comp_code     e.g. "200782"
#   Row 3: username
#   Row 4: division      e.g. "01"
#   Row 5: season_id     e.g. "69"
#   Row 6: seq           e.g. "1"

# article_filter usage (single-product mode):
#   Pass "article_filter" in the event payload to process only one article.
#   Example event:
#     {
#       "detail": { "bucket": {...}, "object": {...} },
#       "article_filter": "123456"
#     }
#   If omitted or null, all articles are processed (normal mode).
# """

# import json
# import logging
# import os
# import re
# import shutil
# import sys
# import types
# from datetime import datetime
# from pathlib import Path
# from urllib.parse import unquote_plus

# import boto3
# import openpyxl

# from audit_logger import AuditLogger
# from metadata_s3 import add_root_metadata_files

# TMP_WORKDIR = "/tmp/stibo_workdir"
# os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# import aldo.main as etl                  # noqa: E402
# import aldo.oor as oor_etl               # noqa: E402
# import aldo.ecommerce as ecommerce_etl   # noqa: E402
# import aldo.oor as oor_etl  

# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s  %(levelname)-8s  %(message)s",
#     handlers=[logging.StreamHandler(sys.stdout)],
# )
# log = logging.getLogger("aldo_lambda")

# s3 = boto3.client("s3")

# RAW_BUCKET       = os.environ["RAW_BUCKET"]
# PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# # ── Aldo specific file type keywords ──────────────────────────────
# FILE_TYPE_KEYWORDS = {
#     "linelist":   ["trenzashop", "linelist", "line list"],
#     "oor_size":   ["oor_size", "oor size run", "size_run", "size run"],
#     "oor":        ["wsl001", "oor report", "oor_report", "oor-report", "oor inline", "oor maa sport inline", "oor maa sport"],
#     "mcr":        ["wsl004", "mcr", "material composition", "material_composition"],
#     "mdd":        ["mdd", "master_data", "master data"],
#     "attributes": ["attributes", "attribute list"],
#     "naming":     ["naming", "naming_convention", "naming convention"],
#     "ecommerce":  ["ecommerce file", "ecommerce"],
# }

# # Only TrenzaShop is mandatory — other files are optional enrichment
# # Only TrenzaShop is mandatory — other files are optional enrichment
# REQUIRED_TYPES = {"linelist"}

# REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
#     "linelist":  {"linelist"},
#     "oor":       {"oor"},
#     "oor_size":  {"oor_size"},
#     "mcr":       {"mcr"},
#     "ecommerce": {"ecommerce"},
# }

# PROCESS_TYPES_BY_TRIGGER: dict[str, set[str]] = {
#     "linelist":  {"linelist"},
#     "oor":       {"oor"},
#     "oor_size":  {"oor_size"},
#     "mcr":       {"mcr"},
#     "ecommerce": {"ecommerce"},
# }

# DEFAULT_SBU = os.environ.get("ALDO_SBU", "FQ")


# def _detect_file_type(filename: str):
#     name = filename.lower()
#     normalized = re.sub(r"[^a-z0-9]+", " ", name).strip()

#     # Legacy ALDO naming: some OOR files are delivered as
#     # "... Line list MAA Sport Inline[-Multi] ...".
#     # Treat "MAA Sport Inline" as OOR family even if "line list" appears.
#     if "maa sport" in normalized:
#         if "size run" in normalized or "size_run" in normalized:
#             return "oor_size"
#         if "oor" in normalized:
#             return "oor"
#         # Fall through to normal keyword detection (will match "line list" → linelist)

#     # Order matters: oor_size before oor, ecommerce last to avoid false matches
#     for ftype in ["oor_size", "oor", "linelist", "mcr", "mdd", "attributes", "naming", "ecommerce"]:
#         keywords = FILE_TYPE_KEYWORDS.get(ftype, [])
#         if any(re.sub(r"[^a-z0-9]+", " ", kw.lower()).strip() in normalized for kw in keywords):
#             return ftype
#     return None


# def _detect_triggered_file_type(key: str) -> str | None:
#     """Derive file type from the S3 key that triggered this invocation."""
#     filename = Path(key).name
#     if filename.startswith(".__"):
#         return None
#     return _detect_file_type(filename)

# def _parse_event(event):
#     detail = event.get("detail", {})
#     bucket = detail.get("bucket", {}).get("name", "")
#     key    = detail.get("object",  {}).get("key",  "")
#     return bucket, str(key)


# def _extract_principal(key):
#     parts = key.strip("/").split("/")
#     if len(parts) >= 4:
#         return parts[2]
#     return None


# def _list_principal_files(bucket, principal, exclude_types=None):
#     exclude_types = exclude_types or set()
#     prefix    = f"raw/metadata/{principal}/"
#     paginator = s3.get_paginator("list_objects_v2")
#     found     = {}

#     for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
#         for obj in page.get("Contents", []):
#             key      = obj["Key"]
#             filename = Path(key).name
#             if not filename:
#                 continue
#             if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
#                 continue
#             ftype = _detect_file_type(filename)
#             if ftype is None:
#                 log.warning("Unrecognised file: %s — skipping", filename)
#                 continue
#             if ftype in exclude_types:          # ← NEW
#                 log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
#                 continue
#             last_modified = obj["LastModified"]
#             if ftype not in found or last_modified > found[ftype]["last_modified"]:
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified
#                 }
#                 log.info("  Classified %-12s ← %s", ftype, filename)

#     GLOBAL_TYPES  = {"mdd", "naming"}
#     ROOT_TYPES    = {"mdd", "attributes"}
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
#             key      = obj["Key"]
#             filename = Path(key).name
#             if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
#                 continue
#             ftype = _detect_file_type(filename)
#             if ftype not in GLOBAL_TYPES:
#                 continue
#             if ftype in exclude_types:          # ← NEW
#                 continue
#             last_modified = obj["LastModified"]
#             if ftype not in found:
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified
#                 }
#                 log.info("  Classified %-12s ← %s [global]", ftype, filename)

#     return found


# def _prepare_tmp_dirs():
#     base       = Path(TMP_WORKDIR)
#     input_base = base / "input"
#     if input_base.exists():
#         shutil.rmtree(str(input_base))
#     xml_out = base / "output" / "xml"
#     if xml_out.exists():
#         shutil.rmtree(str(xml_out))
#     xml_out.mkdir(parents=True, exist_ok=True)

#     dirs = {
#         "linelist":   base / "input" / "linelist",
#         "oor":        base / "input" / "oor",
#         "oor_size":   base / "input" / "oor_size",
#         "mcr":        base / "input" / "mcr",
#         "mdd":        base / "input" / "mdd",
#         "attributes": base / "input" / "attributes",
#         "naming":     base / "input" / "naming",
#         "ecommerce":  base / "input" / "ecommerce",
#     }
#     for d in dirs.values():
#         d.mkdir(parents=True, exist_ok=True)
#     return dirs


# def _download_files(bucket, found, dirs, auditor):
#     for ftype, info in found.items():
#         if ftype not in dirs:
#             continue
#         local_path = dirs[ftype] / info["filename"]
#         log.info("  Downloading %-12s ← s3://%s/%s", ftype, bucket, info["key"])
#         try:
#             s3.download_file(bucket, info["key"], str(local_path))
#             auditor.record_download(ftype, info["filename"], status="ok")
#         except Exception as exc:
#             log.error("  Download FAILED: %s", exc)
#             auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
#             raise


# def _upload_xml_outputs(principal):
#     xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
#     uploaded = []
#     for xml_file in xml_dir.glob("*.xml"):
#         s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
#         log.info("  Uploading → s3://%s/%s", PROCESSED_BUCKET, s3_key)
#         s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
#         uploaded.append(s3_key)
#         log.info("  Uploaded: %s", s3_key)
#     return uploaded


# def _parse_args_from_trenzashop(linelist_dir: Path, args):
#     """
#     Read comp_code, season, seq directly from the TrenzaShop .xlsm PARAMS sheet.
#     TrenzaShop is self-contained — naming file is NOT used for Aldo.
#     """
#     xlsm_files = list(linelist_dir.glob("*.xlsm")) + list(linelist_dir.glob("*.xlsx"))
#     if not xlsm_files:
#         log.warning("_parse_args_from_trenzashop: no file found in %s", linelist_dir)
#         return

#     path = xlsm_files[0]
#     log.info("Parsing args from TrenzaShop file: %s", path.name)

#     try:
#         wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
#     except Exception as exc:
#         log.warning("Cannot open workbook for args parsing: %s", exc)
#         return

#     try:
#         if "PARAMS" not in wb.sheetnames:
#             log.warning("PARAMS sheet not found — using defaults")
#             wb.close()
#             return

#         params_rows = list(wb["PARAMS"].iter_rows(values_only=True))

#         def _param(idx):
#             if idx < len(params_rows) and params_rows[idx][0] is not None:
#                 return str(params_rows[idx][0]).strip()
#             return ""

#         comp_code = _param(2)
#         season_id = _param(5)
#         seq_raw   = _param(6)

#         if comp_code:
#             args.comp_code = comp_code
#             log.info("  comp_code from PARAMS: %s", comp_code)

#         if seq_raw:
#             try:
#                 args.seq = int(seq_raw)
#             except ValueError:
#                 pass

#         if season_id and "SellingSeason" in wb.sheetnames:
#             season_name = ""
#             for row in wb["SellingSeason"].iter_rows(values_only=True):
#                 if row[0] and f"[{season_id}]" in str(row[0]):
#                     season_name = str(row[0]).strip()
#                     break

#             if season_name:
#                 log.info("  Season name from SellingSeason: %s", season_name)
#                 year_match = re.search(r'20(\d{2})', season_name)
#                 season_yy  = year_match.group(1) if year_match else ""
#                 ss = season_name.upper()
#                 if "SPRING" in ss or " SS" in ss:
#                     season_code = "SS"
#                 elif "FALL" in ss or "FW" in ss or "FL" in ss or "AUTUMN" in ss or "AW" in ss:
#                     season_code = "FL"   # preserve FL if that's how Aldo names it
#                 else:
#                     season_code = "SS"
#                 if season_yy:
#                     args.season = f"{season_code}{season_yy}"
#                     log.info("  season from SellingSeason: %s", args.season)
#             else:
#                 log.warning("  season_id [%s] not found in SellingSeason", season_id)

#     except Exception as exc:
#         log.warning("Error parsing args from TrenzaShop: %s", exc)
#     finally:
#         wb.close()


# def _parse_metadata_from_linelist_filename(linelist_dir: Path) -> dict:
#     """Extract article type, season, and country from linelist filename."""
#     files = list(linelist_dir.glob("*.xlsm")) + list(linelist_dir.glob("*.xlsx"))
#     if not files:
#         return {}

#     stem = files[0].stem
#     upper = stem.upper()

#     article_type = ""
#     if "LICENSE" in upper or "LICENSED" in upper:
#         article_type = "License"
#     elif "INLINE" in upper:
#         article_type = "Inline"

#     season = ""
#     m = re.search(r"(?<![A-Z])([A-Z]{2})(20\d{2}|\d{2})(?![0-9A-Z])", upper)
#     if m:
#         raw_code = m.group(1)   # preserve raw: FL, SS, FW, etc.
#         yy = m.group(2)[-2:]
#         season = f"{raw_code}{yy}"   # e.g. FL29, SS27, FW28

#     country = ""
#     parts = [p for p in re.split(r"[-_\s]+", upper) if p]
#     for token in reversed(parts):
#         if re.fullmatch(r"[A-Z]{2}", token) and not re.search(r"(?<![A-Z])" + token + r"(20\d{2}|\d{2})(?![0-9A-Z])", upper):
#             country = token
#             break

#     return {
#         "article_type_from_filename": article_type,
#         "season": season,
#         "country_code": country,
#     }


# def lambda_handler(event, context, auditor=None):
#     if auditor is None:
#         auditor = AuditLogger(event=event, context=context)
#         auditor.set_router_context(
#             trigger_type="direct_invocation",
#             original_key="",
#             detected_brand="aldo",
#             routing_decision="aldo lambda_handler direct",
#         )
#         _direct = True
#     else:
#         _direct = False


#     # ── article_filter: optional single-product mode ──────────────
#     article_filter = event.get("article_filter") or None
#     if article_filter:
#         article_filter = str(article_filter).strip()
#         log.info("Single-article mode: article_filter=%s", article_filter)

#     bucket, key = _parse_event(event)
#     if not bucket or not key:
#         auditor.set_lambda_status("error", error="Unparseable event")
#         return {"statusCode": 400, "error": "Unparseable event"}

#     key = unquote_plus(key)
#     if not key.startswith("raw/metadata/"):
#         auditor.set_lambda_status("skipped_outside_prefix")
#         return {"statusCode": 200, "skipped": True}

#     principal = _extract_principal(key)
#     if not principal:
#         auditor.set_lambda_status("error", error="Cannot determine principal")
#         return {"statusCode": 400, "error": "Cannot determine principal"}

#     log.info("Principal: %s", principal)

#     # ── Detect triggered file type ────────────────────────────────
#     triggered_type = _detect_triggered_file_type(key)
#     log.info("Triggered file type: %s", triggered_type or "unknown")

#     # ── List files, excluding the triggered type (we pin it below) ─
#     found = _list_principal_files(bucket, principal,
#                                 exclude_types={triggered_type} if triggered_type else set())

#     # ── Pin the trigger file as authoritative for its type ────────
#     trigger_filename = Path(key).name
#     if triggered_type and not trigger_filename.startswith(".__"):
#         found[triggered_type] = {
#             "key":           key,
#             "filename":      trigger_filename,
#             "last_modified": None,
#         }
#         log.info("Pinned trigger file → type='%s'  file='%s'", triggered_type, trigger_filename)

#     log.info("Files found: %s", {k: v["filename"] for k, v in found.items()})
#     auditor.set_files_classified(found)

#     # ── Check mandatory files scoped to triggered type ────────────
#     required_types = REQUIRED_TYPES_BY_TRIGGER.get(triggered_type, REQUIRED_TYPES)
#     missing = required_types - set(found.keys())
#     if missing:
#         log.info("Waiting for: %s", missing)
#         auditor.set_files_missing(sorted(missing))
#         auditor.set_lambda_status("waiting_for_mandatory_files")
#         if _direct:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "missing":    sorted(missing),
#             "status":     "waiting_for_mandatory_files",
#         }

#     dirs = _prepare_tmp_dirs()
#     _download_files(bucket, found, dirs, auditor)

#     # ── Build args: safe defaults ─────────────────────────────────
#     args = types.SimpleNamespace(
#         brand      = "Aldo",
#         brand_code = "ALD",
#         comp_code  = "200782",
#         sbu        = DEFAULT_SBU,
#         season     = "SS26",
#         seq        = 1,
#     )

#     # ── Parse from TrenzaShop PARAMS sheet (only for linelist trigger) ─
#     if triggered_type == "linelist":
#         _parse_args_from_trenzashop(dirs["linelist"], args)
#         filename_meta = _parse_metadata_from_linelist_filename(dirs["linelist"])
#     # After
#     elif triggered_type in dirs:
#         filename_meta = _parse_metadata_from_linelist_filename(dirs[triggered_type])
#     else:
#         filename_meta = {}

#     if filename_meta.get("season"):
#         args.season = filename_meta["season"]
#     else:
#         # Fallback: extract raw season token directly from filename — no whitelist
#         stem = Path(key).stem.upper()
#         m = re.search(r"([A-Z]{2})(20\d{2}|\d{2})(?=[-_\s.]|$)", stem)
#         if m:
#             raw_code = m.group(1)
#             year = m.group(2)
#             if len(year) == 2:
#                 year = f"20{year}"
#             args.season = f"{raw_code}{year}"
#             log.info("Season extracted from filename (fallback): %s", args.season)
#     args.article_type_from_filename = filename_meta.get("article_type_from_filename", "")
#     args.country_code = filename_meta.get("country_code", "")

#     log.info("Running Aldo ETL — args: %s  article_filter: %s",
#              vars(args), article_filter)
#     auditor.set_metadata_parsed(vars(args))

#     optional_found = [ft for ft in ["oor", "oor_size", "mcr", "ecommerce"] if ft in found]
#     if optional_found:
#         log.info("Optional files to process: %s", optional_found)
#     else:
#         log.info("No optional files (oor/oor_size/mcr/ecommerce) found — TrenzaShop only")

#     # ── 1. Run source-specific ETL scoped to triggered type only ─
#     process_types = PROCESS_TYPES_BY_TRIGGER.get(triggered_type, set(found.keys()))
#     log.info("Running ETL -- triggered_type=%s  process_types=%s", triggered_type, process_types)
#     if triggered_type == "linelist":
#         etl.run(args, auditor=auditor, article_filter=article_filter, process_types=process_types)
#     elif triggered_type in {"oor", "oor_size"}:
#         export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#         season_code = args.season[:2].upper()
#         year_part   = args.season[2:]
#         season_year = f"20{year_part}" if len(year_part) == 2 else year_part
#         season_id   = f"CLH_AOD_{season_code}{season_year}"
#         xml_out_dir = Path(TMP_WORKDIR) / "output" / "xml"

#         written = oor_etl.run(
#             args           = args,
#             xml_out_dir    = xml_out_dir,
#             export_time    = export_time,
#             season_code    = season_code,
#             season_year    = season_year,
#             season_id      = season_id,
#             source_dir     = dirs[triggered_type],
#             source_type    = triggered_type,
#             auditor        = auditor,
#             article_filter = article_filter,
#         )
#         if written:
#             log.info("OOR XML generated.")
#         else:
#             log.info("OOR: no articles written (file may be empty or filtered out).")
#     else:
#         log.info("Main ETL not run for triggered_type=%s", triggered_type)

#     export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#     xml_out_dir = Path(TMP_WORKDIR) / "output" / "xml"
#     season_code = args.season[:2].upper() if len(args.season) >= 2 else "SS"
#     year_part   = args.season[2:]
#     season_year = f"20{year_part}" if len(year_part) == 2 else year_part
#     season_id   = f"CLH_AOD_{season_code}{season_year}"

#     if triggered_type in ("linelist", "oor", "mcr"):
#         log.info("Dispatching → aldo.main (process_types=%s)", process_types)
#         etl.run(args, auditor=auditor, article_filter=article_filter,
#                 process_types=process_types)

#     elif triggered_type == "oor_size":
#         log.info("Dispatching → aldo.oor (OOR Size Run)")
#         oor_etl.run(
#             args          = args,
#             xml_out_dir   = xml_out_dir,
#             export_time   = export_time,
#             season_code   = season_code,
#             season_year   = season_year,
#             season_id     = season_id,
#             oor_size_dir  = dirs["oor_size"],
#             auditor       = auditor,
#             article_filter= article_filter,
#         )

#     elif triggered_type == "ecommerce":
#         log.info("Dispatching → aldo.ecommerce")
#         ecommerce_dir = dirs["ecommerce"]
#         if list(ecommerce_dir.glob("*.xls*")):
#             written = ecommerce_etl.run(
#                 args           = args,
#                 xml_out_dir    = xml_out_dir,
#                 export_time    = export_time,
#                 season_code    = season_code,
#                 season_year    = season_year,
#                 season_id      = season_id,
#                 ecommerce_dir  = ecommerce_dir,
#                 auditor        = auditor,
#                 article_filter = article_filter,
#             )
#             if written:
#                 log.info("✅ Ecommerce XML generated.")
#             else:
#                 log.info("Ecommerce: no articles written.")
#         else:
#             log.warning("Ecommerce triggered but no file found in dir — skipping.")

#     else:
#         log.warning("No ETL dispatcher for triggered_type='%s' — skipping.", triggered_type)

#     # ── Upload all XMLs to S3 ─────────────────────────────────────
#     uploaded = _upload_xml_outputs(principal)

#     if not uploaded:
#         log.warning("No XML produced")
#         auditor.set_lambda_status("no_xml_generated")
#         if _direct:
#             auditor.flush(principal)
#         return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

#     log.info("✅ Done. %d XML(s) uploaded.", len(uploaded))
#     auditor.set_xml_uploads(uploaded)
#     auditor.set_lambda_status("ok")
#     if _direct:
#         auditor.flush(principal)

#     return {
#         "statusCode": 200,
#         "principal":  principal,
#         "uploaded":   uploaded,
#         "count":      len(uploaded),
#         "status":     "ok",
#         **({"article_filter": article_filter} if article_filter else {}),
#     }

"""
aldo/lambda_function.py
========================
Aldo brand handler — same structure as adidas/lambda_function.py
Source: TrenzaShop (.xlsm) + OOR Report + OOR Size Run + MCR (.xlsx)

File type detection for Aldo:
  - linelist  : "trenzashop", "line list", "linelist"
  - oor       : "wsl001", "oor report"  (but NOT "size run")
  - oor_size  : "oor_size", "oor size", "size run"
  - mcr       : "wsl004", "mcr", "material composition"
  - mdd       : "mdd", "master data"
  - attributes: "attributes"
  - ecommerce : "ecommerce file", "ecommerce"

PARAMS sheet (inside TrenzaShop .xlsm) layout:
  Row 0: DB connection string
  Row 1: catalog_code  e.g. "ALDLFTWSPRING1S26"
  Row 2: comp_code     e.g. "200782"
  Row 3: username
  Row 4: division      e.g. "01"
  Row 5: season_id     e.g. "69"
  Row 6: seq           e.g. "1"

article_filter usage (single-product mode):
  Pass "article_filter" in the event payload to process only one article.
  Example event:
    {
      "detail": { "bucket": {...}, "object": {...} },
      "article_filter": "123456"
    }
  If omitted or null, all articles are processed (normal mode).
"""

import json
import logging
import os
import re
import shutil
import sys
import types
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote_plus

import boto3
import openpyxl

from audit_logger import AuditLogger
from metadata_s3 import add_root_metadata_files

TMP_WORKDIR = "/tmp/stibo_workdir"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

import aldo.main as etl                  
import aldo.ecommerce as ecommerce_etl  
import aldo.oor as oor_etl  

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("aldo_lambda")

s3 = boto3.client("s3")

RAW_BUCKET       = os.environ["RAW_BUCKET"]
PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# ── Aldo specific file type keywords ──────────────────────────────
FILE_TYPE_KEYWORDS = {
    "linelist":   ["trenzashop", "linelist", "line list"],
    "oor_size":   ["oor_size", "oor size run", "size_run", "size run"],
    "oor":        ["wsl001", "oor report", "oor_report", "oor-report", "oor inline", "oor maa sport inline", "oor maa sport"],
    "mcr":        ["wsl004", "mcr", "material composition", "material_composition"],
    "mdd":        ["mdd", "master_data", "master data"],
    "attributes": ["attributes", "attribute list"],
    "ecommerce":  ["ecommerce file", "ecommerce"],
}

# Only TrenzaShop is mandatory — other files are optional enrichment
REQUIRED_TYPES = {"linelist"}

REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist":  {"linelist"},
    "oor":       {"oor"},
    "oor_size":  {"oor_size"},
    "mcr":       {"mcr"},
    "ecommerce": {"ecommerce"},
}

PROCESS_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist":  {"linelist"},
    "oor":       {"oor"},
    "oor_size":  {"oor_size"},
    "mcr":       {"mcr"},
    "ecommerce": {"ecommerce"},
}

DEFAULT_SBU = os.environ.get("ALDO_SBU", "FQ")


def _detect_file_type(filename: str):
    name = filename.lower()
    normalized = re.sub(r"[^a-z0-9]+", " ", name).strip()

    # Legacy ALDO filename handling: some OOR files are delivered as
    # "... Line list MAA Sport Inline[-Multi] ...".
    # Treat "MAA Sport Inline" as OOR family even if "line list" appears.
    if "maa sport" in normalized:
        if "size run" in normalized or "size_run" in normalized:
            return "oor_size"
        if "oor" in normalized:
            return "oor"
        # Fall through to normal keyword detection (will match "line list" → linelist)

    # Order matters: oor_size before oor, ecommerce last to avoid false matches
    for ftype in ["oor_size", "oor", "linelist", "mcr", "mdd", "attributes", "ecommerce"]:
        keywords = FILE_TYPE_KEYWORDS.get(ftype, [])
        if any(re.sub(r"[^a-z0-9]+", " ", kw.lower()).strip() in normalized for kw in keywords):
            return ftype
    return None


def _detect_triggered_file_type(key: str) -> str | None:
    """Derive file type from the S3 key that triggered this invocation."""
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)

def _parse_event(event):
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key):
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    return None


def _list_principal_files(bucket, principal, exclude_types=None):
    exclude_types = exclude_types or set()
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found     = {}

    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key      = obj["Key"]
            filename = Path(key).name
            if not filename:
                continue
            if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype is None:
                log.warning("Unrecognised file: %s — skipping", filename)
                continue
            if ftype in exclude_types:          # ← NEW
                log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or last_modified > found[ftype]["last_modified"]:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified
                }
                log.info("  Classified %-12s ← %s", ftype, filename)

    GLOBAL_TYPES  = {"mdd"}
    ROOT_TYPES    = {"mdd", "attributes"}
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
            key      = obj["Key"]
            filename = Path(key).name
            if not filename.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue
            ftype = _detect_file_type(filename)
            if ftype not in GLOBAL_TYPES:
                continue
            if ftype in exclude_types:          # ← NEW
                continue
            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified
                }
                log.info("  Classified %-12s ← %s [global]", ftype, filename)

    return found


def _prepare_tmp_dirs():
    base       = Path(TMP_WORKDIR)
    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))
    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs = {
        "linelist":   base / "input" / "linelist",
        "oor":        base / "input" / "oor",
        "oor_size":   base / "input" / "oor_size",
        "mcr":        base / "input" / "mcr",
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
        "ecommerce":  base / "input" / "ecommerce",
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
            log.error("  Download FAILED: %s", exc)
            auditor.record_download(ftype, info["filename"], status="failed", error=str(exc))
            raise


def _upload_xml_outputs(principal):
    xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
    uploaded = []
    for xml_file in xml_dir.glob("*.xml"):
        s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
        log.info("  Uploading → s3://%s/%s", PROCESSED_BUCKET, s3_key)
        s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
        uploaded.append(s3_key)
        log.info("  Uploaded: %s", s3_key)
    return uploaded


def _parse_args_from_trenzashop(linelist_dir: Path, args):
    """
    Read comp_code, season, seq directly from the TrenzaShop .xlsm PARAMS sheet.
    """
    xlsm_files = list(linelist_dir.glob("*.xlsm")) + list(linelist_dir.glob("*.xlsx"))
    if not xlsm_files:
        log.warning("_parse_args_from_trenzashop: no file found in %s", linelist_dir)
        return

    path = xlsm_files[0]
    log.info("Parsing args from TrenzaShop file: %s", path.name)

    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        log.warning("Cannot open workbook for args parsing: %s", exc)
        return

    try:
        if "PARAMS" not in wb.sheetnames:
            log.warning("PARAMS sheet not found — using defaults")
            wb.close()
            return

        params_rows = list(wb["PARAMS"].iter_rows(values_only=True))

        def _param(idx):
            if idx < len(params_rows) and params_rows[idx][0] is not None:
                return str(params_rows[idx][0]).strip()
            return ""

        comp_code = _param(2)
        season_id = _param(5)
        seq_raw   = _param(6)

        if comp_code:
            args.comp_code = comp_code
            log.info("  comp_code from PARAMS: %s", comp_code)

        if seq_raw:
            try:
                args.seq = int(seq_raw)
            except ValueError:
                pass

        if season_id and "SellingSeason" in wb.sheetnames:
            season_name = ""
            for row in wb["SellingSeason"].iter_rows(values_only=True):
                if row[0] and f"[{season_id}]" in str(row[0]):
                    season_name = str(row[0]).strip()
                    break

            if season_name:
                log.info("  Season name from SellingSeason: %s", season_name)
                year_match = re.search(r'20(\d{2})', season_name)
                season_yy  = year_match.group(1) if year_match else ""
                ss = season_name.upper()
                if "SPRING" in ss or " SS" in ss:
                    season_code = "SS"
                elif "FALL" in ss or "FW" in ss or "FL" in ss or "AUTUMN" in ss or "AW" in ss:
                    season_code = "FL"   # preserve FL if that's how Aldo names it
                else:
                    season_code = "SS"
                if season_yy:
                    args.season = f"{season_code}{season_yy}"
                    log.info("  season from SellingSeason: %s", args.season)
            else:
                log.warning("  season_id [%s] not found in SellingSeason", season_id)

    except Exception as exc:
        log.warning("Error parsing args from TrenzaShop: %s", exc)
    finally:
        wb.close()


def _parse_metadata_from_linelist_filename(linelist_dir: Path) -> dict:
    """Extract article type, season, and country from linelist filename."""
    files = list(linelist_dir.glob("*.xlsm")) + list(linelist_dir.glob("*.xlsx"))
    if not files:
        return {}

    stem = files[0].stem
    upper = stem.upper()

    article_type = ""
    if "LICENSE" in upper or "LICENSED" in upper:
        article_type = "License"
    elif "INLINE" in upper:
        article_type = "Inline"

    season = ""
    m = re.search(r"(?<![A-Z])([A-Z]{2})(20\d{2}|\d{2})(?![0-9A-Z])", upper)
    if m:
        raw_code = m.group(1)   # preserve raw: FL, SS, FW, etc.
        yy = m.group(2)[-2:]
        season = f"{raw_code}{yy}"   # e.g. FL29, SS27, FW28

    country = ""
    parts = [p for p in re.split(r"[-_\s]+", upper) if p]
    for token in reversed(parts):
        if re.fullmatch(r"[A-Z]{2}", token) and not re.search(r"(?<![A-Z])" + token + r"(20\d{2}|\d{2})(?![0-9A-Z])", upper):
            country = token
            break

    return {
        "article_type_from_filename": article_type,
        "season": season,
        "country_code": country,
    }


def lambda_handler(event, context, auditor=None):
    if auditor is None:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type="direct_invocation",
            original_key="",
            detected_brand="aldo",
            routing_decision="aldo lambda_handler direct",
        )
        _direct = True
    else:
        _direct = False



    # ── article_filter: optional single-product mode ──────────────
    article_filter = event.get("article_filter") or None
    if article_filter:
        article_filter = str(article_filter).strip()
        log.info("Single-article mode: article_filter=%s", article_filter)

    bucket, key = _parse_event(event)
    if not bucket or not key:
        auditor.set_lambda_status("error", error="Unparseable event")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    if not key.startswith("raw/metadata/"):
        auditor.set_lambda_status("skipped_outside_prefix")
        return {"statusCode": 200, "skipped": True}

    principal = _extract_principal(key)
    if not principal:
        auditor.set_lambda_status("error", error="Cannot determine principal")
        return {"statusCode": 400, "error": "Cannot determine principal"}

    log.info("Principal: %s", principal)

    # ── Detect triggered file type ────────────────────────────────
    triggered_type = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_type or "unknown")

    # ── List files, excluding the triggered type (we pin it below) ─
    found = _list_principal_files(bucket, principal,
                                exclude_types={triggered_type} if triggered_type else set())

    # ── Pin the trigger file as authoritative for its type ────────
    trigger_filename = Path(key).name
    if triggered_type and not trigger_filename.startswith(".__"):
        found[triggered_type] = {
            "key":           key,
            "filename":      trigger_filename,
            "last_modified": None,
        }
        log.info("Pinned trigger file → type='%s'  file='%s'", triggered_type, trigger_filename)

    log.info("Files found: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    # ── Check mandatory files scoped to triggered type ────────────
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(triggered_type, REQUIRED_TYPES)
    missing = required_types - set(found.keys())
    if missing:
        log.info("Waiting for: %s", missing)
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "missing":    sorted(missing),
            "status":     "waiting_for_mandatory_files",
        }

    dirs = _prepare_tmp_dirs()
    _download_files(bucket, found, dirs, auditor)

    # ── Build args: safe defaults ─────────────────────────────────
    args = types.SimpleNamespace(
        brand      = "Aldo",
        brand_code = "ALD",
        comp_code  = "200782",
        sbu        = DEFAULT_SBU,
        season     = "SS26",
        seq        = 1,
    )

    # ── Parse from TrenzaShop PARAMS sheet (only for linelist trigger) ─
    if triggered_type == "linelist":
        _parse_args_from_trenzashop(dirs["linelist"], args)
        filename_meta = _parse_metadata_from_linelist_filename(dirs["linelist"])
    # After
    elif triggered_type in dirs:
        filename_meta = _parse_metadata_from_linelist_filename(dirs[triggered_type])
    else:
        filename_meta = {}

    if filename_meta.get("season"):
        args.season = filename_meta["season"]
    else:
        # Fallback: extract raw season token directly from filename — no whitelist
        stem = Path(key).stem.upper()
        m = re.search(r"([A-Z]{2})(20\d{2}|\d{2})(?=[-_\s.]|$)", stem)
        if m:
            raw_code = m.group(1)
            year = m.group(2)
            if len(year) == 2:
                year = f"20{year}"
            args.season = f"{raw_code}{year}"
            log.info("Season extracted from filename (fallback): %s", args.season)
    args.article_type_from_filename = filename_meta.get("article_type_from_filename", "")
    args.country_code = filename_meta.get("country_code", "")

    log.info("Running Aldo ETL — args: %s  article_filter: %s",
             vars(args), article_filter)
    auditor.set_metadata_parsed(vars(args))

    optional_found = [ft for ft in ["oor", "oor_size", "mcr", "ecommerce"] if ft in found]
    if optional_found:
        log.info("Optional files to process: %s", optional_found)
    else:
        log.info("No optional files (oor/oor_size/mcr/ecommerce) found — TrenzaShop only")

    # ── 1. Run source-specific ETL scoped to triggered type only ─
    process_types = PROCESS_TYPES_BY_TRIGGER.get(triggered_type, set(found.keys()))
    log.info("Running ETL -- triggered_type=%s  process_types=%s", triggered_type, process_types)
    if triggered_type == "linelist":
        etl.run(args, auditor=auditor, article_filter=article_filter, process_types=process_types)
    elif triggered_type in {"oor", "oor_size"}:
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        season_code = args.season[:2].upper()
        year_part   = args.season[2:]
        season_year = f"20{year_part}" if len(year_part) == 2 else year_part
        season_id   = f"CLH_AOD_{season_code}{season_year}"
        xml_out_dir = Path(TMP_WORKDIR) / "output" / "xml"

        written = oor_etl.run(
            args           = args,
            xml_out_dir    = xml_out_dir,
            export_time    = export_time,
            season_code    = season_code,
            season_year    = season_year,
            season_id      = season_id,
            source_dir     = dirs[triggered_type],
            source_type    = triggered_type,
            auditor        = auditor,
            article_filter = article_filter,
        )
        if written:
            log.info("OOR XML generated.")
        else:
            log.info("OOR: no articles written (file may be empty or filtered out).")
    else:
        log.info("Main ETL not run for triggered_type=%s", triggered_type)

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    xml_out_dir = Path(TMP_WORKDIR) / "output" / "xml"
    season_code = args.season[:2].upper() if len(args.season) >= 2 else "SS"
    year_part   = args.season[2:]
    season_year = f"20{year_part}" if len(year_part) == 2 else year_part
    season_id   = f"CLH_AOD_{season_code}{season_year}"

    if triggered_type in ("linelist", "oor", "mcr"):
        log.info("Dispatching → aldo.main (process_types=%s)", process_types)
        etl.run(args, auditor=auditor, article_filter=article_filter,
                process_types=process_types)

    elif triggered_type == "oor_size":
        log.info("Dispatching → aldo.oor (OOR Size Run)")
        oor_etl.run(
            args          = args,
            xml_out_dir   = xml_out_dir,
            export_time   = export_time,
            season_code   = season_code,
            season_year   = season_year,
            season_id     = season_id,
            oor_size_dir  = dirs["oor_size"],
            auditor       = auditor,
            article_filter= article_filter,
        )

    elif triggered_type == "ecommerce":
        log.info("Dispatching → aldo.ecommerce")
        ecommerce_dir = dirs["ecommerce"]
        if list(ecommerce_dir.glob("*.xls*")):
            written = ecommerce_etl.run(
                args           = args,
                xml_out_dir    = xml_out_dir,
                export_time    = export_time,
                season_code    = season_code,
                season_year    = season_year,
                season_id      = season_id,
                ecommerce_dir  = ecommerce_dir,
                auditor        = auditor,
                article_filter = article_filter,
            )
            if written:
                log.info("✅ Ecommerce XML generated.")
            else:
                log.info("Ecommerce: no articles written.")
        else:
            log.warning("Ecommerce triggered but no file found in dir — skipping.")

    else:
        log.warning("No ETL dispatcher for triggered_type='%s' — skipping.", triggered_type)

    # ── Upload all XMLs to S3 ─────────────────────────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("No XML produced")
        auditor.set_lambda_status("no_xml_generated")
        if _direct:
            auditor.flush(principal)
        return {"statusCode": 200, "principal": principal, "status": "no_xml_generated"}

    log.info("✅ Done. %d XML(s) uploaded.", len(uploaded))
    auditor.set_xml_uploads(uploaded)
    auditor.set_lambda_status("ok")
    if _direct:
        auditor.flush(principal)

    return {
        "statusCode": 200,
        "principal":  principal,
        "uploaded":   uploaded,
        "count":      len(uploaded),
        "status":     "ok",
        **({"article_filter": article_filter} if article_filter else {}),
    }

