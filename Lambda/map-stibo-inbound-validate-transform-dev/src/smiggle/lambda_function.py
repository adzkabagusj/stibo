# """
# lambda_function.py — Lambda Handler for Smiggle
# ================================================
# Function name : map-stibo-inbound-validate-transform-smiggle-dev

# Trigger:
#     EventBridge rule (S3 Object Created) scoped to:
#         s3://map-stibo-inbound-raw-dev/raw/metadata/smiggle/*

# Smiggle-specific flow vs Adidas:
#     • Only ONE input file from the brand  : linelist (wholesale catalogue)
#     • No TDD, no Backlog — these do not exist for Smiggle
#     • mdd + attributes are shared/global files (raw/metadata/ root)
#     • naming convention file is also global
#     • Required types: {"linelist", "mdd", "attributes", "naming"}
#     • Linelist filename keyword detection now also recognises:
#         "catalogue", "catalog", "mapi", "wholesale"
#     • Linelist diff is still run to detect deletions across catalogue versions

# File-type dispatch:
#     The triggered S3 key determines which ETL module is called:
#         linelist  → smiggle.linelist_main.run()
#         packing   → smiggle.packing_list_main.run()
#         ecommerce → smiggle.ecommerce_main.run()

# S3 layout:
#     raw/metadata/smiggle/{catalogue_file}.xlsx   ← Smiggle uploads here
#     raw/metadata/mdd_file.xlsx                   ← shared global MDD
#     raw/metadata/attributes_list.xlsx            ← shared global attr list
#     raw/metadata/naming_convention.xlsx          ← shared global naming

# Output:
#     s3://map-stibo-inbound-processed-dev/
#         processed/stepxml/smiggle/{filename}.xml
# """

# import json
# import logging
# import os
# import shutil
# import sys
# import types
# import pyxlsb
# import openpyxl
# from pathlib import Path
# from urllib.parse import unquote_plus

# import boto3
# import pandas as pd

# from audit_logger import AuditLogger
# from metadata_s3 import add_root_metadata_files

# # ─────────────────────────────────────────────────────────────────────────────
# # Set LAMBDA_TMP_DIR BEFORE importing smiggle modules so that their
# # module-level BASE_DIR resolves to /tmp/stibo_workdir_smiggle
# # (isolated from adidas).
# # ─────────────────────────────────────────────────────────────────────────────
# TMP_WORKDIR = "/tmp/stibo_workdir_smiggle"
# os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# # ── ETL modules — one per Smiggle file type ──────────────────────────────────
# import smiggle.linelist_main      as linelist_etl       # noqa: E402
# import smiggle.packing_list_main  as packing_etl        # noqa: E402
# import smiggle.ecommerce_main     as ecommerce_etl      # noqa: E402

# # ─────────────────────────────────────────────────────────────────────────────
# # Logging
# # ─────────────────────────────────────────────────────────────────────────────
# logging.basicConfig(
#     level=logging.INFO,
#     format="%(asctime)s  %(levelname)-8s  %(message)s",
#     handlers=[logging.StreamHandler(sys.stdout)],
# )
# log = logging.getLogger("lambda_smiggle")

# # ─────────────────────────────────────────────────────────────────────────────
# # AWS client
# # ─────────────────────────────────────────────────────────────────────────────
# s3 = boto3.client("s3")

# # ─────────────────────────────────────────────────────────────────────────────
# # Environment variables
# # ─────────────────────────────────────────────────────────────────────────────
# RAW_BUCKET       = os.environ["RAW_BUCKET"]
# PROCESSED_BUCKET = os.environ["PROCESSED_BUCKET"]

# # ─────────────────────────────────────────────────────────────────────────────
# # Brand lookup
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
# # File type detection
# # Smiggle adds "catalogue", "catalog", "mapi", "wholesale" keywords
# # ─────────────────────────────────────────────────────────────────────────────
# FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
#     # Smiggle primary input (wholesale catalogue / MAPI)
#     "linelist":   [
#         "line list", "linelist",
#         "catalogue", "catalog",
#         "mapi", "wholesale",
#         "handover",
#     ],
#     # Smiggle packing list
#     "packing":    [
#         "packing list", "packinglist", "packing",
#     ],
#     # Smiggle ecommerce file
#     "ecommerce":  [
#         "ecommerce", "e-commerce", "ecom",
#     ],
#     # Shared / global
#     "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
#     "attributes": ["attributes", "attributes_list", "attribute list"],
#     "naming":     ["naming", "naming_convention", "naming convention"],
#     # Not used by Smiggle — kept so this handler can be extended
#     "backlog":    ["backlog"],
#     "tdd":        ["tdd"],
# }

# # ── Required types per triggered file type ────────────────────────────────────
# # Each set lists what must be present in S3 before the ETL can run.
# # mdd + attributes are always required (shared global files).
# REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
#     "linelist":   {"linelist",  "mdd", "attributes", "naming"},
#     "packing":    {"packing",   "mdd", "attributes", "naming"},
#     "ecommerce":  {"ecommerce", "mdd", "attributes", "naming"},
# }

# # Default (used when triggered by a global file fan-out or unknown type)
# REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes", "naming"}

# # Global types fetched from raw/metadata/ root (not inside smiggle/)
# GLOBAL_TYPES: set[str] = {"mdd", "naming"}
# ROOT_TYPES: set[str] = {"mdd", "attributes"}

# # ── ETL dispatcher ────────────────────────────────────────────────────────────
# # Maps detected file type → ETL module that has a run(args, auditor) function.
# ETL_DISPATCHER: dict[str, object] = {
#     "linelist":  linelist_etl,
#     "packing":   packing_etl,
#     "ecommerce": ecommerce_etl,
# }


# # ═════════════════════════════════════════════════════════════════════════════
# # HELPERS
# # ═════════════════════════════════════════════════════════════════════════════

# def _clean_cell(v):
#     """Strip illegal XML / Excel characters from string cell values."""
#     if not isinstance(v, str):
#         return v
#     return "".join(
#         c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r"
#     )


# def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
#     """Convert .xlsb → .xlsx streaming row-by-row."""
#     xlsx_path  = xlsb_path.with_suffix(".xlsx")
#     wb_out     = openpyxl.Workbook(write_only=True)
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
#     """Convert .csv → .xlsx."""
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


# def _safe_ts_to_epoch(ts) -> int:
#     """Convert ANY timestamp type to Unix epoch seconds. Returns 0 on failure."""
#     try:
#         if hasattr(ts, 'timestamp'):
#             return int(ts.timestamp())
#         elif isinstance(ts, (int, float)):
#             return int(float(ts))
#         elif isinstance(ts, str):
#             from datetime import datetime as _dt
#             d = _dt.fromisoformat(ts.replace('Z', '+00:00'))
#             return int(d.timestamp())
#     except Exception:
#         pass
#     return 0

# def _detect_file_type(filename: str) -> str | None:
#     """Return the file type key for a filename, or None if unrecognised."""
#     name = filename.lower()

#     # Ecommerce must win over broad tokens that might also appear in other
#     # file names (for example: "...Ecommerce File ... Inline-Multi...").
#     # This keeps files such as:
#     #   0888-SP-SMIGGLE-Ecommerce File MAA Sport Inline-Multi-FL2027-KH-1.xlsx
#     # routed to smiggle.ecommerce_main.
#     if any(tok in name for tok in ["ecommerce", "e-commerce", "ecom"]):
#         return "ecommerce"

#     for ftype, keywords in FILE_TYPE_KEYWORDS.items():
#         if any(kw in name for kw in keywords):
#             return ftype
#     return None


# def _detect_triggered_file_type(key: str) -> str | None:
#     """
#     Derive the file type from the S3 key that triggered this invocation.

#     raw/metadata/smiggle/0888-SM-Smiggle-Line List-Multi-SS26-1.xlsx
#                           ↑ filename used for keyword detection

#     Returns a key from FILE_TYPE_KEYWORDS (e.g. "linelist", "packing",
#     "ecommerce") or None if the trigger key is a global fan-out placeholder
#     or is otherwise unrecognised.
#     """
#     filename = Path(key).name
#     # Ignore the synthetic placeholder injected by router.py fan-out
#     if filename.startswith(".__"):
#         return None
#     return _detect_file_type(filename)


# def _parse_event(event: dict) -> tuple[str, str]:
#     """Extract (bucket, key) from an EventBridge S3 Object Created event."""
#     detail = event.get("detail", {})
#     bucket = detail.get("bucket", {}).get("name", "")
#     key    = detail.get("object",  {}).get("key",  "")
#     return bucket, str(key)


# def _extract_principal(key: str) -> str | None:
#     """
#     Derive principal from key:  raw/metadata/{principal}/{filename}
#     Returns the principal (e.g. "smiggle") or None.
#     """
#     parts = key.strip("/").split("/")
#     if len(parts) >= 4:
#         return parts[2]
#     log.warning("Key does not follow raw/metadata/{principal}/{file}: %s", key)
#     return None


# def _list_principal_files(bucket: str, principal: str,
#                           exclude_types: set = None) -> dict[str, dict]:
#     """
#     Scan S3 for all recognised files:
#       1. Brand-specific files  ← raw/metadata/{principal}/
#       2. Global shared files   ← raw/metadata/           (mdd, attributes, naming only)

#     Brand-specific takes priority if both exist for the same type.
#     """
#     # ADD as first line inside the function (after the docstring)
#     exclude_types = exclude_types or set()
#     prefix    = f"raw/metadata/{principal}/"
#     paginator = s3.get_paginator("list_objects_v2")
#     found: dict[str, dict] = {}

#     # ── Step 1: brand-specific folder ──────────────────────────
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
#                 log.warning("  Unrecognised file — skipping: %s", filename)
#                 continue
#             if ftype in exclude_types:
#                 log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
#                 continue   
#             last_modified = obj["LastModified"]
#             if ftype not in found or (
#                 _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
#             ):
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified
#                 }
#                 log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

#     # ── Step 2: global folder (mdd, attributes, naming only) ───
#     add_root_metadata_files(
#         s3_client=s3,
#         bucket=bucket,
#         found=found,
#         include_types=ROOT_TYPES,
#         exclude_types=exclude_types,
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
#             if ftype in exclude_types:
#                 continue
#             last_modified = obj["LastModified"]
#             if ftype not in found:
#                 found[ftype] = {
#                     "key": key, "filename": filename, "last_modified": last_modified
#                 }
#                 log.info("  Classified  %-12s ← %s  [global]", ftype, filename)
#             else:
#                 log.info(
#                     "  Skipping global %s — brand-specific already found: %s",
#                     ftype, found[ftype]["filename"],
#                 )

#     return found


# def _parse_metadata_from_naming_file(naming_dir: Path) -> dict:
#     """
#     Parse comp_code, sbu, brand, brand_code, season, seq from the naming
#     convention file (same logic as Adidas lambda — looks for "For example:"
#     pattern in column D, or a clean filename in column E).
#     """
#     import re

#     xlsx_files = list(naming_dir.glob("*.xlsx"))
#     if not xlsx_files:
#         log.error("No .xlsx in naming dir: %s", naming_dir)
#         return {}

#     naming_file = xlsx_files[0]
#     log.info("Reading metadata from naming file: %s", naming_file.name)

#     try:
#         wb = openpyxl.load_workbook(str(naming_file), read_only=True, data_only=True)
#     except Exception as e:
#         log.error("Cannot open naming file '%s': %s", naming_file.name, e)
#         return {}

#     sheet_name = "Stibo" if "Stibo" in wb.sheetnames else wb.sheetnames[0]
#     ws = wb[sheet_name]
#     log.info("  Using sheet: '%s'", sheet_name)

#     for row in ws.iter_rows(values_only=True):
#         if not row or len(row) < 4:
#             continue

#         stem = None
#         cell_d = row[3]
#         if isinstance(cell_d, str):
#             match = re.search(r"[Ff]or\s+example[:\s]+([^\n]+)", cell_d)
#             if match:
#                 stem = match.group(1).strip()

#         if stem is None:
#             cell_e = row[4] if len(row) >= 5 else None
#             if isinstance(cell_e, str):
#                 stem = cell_e.strip()

#         if not stem:
#             continue

#         for ext in (".csv", ".xlsx", ".xlsb"):
#             if stem.lower().endswith(ext):
#                 stem = stem[: -len(ext)]
#                 break

#         parts = stem.split("-")
#         if len(parts) < 7:
#             continue

#         comp_code = parts[0].strip()
#         sbu       = parts[1].strip()
#         brand     = parts[2].strip()
#         season    = parts[5].strip()
#         seq_str   = parts[6].strip()
#         seq       = int(seq_str) if seq_str.isdigit() else 1

#         if not comp_code or not season:
#             continue

#         brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())
#         log.info(
#             "Parsed: comp_code=%s  sbu=%s  brand=%s  brand_code=%s  season=%s  seq=%s",
#             comp_code, sbu, brand, brand_code, season, seq,
#         )
#         wb.close()
#         return {
#             "comp_code":  comp_code,
#             "sbu":        sbu,
#             "brand":      brand,
#             "brand_code": brand_code,
#             "season":     season,
#             "seq":        seq,
#         }

#     wb.close()
#     log.error(
#         "No parseable example filename found in naming file '%s' (sheet '%s'). "
#         "Ensure a row has: 'For example: {COMP}-{SBU}-{BRAND}-{TYPE}-{MULTI_MONO}-{SEASON}-{SEQ}'",
#         naming_file.name, sheet_name,
#     )
#     return {}


# def _parse_metadata_from_linelist_filename(dirs: dict, triggered_file_type: str = None) -> dict:
#     """
#     Parse metadata from the triggered brand file's filename.
#     triggered_file_type narrows the search to the exact dir that was triggered,
#     avoiding misclassified files in other dirs.
#     """
#     import re

#     SEASON_PATTERN = re.compile(r'^[A-Z]{2}\d{2,4}$', re.IGNORECASE)

#     # ── Build candidate list — triggered dir FIRST, then fallbacks ──
#     if triggered_file_type and triggered_file_type in dirs:
#         candidate_dirs = [
#             dirs.get(triggered_file_type),   # ← exact triggered dir first
#             dirs.get("linelist"),
#             dirs.get("packing"),
#             dirs.get("ecommerce"),
#         ]
#     else:
#         candidate_dirs = [
#             dirs.get("linelist"),
#             dirs.get("packing"),
#             dirs.get("ecommerce"),
#         ]

#     # Deduplicate while preserving order
#     seen = set()
#     candidate_dirs = [
#         d for d in candidate_dirs
#         if d is not None and str(d) not in seen and not seen.add(str(d))
#     ]

#     for folder in candidate_dirs:
#         if not folder or not folder.exists():
#             continue
#         for xlsx_file in folder.glob("*.xlsx"):
#             stem = xlsx_file.stem
#             log.info("Attempting metadata parse from filename: '%s'", stem)

#             parts = re.split(r"\s*-\s*", stem)
#             if len(parts) < 7:
#                 log.warning(
#                     "  Filename '%s' has only %d '-'-separated parts (need ≥7) — skipping",
#                     stem, len(parts),
#                 )
#                 continue

#             comp_code = parts[0].strip()
#             sbu       = parts[1].strip()
#             brand     = parts[2].strip().title()

#             if not comp_code:
#                 log.warning("  comp_code empty — skipping: '%s'", stem)
#                 continue

#             # ── Locate season token by regex (same approach as Adidas) ──
#             # Handles SP26, FW2026, SS2029 etc. regardless of position.
#             season     = None
#             season_idx = None
#             for i, p in enumerate(parts):
#                 if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
#                     season     = p.strip().upper()
#                     season_idx = i
#                     break

#             if not season or season_idx is None or season_idx < 3:
#                 log.warning("  No valid season token in '%s' — skipping", stem)
#                 continue

#             # ── Everything after the season token ──────────────────────
#             trailing = parts[season_idx + 1:]

#             # seq = last all-digit token in trailing
#             seq = 1
#             for p in reversed(trailing):
#                 if p.strip().isdigit():
#                     seq = int(p.strip())
#                     break

#             # country = first 2-letter uppercase token in trailing
#             country_code = ""
#             for p in trailing:
#                 p = p.strip()
#                 if re.match(r'^[A-Z]{2,3}$', p, re.IGNORECASE) and not p.isdigit():
#                     country_code = p.upper()
#                     break

#             # multi_mono = token just before season
#             multi_mono = parts[season_idx - 1].strip() if season_idx >= 1 else "Multi"

#             brand_code = BRAND_NAME_TO_CODE.get(brand.upper(), brand[:3].upper())

#             # ── article_type: scan tokens between parts[3] and season ──
#             article_type_from_filename = ""
#             file_type_tokens = parts[3:season_idx]
#             for token in file_type_tokens:
#                 t = token.strip().lower()
#                 if "inline" in t:
#                     article_type_from_filename = "Inline"
#                     break
#                 if "license" in t or "licence" in t:
#                     article_type_from_filename = "License"
#                     break
#                 if "sse" in t:
#                     article_type_from_filename = "SSE"
#                     break

#             log.info(
#                 "Parsed from filename: comp_code=%s  sbu=%s  brand=%s  "
#                 "brand_code=%s  season=%s  season_idx=%s  seq=%s  "
#                 "multi_mono=%s  country=%s  article_type=%s",
#                 comp_code, sbu, brand, brand_code, season, season_idx,
#                 seq, multi_mono, country_code, article_type_from_filename,
#             )
#             return {
#                 "comp_code":  comp_code,
#                 "sbu":        sbu,
#                 "brand":      brand,
#                 "brand_code": brand_code,
#                 "season":     season,
#                 "seq":        seq,
#                 "multi_mono": multi_mono,
#                 "country_code":              country_code,
#                 "article_type_from_filename": article_type_from_filename,
#             }

#     log.error("No valid brand file found in candidate dirs.")
#     return {}

# def _prepare_tmp_dirs() -> dict[str, Path]:
#     """
#     Create clean /tmp/stibo_workdir_smiggle/input/{type}/ dirs.
#     Only linelist, mdd, attributes, naming are needed for Smiggle.
#     backlog and tdd dirs are created as stubs (ETL expects them to exist
#     even if empty).
#     """
#     base = Path(TMP_WORKDIR)

#     input_base = base / "input"
#     if input_base.exists():
#         shutil.rmtree(str(input_base))

#     xml_out = base / "output" / "xml"
#     if xml_out.exists():
#         shutil.rmtree(str(xml_out))
#     xml_out.mkdir(parents=True, exist_ok=True)

#     dirs: dict[str, Path] = {
#         "linelist":   base / "input" / "linelist",
#         "packing":    base / "input" / "packing",
#         "ecommerce":  base / "input" / "ecommerce",
#         "mdd":        base / "input" / "mdd",
#         "attributes": base / "input" / "attributes",
#         "naming":     base / "input" / "naming",
#         # Stubs — not used by Smiggle ETL but created to avoid OS errors
#         "backlog":    base / "input" / "backlog",
#         "tdd":        base / "input" / "tdd",
#     }
#     for d in dirs.values():
#         d.mkdir(parents=True, exist_ok=True)

#     return dirs


# def _download_files(
#     bucket:  str,
#     found:   dict[str, dict],
#     dirs:    dict[str, Path],
#     auditor: AuditLogger,
# ) -> None:
#     """Download each classified S3 file to its local sub-folder."""
#     for ftype, info in found.items():
#         if ftype not in dirs:
#             log.warning("  No local dir for type '%s' — skipping download", ftype)
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
#     Upload every .xml under TMP_WORKDIR/output/xml/ to:
#         s3://{PROCESSED_BUCKET}/processed/stepxml/{principal}/{filename}
#     Returns list of uploaded S3 keys.
#     """
#     xml_dir  = Path(TMP_WORKDIR) / "output" / "xml"
#     uploaded = []
#     for xml_file in xml_dir.glob("*.xml"):
#         s3_key = f"processed/stepxml/{principal}/{xml_file.name}"
#         log.info("  Uploading %s → s3://%s/%s", xml_file.name, PROCESSED_BUCKET, s3_key)
#         s3.upload_file(str(xml_file), PROCESSED_BUCKET, s3_key)
#         uploaded.append(s3_key)
#     return uploaded


# # ═════════════════════════════════════════════════════════════════════════════
# # LAMBDA HANDLER
# # ═════════════════════════════════════════════════════════════════════════════

# def lambda_handler(event, context, auditor: AuditLogger = None):
#     """
#     Main Lambda entry point.

#     auditor is injected by router.py in the shared platform;
#     if invoked directly (local test / manual trigger), we create our own.
#     """
#     _direct_invocation = auditor is None
#     if _direct_invocation:
#         auditor = AuditLogger(event=event, context=context)
#         auditor.set_router_context(
#             trigger_type     = "direct_invocation",
#             original_key     = "",
#             detected_brand   = "smiggle",
#             routing_decision = "smiggle lambda_handler invoked directly",
#         )

#     log.info("Event received: %s", json.dumps(event))

#     # ── 1. Parse event ────────────────────────────────────────────
#     bucket, key = _parse_event(event)
#     if not bucket or not key:
#         log.error("Cannot parse bucket/key from event")
#         auditor.set_lambda_status("error", error="Unparseable event")
#         if _direct_invocation:
#             auditor.flush("unknown")
#         return {"statusCode": 400, "error": "Unparseable event"}

#     key = unquote_plus(key)
#     log.info("Triggered by: s3://%s/%s", bucket, key)

#     # ── 2. Guard: only raw/metadata/ keys ────────────────────────
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

#     log.info("Principal: %s", principal)

#     # ── 3b. Detect which file type triggered this invocation ──────
#     triggered_file_type = _detect_triggered_file_type(key)
#     log.info("Triggered file type: %s", triggered_file_type or "unknown/global-fanout")

#     # Resolve the ETL module and required-types set for this trigger
#     etl_module    = ETL_DISPATCHER.get(triggered_file_type)       # may be None for global fan-out
#     required_types = REQUIRED_TYPES_BY_TRIGGER.get(
#         triggered_file_type, REQUIRED_TYPES_DEFAULT
#     )

#     # ── 4. List + classify files ──────────────────────────────────
#     found = _list_principal_files(bucket, principal,
#                                   exclude_types={triggered_file_type} if triggered_file_type else set())

#     # Always pin the trigger file — it is the authoritative source for its type
#     trigger_filename = Path(key).name
#     if triggered_file_type and not trigger_filename.startswith(".__"):
#         found[triggered_file_type] = {
#             "key":           key,
#             "filename":      trigger_filename,
#             "last_modified": None,
#         }
#         log.info("Pinned trigger file → type='%s'  file='%s'",
#                  triggered_file_type, trigger_filename)

#     log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
#     auditor.set_files_classified(found)

#     # ── 4b. Collect ALL linelist versions for diff ────────────────
#     all_linelists: list[dict] = []
#     for page in s3.get_paginator("list_objects_v2").paginate(
#         Bucket=bucket, Prefix=f"raw/metadata/{principal}/"
#     ):
#         for obj in page.get("Contents", []):
#             k_ = obj["Key"]
#             fn = Path(k_).name.lower()
#             if fn.endswith((".xlsx", ".xlsb")) and any(
#                 kw in fn for kw in ("catalogue", "catalog", "mapi", "linelist", "wholesale", "handover")
#             ):
#                 all_linelists.append({
#                     "key":           k_,
#                     "filename":      Path(k_).name,
#                     "last_modified": obj["LastModified"],
#                 })
#     log.info("[Diff] Linelist/catalogue files found: %s",
#              [o["filename"] for o in all_linelists])

#     # ── 4c. If triggered by global fan-out, pick ETL module from
#     #        whatever brand file is present in S3 ─────────────────
#     if etl_module is None:
#         for ftype in ("linelist", "packing", "ecommerce"):
#             if ftype in found:
#                 etl_module         = ETL_DISPATCHER[ftype]
#                 triggered_file_type = ftype
#                 log.info(
#                     "Global fan-out: resolved ETL module from present files → '%s'",
#                     ftype,
#                 )
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
#     missing = required_types - set(found.keys())
#     if missing:
#         log.info("Mandatory files not yet present for '%s' — waiting for: %s",
#                  principal, sorted(missing))
#         auditor.set_files_missing(sorted(missing))
#         auditor.set_lambda_status("waiting_for_mandatory_files")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 200,
#             "principal":  principal,
#             "present":    sorted(found.keys()),
#             "missing":    sorted(missing),
#             "status":     "waiting_for_mandatory_files",
#         }

#     # ── 6. Prepare /tmp/ directories ─────────────────────────────
#     dirs = _prepare_tmp_dirs()

#     # ── 7. Download files ─────────────────────────────────────────
#     log.info("Downloading %d files ...", len(found))
#     log.info("Downloading %d files ...", len(found))
#     try:
#         _download_files(bucket, found, dirs, auditor)
#     except Exception as exc:
#         auditor.set_lambda_status("error", error=f"Download failed: {exc}")
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {"statusCode": 500, "principal": principal,
#                 "error": f"Download failed: {exc}"}

#     # ── 7b. Convert .xlsb / .csv → .xlsx ─────────────────────────
#     for type_dir in dirs.values():
#         for xlsb_file in type_dir.glob("*.xlsb"):
#             try:
#                 _convert_xlsb_to_xlsx(xlsb_file)
#                 auditor.record_conversion(
#                     xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
#                 )
#             except Exception as e:
#                 log.error("xlsb conversion failed: %s — %s", xlsb_file.name, e)
#                 auditor.record_conversion(
#                     xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx",
#                     status="failed", error=str(e),
#                 )
#                 auditor.set_lambda_status(
#                     "error", error=f"xlsb conversion failed: {xlsb_file.name}: {e}"
#                 )
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {
#                     "statusCode": 500, "principal": principal,
#                     "error": f"xlsb conversion failed: {xlsb_file.name}: {e}",
#                 }

#         for csv_file in type_dir.glob("*.csv"):
#             try:
#                 _convert_csv_to_xlsx(csv_file)
#                 auditor.record_conversion(
#                     csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
#                 )
#             except Exception as e:
#                 log.error("csv conversion failed: %s — %s", csv_file.name, e)
#                 auditor.record_conversion(
#                     csv_file.name, from_ext=".csv", to_ext=".xlsx",
#                     status="failed", error=str(e),
#                 )
#                 auditor.set_lambda_status(
#                     "error", error=f"csv conversion failed: {csv_file.name}: {e}"
#                 )
#                 if _direct_invocation:
#                     auditor.flush(principal)
#                 return {
#                     "statusCode": 500, "principal": principal,
#                     "error": f"csv conversion failed: {csv_file.name}: {e}",
#                 }

#     # ── 8. Parse pipeline metadata from linelist filename ─────────
#     meta = _parse_metadata_from_linelist_filename(dirs,triggered_file_type=triggered_file_type)
#     if not meta:
#         log.error(
#             "Cannot parse metadata from linelist filename. "
#             "Ensure filename follows: {CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Seq}.xlsx"
#         )
#         auditor.set_lambda_status(
#             "error", error="Unparseable metadata in linelist filename"
#         )
#         if _direct_invocation:
#             auditor.flush(principal)
#         return {
#             "statusCode": 400, "principal": principal,
#             "error": "Unparseable metadata in linelist filename",
#         }

#     auditor.set_metadata_parsed(meta)

#     # Verify brand_code matches what we expect (safety guard)
#     if meta["brand_code"] != "SMI":
#         log.warning(
#             "Naming file resolved brand_code='%s' — expected 'SMI'. "
#             "Proceeding but verify naming convention file content.",
#             meta["brand_code"],
#         )

#     args = types.SimpleNamespace(
#         brand      = meta["brand"],
#         brand_code = meta["brand_code"],
#         comp_code  = meta["comp_code"],
#         sbu        = meta["sbu"],
#         season     = meta["season"],
#         seq        = meta["seq"],
#         multi_mono = meta.get("multi_mono", "Multi"),
#         country_code = meta.get("country_code", ""),
#         article_type_from_filename = meta.get("article_type_from_filename", ""),
#     )

#     # ── 9. Run ETL — dispatched to the correct module ─────────────
#     log.info(
#         "Running Smiggle ETL → %s (triggered_file_type=%s) ...",
#         etl_module.__name__, triggered_file_type,
#     )
#     etl_module.run(args, auditor=auditor)

#     # ── 10. Run linelist diff (detect deleted articles) ───────────
#     try:
#         from smiggle.linelist_diff import run_diff_and_upload as smiggle_diff
#         diff_result = smiggle_diff(
#             bucket                = bucket,
#             principal             = principal,
#             processed_bucket      = PROCESSED_BUCKET,
#             all_linelist_objects  = all_linelists,
#             brand_code            = args.brand_code,
#             brand_name            = args.brand,
#             season_id             = args.season,
#             comp_code             = args.comp_code,
#             sbu                   = args.sbu,
#             seq                   = args.seq,
#             season                = args.season,
#             tmp_dir               = Path(TMP_WORKDIR),
#         )
#         log.info("[Diff] Result: %s", diff_result)
#         auditor.set_diff_result(diff_result)
#     except ImportError:
#         log.info("[Diff] smiggle.linelist_diff not found — skipping diff step.")
#         diff_result = {"skipped": True, "reason": "linelist_diff module not available"}

#     # ── 11. Upload generated XMLs ─────────────────────────────────
#     uploaded = _upload_xml_outputs(principal)

#     if not uploaded:
#         log.warning("ETL completed but no XML files were produced.")
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
#         "statusCode":          200,
#         "principal":           principal,
#         "triggered_file_type": triggered_file_type,
#         "linelist_used":       linelist_info.get("filename"),
#         "linelist_s3_key":     linelist_info.get("key"),
#         "uploaded":            uploaded,
#         "count":               len(uploaded),
#         "status":              "ok",
#         "diff_result":         diff_result,
#     }



"""
lambda_function.py — Lambda Handler for Smiggle
================================================
Function name : map-stibo-inbound-validate-transform-smiggle-dev

Trigger:
    EventBridge rule (S3 Object Created) scoped to:
        s3://map-stibo-inbound-raw-dev/raw/metadata/smiggle/*

Smiggle-specific flow vs Adidas:
    • Only ONE input file from the brand  : linelist (wholesale catalogue)
    • No TDD, no Backlog — these do not exist for Smiggle
    • mdd + attributes are shared/global files (raw/metadata/ root)
    • Required types: {"linelist", "mdd", "attributes"}
    • Linelist filename keyword detection now also recognises:
        "catalogue", "catalog", "mapi", "wholesale"

File-type dispatch:
    The triggered S3 key determines which ETL module is called:
        linelist  → smiggle.linelist_main.run()
        packing   → smiggle.packing_list_main.run()
        ecommerce → smiggle.ecommerce_main.run()

S3 layout:
    raw/metadata/smiggle/{catalogue_file}.xlsx   ← Smiggle uploads here
    raw/metadata/mdd_file.xlsx                   ← shared global MDD
    raw/metadata/attributes_list.xlsx            ← shared global attr list

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/smiggle/{filename}.xml
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
# Set LAMBDA_TMP_DIR BEFORE importing smiggle modules so that their
# module-level BASE_DIR resolves to /tmp/stibo_workdir_smiggle
# (isolated from adidas).
# ─────────────────────────────────────────────────────────────────────────────
TMP_WORKDIR = "/tmp/stibo_workdir_smiggle"
os.environ["LAMBDA_TMP_DIR"] = TMP_WORKDIR

# ── ETL modules — one per Smiggle file type ──────────────────────────────────
import smiggle.linelist_main      as linelist_etl       # noqa: E402
import smiggle.packing_list_main  as packing_etl        # noqa: E402
import smiggle.ecommerce_main     as ecommerce_etl      # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("lambda_smiggle")

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
}

# ─────────────────────────────────────────────────────────────────────────────
# File type detection
# Smiggle adds "catalogue", "catalog", "mapi", "wholesale" keywords
# ─────────────────────────────────────────────────────────────────────────────
FILE_TYPE_KEYWORDS: dict[str, list[str]] = {
    # Smiggle primary input (wholesale catalogue / MAPI)
    "linelist":   [
        "line list", "linelist",
        "catalogue", "catalog",
        "mapi", "wholesale",
        "handover",
    ],
    # Smiggle packing list
    "packing":    [
        "packing list", "packinglist", "packing",
    ],
    # Smiggle ecommerce file
    "ecommerce":  [
        "ecommerce", "e-commerce", "ecom",
    ],
    # Shared / global
    "mdd":        ["mdd", "master_data", "master data", "master data dictionary"],
    "attributes": ["attributes", "attributes_list", "attribute list"],
    # Not used by Smiggle — kept so this handler can be extended
    "backlog":    ["backlog"],
    "tdd":        ["tdd"],
}

# ── Required types per triggered file type ────────────────────────────────────
REQUIRED_TYPES_BY_TRIGGER: dict[str, set[str]] = {
    "linelist":   {"linelist",  "mdd", "attributes"},
    "packing":    {"packing",   "mdd", "attributes"},
    "ecommerce":  {"ecommerce", "mdd", "attributes"},
}

# Default (used when triggered by a global file fan-out or unknown type)
REQUIRED_TYPES_DEFAULT: set[str] = {"linelist", "mdd", "attributes"}

# Global types fetched from raw/metadata/ root (not inside smiggle/)
GLOBAL_TYPES: set[str] = {"mdd"}
ROOT_TYPES: set[str] = {"mdd", "attributes"}

# ── ETL dispatcher ────────────────────────────────────────────────────────────
ETL_DISPATCHER: dict[str, object] = {
    "linelist":  linelist_etl,
    "packing":   packing_etl,
    "ecommerce": ecommerce_etl,
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _clean_cell(v):
    """Strip illegal XML / Excel characters from string cell values."""
    if not isinstance(v, str):
        return v
    return "".join(
        c for c in v if 0x20 <= ord(c) <= 0xFFFD or c in "\t\n\r"
    )


def _convert_xlsb_to_xlsx(xlsb_path: Path) -> Path:
    """Convert .xlsb → .xlsx streaming row-by-row."""
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
    """Convert .csv → .xlsx."""
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


def _safe_ts_to_epoch(ts) -> int:
    """Convert ANY timestamp type to Unix epoch seconds. Returns 0 on failure."""
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
    """Return the file type key for a filename, or None if unrecognised."""
    name = filename.lower()

    if any(tok in name for tok in ["ecommerce", "e-commerce", "ecom"]):
        return "ecommerce"

    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw in name for kw in keywords):
            return ftype
    return None


def _detect_triggered_file_type(key: str) -> str | None:
    """
    Derive the file type from the S3 key that triggered this invocation.
    Returns a key from FILE_TYPE_KEYWORDS or None.
    """
    filename = Path(key).name
    if filename.startswith(".__"):
        return None
    return _detect_file_type(filename)


def _parse_event(event: dict) -> tuple[str, str]:
    """Extract (bucket, key) from an EventBridge S3 Object Created event."""
    detail = event.get("detail", {})
    bucket = detail.get("bucket", {}).get("name", "")
    key    = detail.get("object",  {}).get("key",  "")
    return bucket, str(key)


def _extract_principal(key: str) -> str | None:
    """
    Derive principal from key:  raw/metadata/{principal}/{filename}
    Returns the principal (e.g. "smiggle") or None.
    """
    parts = key.strip("/").split("/")
    if len(parts) >= 4:
        return parts[2]
    log.warning("Key does not follow raw/metadata/{principal}/{file}: %s", key)
    return None


def _list_principal_files(bucket: str, principal: str,
                          exclude_types: set = None) -> dict[str, dict]:
    """
    Scan S3 for all recognised files:
      1. Brand-specific files  ← raw/metadata/{principal}/
      2. Global shared files   ← raw/metadata/           (mdd, attributes only)
    """
    exclude_types = exclude_types or set()
    prefix    = f"raw/metadata/{principal}/"
    paginator = s3.get_paginator("list_objects_v2")
    found: dict[str, dict] = {}

    # ── Step 1: brand-specific folder ──────────────────────────
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
                log.warning("  Unrecognised file — skipping: %s", filename)
                continue
            if ftype in exclude_types:
                log.info("  Skipping '%s' (excluded type '%s')", filename, ftype)
                continue
            last_modified = obj["LastModified"]
            if ftype not in found or (
                _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(found[ftype]["last_modified"])
            ):
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified
                }
                log.info("  Classified  %-12s ← %s  [brand-specific]", ftype, filename)

    # ── Step 2: global folder (mdd, attributes only) ───
    add_root_metadata_files(
        s3_client=s3,
        bucket=bucket,
        found=found,
        include_types=ROOT_TYPES,
        exclude_types=exclude_types,
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
            if ftype in exclude_types:
                continue
            last_modified = obj["LastModified"]
            if ftype not in found:
                found[ftype] = {
                    "key": key, "filename": filename, "last_modified": last_modified
                }
                log.info("  Classified  %-12s ← %s  [global]", ftype, filename)
            else:
                log.info(
                    "  Skipping global %s — brand-specific already found: %s",
                    ftype, found[ftype]["filename"],
                )

    return found


def _parse_metadata_from_linelist_filename(dirs: dict, triggered_file_type: str = None) -> dict:
    """
    Parse metadata from the triggered brand file's filename.
    """
    import re

    SEASON_PATTERN = re.compile(r'^[A-Z]{2}\d{2,4}$', re.IGNORECASE)

    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs = [
            dirs.get(triggered_file_type),
            dirs.get("linelist"),
            dirs.get("packing"),
            dirs.get("ecommerce"),
        ]
    else:
        candidate_dirs = [
            dirs.get("linelist"),
            dirs.get("packing"),
            dirs.get("ecommerce"),
        ]

    seen = set()
    candidate_dirs = [
        d for d in candidate_dirs
        if d is not None and str(d) not in seen and not seen.add(str(d))
    ]

    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        for xlsx_file in folder.glob("*.xlsx"):
            stem = xlsx_file.stem
            log.info("Attempting metadata parse from filename: '%s'", stem)

            parts = re.split(r"\s*-\s*", stem)
            if len(parts) < 7:
                log.warning(
                    "  Filename '%s' has only %d '-'-separated parts (need ≥7) — skipping",
                    stem, len(parts),
                )
                continue

            comp_code = parts[0].strip()
            sbu       = parts[1].strip()
            brand     = parts[2].strip().title()

            if not comp_code:
                log.warning("  comp_code empty — skipping: '%s'", stem)
                continue

            season     = None
            season_idx = None
            for i, p in enumerate(parts):
                if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
                    season     = p.strip().upper()
                    season_idx = i
                    break

            if not season or season_idx is None or season_idx < 3:
                log.warning("  No valid season token in '%s' — skipping", stem)
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
                if "sse" in t:
                    article_type_from_filename = "SSE"
                    break

            log.info(
                "Parsed from filename: comp_code=%s  sbu=%s  brand=%s  "
                "brand_code=%s  season=%s  season_idx=%s  seq=%s  "
                "multi_mono=%s  country=%s  article_type=%s",
                comp_code, sbu, brand, brand_code, season, season_idx,
                seq, multi_mono, country_code, article_type_from_filename,
            )
            return {
                "comp_code":  comp_code,
                "sbu":        sbu,
                "brand":      brand,
                "brand_code": brand_code,
                "season":     season,
                "seq":        seq,
                "multi_mono": multi_mono,
                "country_code":              country_code,
                "article_type_from_filename": article_type_from_filename,
            }

    log.error("No valid brand file found in candidate dirs.")
    return {}


def _prepare_tmp_dirs() -> dict[str, Path]:
    """
    Create clean /tmp/stibo_workdir_smiggle/input/{type}/ dirs.
    """
    base = Path(TMP_WORKDIR)

    input_base = base / "input"
    if input_base.exists():
        shutil.rmtree(str(input_base))

    xml_out = base / "output" / "xml"
    if xml_out.exists():
        shutil.rmtree(str(xml_out))
    xml_out.mkdir(parents=True, exist_ok=True)

    dirs: dict[str, Path] = {
        "linelist":   base / "input" / "linelist",
        "packing":    base / "input" / "packing",
        "ecommerce":  base / "input" / "ecommerce",
        "mdd":        base / "input" / "mdd",
        "attributes": base / "input" / "attributes",
        "backlog":    base / "input" / "backlog",
        "tdd":        base / "input" / "tdd",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    return dirs


def _download_files(
    bucket:  str,
    found:   dict[str, dict],
    dirs:    dict[str, Path],
    auditor: AuditLogger,
) -> None:
    """Download each classified S3 file to its local sub-folder."""
    for ftype, info in found.items():
        if ftype not in dirs:
            log.warning("  No local dir for type '%s' — skipping download", ftype)
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
    Upload every .xml under TMP_WORKDIR/output/xml/ to:
        s3://{PROCESSED_BUCKET}/processed/stepxml/{principal}/{filename}
    Returns list of uploaded S3 keys.
    """
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
    """
    Main Lambda entry point.

    auditor is injected by router.py in the shared platform;
    if invoked directly (local test / manual trigger), we create our own.
    """
    _direct_invocation = auditor is None
    if _direct_invocation:
        auditor = AuditLogger(event=event, context=context)
        auditor.set_router_context(
            trigger_type     = "direct_invocation",
            original_key     = "",
            detected_brand   = "smiggle",
            routing_decision = "smiggle lambda_handler invoked directly",
        )

    # ── 1. Parse event ────────────────────────────────────────────
    bucket, key = _parse_event(event)
    if not bucket or not key:
        log.error("Cannot parse bucket/key from event")
        auditor.set_lambda_status("error", error="Unparseable event")
        if _direct_invocation:
            auditor.flush("unknown")
        return {"statusCode": 400, "error": "Unparseable event"}

    key = unquote_plus(key)
    log.info("Triggered by: s3://%s/%s", bucket, key)

    # ── 2. Guard: only raw/metadata/ keys ────────────────────────
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

    log.info("Principal: %s", principal)

    # ── 3b. Detect which file type triggered this invocation ──────
    triggered_file_type = _detect_triggered_file_type(key)
    log.info("Triggered file type: %s", triggered_file_type or "unknown/global-fanout")

    etl_module    = ETL_DISPATCHER.get(triggered_file_type)
    required_types = REQUIRED_TYPES_BY_TRIGGER.get(
        triggered_file_type, REQUIRED_TYPES_DEFAULT
    )

    # ── 4. List + classify files ──────────────────────────────────
    found = _list_principal_files(bucket, principal,
                                  exclude_types={triggered_file_type} if triggered_file_type else set())

    trigger_filename = Path(key).name
    if triggered_file_type and not trigger_filename.startswith(".__"):
        found[triggered_file_type] = {
            "key":           key,
            "filename":      trigger_filename,
            "last_modified": None,
        }
        log.info("Pinned trigger file → type='%s'  file='%s'",
                 triggered_file_type, trigger_filename)

    log.info("Files classified: %s", {k: v["filename"] for k, v in found.items()})
    auditor.set_files_classified(found)

    # ── 4b. If triggered by global fan-out, pick ETL module from
    #        whatever brand file is present in S3 ─────────────────
    if etl_module is None:
        for ftype in ("linelist", "packing", "ecommerce"):
            if ftype in found:
                etl_module          = ETL_DISPATCHER[ftype]
                triggered_file_type = ftype
                log.info(
                    "Global fan-out: resolved ETL module from present files → '%s'",
                    ftype,
                )
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
    missing = required_types - set(found.keys())
    if missing:
        log.info("Mandatory files not yet present for '%s' — waiting for: %s",
                 principal, sorted(missing))
        auditor.set_files_missing(sorted(missing))
        auditor.set_lambda_status("waiting_for_mandatory_files")
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 200,
            "principal":  principal,
            "present":    sorted(found.keys()),
            "missing":    sorted(missing),
            "status":     "waiting_for_mandatory_files",
        }

    # ── 6. Prepare /tmp/ directories ─────────────────────────────
    dirs = _prepare_tmp_dirs()

    # ── 7. Download files ─────────────────────────────────────────
    log.info("Downloading %d files ...", len(found))
    try:
        _download_files(bucket, found, dirs, auditor)
    except Exception as exc:
        auditor.set_lambda_status("error", error=f"Download failed: {exc}")
        if _direct_invocation:
            auditor.flush(principal)
        return {"statusCode": 500, "principal": principal,
                "error": f"Download failed: {exc}"}

    # ── 7b. Convert .xlsb / .csv → .xlsx ─────────────────────────
    for type_dir in dirs.values():
        for xlsb_file in type_dir.glob("*.xlsb"):
            try:
                _convert_xlsb_to_xlsx(xlsb_file)
                auditor.record_conversion(
                    xlsb_file.name, from_ext=".xlsb", to_ext=".xlsx", status="ok"
                )
            except Exception as e:
                log.error("xlsb conversion failed: %s — %s", xlsb_file.name, e)
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
                    "error": f"xlsb conversion failed: {xlsb_file.name}: {e}",
                }

        for csv_file in type_dir.glob("*.csv"):
            try:
                _convert_csv_to_xlsx(csv_file)
                auditor.record_conversion(
                    csv_file.name, from_ext=".csv", to_ext=".xlsx", status="ok"
                )
            except Exception as e:
                log.error("csv conversion failed: %s — %s", csv_file.name, e)
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
                    "error": f"csv conversion failed: {csv_file.name}: {e}",
                }

    # ── 8. Parse pipeline metadata from linelist filename ─────────
    meta = _parse_metadata_from_linelist_filename(dirs, triggered_file_type=triggered_file_type)
    if not meta:
        log.error(
            "Cannot parse metadata from linelist filename. "
            "Ensure filename follows: {CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Seq}.xlsx"
        )
        auditor.set_lambda_status(
            "error", error="Unparseable metadata in linelist filename"
        )
        if _direct_invocation:
            auditor.flush(principal)
        return {
            "statusCode": 400, "principal": principal,
            "error": "Unparseable metadata in linelist filename",
        }

    auditor.set_metadata_parsed(meta)

    args = types.SimpleNamespace(
        brand      = meta["brand"],
        brand_code = meta["brand_code"],
        comp_code  = meta["comp_code"],
        sbu        = meta["sbu"],
        season     = meta["season"],
        seq        = meta["seq"],
        multi_mono = meta.get("multi_mono", "Multi"),
        country_code = meta.get("country_code", ""),
        article_type_from_filename = meta.get("article_type_from_filename", ""),
    )

    # ── 9. Run ETL — dispatched to the correct module ─────────────
    log.info(
        "Running Smiggle ETL → %s (triggered_file_type=%s) ...",
        etl_module.__name__, triggered_file_type,
    )
    etl_module.run(args, auditor=auditor)

    # ── 10. Upload generated XMLs ─────────────────────────────────
    uploaded = _upload_xml_outputs(principal)

    if not uploaded:
        log.warning("ETL completed but no XML files were produced.")
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
        "statusCode":          200,
        "principal":           principal,
        "triggered_file_type": triggered_file_type,
        "linelist_used":       linelist_info.get("filename"),
        "linelist_s3_key":     linelist_info.get("key"),
        "uploaded":            uploaded,
        "count":               len(uploaded),
        "status":              "ok",
    }