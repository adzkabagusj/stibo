"""
╔══════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — New Balance Inline Footwear      ║
║  Ecommerce File (Item tab / MAP_S226_FTW) → Stibo STEP XML      ║
║  v1.0                                                            ║
╚══════════════════════════════════════════════════════════════════╝

New Balance Inline Footwear Ecommerce differences vs Licensed ecommerce:

Source file
  • Input file  : NB Inline FTW Ecommerce export (e.g. MAP_S226_FTW.xlsx)
  • Active sheet: "Item"
  • Header row  : index 0 (row 1 in Excel) — column names in row 1
  • Data starts  : index 1 (row 2 in Excel)
  • No explicit channel filter — all rows are processed (status not used
    as hard filter since FTW file uses "BuyReady"/"Extremes"/"Development")

Key FTW-specific column mapping (by column name):
    ItemStyleNumber           → style+color SKU (e.g. PZ530SB1)
    ItemParentProductNumber   → style root with sub-model (e.g. PZ530V1-40418)
    ItemGender                → extended gender labels (PreBoys, GradeGirls,
                                InfantBoys, Mens, Womens, Unisex, etc.)
    ProductGender             → product-level gender label
    ProductDisplayName        → product name / style description
    ItemColor                 → "WHITE (100)" format
    ItemFirstColor_en         → "NB WHITE"
    ItemFirstGenericColor     → "White"
    ItemSeasons               → "2026S2;2026S1" (semicolon-separated)
    ItemLatestSeason          → "2027S1" (latest season token)
    ItemAPACInLineIntroDate   → preferred APAC launch date
    ItemIntroDate             → fallback intro date
    ItemDTTeCommPrice         → DTC ecomm price (RRP)
    ItemDTCeCommJNBORetailPrice → JNBO retail price (secondary RRP)
    ItemUpcSyndication        → JSON array with per-size UPC/EAN/Width/Size
    ItemUniqueIdentifiers     → alternative JSON (fallback)
    ProductPrimaryMaterial    → primary material
    ProductMaterialPercentages→ material percentages string
    ProductBullet1_en … ProductBullet10_en → feature bullets
    ProductSizeFitBullet1_en … ProductSizeFitBullet7_en → fit/size bullets
    ProductMaterialBullet1_en … ProductMaterialBullet5_en → material bullets
    ProductDTCeCommDescription_en → ecom long description EN

FTW-specific article structure
  • Article Type   : "Inline" (NB Inline, not Licensed)
  • SBU            : "FW" (Footwear)
  • Each row       = one Generic product (ItemStyleNumber = style+color)
    • Generic code   : NEW + clean(ItemStyleNumber)[:9] + color code from ItemColor
  • SAP Style code : max 8-char ItemStyleNumber + 1-char primary width
    • Variants       : parsed from ItemUpcSyndication JSON
                                         each size within the article → one Variant

FTW gender mapping (extended NB gender labels):
    Mens / Unisex           → SAP M / U  |  Age AD (Adult)
    Womens                  → SAP F       |  Age AD (Adult)
    PreBoys / GradeBoys     → SAP M       |  Age CH (Children)
    PreGirls / GradeGirls   → SAP F       |  Age CH (Children)
    InfantBoys              → SAP M       |  Age IN (Infant)
    InfantGirls             → SAP F       |  Age IN (Infant)
    Girls / Boys            → SAP F/M     |  Age CH (Children)
    Kids (ProductGender)    → SAP U       |  Age CH (Children)

Footwear size handling:
    NB sizes are numeric: "01", "015", "02" ... "14", "15"
    SAP size code = numeric NB size zero-padded to 3 chars (e.g. "01" → "010",
    "045" → "045", "14" → "140")
    Half sizes ending in 5 are kept as-is (e.g. "075" stays "075")
"""

import re
import sys
import json
import logging
import argparse
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd
import os

# ─────────────────────────────────────────────────────────────────────────────
# Directory layout  (mirrors the unified lambda structure)
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
ECOMM_DIR = INPUT_DIR / "ecommerce_inline"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR   = BASE_DIR / "output"
XML_OUT_DIR  = OUTPUT_DIR / "xml"
LOG_DIR      = OUTPUT_DIR / "logs"

for _d in [ECOMM_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

_log_path = LOG_DIR / f"nb_ftw_ecomm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(_log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════
# SECTION 0 — LOV TABLES
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS",       "NIK": "NIKE",         "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE",      "ALD": "ALDO",          "CRO": "CROCS",
    "LOT": "LOTTO",        "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults",     "CH": "Children",   "IN": "Infant",
    "AA": "All Ages",   "JR": "Junior",
}
LOV_BY_AGE = {
    "AD": "Adult",     "ADULT": "Adult",    "CH": "Children",
    "IN": "Infant",    "AA": "All Ages",    "JR": "Junior",
    "KIDS": "Kids",
}
LOV_SEASON = {
    "SS": "Spring-Summer",  "FW": "Fall-Winter",
    "AL": "All Season",     "HO": "Holiday",
    "AW": "Autumn-Winter",
}
LOV_SBU = {
    "SP": "Sports",   "FQ": "Footlocker",  "FL": "Fashion Footwear",
    "SM": "Smiggle",  "FW": "Footwear",
}

# ── Extended NB FTW gender → SAP Gender code ─────────────────────
NB_FTW_GENDER_TO_SAP: dict[str, str] = {
    "MENS":        "M",
    "WOMENS":      "F",
    "UNISEX":      "U",
    "PREBOYS":     "M",
    "PREGIRLS":    "F",
    "GRADEBBOYS":  "M",  # typo-safe
    "GRADEBOYS":   "M",
    "GRADEGIRLS":  "F",
    "INFANTBOYS":  "M",
    "INFANTGIRLS": "F",
    "BOYS":        "M",
    "GIRLS":       "F",
    "KIDS":        "U",
    "GENERAL":     "U",
}

# ── Extended NB FTW gender → SAP Age code ────────────────────────
NB_FTW_GENDER_TO_AGE: dict[str, str] = {
    "MENS":        "AD",
    "WOMENS":      "AD",
    "UNISEX":      "AD",
    "PREBOYS":     "CH",
    "PREGIRLS":    "CH",
    "GRADEBOYS":   "CH",
    "GRADEGIRLS":  "CH",
    "INFANTBOYS":  "IN",
    "INFANTGIRLS": "IN",
    "BOYS":        "CH",
    "GIRLS":       "CH",
    "KIDS":        "CH",
    "GENERAL":     "AA",
}

# ── Age code → BY LOV ID ─────────────────────────────────────────
NB_AGE_CODE_TO_BY_LOV_ID: dict[str, str] = {
    "AD": "ADULT",
    "CH": "KIDS",
    "IN": "INFANT",
    "AA": "ALL AGES",
    "JR": "KIDS",
}

# ── Footwear width → single-char SAP suffix ──────────────────────
# Used to build the SAP Style Code (ItemStyleNumber + width suffix)
NB_WIDTH_TO_SUFFIX: dict[str, str] = {
    "2A": "2",   "A":  "A",   "B":  "B",   "C":  "C",
    "D":  "D",   "E":  "E",   "EE": "E",   "2E": "E",
    "M":  "M",   "N":  "N",   "W":  "W",   "X":  "X",
    "4E": "4",   "SW": "S",   "K":  "K",
}


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """Loads Core Attributes + LOVs from the MDD Excel."""

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb   = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws   = wb["Core Attributes"]
        rows = list(ws.iter_rows(values_only=True))
        hdr  = rows[1]

        col: dict[str, int] = {}
        for i, h in enumerate(hdr):
            if h:
                col[str(h).split("\n")[0].strip()] = i
                col[str(h).strip()] = i

        for row in rows[2:]:
            aid = row[7] if len(row) > 7 else None
            if not aid:
                continue
            aid = str(aid).strip()

            def _v(key, _row=row, _col=col):
                idx = _col.get(key)
                if idx is None or idx >= len(_row):
                    return None
                v = _row[idx]
                return str(v).strip() if v else None

            self.attributes[aid] = {
                "id":          aid,
                "name":        _v("PIM Attribute Name\n(As MAA want it to be called in STEP)"),
                "validation":  _v("Validation Base Type"),
                "cardinality": _v("Source Application Name"),   # column mapping per MDD layout
            }

        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                raw_id = str(val_id).strip() if val_id else str(val_name).strip()
                if raw_id.isdigit():
                    raw_id = raw_id.zfill(3)
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = raw_id

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                code, name = row[0], row[1]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )
        # ── Color Code LOV special handler ────────────────────────
        if "Color Code LOV" in wb.sheetnames:
            lov = {}
            for row in list(wb["Color Code LOV"].iter_rows(values_only=True))[1:]:
                if not row or len(row) < 2:
                    continue
                code, name = row[0], row[1]
                if code and name:
                    raw_code = str(code).strip()
                    lov[str(name).strip().upper()] = raw_code
            self.lovs["Color Code"] = lov


class AttributesListLoader:
    """Loads the NEW BALANCE tab from the Attributes List workbook."""

    def __init__(self, path: Path, brand: str = "NEW BALANCE"):
        self.path      = path
        self.brand     = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames
             if s.upper() == self.brand and "V4" not in s.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames
             if self.brand in s.upper() and "V4" not in s.upper()),
            wb.sheetnames[0],
        )

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows) if r and any(
                str(v).strip() == "Attributes" for v in r if v
            )),
            None,
        )
        if hdr_idx is None:
            log.warning("[AttrList] Cannot find header row in '%s'", sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}

        for row in rows[hdr_idx + 1:]:
            attr = row[1] if len(row) > 1 else None
            if not attr:
                continue
            self.attr_map.append({
                "attribute":     str(attr).strip(),
                "cluster":       self._s(row, col, "Cluster"),
                "indicator":     self._s(row, col, "Indicator"),
                "mapping_logic": self._s(row, col, "Field Name / Mapping Logic"),
            })

        wb.close()
        log.info("[AttrList] %d attributes from sheet '%s'", len(self.attr_map), sheet_name)

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class NBFTWEcommLoader:
    """
    Loads the NB Inline Footwear Ecommerce export (MAP_S226_FTW.xlsx).

    Sheet    : "Item"
    Header   : row 1 (index 0)
    Data     : row 2+ (index 1+)

    Rows with ItemStatus == "Inactive" or "Canceled" are dropped.
    Rows without ItemStyleNumber are dropped.
    """

    SHEET_NAME = "Item"

    def __init__(self, path: Path):
        self.path = path
        self.df   = pd.DataFrame()
        self._load()

    def _load(self):
        log.info("[FTW-Ecomm] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in wb.sheetnames if s.upper() == self.SHEET_NAME.upper()),
            None,
        ) or wb.sheetnames[0]

        log.info("[FTW-Ecomm] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            log.error("[FTW-Ecomm] Empty sheet '%s'", target)
            wb.close()
            return

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[0])
        ]
        df = pd.DataFrame(rows[1:], columns=header)

        # Drop clearly inactive / canceled
        if "ItemStatus" in df.columns:
            before = len(df)
            df = df[~df["ItemStatus"].astype(str).str.strip().str.upper().isin(
                ["INACTIVE", "CANCELED"]
            )]
            log.info(
                "[FTW-Ecomm] Dropped Inactive/Canceled: %d → %d",
                before, len(df),
            )

        # Drop rows without style number
        if "ItemStyleNumber" in df.columns:
            df = df[
                df["ItemStyleNumber"].notna()
                & (~df["ItemStyleNumber"].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        wb.close()
        self.df = df.reset_index(drop=True)
        log.info("[FTW-Ecomm] %d active rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — HELPERS
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "N/A", "none") else s


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw:
        return ""
    for fmt in (
        "%m/%d/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M:%S %p",
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(raw.split(" ")[0], fmt.split(" ")[0]).strftime(
                "%d-%b-%Y"
            ).lower()
        except (ValueError, TypeError):
            pass
    return raw


def _parse_season_from_tokens(seasons_str: str, latest_season: str) -> str:
    """
    Normalise NB season tokens → 2+2 season code (e.g. SS26, FW26).

    Handles:
      2026S1  → SS26    2026S2  → FW26
      S24     → SS24    F25     → FW25
      SS26    → SS26    FW26    → FW26
      Fall26  → FW26    Spring26 → SS26
    """
    token = _s(latest_season)
    if not token and seasons_str:
        parts = [t.strip() for t in _s(seasons_str).split(";") if t.strip()]
        token = parts[-1] if parts else ""
    if not token:
        return ""

    # Pattern: 2026S1 / 2026S2
    m = re.match(r"^(\d{4})S([12])$", token, re.IGNORECASE)
    if m:
        year   = m.group(1)[2:]
        prefix = "SS" if m.group(2) == "1" else "FW"
        return f"{prefix}{year}"

    # Pattern: S24, F25
    m = re.match(r"^([SF])(\d{2,4})$", token, re.IGNORECASE)
    if m:
        prefix = "SS" if m.group(1).upper() == "S" else "FW"
        return f"{prefix}{m.group(2)[-2:]}"

    # Pattern: FW26, SS26, AW26, Fall26, Spring26
    m = re.match(
        r"^(FW|SS|AW|HO|Fall|Spring|Summer|Winter)[\s_]?(\d{2,4})$",
        token, re.IGNORECASE,
    )
    if m:
        raw_prefix = m.group(1).upper()
        pmap = {"FALL": "FW", "WINTER": "FW", "SPRING": "SS", "SUMMER": "SS"}
        prefix = pmap.get(raw_prefix, raw_prefix[:2])
        return f"{prefix}{m.group(2)[-2:]}"

    return token


def _parse_upc_json(json_str: str) -> list[dict]:
    """
    Parse ItemUpcSyndication JSON into a list of variant dicts.

    FTW UPC JSON format per entry:
    {
      "Style": "PZ530SB1",
      "UPC": "196432257133",
      "Ean": "0196432257133",
      "NewBalanceSize": "01",
      "Width": "M",
      "Gender": "K",
      "WidthDescription": "Medium",
      "UsaSize": "01",
      "UkSize": "13.5",
      "EuropeanSize": "32.5",
      "CmSize": "19",
      "SKU": "PZ530SB1-M-01",
      "Pckg Weight": "0.980",
      "Pckg Length": "10.80",
      "Pckg Width": "",
      "Pckg Height": "4.30"
    }
    """
    if not json_str or _s(json_str) in ("", "None", "nan"):
        return []
    try:
        data = json.loads(str(json_str))
    except (json.JSONDecodeError, TypeError):
        log.warning("Failed to parse UPC JSON (first 120 chars): %s", str(json_str)[:120])
        return []

    variants = []
    for item in data:
        if not isinstance(item, dict):
            continue
        # ItemUpcSyndication uses "NewBalanceSize"
        size  = _s(item.get("NewBalanceSize") or item.get("Size", ""))
        width = _s(item.get("Width", ""))
        # Skip "SS" (Sample Size)
        if size.upper() == "SS" or width.upper() == "SW":
            continue
        variants.append({
            "size":        size,
            "width":       width,
            "upc":         _s(item.get("UPC", "")),
            "ean":         _s(item.get("Ean", "")),
            "usa_size":    _s(item.get("UsaSize", "")),
            "uk_size":     _s(item.get("UkSize", "")),
            "eu_size":     _s(item.get("EuropeanSize", "")),
            "cm_size":     _s(item.get("CmSize", "")),
            "sku":         _s(item.get("SKU", "")),
            "pckg_weight": _s(item.get("Pckg Weight", "")),
            "pckg_length": _s(item.get("Pckg Length", "")),
            "pckg_width":  _s(item.get("Pckg Width", "")),
            "pckg_height": _s(item.get("Pckg Height", "")),
        })
    return variants


def _parse_uid_json(json_str: str) -> list[dict]:
    """
    Parse ItemUniqueIdentifiers JSON (fallback / alternative format).

    Format uses "AdditionnalAttributes" + "Variants" sub-objects.
    """
    if not json_str or _s(json_str) in ("", "None", "nan"):
        return []
    try:
        data = json.loads(str(json_str))
    except (json.JSONDecodeError, TypeError):
        return []

    variants = []
    for item in data:
        if not isinstance(item, dict):
            continue
        addl     = item.get("AdditionnalAttributes", {}) or {}
        var_info = item.get("Variants", {}) or {}
        size     = _s(var_info.get("Size", ""))
        width    = _s(var_info.get("Width", ""))
        if size.upper() == "SS" or width.upper() == "SW":
            continue
        variants.append({
            "size":        size,
            "width":       width,
            "upc":         _s(addl.get("UPC", "")),
            "ean":         _s(addl.get("Ean", "")),
            "usa_size":    _s(var_info.get("UsaSize", "")),
            "uk_size":     _s(var_info.get("UkSize", "")),
            "eu_size":     _s(var_info.get("EuropeanSize", "")),
            "cm_size":     _s(var_info.get("CmSize", "")),
            "sku":         "",
            "pckg_weight": _s(addl.get("Pckg Weight", "")),
            "pckg_length": _s(addl.get("Pckg Length", "")),
            "pckg_width":  _s(addl.get("Pckg Width", "")),
            "pckg_height": _s(addl.get("Pckg Height", "")),
        })
    return variants


def _ftw_sap_size_code(size_raw: str) -> str:
    """
    Convert a NB footwear size to a 3-char SAP size code.

    NB footwear sizes are numeric:
      "01"  → "010"
      "015" → "015"
      "045" → "045"
      "14"  → "140"
      "4"   → "040"
    Rule: if 2 chars → append "0"; if 1 char → left-pad to 3 with "0";
          if already 3 chars → keep as-is.
    """
    s = size_raw.strip()
    if not s:
        return "000"
    # Remove any non-numeric chars for safety
    digits = re.sub(r"[^0-9]", "", s)
    if not digits:
        return s[:3].ljust(3, "0")
    if len(digits) == 1:
        return digits.zfill(3)
    if len(digits) == 2:
        return digits + "0"
    return digits[:3]


def _color_to_sap_token(color_str: str) -> str:
    """
    Derive 3-char SAP color token from NB color string.

    "WHITE (100)" → "100"
    "MOONBEAM (121)" → "121"
    "BEIGE (250)" → "250"
    "BLACK" → "BLK"   (no parenthesis fallback)
    """
    m = re.search(r"\(([A-Z0-9]+)\)", (color_str or "").upper())
    if m:
        code = m.group(1)[:3].ljust(3, "X")
        return code
    cleaned = re.sub(r"[^A-Z0-9]", "", (color_str or "").upper())[:3].ljust(3, "X")
    return cleaned


def _sap_style_code_ftw(item_style: str, primary_width: str) -> str:
    """
    Build FTW SAP Style Code: up to 8-char ItemStyleNumber + 1-char width suffix.

    Per MDD: "Max 8 Digits Principal Style Code + 1 digit Width"
    e.g. PZ530SB1 + M → PZ530SB1M (9 chars max)
    """
    style_clean = re.sub(r"[^A-Z0-9]", "", item_style.upper())[:8]
    width_suffix = NB_WIDTH_TO_SUFFIX.get(primary_width.upper(), primary_width[:1].upper() or "D")
    return style_clean + width_suffix


def _excel_color_code(color_str: str, fallback: str = "") -> str:
    """Extract the source color code from ItemColor, e.g. 'WHITE (100)' -> '100'."""
    m = re.search(r"\(([A-Z0-9]+)\)", (color_str or "").upper())
    raw = m.group(1) if m else (fallback or color_str)
    return re.sub(r"[^A-Z0-9]", "", raw.upper())


def _excel_color_name(color_str: str, fallback: str = "") -> str:
    """Extract the color name from ItemColor, e.g. 'WHITE (100)' -> 'WHITE'."""
    raw = re.sub(r"\s*\([^)]*\)", "", color_str or "").strip()
    if not raw:
        raw = fallback or ""
    raw = re.sub(r"^NB\s+", "", raw.strip(), flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", raw).strip()


def _build_generic_code(brand_code: str, item_style: str, color_code: str = "") -> str:
    style = re.sub(r"[^A-Z0-9]", "", item_style.upper())[:9]
    return f"{brand_code}{style}"


def _join_bullets(row: dict, prefix: str, count: int) -> str:
    """Join non-empty bullet columns into pipe-separated string."""
    parts = []
    for i in range(1, count + 1):
        col = f"{prefix}{i}_en"
        v   = _s(row.get(col, ""))
        if v:
            parts.append(v)
    return " | ".join(parts)


def _price(v) -> str:
    if v is None:
        return ""
    try:
        f = float(v)
        return "" if f == 0.0 else str(round(f, 2))
    except (TypeError, ValueError):
        m = re.search(r"[\d.]+", str(v))
        return m.group() if m else str(v)


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_article_nb_ftw(row: dict, brand_code: str = "NEW") -> dict:
    """
    Map one NB FTW Ecommerce row → unified article dict.

    Key FTW decisions:
    ─────────────────
    • article_no      = ItemStyleNumber (e.g. PZ530SB1)
    • sap_style_code  = clean(ItemStyleNumber)[:8] + primary_width_suffix
    • generic_code    = NEW + clean(ItemStyleNumber)[:9] + ItemColor code
    • color_token     = 3-char from ItemColor parenthesis "(100)" → "100"
    • variants        = from ItemUpcSyndication JSON (Width-grouped)
    • gender/age      = mapped from extended NB gender labels
    • season          = from ItemLatestSeason / ItemSeasons tokens
    • rrp             = ItemDTTeCommPrice (primary) or ItemDTCeCommJNBORetailPrice
    """
    article_no     = _s(row.get("ItemStyleNumber"))
    parent_product = _s(row.get("ItemParentProductNumber"))
    gender_raw     = _s(row.get("ItemGender"))
    prod_gender    = _s(row.get("ProductGender"))
    model_name     = _s(row.get("ProductDisplayName"))
    color_str      = _s(row.get("ItemColor"))           # "WHITE (100)"
    color_name_en  = _s(row.get("ItemFirstColor_en"))   # "NB WHITE"
    color_family   = _s(row.get("ItemFirstGenericColor")) # "White"
    seasons_str    = _s(row.get("ItemSeasons", ""))
    latest_season  = _s(row.get("ItemLatestSeason", ""))
    apac_intro     = row.get("ItemAPACInLineIntroDate")
    intro_date     = row.get("ItemIntroDate")
    mat_primary    = _s(row.get("ProductPrimaryMaterial", ""))
    mat_pct        = _s(row.get("ProductMaterialPercentages", ""))
    ecomm_desc     = _s(row.get("ProductDTCeCommDescription_en", ""))

    rrp_dtc  = row.get("ItemDTTeCommPrice")
    rrp_jnbo = row.get("ItemDTCeCommJNBORetailPrice")

    # Bullets
    tech_bullets = _join_bullets(row, "ProductBullet", 10)
    fit_bullets  = _join_bullets(row, "ProductSizeFitBullet", 7)
    mat_bullets  = _join_bullets(row, "ProductMaterialBullet", 5)

    # Gender / Age mapping (extended NB gender labels)
    g_key       = gender_raw.upper().replace(" ", "")
    gender_code = NB_FTW_GENDER_TO_SAP.get(g_key, "U")
    age_code    = NB_FTW_GENDER_TO_AGE.get(g_key, "AD")

    # Color token
    color_token = _color_to_sap_token(color_str or color_name_en)
    generic_color_code = _excel_color_code(color_str, color_token)
    principal_color_name = _excel_color_name(color_str, color_name_en or color_family)

    # Launching date
    launching_date = _fmt_date(apac_intro) or _fmt_date(intro_date)

    # Season
    season = _parse_season_from_tokens(seasons_str, latest_season)

    # Price
    rrp = _price(rrp_dtc) or _price(rrp_jnbo)

    # Variants from UPC Syndication (preferred) → fallback UniqueIdentifiers
    upc_json = row.get("ItemUpcSyndication")
    uid_json = row.get("ItemUniqueIdentifiers")
    variants = _parse_upc_json(upc_json)
    if not variants:
        variants = _parse_uid_json(uid_json)

    # Derive primary width from first variant
    primary_width = variants[0]["width"] if variants else "D"

    # SAP Style Code: ItemStyleNumber (up to 8 chars) + 1-char width suffix
    sap_style = _sap_style_code_ftw(article_no, primary_width)

    # Generic code follows NB linelist pattern: brand + style max 9 + source color code.
    generic_code = _build_generic_code(brand_code, article_no, generic_color_code)

    return {
        # Core identifiers
        "article_no":    article_no,
        "parent_product": parent_product,
        "brand_code":    brand_code,

        # Colour
        "color_str":     color_str,
        "color_name_en": color_name_en,
        "color_family":  color_family,
        "color_token":   color_token,
        "generic_color_code": generic_color_code,
        "principal_color_name": principal_color_name,

        # Gender / Age
        "gender_code":   gender_code,
        "age_code":      age_code,
        "gender_raw":    gender_raw,
        "prod_gender":   prod_gender,

        # Descriptions
        "model_name":    model_name,
        "ecomm_desc_en": ecomm_desc,

        # Bullets / Content
        "tech_bullets":  tech_bullets,
        "fit_bullets":   fit_bullets,
        "mat_bullets":   mat_bullets,
        "mat_primary":   mat_primary,
        "mat_pct":       mat_pct,

        # Commercial
        "article_type":  "Inline",   # NB Inline (not Licensed)
        "art_category":  "1",        # Generic
        "rrp":           rrp,

        # Dates & Season
        "launching_date": launching_date,
        "season":         season,

        # Variants (all widths combined; will be split per-width in XML builder)
        "variants":       variants,
        "primary_width":  primary_width,

        # Derived codes
        "sap_style_code": sap_style,
        "generic_code":   generic_code,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Basic mandatory-field check using MDD attribute metadata."""
    warns = []
    art   = mapped["article_no"]
    for field, label in [
        ("article_no",   "ItemStyleNumber"),
        ("gender_code",  "SAP Gender"),
        ("age_code",     "SAP Age"),
        ("brand_code",   "Brand"),
        ("season",       "Season"),
    ]:
        val = mapped.get(field, "")
        if not val or str(val).strip() in ("", "None", "nan"):
            warns.append(f"[{art}] MISSING: {label}")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element | None:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip()  if value  else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v  = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


# ── Placeholder attributes (written as empty values) ─────────────
NB_FTW_GENERIC_PLACEHOLDERS = [
    # "AT_SPUGrouping",            "AT_SPUGroupingName",
    # "AT_Collection1",            "AT_Collection2",
    # "AT_Franchise",
    # "AT_EComProductNameID",      "AT_EComProductNamePH",
    # "AT_EComProductNameMY",      "AT_EComProductNameVN",
    # "AT_EComProductNameTH",      "AT_EComProductNameKH",
    # "AT_ShortDescriptionID",     "AT_ShortDescriptionPH",
    # "AT_ShortDescriptionMY",     "AT_ShortDescriptionVN",
    # "AT_ShortDescriptionTH",     "AT_ShortDescriptionKH",
    # "AT_LongDescriptionID",      "AT_LongDescriptionPH",
    # "AT_LongDescriptionMY",      "AT_LongDescriptionVN",
    # "AT_LongDescriptionTH",      "AT_LongDescriptionKH",
    # "AT_CareInstructionEN",      "AT_CareInstructionID",
    # "AT_CareInstructionPH",      "AT_CareInstructionMY",
    # "AT_CareInstructionVN",      "AT_CareInstructionTH",
    # "AT_CareInstructionKH",
    # "AT_PrincipalColorName",
    # "AT_EComAgesCategory",
    # "AT_CertificateNumber",
    # "AT_ProductWeight",
    # "AT_ProductLengthWidthHeight",
    # "AT_ImagesSource",
    # "AT_EstimatedLandedCost",
    # "AT_ExchangeRate",
    # "AT_FOB",
    # "AT_FOBCurrency",
    # "AT_MainEANIndicator",
    # "AT_HSCode",
    # "AT_Royalty",
    # "AT_FreightCost",
    # "AT_MarketingFee",
    # "AT_SGS",
    # "AT_HaddadOfficeCharge",
    # "AT_HaddadOfficeChargeDeduction",
    # "AT_Others",
    # "AT_Commission",
    # "AT_Createdon",
    # "AT_Updatedon",
    # "AT_Updatedby",
    # "AT_Width",                  # Footwear-specific BY attribute
    # "AT_Style",                  # Footwear silhouette style
    # "AT_StyleType",
    # "AT_GenderSizeChartUsage",
    # "AT_GenderSizeChart",
]


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
) -> None:
    """Write all Generic-level <Value> elements for a NB FTW article."""
    written: set[str] = set()

    def _w(attr_id, value="", id_val=""):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val)
            written.add(attr_id)

    def _mw(attr_id, id_val):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── Organisational ───────────────────────────────────────────
    _mw("AT_SBU", sbu)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _w("AT_Brand",      id_val=b_code)
    _w("AT_BrandGroup", id_val=b_label)

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["article_no"])
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["generic_color_code"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
    _w("AT_InboundGenericCode",                   art["generic_code"])

    # ── Color ────────────────────────────────────────────────────
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        # Try the numeric code extracted from parenthesis first
        paren_code = art["color_token"]
        # Look up by the EN name (uppercase) and by color family
        colour_id = (
            color_lov.get(art["color_name_en"].upper(), "")
            or color_lov.get(art["color_family"].upper(), "")
            or color_lov.get(paren_code, "")
        )
    if not colour_id:
        paren_match = re.search(r"\((\d+)\)", art["color_str"])
        if paren_match:
            colour_id = paren_match.group(1) 
    # if colour_id and colour_id.isdigit():
    #     colour_id = colour_id.zfill(3)

    _w("AT_Color", id_val=colour_id)
    _w("AT_PrincipalColorCode", art["generic_color_code"])
    _w("AT_PrincipalColorName", art["principal_color_name"])
    # _w("AT_PrincipalColorNameEN", art["color_name_en"])

    # ── Gender ───────────────────────────────────────────────────
    g_code, _ = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",                     id_val=g_code)
    _w("AT_BYGender",                   id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age ──────────────────────────────────────────────────────
    sap_age_code = art["age_code"]
    by_age_lov   = NB_AGE_CODE_TO_BY_LOV_ID.get(sap_age_code, "ADULT")
    _w("AT_SAPAge", id_val=sap_age_code)
    _w("AT_BYAge",  id_val=by_age_lov)
    _w("AT_PrincipalAgeDescription", art["gender_raw"])

    # ── Season ───────────────────────────────────────────────────
    # ── Season  ── FIX: handles SS26/SS2026; writes label as text + prefix as ID
    # sea_raw = art.get("season", "")
    # sea_m   = re.match(r'^([A-Z]{2})(\d{2,4})$', sea_raw.strip().upper())
    # if sea_m:
    #     sea_prefix = sea_m.group(1)                                               # "SS"
    #     sea_digits = sea_m.group(2)                                               # "26" or "2026"
    #     sea_year_4 = sea_digits if len(sea_digits) == 4 else f"20{sea_digits}"    # "2026"
    # else:
    #     sea_prefix = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    #     sea_year_4 = ""

    file_sea_code = art.get("file_season_code", "")
    file_sea_year = art.get("file_season_year", "")
    if file_sea_code and file_sea_year:
        sea_prefix = file_sea_code
        sea_year_4 = f"20{file_sea_year}" if len(file_sea_year) == 2 else file_sea_year
    else:
        sea_raw = art.get("season", "")
        sea_m   = re.match(r'^([A-Z]{2})(\d{2,4})$', sea_raw.strip().upper())
        if sea_m:
            sea_prefix = sea_m.group(1)
            sea_digits = sea_m.group(2)
            sea_year_4 = sea_digits if len(sea_digits) == 4 else f"20{sea_digits}"
        else:
            sea_prefix = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
            sea_year_4 = ""

    LOV_SEASON_LABEL = {
        "SS": "Spring-Summer", "FW": "Fall-Winter",
        "AL": "All Season",    "HO": "Holiday",   "AW": "Autumn-Winter",
    }
    sea_label = LOV_SEASON_LABEL.get(sea_prefix, sea_prefix)
    # _w("AT_Season",     sea_label, id_val=sea_prefix)
    _w("AT_Season", id_val=sea_prefix)
    _w("AT_SeasonYear", sea_year_4)

    # ── SAP Article Category ─────────────────────────────────────
    _w("AT_SAPArticleCategory", id_val="0")

    # ── BY Article Type ──────────────────────────────────────────
    _w("AT_BYArticleType", id_val="Inline")

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   id_val="Y")
    _w("AT_SAPIndicator",  id_val="N")
    _w("AT_EcomIndicator", id_val="Y")

    # ── UOM ──────────────────────────────────────────────────────
    _w("AT_UOM", id_val="EA")

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice",       art["rrp"])
    _w("AT_CurrentPrice",        art["rrp"])
    _w("AT_RetailPriceCurrency", id_val="USD")

    # ── Ecommerce content (English) ──────────────────────────────
    _w("AT_EComProductNameEN",  art["model_name"])
    _w("AT_ShortDescriptionEN", art["ecomm_desc_en"])
    _w("AT_LongDescriptionEN",  art["ecomm_desc_en"])

    # ── Technology / Feature bullets ─────────────────────────────
    _w("AT_TechnologyUsed", art["tech_bullets"])


    # ── Main Vendor Identification ───────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # AT_Country — populated from filename country token (e.g. MY, ID, PH)
    country_val = art.get("country_code", "")
    if country_val:
        # _w("AT_Country", country_val, id_val=country_val)
        _w("AT_Country", id_val=country_val)

    # ── Launching Date ───────────────────────────────────────────
    _w("AT_LaunchingDate", art["launching_date"])

    # ── Remaining placeholders ───────────────────────────────────
    for attr_id in NB_FTW_GENERIC_PLACEHOLDERS:
        if attr_id not in written:
            _val(vals_el, attr_id)
            written.add(attr_id)


def _add_variant_values(
    vals_el:      ET.Element,
    art:          dict,
    variant:      dict,
    sap_size_code: str,
) -> None:
    """Write Variant-level <Value> elements for one size/width combo."""
    size_raw = variant["size"]
    width    = variant["width"]

    def _w(attr_id, value="", id_val=""):
        _val(vals_el, attr_id, value=value, id_val=id_val)

    # ── Key variant-defining attributes ──────────────────────────
    _w("AT_Brand",              id_val=art.get("brand_code", "NEW"))
    _w("AT_PrincipalStyleCode", art["article_no"])
    _w("AT_Size",               sap_size_code, id_val=sap_size_code)
    _w("AT_Color",              id_val=art.get("color_token", ""))
    # _w("AT_Width",              id_val=width)

    # ── Size attributes ──────────────────────────────────────────
    _w("AT_PrincipalSizeCode", size_raw)
    _w("AT_SAPSize",           sap_size_code, id_val=sap_size_code)
    _w("AT_PrincipalSize",     variant.get("usa_size") or size_raw)
    _w("AT_EComSize",          f"US {variant.get('usa_size') or size_raw}".strip())

    # ── Country sizes ─────────────────────────────────────────────
    if variant.get("uk_size"):
        _w("AT_UKSize", variant["uk_size"])
    if variant.get("eu_size"):
        _w("AT_EUSize", variant["eu_size"])
    if variant.get("cm_size"):
        _w("AT_CMSize", variant["cm_size"])

    # ── UPC / EAN ─────────────────────────────────────────────────
    if variant.get("upc"):
        _w("AT_PrincipalBarcode", variant["upc"])
    if variant.get("ean"):
        _w("AT_EAN", variant["ean"])

    # ── Packaging ─────────────────────────────────────────────────
    if variant.get("pckg_weight") and variant["pckg_weight"] not in ("0.000", "0", ""):
        _w("AT_PackagingWeight", variant["pckg_weight"])
    if variant.get("pckg_length") and variant["pckg_length"] not in ("0.00", "0", ""):
        _w("AT_PackagingLength", variant["pckg_length"])
    if variant.get("pckg_width") and variant["pckg_width"] not in ("0.00", "0", ""):
        _w("AT_PackagingWidth",  variant["pckg_width"])
    if variant.get("pckg_height") and variant["pckg_height"] not in ("0.00", "0", ""):
        _w("AT_PackagingHeight", variant["pckg_height"])


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_s = season_code[2:] if len(season_code) > 2 else ""
    sea_year   = f"20{sea_year_s}" if len(sea_year_s) == 2 else sea_year_s

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent   = f"CLH_{brand.replace(' ', '')}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_parent)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed,    f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed,    f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


def build_product_xml(
    art:        dict,
    brand:      str,
    brand_code: str,
    comp_code:  str,
    sbu:        str,
    season_id:  str,
    mdd=None,
) -> str:
    """
    Build a <Product> XML fragment for one NB FTW Ecommerce article.

    Structure:
      Product (Generic, UserTypeID = PRD_GenericArticle)
        └─ Product (Variant, UserTypeID = PRD_VariantArticle) × n_sizes

        For FTW, each source article row becomes one Generic and each size becomes
        one Variant under that Generic.

    ParentID for Generic : PPH_E-TempSubCat  (Footwear → division FW)
    KeyID for Generic    : KEY_InboundArticle → {generic_code}
    KeyID for Variant    : KEY_Variant → {generic_code}{sap_color_token}{sap_size_code}
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    parent_id  = "PPH_E-TempSubCat"   # Footwear division placeholder
    key_generic = art["generic_code"]

    # ── Generic element ──────────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["model_name"] or article_no

    # ── Classification references ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Generic-level Values ─────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    # ── Variant sub-products ─────────────────────────────────────
    color_token = art.get("color_token", "")

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    row, brand_code, mdd, season, country_code, file_meta = row_tuple
    mapped = map_article_nb_ftw(row, brand_code=brand_code)  
    if not mapped.get("season"):
        mapped["season"] = season
    mapped["country_code"] = country_code
    if file_meta.get("file_season_code"):                    
        mapped["file_season_code"] = file_meta["file_season_code"]
        mapped["file_season_year"] = file_meta["file_season_year"]
    warns = validate(mapped, mdd)
    return mapped, warns

# ══════════════════════════════════════════════════════════════════
# SECTION 9 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point — called by the unified lambda_function.py dispatcher.

    args must have:
        brand       str   e.g. "New Balance"
        brand_code  str   e.g. "NEW"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "FW"  (or "SP")
        season      str   e.g. "SS26"  (pipeline fallback season)
        seq         int   e.g. 1
        file_type   str   e.g. "Ecommerce File"
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = sorted(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ec_files = sorted(
        f for f in ECOMM_DIR.glob("*.xlsx")
        if "inline" in f.name.lower() or "ecommerce file inline" in f.name.lower().replace("eommerce", "ecommerce")
    )

    # Validate inputs
    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Ecommerce", ec_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise FileNotFoundError(f"No {label} file in expected input directory.")

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="NEW BALANCE")

    log.info(
        "Pipeline → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Build season classification ID ───────────────────────────
    sea_prefix = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_s = args.season[2:] if len(args.season) > 2 else ""
    sea_year   = f"20{sea_year_s}" if len(sea_year_s) == 2 else sea_year_s
    season_id  = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for ec_path in ec_files:
        log.info("─── Processing FTW Ecommerce file: %s ───", ec_path.name)

        ec = NBFTWEcommLoader(ec_path)
        if ec.df.empty:
            log.warning("Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in ec.df.iterrows()]
        # test for only 5 product
        # rows=rows[:5]
        total_rows = len(rows)
        log.info("Total active rows: %d", total_rows)

        # ── Parse file_type and multi_mono from input filename ────
        _stem  = ec_path.stem
        _parts = re.split(r"\s*-\s*", _stem)
        file_type_str  = _parts[3] if len(_parts) >= 7 else getattr(args, "file_type", "Ecommerce File")
        multi_mono_str = _parts[4] if len(_parts) >= 7 else getattr(args, "multi_mono", "Multi")

        # ── Output filename ───────────────────────────────────────
        # FIX: output filename = input filename (stem) + .xml
        out_name = f"{ec_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name
        file_meta = {}
        m_sea = re.search(r'(SS|FW|AW|HO|SM|SP|AL)(\d{4}|\d{2})', ec_path.stem.upper())
        if m_sea:
            file_meta["file_season_code"] = m_sea.group(1)
            file_meta["file_season_year"] = m_sea.group(2)
        # ── Pass 1: parallel map + validate ──────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args = [
            # (row, args.brand_code, mdd, args.season, getattr(args, "country_code", ""))
            (row, args.brand_code, mdd, args.season, getattr(args, "country_code", ""), file_meta)
            for row in rows
        ]

        log.info(
            "Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
            total_rows, num_workers,
        )

        ordered: list[tuple[int, dict, list]] = []
        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {
                pool.submit(_process_article, t): i
                for i, t in enumerate(task_args)
            }
            for fut in as_completed(futures):
                idx           = futures[fut]
                mapped, warns = fut.result()
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered]
        del ordered

        log.info("Pass 1 done — mapped %d articles", len(mapped_articles))

        # ── Pass 2: stream XML ────────────────────────────────────
        log.info("Pass 2/2 — writing XML → %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cls_el  = build_classifications(
            args.brand, args.brand_code, args.comp_code, args.sbu, args.season,
        )
        cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
        del cls_el

        written_count = 0
        with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write(
                f'<STEP-ProductInformation'
                f' xmlns="{STIBO_NS}"'
                f' xmlns:xsi="{STIBO_XSI}"'
                f' xsi:schemaLocation="{STIBO_SCHEMA}"'
                f' ExportTime="{export_time}"'
                f' ExportContext="Context1"'
                f' ContextID="Context1"'
                f' WorkspaceID="Main"'
                f' UseContextLocale="false">\n\n'
            )
            f.write(f"  {cls_str}\n\n")
            del cls_str

            f.write("  <Products>\n")
            for art in mapped_articles:
                if not art.get("article_no"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, args.brand_code,
                    args.comp_code, args.sbu, season_id, mdd=mdd,
                )
                product_xml = _XMLNS_RE.sub("", product_xml)
                f.write(f"    {product_xml}\n")
                del product_xml
                written_count += 1

            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        del mapped_articles

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%d KB)", out_path, file_kb)

        print("═══ ARTICLE SUMMARY (NEW BALANCE INLINE FOOTWEAR ECOMMERCE) ═══", flush=True)
        print(f"  Active rows processed  : {total_rows}",    flush=True)
        print(f"  Generics written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb} KB",    flush=True)
        print("══════════════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ─────────────────────────────────────────
    rpt_path = LOG_DIR / f"nb_ftw_validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")

    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 10 — CLI  (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — NB Inline Footwear Ecommerce v1.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", dest="brand_code", default="NEW")
    p.add_argument("--comp-code",  dest="comp_code",  default="0888")
    p.add_argument("--sbu",        default="FW")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    p.add_argument("--file-type",  dest="file_type",  default="Ecommerce File")
    args = p.parse_args()
    run(args)


if __name__ == "__main__":
    main()
