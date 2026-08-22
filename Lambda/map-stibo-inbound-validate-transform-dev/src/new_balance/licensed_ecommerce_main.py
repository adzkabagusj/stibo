"""
╔══════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — New Balance Licensed  v1.0       ║
║  Ecommerce File (Item tab) → Stibo STEP XML                     ║
╚══════════════════════════════════════════════════════════════════╝

New Balance Licensed Ecommerce-specific differences vs linelist_main:

Source file
  • Single input file  : NB Ecommerce export (e.g. MAP_S226_APP.xlsx)
  • Active sheet       : "Item"
  • Header row         : index 0 (row 1 in Excel) — column names are in row 1
  • Data starts        : index 1 (row 2 in Excel)
  • No channel filter  — all rows are processed

Key column mapping (by column name, not positional index):
    ItemStyleNumber           → Principal Style Code (style+color, e.g. WS41225BK)
    ItemParentProductNumber   → SAP Style Code (style root, e.g. WS41225)
    ItemColor                 → Principal Color Code (e.g. "BLACK (001)")
    ItemFirstColor_en         → Color name (e.g. "BLACK")
    ItemFirstGenericColor     → Color family (e.g. "Black")
    ItemGender                → SAP Gender + BY Gender (Womens→F/Female, Mens→M/Male)
    ProductGender             → Principal Gender Description
    ProductDisplayName        → Principal Style Description / Ecom Product Name EN
    ProductDTCeCommDescription_en → Ecom Short & Long Description EN
    ItemIntroDate             → Launching Date (fallback)
    ItemAPACInLineIntroDate   → Launching Date (preferred)
    ItemSeasons               → Season tokens (parsed to derive season code)
    ItemLatestSeason          → Latest season for season derivation
    ProductPrimaryMaterial    → Material content
    ProductMaterialPercentages → Material percentages
    ProductBullet1_en … ProductBullet10_en → Feature bullets (joined as TechnologyUsed)
    ProductSizeFitBullet1_en … → Size/Fit bullets
    ProductMaterialBullet1_en … → Material bullets
    ItemUpcSyndication        → JSON array: per-variant UPC, EAN, Size, packaging dims
    ItemUniqueIdentifiers     → JSON array: per-variant Size, Width, UPC, EAN, pkg dims
    ItemDTTeCommPrice         → Retail price (JNBO/DTC ecommerce price used as RRP fallback)

Article structure
  • Article Type   : "License"  (same as linelist_main — Licensed channel)
  • Article Category: Generic (1) — each ItemStyleNumber = one Generic product
  • Each row = one Generic (one ItemStyleNumber = style+color combination)
  • Variants       : derived by parsing ItemUpcSyndication JSON
                     Each size entry → one Variant sub-product

Generic code  : NEW + ItemParentProductNumber[:9] + ItemColor code (e.g. NEWWS41225001)
Variant code  : Generic code + 3-char SAP color token + 3-char SAP size code
"""

import re
import sys
import json
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

import os

BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR     = BASE_DIR / "input"
ECOMM_DIR = INPUT_DIR / "ecommerce_licensed"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [ECOMM_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"ecomm_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def _parse_metadata_from_input_filename(stem: str) -> dict:
    """Parse season, Inline/License/SSE article type, and country from input filename."""
    parts = re.split(r"\s*-\s*", stem)
    season = ""
    season_idx = None
    for i, p in enumerate(parts):
        tok = p.strip()
        if re.match(r"^[A-Z]{2}\d{2,4}$", tok, re.IGNORECASE):
            season = tok.upper()
            season_idx = i
            break

    article_type = ""
    for part in parts:
        file_token = part.strip()
        for at in ("Inline", "License", "SSE"):
            if at.lower() in file_token.lower():
                article_type = at
                break
        if article_type:
            break

    country = ""
    if season_idx is not None:
        for p in parts[season_idx + 1:]:
            tok = p.strip()
            if re.match(r"^[A-Z]{2,3}$", tok, re.IGNORECASE) and not tok.isdigit():
                country = tok.upper()
                break

    return {"season": season, "article_type": article_type, "country": country}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES  (shared with linelist_main — duplicated here so this
# module is fully self-contained and importable independently)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS",      "NIK": "NIKE",      "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE",     "ALD": "ALDO",      "CRO": "CROCS",
    "LOT": "LOTTO",       "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults",    "CH": "Children", "IN": "Infant",
    "AA": "All Ages",  "JR": "Junior",
}
LOV_BY_AGE = {
    "AD": "Adult", "ADULT": "Adult", "ADULTS": "Adult",
    "CH": "Children", "CHILDREN": "Children", "CHILD": "Child",
    "AA": "All Ages", "ALL AGES": "All Ages",
    "JR": "Junior",   "JUNIOR": "Junior",
    "MENS": "Adult",  "WOMENS": "Adult",
    "UNISEX": "Adult", "YOUTH": "Kids",
}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter",
}
LOV_COUNTRY_ORIGIN = {
    "CN": "China",      "VN": "Vietnam",    "ID": "Indonesia",
    "KH": "Cambodia",   "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia",   "TH": "Thailand",   "PK": "Pakistan",
    "LK": "Sri Lanka",  "JO": "Jordan",     "SG": "Singapore",
}
LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports",  "FQ": "Footlocker", "FL": "Fashion Footwear",
    "SM": "Smiggle",
}

# ── NB Ecommerce-specific lookup tables ──────────────────────────

# Ecommerce ItemGender → SAP Gender code
NB_ECOMM_GENDER_TO_SAP: dict[str, str] = {
    "WOMENS": "F",
    "MENS":   "M",
    "UNISEX": "U",
    "KIDS":   "U",
    "YOUTH":  "U",
    "BOYS":   "M",
    "GIRLS":  "F",
}

# Ecommerce ItemGender → SAP Age code
NB_ECOMM_GENDER_TO_AGE: dict[str, str] = {
    "WOMENS": "AD",
    "MENS":   "AD",
    "UNISEX": "AD",
    "KIDS":   "CH",
    "YOUTH":  "CH",
    "BOYS":   "CH",
    "GIRLS":  "CH",
}

# Standard SAP size code mapping for NB apparel / accessories
NB_SIZE_TO_SAP_CODE: dict[str, str] = {
    "OSZ":   "000",
    "2XS":   "2XS",
    "XS":    "XSX",
    "S":     "SXX",
    "M":     "MXX",
    "L":     "LXX",
    "XL":    "XLX",
    "2XL":   "2XL",
    "3XL":   "3XL",
    "4XL":   "4XL",
    "5XL":   "5XL",
    "M/L":   "MXL",
    "LXL":   "LXX",
    "S/M":   "SMX",
    "SS":    "SXX",   # "Special Size" fallback — map as S equivalent
    "XXS":   "2XS",
    # Youth
    "3-4Y":   "3Y4",
    "5-6Y":   "5Y6",
    "6-7Y":   "6Y7",
    "7-8Y":   "7Y8",
    "9-10Y":  "9Y0",
    "10-11Y": "AY1",
    "12-13Y": "BY3",
    "14-15Y": "CY5",
    "16Y+":   "DYP",
}

# NB season token normalisation:  "2026S1" → "SS26",  "F25" → "FW25",  "S24" → "SS24"
_SEASON_NORMALIZE_RE = re.compile(
    r"^(\d{4})S([12])$|^([FS])(\d{2,4})$|^(FW|SS|AW|HO|Fall|Spring|Summer|Winter)[\s_]?(\d{2,4})$",
    re.IGNORECASE,
)

NB_AGE_CODE_TO_BY_LOV_ID: dict[str, str] = {
    "AD": "ADULT",
    "CH": "KIDS",
    "AA": "ALL AGES",
    "IN": "INFANT",
    "JR": "KIDS",
}



# ══════════════════════════════════════════════════════════════════
# STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
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

            def _v(key):
                idx = col.get(key)
                if idx is None or idx >= len(row):
                    return None
                v = row[idx]
                return str(v).strip() if v else None

            self.attributes[aid] = {
                "id":           aid,
                "name":         _v("PIM Attribute Name"),
                "source_name":  _v("Source Attribute Name"),
                "cardinality":  _v("Cardinality"),
                "validation":   _v("Validation Base Type"),
                "multi_valued": _v("Multi Valued"),
                "lov_name":     _v("Name of LOV"),
                "max_chars":    _v("Max Characters"),
                "group":        _v("PIM Attribute Group"),
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
                # Zero-pad purely numeric IDs to 3 digits (e.g. 5 → "005", 12 → "012")
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
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"),
            None,
        )
        if hdr_idx is None:
            log.warning("[AttrList] Cannot find header row in sheet '%s'", sheet_name)
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
        log.info(
            "[AttrList] %d attributes loaded from sheet '%s'",
            len(self.attr_map), sheet_name,
        )

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class NBEcommLoader:
    """
    Loads the New Balance Ecommerce export file.

    Active sheet  : "Item"
    Header row    : row 1 (index 0, 0-based) — column names are headers
    Data rows     : row 2+ (index 1+)
    No filter     — all active rows are processed

    Key columns (by name):
        ItemStyleNumber           → style+color SKU key (e.g. WS41225BK)
        ItemParentProductNumber   → style root (e.g. WS41225)
        ItemColor                 → full color string  "BLACK (001)"
        ItemFirstColor_en         → color name "BLACK"
        ItemFirstGenericColor     → generic color "Black"
        ItemGender                → gender label "Womens"/"Mens"
        ProductGender             → product-level gender label
        ProductDisplayName        → product name
        ProductDTCeCommDescription_en → short/long ecom description
        ItemIntroDate             → intro date (fallback)
        ItemAPACInLineIntroDate   → APAC intro date (preferred)
        ItemLatestSeason          → latest season code like "2026S2"
        ItemSeasons               → semi-colon separated seasons "S24;F24;2026S1"
        ItemUpcSyndication        → JSON: per-variant UPC/EAN/Size/packaging
        ItemUniqueIdentifiers     → JSON: per-variant UPC/EAN/Size/packaging (alt)
        ItemDTCeCommJapanRetailPrice  → retail price (IDR/local currency fallback)
        ItemDTTeCommPrice         → DTC ecomm price
        ProductBullet1_en … ProductBullet10_en → feature bullets
        ProductSizeFitBullet1_en … ProductSizeFitBullet7_en → fit bullets
        ProductMaterialBullet1_en … ProductMaterialBullet5_en → material bullets
        ProductPrimaryMaterial    → primary material code
        ProductMaterialPercentages → material percentage string
        ItemStatus                → "Active" / "Inactive"
    """

    SHEET_NAME = "Item"

    def __init__(self, path: Path):
        self.path = path
        self.df   = pd.DataFrame()
        self._load()

    def _load(self):
        log.info("[Ecomm-NBL] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in wb.sheetnames if s.upper() == self.SHEET_NAME.upper()),
            None,
        ) or wb.sheetnames[0]

        log.info("[Ecomm-NBL] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            log.error("[Ecomm-NBL] Empty sheet '%s'", target)
            wb.close()
            return

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[0])
        ]
        df = pd.DataFrame(rows[1:], columns=header)

        # Keep only Active items
        if "ItemStatus" in df.columns:
            before = len(df)
            df = df[
                df["ItemStatus"].astype(str).str.strip().str.upper() == "ACTIVE"
            ]
            log.info(
                "[Ecomm-NBL] Filtered Active rows: %d → %d (dropped %d inactive)",
                before, len(df), before - len(df),
            )

        # Drop rows where ItemStyleNumber is blank
        if "ItemStyleNumber" in df.columns:
            df = df[
                df["ItemStyleNumber"].notna()
                & (~df["ItemStyleNumber"].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        wb.close()
        self.df = df.reset_index(drop=True)
        log.info("[Ecomm-NBL] %d rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — HELPERS
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "N/A") else s


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(raw.split(" ")[0], fmt.split(" ")[0]).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _parse_season_from_tokens(seasons_str: str, latest_season: str) -> str:
    """
    Derive a normalised 2+2 season code (e.g. "SS26", "FW25") from the
    ItemLatestSeason / ItemSeasons column values.

    NB uses two token formats:
      • Old:  "S24", "F24", "S25", "F25"   → SS24, FW24, SS25, FW25
      • New:  "2026S1", "2026S2"             → SS26, FW26
    """
    # Prefer latest season
    token = _s(latest_season) or ""
    if not token and seasons_str:
        parts = [t.strip() for t in _s(seasons_str).split(";") if t.strip()]
        token = parts[-1] if parts else ""

    if not token:
        return ""

    # Pattern: 2026S1 or 2026S2
    m = re.match(r"^(\d{4})S([12])$", token, re.IGNORECASE)
    if m:
        year  = m.group(1)[2:]   # last 2 digits: "26"
        seq   = m.group(2)       # "1" or "2"
        prefix = "SS" if seq == "1" else "FW"
        return f"{prefix}{year}"

    # Pattern: S24, F24, F25
    m = re.match(r"^([SF])(\d{2,4})$", token, re.IGNORECASE)
    if m:
        letter = m.group(1).upper()
        year   = m.group(2)[-2:]   # keep last 2 digits
        prefix = "SS" if letter == "S" else "FW"
        return f"{prefix}{year}"

    # Pattern: FW26, SS26, Fall26, etc.
    m = re.match(
        r"^(FW|SS|AW|HO|Fall|Spring|Summer|Winter)[\s_]?(\d{2,4})$",
        token,
        re.IGNORECASE,
    )
    if m:
        raw_prefix = m.group(1).upper()
        raw_year   = m.group(2)
        map_prefix = {
            "FALL": "FW", "WINTER": "FW", "SPRING": "SS", "SUMMER": "SS",
        }
        prefix = map_prefix.get(raw_prefix, raw_prefix[:2])
        year   = raw_year[-2:]
        return f"{prefix}{year}"

    return token   # return as-is if nothing matched


def _parse_upc_syndication(json_str: str) -> list[dict]:
    """
    Parse ItemUpcSyndication JSON into a list of variant dicts.
    Each dict has: Size, Width, UPC, EAN, UsaSize, NewBalanceSize,
                   PckgWeight, PckgLength, PckgWidth, PckgHeight
    Falls back to ItemUniqueIdentifiers if passed instead.
    """
    if not json_str or str(json_str).strip() in ("", "None", "nan"):
        return []
    try:
        data = json.loads(str(json_str))
    except (json.JSONDecodeError, TypeError):
        log.warning("Failed to parse UPC JSON: %s", str(json_str)[:120])
        return []

    variants = []
    for item in data:
        if isinstance(item, dict):
            # ItemUpcSyndication format
            size = _s(item.get("NewBalanceSize") or item.get("Size", ""))
            variants.append({
                "size":         size,
                "width":        _s(item.get("Width", "")),
                "upc":          _s(item.get("UPC", "")),
                "ean":          _s(item.get("Ean", "")),
                "usa_size":     _s(item.get("UsaSize", "")),
                "sku":          _s(item.get("SKU", "")),
                "pckg_weight":  _s(item.get("Pckg Weight", "")),
                "pckg_length":  _s(item.get("Pckg Length", "")),
                "pckg_width":   _s(item.get("Pckg Width", "")),
                "pckg_height":  _s(item.get("Pckg Height", "")),
            })
    return variants


def _size_to_sap_code(size_raw: str) -> str:
    key = size_raw.strip().upper()
    if key in NB_SIZE_TO_SAP_CODE:
        return NB_SIZE_TO_SAP_CODE[key]
    cleaned = re.sub(r"[^A-Z0-9]", "", key)[:3].ljust(3, "X")
    return cleaned


def _color_to_sap_token(color_str: str) -> str:
    """
    Derive a 3-char SAP color token from the color string.
    "BLACK (001)" → "001";  "ASH HEATHER (047)" → "047";  "BLACK" → "BLK"
    """
    # Extract code from parentheses, e.g. "BLACK (001)" → "001"
    m = re.search(r"\(([A-Z0-9]+)\)", (color_str or "").upper())
    if m:
        code = m.group(1)[:3].ljust(3, "X")
        return code
    # Fallback: first 3 alphanum chars of the color name
    cleaned = re.sub(r"[^A-Z0-9]", "", (color_str or "").upper())[:3].ljust(3, "X")
    return cleaned


def _excel_color_code(color_str: str, fallback: str = "") -> str:
    m = re.search(r"\(([A-Z0-9]+)\)", (color_str or "").upper())
    raw = m.group(1) if m else (fallback or color_str)
    return re.sub(r"[^A-Z0-9]", "", raw.upper())


def _join_bullets(row: dict, prefix: str, count: int) -> str:
    """Join non-empty bullet columns into a single pipe-separated string."""
    parts = []
    for i in range(1, count + 1):
        col = f"{prefix}{i}_en"
        v = _s(row.get(col, ""))
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
# SECTION 3 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_article_nb_ecomm(row: dict, brand_code: str = "NEW") -> dict:
    """
    Map one NB Ecommerce Item row → unified article dict.

    Key decisions:
      - article_no          = ItemStyleNumber  (style+color, e.g. WS41225BK)
      - product_no          = ItemParentProductNumber  (style root, e.g. WS41225)
      - color_str           = ItemColor  (e.g. "BLACK (001)")
      - color_name          = ItemFirstColor_en  (e.g. "BLACK")
      - color_family        = ItemFirstGenericColor  (e.g. "Black")
      - color_token         = 3-char SAP code derived from ItemColor
      - gender_code         = mapped from ItemGender (Womens→F, Mens→M, etc.)
      - age_code            = mapped from ItemGender (Womens→AD, Kids→CH)
      - model_name          = ProductDisplayName
      - ecomm_desc_en       = ProductDTCeCommDescription_en
      - technology_bullets  = ProductBullet1..10 joined
      - size_fit_bullets    = ProductSizeFitBullet1..7 joined
      - material_bullets    = ProductMaterialBullet1..5 joined
      - material_primary    = ProductPrimaryMaterial
      - material_pct        = ProductMaterialPercentages
      - launching_date      = ItemAPACInLineIntroDate (preferred) else ItemIntroDate
      - season              = derived from ItemLatestSeason / ItemSeasons
      - rrp                 = ItemDTTeCommPrice (DTC ecomm price)
            - variants            = parsed from ItemUpcSyndication JSON
            - generic_code        = NEW + ItemStyleNumber[:9]
    """
    article_no     = _s(row.get("ItemStyleNumber"))
    product_no     = _s(row.get("ItemParentProductNumber"))
    color_str      = _s(row.get("ItemColor"))          # "BLACK (001)"
    color_name     = _s(row.get("ItemFirstColor_en"))  # "BLACK"
    color_family   = _s(row.get("ItemFirstGenericColor"))  # "Black"
    gender_raw     = _s(row.get("ItemGender"))
    prod_gender    = _s(row.get("ProductGender"))
    model_name     = _s(row.get("ProductDisplayName"))
    ecomm_desc     = _s(row.get("ProductDTCeCommDescription_en"))
    intro_date     = row.get("ItemIntroDate")
    apac_intro     = row.get("ItemAPACInLineIntroDate")
    seasons_str    = _s(row.get("ItemSeasons", ""))
    latest_season  = _s(row.get("ItemLatestSeason", ""))
    mat_primary    = _s(row.get("ProductPrimaryMaterial", ""))
    mat_pct        = _s(row.get("ProductMaterialPercentages", ""))
    rrp_raw        = row.get("ItemDTTeCommPrice")

    # Bullets
    tech_bullets   = _join_bullets(row, "ProductBullet", 10)
    fit_bullets    = _join_bullets(row, "ProductSizeFitBullet", 7)
    mat_bullets    = _join_bullets(row, "ProductMaterialBullet", 5)

    # Gender / Age
    g_upper      = gender_raw.upper()
    gender_code  = NB_ECOMM_GENDER_TO_SAP.get(g_upper, "U")
    age_code     = NB_ECOMM_GENDER_TO_AGE.get(g_upper, "AD")

    # Color token
    color_token = _color_to_sap_token(color_str or color_name)
    generic_color_code = _excel_color_code(color_str, color_token)

    # Launching Date: prefer APAC intro date
    launching_date = _fmt_date(apac_intro) or _fmt_date(intro_date)

    # Season
    season = _parse_season_from_tokens(seasons_str, latest_season)

    # Prices
    rrp = _price(rrp_raw)

    # Generic code: brand_code + Item Style Number (cleaned and truncated to 9 chars)
    style_clean = re.sub(r"[^A-Z0-9]", "", article_no.upper())[:9]
    generic_code = f"{brand_code}{style_clean}"
    product_no_clean = re.sub(r"[^A-Z0-9]", "", product_no.upper())[:9]

    # Parse variants from UPC Syndication JSON
    upc_json  = row.get("ItemUpcSyndication")
    uid_json  = row.get("ItemUniqueIdentifiers")
    variants  = _parse_upc_syndication(upc_json)
    if not variants:
        variants = _parse_upc_syndication(uid_json)

    return {
        # Core identifiers
        "article_no":    article_no,
        "product_no":    product_no,
        "brand_code":    brand_code,

        # Colour
        "color_str":     color_str,
        "color_name":    color_name,
        "color_family":  color_family,
        "color_token":   color_token,

        # Gender / Age
        "gender_code":   gender_code,
        "age_code":      age_code,
        "gender_raw":    gender_raw,   # original label
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
        "article_type":  "License",   # always License
        "art_category":  "1",         # Generic
        "rrp":           rrp,

        # Dates & Season
        "launching_date": launching_date,
        "season":         season,

        # Variants
        "variants":       variants,

        # Derived codes
        "generic_code":       generic_code,
        "generic_color_code": generic_color_code,
        "sap_style_code":     product_no_clean,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    warns = []
    art   = mapped["article_no"]
    field_checks = {
        "sap_style_code": "AT_PrincipalStyleCode",
        "gender_code": "AT_Gender",
        "age_code":    "AT_SAPAge",
        "brand_code":  "AT_Brand",
    }
    for field, at_id in field_checks.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id}")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(parent, attr_id, value="", id_val="", derived=False):
    if derived:
        return None
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
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


# ── Ecommerce Generic-level placeholder attributes
NB_ECOMM_GENERIC_PLACEHOLDERS = [
    "AT_SPUGrouping",            "AT_SPUGroupingName",
    "AT_Collection1",            "AT_Collection2",
    "AT_Franchise",
    "AT_EComProductNameID",      "AT_EComProductNamePH",
    "AT_EComProductNameMY",      "AT_EComProductNameVN",
    "AT_EComProductNameTH",      "AT_EComProductNameKH",
    "AT_ShortDescriptionID",     "AT_ShortDescriptionPH",
    "AT_ShortDescriptionMY",     "AT_ShortDescriptionVN",
    "AT_ShortDescriptionTH",     "AT_ShortDescriptionKH",
    "AT_LongDescriptionID",      "AT_LongDescriptionPH",
    "AT_LongDescriptionMY",      "AT_LongDescriptionVN",
    "AT_LongDescriptionTH",      "AT_LongDescriptionKH",
    "AT_CareInstructionEN",      "AT_CareInstructionID",
    "AT_CareInstructionPH",      "AT_CareInstructionMY",
    "AT_CareInstructionVN",      "AT_CareInstructionTH",
    "AT_CareInstructionKH",
    "AT_PrincipalColorName",
    "AT_EComAgesCategory",
    "AT_CertificateNumber",
    "AT_ProductWeight",
    "AT_ProductLengthWidthHeight",
    "AT_ImagesSource",
    "AT_EstimatedLandedCost",
    "AT_ExchangeRate",
    "AT_FOB",
    "AT_FOBCurrency",
    "AT_MainEANIndicator",
    "AT_HSCode",
    "AT_Royalty",
    "AT_FreightCost",
    "AT_MarketingFee",
    "AT_SGS",
    "AT_HaddadOfficeCharge",
    "AT_HaddadOfficeChargeDeduction",
    "AT_Others",
    "AT_Commission",
    "AT_Createdon",
    "AT_Updatedon",
    "AT_Updatedby",
]


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
) -> None:
    """Write all Generic-level <Value> elements for a NB Ecommerce article."""
    written: set[str] = set()

    def _w(attr_id, value="", id_val="", derived=False):
        clean_id  = str(id_val).strip() if id_val is not None else ""
        clean_val = str(value).strip() if value is not None else ""
        if not clean_id and not clean_val:
            return
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val, derived=derived)
            written.add(attr_id)

    def _mw(attr_id, id_val):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    sbu_code, _ = _lov(sbu, LOV_SBU, sbu)
    _mw("AT_SBU", sbu_code)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _w("AT_Brand",      id_val=b_code)    # ✅ LOV field — only ID needed
    _w("AT_BrandGroup", id_val=b_label)   # ✅ uses label as the LOV ID

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode", art["sap_style_code"])
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["generic_color_code"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
    _w("AT_InboundGenericCode",                   art["generic_code"])

    # ── Colour ───────────────────────────────────────────────────
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id = color_lov.get(art["color_name"], "") \
                or color_lov.get(art["color_name"].upper(), "") \
                or color_lov.get(art["color_family"], "") \
                or color_lov.get(art["color_family"].upper(), "")
    # Fallback: send empty rather than guess an invalid LOV value
    if not colour_id:
        log.warning(
            "[AT_Color] No LOV match for color_name='%s' / color_family='%s' — sending empty",
            art["color_name"], art["color_family"],
        )
        colour_id = ""
    # Zero-pad safety net (e.g. "5" → "005")
    if colour_id.isdigit():
        colour_id = colour_id.zfill(3)
    _w("AT_Color", id_val=colour_id)          # ← correct: LOV field needs id_val, not value
    _w("AT_PrincipalColorCode", art["generic_color_code"])
    _w("AT_PrincipalColorName", art["color_name"])

    # ── Gender ───────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",          id_val=g_code)
    _w("AT_BYGender",        id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age ──────────────────────────────────────────────────────
    sap_age_code  = art["age_code"]
    sap_age_label = LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge",       id_val=sap_age_code)
    by_age_label = LOV_BY_AGE.get(art["gender_raw"].upper(), "Adult")
    by_age_lov_id = NB_AGE_CODE_TO_BY_LOV_ID.get(sap_age_code, "ADULT")   # add this dict too — see note below
    _w("AT_BYAge",        id_val=by_age_lov_id)
    _w("AT_PrincipalAgeDescription", art["gender_raw"])

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season_from_filename") or art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    if not year_match:
        year_tail = sea_raw[2:]
        full_year = f"20{year_tail}" if len(year_tail) == 2 else year_tail
        _w("AT_SeasonYear", full_year)
    else:
        _w("AT_SeasonYear", year_match.group())
    country_code = art.get("country_from_filename", "")
    _w("AT_Country",  id_val=country_code)

    # ── SAP Article Category ─────────────────────────────────────
    _w("AT_SAPArticleCategory", id_val="0")

    # ── BY Article Type ──────────────────────────────────────────
    _w("AT_BYArticleType", id_val=art.get("article_type_from_filename") or art.get("article_type") or "License")

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
    _w("AT_EComProductNameEN",   art["model_name"])
    _w("AT_ShortDescriptionEN",  art["ecomm_desc_en"])
    _w("AT_LongDescriptionEN",   art["ecomm_desc_en"])

    # ── Technology / Material ────────────────────────────────────
    _w("AT_TechnologyUsed",      art["tech_bullets"])

    # ── Main Vendor Identification ───────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Launching Date ───────────────────────────────────────────
    _w("AT_LaunchingDate", art["launching_date"])

    # ── Remaining placeholders intentionally omitted when empty ──


def _add_variant_values(
    vals_el: ET.Element,
    art:     dict,
    variant: dict,
    sap_size_code: str,
) -> None:
    """Write Variant-level <Value> elements for one size from UPC Syndication."""
    size_raw = variant["size"]

    def _w(attr_id, value="", id_val=""):
        _val(vals_el, attr_id, value=value, id_val=id_val)

    # ── KEY_Variant defining attributes (REQUIRED for key resolution) ──
    _w("AT_Brand",              id_val=art.get("brand_code", "NEW"))
    _w("AT_PrincipalStyleCode", art["sap_style_code"])
    _w("AT_Size",               sap_size_code, id_val=sap_size_code)
    _w("AT_Color",              id_val=art.get("color_token", ""))

    # ── Size attributes ───────────────────────────────────────────
    _w("AT_PrincipalSizeCode",  size_raw)
    _w("AT_SAPSize",            sap_size_code, id_val=sap_size_code)
    _w("AT_PrincipalSize",      variant.get("usa_size") or size_raw)
    _w("AT_EComSize",           f"US {sap_size_code}".strip())

    # UPC / EAN
    if variant.get("upc"):
        _w("AT_PrincipalBarcode", variant["upc"])
    if variant.get("ean"):
        _w("AT_EAN", variant["ean"])

    # Packaging dimensions
    if variant.get("pckg_weight"):
        _w("AT_PackagingWeight", variant["pckg_weight"])
    if variant.get("pckg_length"):
        _w("AT_PackagingLength", variant["pckg_length"])
    if variant.get("pckg_width"):
        _w("AT_PackagingWidth",  variant["pckg_width"])
    if variant.get("pckg_height"):
        _w("AT_PackagingHeight", variant["pckg_height"])


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand.replace(' ', '')}Batches"

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
    ET.SubElement(confirmed,    f"{{{STIBO_NS}}}Name").text = (
        f"{season_short} Confirmed Articles"
    )

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed,    f"{{{STIBO_NS}}}Name").text = (
        f"{season_short} Unconfirmed Articles"
    )

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
    Build a <Product> XML fragment for one NB Ecommerce article.

    Structure:
      Product (Generic, UserTypeID = PRD_GenericArticle)
        └─ Product (Variant, UserTypeID = PRD_VariantArticle) × n_sizes

    ParentID for Generic  : PPH_E-TempSubCat  (Ecommerce → division E fallback)
    KeyID for Generic     : KEY_InboundArticle -> {generic_code}
    Generic Code attr     : AT_InboundGenericCode -> {generic_code}
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = art.get("brand_code", brand_code) or brand_code
    parent_id  = "PPH_E-TempSubCat"   # Ecommerce articles default to division E

    key_generic = art["generic_code"]

    # ── Generic element ──────────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["model_name"] or article_no
    )

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

   

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one NB Ecommerce row."""
    row, brand_code, mdd, season = row_tuple
    mapped = map_article_nb_ecomm(row, brand_code=brand_code)
    # Override season with pipeline arg if article has none
    if not mapped.get("season"):
        mapped["season"] = season
    warns = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

    args must have:
        brand       str   e.g. "New Balance"
        brand_code  str   e.g. "NEW"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SP"
        season      str   e.g. "SS26"   (used as fallback when article has no season)
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f     = first(MDD_DIR)
    attr_f    = first(ATTR_DIR)
    ec_files = [
        f for f in ECOMM_DIR.glob("*.xlsx")
        if "inline" not in f.name.lower()
    ]

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Ecommerce", ec_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise FileNotFoundError(f"No {label} file found in expected input directory.")

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="NEW BALANCE")

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Season ID ────────────────────────────────────────────────
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for ec_path in ec_files:
        log.info("─── Processing Ecommerce file: %s ───", ec_path.name)
        filename_meta = _parse_metadata_from_input_filename(ec_path.stem)

        ec = NBEcommLoader(ec_path)
        if ec.df.empty:
            log.warning("[Ecomm-NBL] Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in ec.df.iterrows()]
        # ── TEST MODE: limit to 10 products ──────────────────────
        # rows = rows[:10]
        # ────────────────────────────────────────────────────────
        total_rows = len(rows)
        log.info("Total Active rows: %d", total_rows)

        # ── Output filename: keep exactly same as input stem ─────
        out_name = f"{ec_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (row, args.brand_code, mdd, args.season) for row in rows
        ]
        ordered: list[tuple[int, dict, list]] = []

        log.info(
            "Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
            total_rows, num_workers,
        )

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
        for _m in mapped_articles:
            _m["article_type_from_filename"] = filename_meta.get("article_type", "")
            _m["season_from_filename"] = filename_meta.get("season", "")
            _m["country_from_filename"] = filename_meta.get("country", "")
        del ordered

        log.info("Pass 1 done — mapped=%d articles", len(mapped_articles))

        # ── Pass 2: stream XML to file ───────────────────────────
        log.info("Pass 2/2 — streaming XML to %s …", out_name)
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
        log.info("✓ XML written → %s  (%dKB)", out_path, file_kb)

        print("═══ ARTICLE SUMMARY (NEW BALANCE LICENSED ECOMMERCE) ═════════", flush=True)
        print(f"  Ecommerce Active rows  : {total_rows}",    flush=True)
        print(f"  Generics written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb}KB",     flush=True)
        print("══════════════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"ecomm_validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — CLI (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — New Balance Licensed Ecommerce v1.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
