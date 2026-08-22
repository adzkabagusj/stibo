"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — New Balance Licensed  v1.0     ║
║   Pricelist (Pricelist tab, Licensed rows) → Stibo STEP XML    ║
╚══════════════════════════════════════════════════════════════════╝

New Balance Licensed-specific differences vs Smiggle / Adidas:

Source file
  • Single input file  : "Inline Apparel & Licensed Accessories Price List – APAC"
  • Active sheet       : "Pricelist"
  • Header row         : index 3 (0-based)   → row 4 in Excel
  • Data starts        : index 4 (0-based)   → row 5 in Excel
  • Filter             : Column "Channel Type" == "Licensed"
    (the same pricelist contains Inline rows which are ignored here)

Column mapping (0-indexed positions confirmed from file inspection):
    col[2]  = Product Number       → Principal Style Code
    col[2]  = Product Number       → SAP Style Code (style without colour)
    col[3]  = Color Code           → Principal Color Code
    col[4]  = GBU                  → (unused)
    col[21] = Product Line         → Principal Merchandise Hierarchy L1
    col[5]  = Line Plan Business   → SAP Product Group / Sports Category
    col[6]  = Category             → Principal Merchandise Hierarchy L2
    col[7]  = Channel Type         → filter = "Licensed"
    col[8]  = LPA Category         → BY Sub Category / franchise grouping
    col[9]  = Merchandise Category → SAP Product Category cross-check
    col[12] = Size Profile         → SAP Gender + SAP Age
    col[13] = Global Intro Date    → Launching Date (fallback)
    col[14] = Region Intro Date    → Launching Date (preferred)
    col[16] = Product Display Name → Principal Style Description
    col[17] = Item - CarryOver/New → Nature of Article
    col[18] = Color Name           → Principal Color Description / SAP Color
    col[19] = Color Family         → Principal Color Description (simplified)
    col[21] = Product Line         → SAP Product Division (Apparel=A, Accessories=E)
    col[22] = Collection           → Collection 1
    col[23] = Silhouette           → SAP Product Category / BY Sub Category
    col[24] = Sizes                → Principal Size Code (comma-separated)
    col[29] = Earliest X-Factory Date → additional date reference
    col[30] = Fit Version          → Country Size mapping (Western Fit=US, Asia Fit=ASIA)
    col[31] = Technologies         → Technology Used
    col[32] = Primary Fabric Content  → Material (Content)
    col[34] = Global Retail Price  → RRP (USD)
    col[35] = NBIL Price           → FOB (USD)
    col[36] = Supplier             → vendor name reference
    col[37] = Production Address Name → vendor address code
    col[38] = COO                  → Country of Origin (full name, mapped to ISO)
    col[43] = Date Added           → date reference

Article structure
  • Article Type   : always "License"  (Channel Type = Licensed)
  • Article Category: Generic (1) — NB accessories/apparel have a size run
  • Each row = one Generic (Product Number + Color Code combination)
  • Variants       : derived by splitting the comma-separated "Sizes" cell
                     Each size → one Variant sub-product

Generic code  : NEW + Product Number (e.g. NEWLAS51363)
Variant code  : Generic + 3-char SAP color token + 3-char SAP size code
"""

import re
import sys
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

INPUT_DIR      = BASE_DIR / "input"
PRICELIST_DIR  = INPUT_DIR / "linelist_licensed"
MDD_DIR        = INPUT_DIR / "mdd"
ATTR_DIR       = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PRICELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════
# LOV TABLES
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
    # NB pricelist Size Profile values
    "MENS": "Adult",   "WOMENS": "Adult",
    "UNISEX": "Adult", "YOUTH": "Kids",
}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SAP_ARTICLE_CATEGORY = {
    "1": "Generic", "0": "Single", "10": "Sell set (Hampers)"
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

# ── New Balance Licensed-specific lookup tables ───────────────────

# COO country full name → ISO 2-char code (from pricelist COO column)
NB_COO_TO_ISO: dict[str, str] = {
    "CHINA":        "CN",
    "VIETNAM":      "VN",
    "INDONESIA":    "ID",
    "CAMBODIA":     "KH",
    "BANGLADESH":   "BD",
    "INDIA":        "IN",
    "MALAYSIA":     "MY",
    "THAILAND":     "TH",
    "PAKISTAN":     "PK",
    "SRI LANKA":    "LK",
    "JORDAN":       "JO",
    "JORDON":       "JO",     # typo seen in source data
    "SINGAPORE":    "SG",
}

# Size Profile column → SAP Gender code
NB_SIZE_PROFILE_TO_GENDER: dict[str, str] = {
    "MENS":   "M",
    "WOMENS": "F",
    "UNISEX": "U",
    "YOUTH":  "U",     # Youth = Unisex gender in SAP
}

# Size Profile column → SAP Age code
NB_SIZE_PROFILE_TO_AGE: dict[str, str] = {
    "MENS":   "AD",
    "WOMENS": "AD",
    "UNISEX": "AD",
    "YOUTH":  "CH",
}

# NEW — add this near the other NB lookup tables
NB_AGE_CODE_TO_BY_LOV_ID: dict[str, str] = {
    "AD": "ADULT",
    "CH": "KIDS",
    "AA": "ALL AGES",
    "IN": "INFANT",
    "JR": "KIDS",
}

# Product Line → SAP Division letter
NB_PRODUCT_LINE_TO_DIVISION: dict[str, str] = {
    "APPAREL":     "A",
    "ACCESSORIES": "E",
    "FOOTWEAR":    "F",
}

# Line Plan Business → SAP Product Group (2-char)
NB_LINE_PLAN_TO_PROD_GROUP: dict[str, str] = {
    "RUNNING":          "RU",
    "TRAINING":         "TR",
    "BASKETBALL":       "BK",
    "LIFESTYLE":        "LS",
    "OTHER FOP":        "OT",
    "KIDS LIFESTYLE":   "KL",
    "KIDS PERFORMANCE": "KP",
    "TENNIS":           "TE",
    "BASEBALL":         "BA",
    "GOLF":             "GO",
    "SWIMMING":         "SW",
    "SOFTBALL":         "SF",
    "SKATE":            "SK",
}

# Silhouette → SAP Product Category (2-char)
NB_SILHOUETTE_TO_PROD_CATEGORY: dict[str, str] = {
    "S/S TOP":                   "ST",
    "L/S TOP":                   "LT",
    "PANT":                      "PT",
    "SHORT":                     "SH",
    "TIGHT":                     "TG",
    "DRESS":                     "DR",
    "HOODIES & SWEATSHIRTS":     "HS",
    "JACKET":                    "JK",
    "TANKS/SLEEVELESS/SINGLET":  "TK",
    "FITNESS":                   "FT",
    "SOCKS":                     "SK",
    "HEADWEAR":                  "HW",
    "BAG":                       "BG",
    "ARM SLEEVE":                "AS",
    "UNDERWEAR":                 "UW",
    "GLOVE":                     "GL",
    "VEST":                      "VT",
}

# Fit Version → Country Size (for AT_CountrySize)
NB_FIT_TO_COUNTRY_SIZE: dict[str, str] = {
    "WESTERN FIT": "US",
    "ASIA FIT":    "ASIA",
    # "N/A" → no value (attribute not sent when blank)
}

# Item CarryOver/New → Nature of Article
NB_ITEM_STATUS_TO_NOA: dict[str, str] = {
    "NEW":      "REG",    # Regular
    "CARRYOVER": "CO",    # Carry Over
}

# Line Plan Business → Sports Category EN (for AT_SportsCategory / BY)
NB_LINE_PLAN_TO_SPORTS_CATEGORY: dict[str, str] = {
    "RUNNING":          "Running",
    "TRAINING":         "Running",
    "OTHER FOP":        "Other",
    "LIFESTYLE":        "Lifestyle / Casual",
    "BASKETBALL":       "Basketball",
    "KIDS LIFESTYLE":   "Lifestyle / Casual",
    "KIDS PERFORMANCE":      "Running",
    "TENNIS":               "Tennis / Padel",
    "BASKETBALL AND SOFTBALL": "Other",
    "SKATE":                "Skateboarding",
}

# Standard SAP size code mapping for NB apparel / accessories
# OSZ = One Size; S/M/L/XL mapped directly; numeric youth sizes get prefix Y
NB_SIZE_TO_SAP_CODE: dict[str, str] = {
    "OSZ":   "000",   # One Size / No Size
    "2XS":   "2XS",
    "XS":    "XSX",
    "S":     "SXX",
    "M":     "MXX",
    "L":     "LXX",
    "XL":    "XLX",
    "2XL":   "2XL",
    "3XL":   "3XL",
    "M/L":   "MXL",
    "LXL":   "LXX",
    "S/M":   "SMX",
    # Youth sizes
    "3-4Y":    "3Y4",
    "5-6Y":    "5Y6",
    "6-7Y":    "6Y7",
    "7-8Y":    "7Y8",
    "9-10Y":   "9Y0",
    "10-11Y":  "AY1",
    "12-13Y":  "BY3",
    "14-15Y":  "CY5",
    "16Y+":    "DYP",
    # Sock sizes
    "S,M,L,XL": "SMX",   # fallback if whole string lands here
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

        # Prefer exact match for "NEW BALANCE", skip v4 variant
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


class NBLicensedPricelistLoader:
    """
    Loads the New Balance "Inline Apparel & Licensed Accessories" pricelist.

    Active sheet  : "Pricelist"
    Header row    : index 3 (0-based)
    Data rows     : index 4+ (0-based)
    Filter        : Channel Type (col[7]) == "Licensed"

    Each row represents one item number (Product Number + Color Code),
    which maps to one Generic in STEP with Variants per size.
    """

    SHEET_NAME = "Pricelist"
    HDR_ROW    = 3   # 0-based

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = self.SHEET_NAME
        self._load()

    def _load(self):
        log.info("[PriceList-NBL] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Find target sheet (exact match first, then partial)
        target = next(
            (s for s in wb.sheetnames if s.upper() == self.SHEET_NAME.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames if "PRICE" in s.upper()),
            wb.sheetnames[0],
        )

        log.info("[PriceList-NBL] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        if len(rows) <= self.HDR_ROW:
            log.error("[PriceList-NBL] Not enough rows in sheet '%s'", target)
            wb.close()
            return

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[self.HDR_ROW])
        ]
        df = pd.DataFrame(rows[self.HDR_ROW + 1:], columns=header)

        # Filter to Licensed rows only
        channel_col = next(
            (c for c in df.columns if "channel" in c.lower()), None
        )
        if channel_col:
            before = len(df)
            df = df[
                df[channel_col].astype(str).str.strip().str.upper() == "LICENSED"
            ]
            log.info(
                "[PriceList-NBL] Filtered Licensed rows: %d → %d (dropped %d non-Licensed)",
                before, len(df), before - len(df),
            )
        else:
            log.warning(
                "[PriceList-NBL] 'Channel Type' column not found — processing all rows"
            )

        # Drop rows where Item Number is blank
        item_col = next(
            (c for c in df.columns if "item number" in c.lower()), None
        )
        if item_col:
            df = df[
                df[item_col].notna()
                & (~df[item_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info("[PriceList-NBL] %d Licensed rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — HELPERS
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "N/A") else s


def _s_keep_na(v) -> str:
    """Like _s but keeps 'N/A' as empty (alias for clarity)."""
    return _s(v)


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _size_to_sap_code(size_raw: str) -> str:
    """
    Convert a single NB size string to a 3-char SAP size code.
    Falls back to a truncated/padded version of the raw value.
    """
    key = size_raw.strip().upper()
    if key in NB_SIZE_TO_SAP_CODE:
        return NB_SIZE_TO_SAP_CODE[key]
    # Generic fallback: strip non-alphanumeric, take first 3 chars, pad to 3
    cleaned = re.sub(r"[^A-Z0-9]", "", key)[:3].ljust(3, "X")
    return cleaned


def _color_to_sap_token(color_name: str, color_code: str) -> str:
    """
    Derive a 3-char SAP Color token from the color code.
    e.g. "BM4" → "BM4", "AS1" → "AS1", "BK" → "BKX"
    """
    raw = (color_code or color_name or "").strip().upper()
    token = re.sub(r"[^A-Z0-9]", "", raw)[:3].ljust(3, "X")
    return token


def _coo_to_iso(coo_raw: str) -> str:
    """Map COO full country name (from pricelist) to ISO 2-char code."""
    return NB_COO_TO_ISO.get(coo_raw.strip().upper(), "CN")  # default CN


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_article_nb_licensed(row: dict, brand_code: str = "NEW") -> dict:
    """
    Map one New Balance Licensed pricelist row → unified article dict.

    Key decisions (per New Balance attributes sheet):
      - article_no         = Item Number  (Product Number + "_" + Color Code)
      - sap_style_code     = Product Number   (9-digit style)
      - color_code         = Color Code   (e.g. BK, WT, AS1)
      - model_name         = Product Display Name
      - gender_code        = from Size Profile (Mens→M, Womens→F, Unisex/Youth→U)
      - age_code           = from Size Profile (Mens/Womens/Unisex→AD, Youth→CH)
      - coo                = from COO column (full name → ISO code)
      - div_letter         = from Product Line (Apparel→A, Accessories→E)
      - prod_group         = from Line Plan Business (Running→RU, etc.)
      - prod_category      = from Silhouette (S/S Top→ST, Socks→SK, etc.)
      - article_type       = always "License"  (Licensed channel)
      - art_category       = "1"  (Generic — has size variants)
      - rrp                = Global Retail Price (USD)
      - fob                = NBIL Price (USD)
      - launching_date     = Region Intro Date (preferred) else Global Intro Date
      - country_size       = from Fit Version (Western Fit→US, Asia Fit→ASIA)
      - technology         = Technologies column
      - collection1        = Collection column
      - lpa_category       = LPA Category (sub-franchise grouping)
      - sizes_raw          = Sizes column (comma-separated, used to build variants)
      - nature_of_article  = CarryOver→CO, New→REG
    """
    item_number  = _s(row.get("Item Number"))
    product_no   = _s(row.get("Product Number"))
    color_code   = _s(row.get("Color Code"))
    gbu          = _s(row.get("GBU"))
    line_plan    = _s(row.get("Line Plan Business"))
    category     = _s(row.get("Category"))
    lpa_category = _s(row.get("LPA Category"))
    merch_cat    = _s(row.get("Merchandise Category"))
    size_profile = _s(row.get("Size Profile"))
    global_intro = row.get("Global Intro Date")
    region_intro = row.get("Region Intro Date")
    model_name   = _s(row.get("Product Display Name"))
    item_status  = _s(row.get("Item - CarryOver/New"))
    color_name   = _s(row.get("Color Name"))
    color_family = _s(row.get("Color Family"))
    product_line = _s(row.get("Product Line"))
    collection   = _s(row.get("Collection"))
    silhouette   = _s(row.get("Silhouette"))
    sizes_raw    = _s(row.get("Sizes"))
    fit_version  = _s(row.get("Fit Version"))
    technologies = _s(row.get("Technologies"))
    fabric_primary = _s(row.get("Primary Fabric Content "))    # trailing space in col name
    rrp_raw      = row.get("Global Retail Price")
    fob_raw      = row.get("NBIL Price")
    coo_raw      = _s(row.get("COO"))

    # ── Derived fields ───────────────────────────────────────────

    # Gender + Age from Size Profile
    sp_upper     = size_profile.upper()
    gender_code  = NB_SIZE_PROFILE_TO_GENDER.get(sp_upper, "U")
    age_code     = NB_SIZE_PROFILE_TO_AGE.get(sp_upper, "AD")

    # SAP Product Division letter
    div_letter = NB_PRODUCT_LINE_TO_DIVISION.get(
        product_line.upper(), "A"
    )

    # SAP Product Group (2-char) from Line Plan Business
    prod_group = NB_LINE_PLAN_TO_PROD_GROUP.get(
        line_plan.upper().strip(), "OT"
    )

    # SAP Product Category (2-char) from Silhouette
    prod_category = NB_SILHOUETTE_TO_PROD_CATEGORY.get(
        silhouette.upper().strip(), "OT"
    )

    # COO → ISO code
    coo = _coo_to_iso(coo_raw)

    # Launching Date: prefer Region Intro Date
    launching_date = _fmt_date(region_intro) or _fmt_date(global_intro)

    # Country Size from Fit Version (empty default = attribute not sent when N/A)
    country_size = NB_FIT_TO_COUNTRY_SIZE.get(
        fit_version.upper().strip(), ""
    )

    # Nature of Article
    noa_code = NB_ITEM_STATUS_TO_NOA.get(item_status.upper().strip(), "REG")

    # Sports Category
    sports_category = NB_LINE_PLAN_TO_SPORTS_CATEGORY.get(
        line_plan.upper().strip(), "Other"
    )

    # Prices — extract numeric part
    def _price(v) -> str:
        if v is None:
            return ""
        try:
            f = float(v)
            return "" if f == 0.0 else str(round(f, 2))
        except (TypeError, ValueError):
            m = re.search(r"[\d.]+", str(v))
            return m.group() if m else str(v)

    rrp = _price(rrp_raw)
    fob = _price(fob_raw)

    # SAP Color token (3-char)
    color_token = _color_to_sap_token(color_name, color_code)

    # Generic code: brand_code (3) + Product Number (up to 9 chars) + Color Code (full)
    # Updated to match accessories/footwear pattern
    product_no_clean = re.sub(r"[^A-Z0-9]", "", product_no.upper())[:9]
    color_clean = re.sub(r"[^A-Z0-9]", "", color_code.upper())
    generic_code = f"{brand_code}{product_no_clean}{color_clean}"

    # Parse sizes into a list
    sizes_list: list[str] = []
    if sizes_raw:
        sizes_list = [s.strip() for s in sizes_raw.split(",") if s.strip()]

    return {
        # Core identifiers
        "article_no":       item_number,          # full Item Number = style_colorcode
        "product_no":       product_no,           # Product Number (style)
        "color_code":       color_code,           # Color Code raw (e.g. BK, AS1)
        "model_name":       model_name,
        "brand_code":       brand_code,

        # Colour
        "color_name":       color_name,           # e.g. "BLACK (001)"
        "color_family":     color_family,         # e.g. "BLACK"
        "color_token":      color_token,          # 3-char SAP token

        # Gender / Age
        "gender_code":      gender_code,
        "age_code":         age_code,
        "size_profile":     size_profile,         # raw principal gender description

        # Classification
        "div_letter":       div_letter,           # SAP Product Division letter
        "prod_group":       prod_group,           # SAP Product Group 2-char
        "prod_category":    prod_category,        # SAP Product Category 2-char
        "silhouette":       silhouette,           # raw silhouette label
        "product_line":     product_line,         # raw Product Line

        # Hierarchy
        "gbu":              gbu,                  # GBU (not used for hierarchy)
        "line_plan":        line_plan,            # Line Plan Business = L2
        "lpa_category":     lpa_category,         # LPA Category
        "merch_cat":        merch_cat,            # Merchandise Category
        "collection1":      collection if collection and collection != "N/A" else "",

        # Commercial
        "article_type":     "License",            # always License for Licensed channel
        "art_category":     "1",                  # Generic (has variants)
        "noa_code":         noa_code,             # Nature of Article
        "item_status":      item_status,          # CarryOver / New

        # Pricing (USD)
        "rrp":              rrp,
        "fob":              fob,

        # Origin + fit
        "coo":              coo,
        "coo_raw":          coo_raw,
        "country_size":     country_size,
        "fit_version":      fit_version,

        # Dates
        "launching_date":   launching_date,

        # Technology + Material
        "technologies":     technologies if technologies != "N/A" else "",
        "fabric_primary":   fabric_primary,
        "sports_category":  sports_category,

        # Sizes
        "sizes_raw":        sizes_raw,
        "sizes_list":       sizes_list,

        # Derived codes
        "generic_code":     generic_code,
        "sap_style_code":   product_no_clean,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Validate mandatory MDD attributes against the mapped article dict."""
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

def _val(
    parent:  ET.Element,
    attr_id: str,
    value:   str  = "",
    id_val:  str  = "",
    derived: bool = False,
) -> ET.Element | None:
    if derived:
        return None
    clean_id = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip() if value else ""
    if not clean_id and not clean_val:
        return None
    if clean_id in ("None", "nan"):
        clean_id = ""
    if clean_val in ("None", "nan"):
        clean_val = ""
    if not clean_id and not clean_val:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
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


def _parse_metadata_from_filename(path: Path) -> dict[str, str]:
    """
    Parse metadata from a standard inbound filename:
    0888-SP-New Balance-Line List Licensed-Multi-FW26-MY-1.xlsx
    """
    stem = path.stem
    parts = re.split(r"\s*-\s*", stem)
    data = {
        "file_type": "Line List",
        "multi_mono": "Multi",
        "season": "",
        "country": "",
    }
    if len(parts) >= 8:
        data["file_type"] = parts[3].strip() or data["file_type"]
        data["multi_mono"] = parts[4].strip() or data["multi_mono"]
        data["season"] = parts[5].strip()
        data["country"] = parts[6].strip().upper()
    return data


# ── NB Licensed Generic-level placeholder attributes
NB_LICENSED_GENERIC_PLACEHOLDERS = [
    "AT_SPUGrouping",            "AT_SPUGroupingName",
    "AT_Collection1",            "AT_Collection2",
    "AT_Franchise",
    "AT_EComProductNameEN",      "AT_EComProductNameID",
    "AT_EComProductNamePH",      "AT_EComProductNameMY",
    "AT_EComProductNameVN",      "AT_EComProductNameTH",
    "AT_EComProductNameKH",
    "AT_ShortDescriptionEN",     "AT_ShortDescriptionID",
    "AT_ShortDescriptionPH",     "AT_ShortDescriptionMY",
    "AT_ShortDescriptionVN",     "AT_ShortDescriptionTH",
    "AT_ShortDescriptionKH",
    "AT_LongDescriptionEN",      "AT_LongDescriptionID",
    "AT_LongDescriptionPH",      "AT_LongDescriptionMY",
    "AT_LongDescriptionVN",      "AT_LongDescriptionTH",
    "AT_LongDescriptionKH",
    "AT_CareInstructionEN",      "AT_CareInstructionID",
    "AT_CareInstructionPH",      "AT_CareInstructionMY",
    "AT_CareInstructionVN",      "AT_CareInstructionTH",
    "AT_CareInstructionKH",
    "AT_PrincipalColorName",
    "AT_EComAgesCategory",
    "AT_TechnologyUsed",
    "AT_CertificateNumber",
    "AT_ProductWeight",
    "AT_ProductLengthWidthHeight",
    "AT_PackagingLength",        "AT_PackagingWidth",
    "AT_PackagingHeight",        "AT_PackagingWeight",
    "AT_ImagesSource",
    "AT_EstimatedLandedCost",
    "AT_ExchangeRate",
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
    "AT_LaunchingDate",
    "AT_Createdon",
    "AT_Updatedon",
    "AT_Updatedby",
]


def _add_generic_values(
    vals_el:   ET.Element,
    art:       dict,
    brand_name: str,
    comp_code: str,
    sbu:       str,
    mdd=None
) -> None:
    """
    Write all Generic-level <Value> elements for a New Balance Licensed article.
    Order mirrors the Stibo attribute group sequence in the MDD.
    """
    written: set[str] = set()

    def _w(attr_id, value="", id_val="", derived=False):
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
    _w("AT_Brand", id_val=b_code)
    _w("AT_BrandGroup",id_val=b_label)

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["sap_style_code"])
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["color_code"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
    _w("AT_InboundGenericCode",                   art["generic_code"])

    # ── Colour (SAP Color) ───────────────────────────────────────
    color_display = art["color_name"].upper()  # value stays the same
    # Look up ID from MDD Color Code LOV by matching display name
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        # color_lov is {display_name: id_code}
        colour_id = color_lov.get(art["color_name"], "") \
                or color_lov.get(art["color_name"].upper(), "") \
                or color_lov.get(art["color_family"], "") \
                or color_lov.get(art["color_family"].upper(), "")
    # Fallback if still not found
    if not colour_id:
        colour_id = re.sub(r"[^A-Z0-9]", "", art["color_family"].upper())[:6] \
                    or art["color_code"].upper()
    if colour_id.isdigit():
        colour_id = colour_id.zfill(3)
    _w("AT_Color", id_val=colour_id)
    _w("AT_PrincipalColorCode",   art["color_code"])

    # ── Gender ───────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",     id_val=g_code)
    _w("AT_BYGender",   id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["size_profile"])     # raw Size Profile as principal gender description

    # ── Age ──────────────────────────────────────────────────────
    sap_age_code  = art["age_code"]
    sap_age_label = LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge",   id_val=sap_age_code)
    by_age_label = LOV_BY_AGE.get(art["size_profile"].upper(), "Adult")
    by_age_lov_id = NB_AGE_CODE_TO_BY_LOV_ID.get(sap_age_code, "A")
    _w("AT_BYAge", id_val=by_age_lov_id)   # sends "A" or "K" — correct
    _w("AT_PrincipalAgeDescription", art["size_profile"])

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season",  id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    _w("AT_SeasonYear", year_match.group() if year_match else "")

    # ── Country of Origin ────────────────────────────────────────
    coo_code, coo_label = _lov(art["coo"], LOV_COUNTRY_ORIGIN, art["coo"])
    _w("AT_CountryOrigin", id_val=coo_code)

    # ── SAP Article Category — Generic (1) — has size variants ──
    _w("AT_SAPArticleCategory", id_val="1")

    # ── BY Article Type — always License for Licensed channel ────
    _w("AT_BYArticleType", id_val="License")

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   id_val="Y")
    _w("AT_SAPIndicator", id_val="N")
    _w("AT_EcomIndicator", "")

    # ── UOM — always EA for accessories/apparel ──────────────────
    _w("AT_UOM",id_val="EA")

    # ── Material Type — ZINA ─────────────────────────────────────
    _w("AT_MaterialType", id_val="ZINA")

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice",       art["rrp"])
    _w("AT_CurrentPrice",        art["rrp"])
    currency_lov = mdd.lovs.get("Retail Price Currency", {}) if mdd else {}
    # currency_lov is {description: id}, e.g. {"United States Dollar": "USD"}
    usd_id = currency_lov.get("United States Dollar", "USD")
    _w("AT_RetailPriceCurrency", id_val=usd_id)
    _w("AT_FOBCurrency",         id_val=usd_id)
    _w("AT_FOB",                 art["fob"])
    

    # ── Nature of Article ────────────────────────────────────────
    noa_label = {"REG": "Regular", "CO": "Carry Over"}.get(
        art["noa_code"], "Regular"
    )
    _w("AT_NatureOfArticle", id_val=art["noa_code"])

    # ── Vendor identification ─────────────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Hierarchy / Classification ───────────────────────────────
    _w("AT_PrincipalMerchandiseHierarchyL1", art["product_line"])
    _w("AT_PrincipalMerchandiseHierarchyL2", art["line_plan"])
    _w("AT_PrincipalMerchandiseHierarchyL3", art["silhouette"])

    # ── Collections ──────────────────────────────────────────────
    _w("AT_Collection1", art["collection1"])
    # _w("AT_Franchise",   art["lpa_category"])    # LPA Category as Franchise grouping

    # ── Launching Date ───────────────────────────────────────────
    _w("AT_LaunchingDate", art["launching_date"])

    # ── Technology ───────────────────────────────────────────────
    _w("AT_TechnologyUsed", art["technologies"])

    # ── Country Size (from Fit Version) ─────────────────────────
    # Only write if value exists (skip when blank/N/A)
    if art["country_size"]:
        _w("AT_CountrySize", art["country_size"])

    # ── Sports Category ──────────────────────────────────────────
    _w("AT_SportsCategoryEN", art["sports_category"])

    # ── SAP Product Division / Group / Category ──────────────────
    # _w("AT_SAPProductDivision", art["div_letter"])
    # _w("AT_SAPProductGroup",    art["prod_group"])
    # _w("AT_SAPProductCategory", art["prod_category"])

    # ── Principal Size Code (comma-sep list from sizes_raw) ──────
    # _w("AT_PrincipalSizeCode", art["sizes_raw"])

    # ── Country from input filename token (ID only) ─────────────
    country_code = _s(art.get("country_code", "")).upper()
    if country_code:
        _w("AT_Country", id_val=country_code)

    # ── Remaining placeholders (skip empty by design) ───────────
    for attr_id in NB_LICENSED_GENERIC_PLACEHOLDERS:
        if attr_id not in written:
            _val(vals_el, attr_id)
            written.add(attr_id)


def _add_variant_values(
    vals_el: ET.Element,
    art:     dict,
    size:    str,
    sap_size_code: str,
) -> None:
    """Write Variant-level <Value> elements for one size."""

    def _w(attr_id, value="", id_val=""):
        _val(vals_el, attr_id, value=value, id_val=id_val)

    # ── KEY_Variant defining attributes (required for key resolution) ──
    b_code  = art.get("brand_code", "NEW")
    b_label = LOV_BRAND.get(b_code, b_code)
    _w("AT_Brand",       id_val=b_code)
    _w("AT_PrincipalStyleCode", art["sap_style_code"])
    _w("AT_Size",               sap_size_code,         id_val=sap_size_code)

    # ── Size attributes ────────────────────────────────────────
    _w("AT_PrincipalSizeCode", size)
    _w("AT_SAPSize",           sap_size_code, id_val=sap_size_code)
    _w("AT_PrincipalSize",     size)
    _w("AT_EComSize",          f"{art['country_size']} {sap_size_code}".strip())


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
    batches_parent = f"CLH_{brand.title().replace(' ', '')}Batches"

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
    mdd=None
) -> str:
    """
    Build a <Product> XML fragment for one NB Licensed article.

    Structure:
      Product (Generic, UserTypeID = PRD_GenericArticle)
        └─ Product (Variant, UserTypeID = PRD_VariantArticle) × n_sizes

    ParentID for Generic  : PPH_{div_letter}-TempSubCat
    KeyID for Generic     : KEY_Article  → {generic_code}{color_token}
    KeyID for Variant     : KEY_Article  → {generic_code}{color_token}{sap_size_code}

    ClassificationReferences on Generic:
        CPL_Merchandiser   → CLH_{brand}Articles
        CPL_BYHierarchy    → MA_{brand_code}_{div_letter}_{prod_group}
        CPL_SAPHierarchy   → CLH_{div_letter}
        CPL_UnConfirmedForSeason → {season_id}UA
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = art.get("brand_code", brand_code) or brand_code
    div_letter = art.get("div_letter", "A")
    parent_id  = f"PPH_{div_letter}-TempSubCat"

    # Generic key = generic_code only (brand_code + sap_style_code)
    key_generic = art["generic_code"]

    # Product Group for BY Hierarchy classification reference
    prod_group = art.get("prod_group", "OT")

    # ── Generic Product element ──────────────────────────────────
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
    cr_merch.set("ClassificationID", f"CLH_{brand.title().replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{b_code}_{div_letter}_{prod_group}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{div_letter}")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

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
    """Thread worker: map + validate one NB Licensed row."""
    row, brand_code, mdd, season, country_code = row_tuple
    mapped = map_article_nb_licensed(row, brand_code=brand_code)
    mapped["season"] = season
    mapped["country_code"] = country_code
    warns  = validate(mapped, mdd)
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
        season      str   e.g. "FW26"
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    pl_files = list(PRICELIST_DIR.glob("*.xlsx"))

    # ── Mandatory file checks ────────────────────────────────────
    # CHANGE TO — raises catchable exception instead
    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Pricelist", pl_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise FileNotFoundError(f"No {label} file found in expected input directory.")

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="NEW BALANCE")

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Season ID (CLH_NEW_FW2026) ───────────────────────────────
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for pl_path in pl_files:
        log.info("─── Processing Pricelist: %s ───", pl_path.name)

        pl = NBLicensedPricelistLoader(pl_path)
        if pl.df.empty:
            log.warning("[PriceList-NBL] Empty dataframe — skipping.")
            continue

        rows = [row for _, row in pl.df.iterrows()]
        # ── TEST MODE: limit to 5 products ──────────────────────
        # rows = rows[:2]
        # ────────────────────────────────────────────────────────
        total_rows = len(rows)
        log.info("Total valid Licensed rows: %d", total_rows)

        filename_meta = _parse_metadata_from_filename(pl_path)
        out_name = f"{pl_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        season_for_rows = filename_meta["season"] or args.season
        country_for_rows = filename_meta["country"]
        task_args   = [
            (row, args.brand_code, mdd, season_for_rows, country_for_rows) for row in rows
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

        print("═══ ARTICLE SUMMARY (NEW BALANCE LICENSED) ═════════", flush=True)
        print(f"  Pricelist Licensed rows  : {total_rows}",   flush=True)
        print(f"  Generics written         : {written_count}", flush=True)
        print(f"  XML file size            : {file_kb}KB",     flush=True)
        print("════════════════════════════════════════════════════",  flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
        description="Stibo Inbound XML Generator — New Balance Licensed v1.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="FW26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
