"""
╔══════════════════════════════════════════════════════════════════╗
║     STIBO INBOUND XML GENERATOR — NEW BALANCE  v1.14            ║
║     Price List (Footwear) → Stibo STEP XML                      ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : S2 2026 Footwear Price List APAC
Primary tab  : All Sourced  (header row index 3, data starts row index 4)

Key column → AT_ attribute mappings  (Footwear-specific)
─────────────────────────────────────────────────────────────────
  item_number       → AT_PrincipalStyleCode
  Item Number          → AT_PrincipalStyleDescription
  Product Number       → AT_SAPStyleCode
  NRF Color            → AT_PrincipalColorCode / AT_Color
  Product Name         → display name / AT_PrincipalStyleDescription
  Size Profile         → AT_Gender / AT_SAPAge / AT_BYAge
  Sizes - Region       → size variants (AT_Size / AT_PrincipalSize)
  Country of Origin    → AT_CountryOrigin
  NBIL Price           → AT_FOB
  Retail Price         → AT_OriginalPrice / AT_CurrentPrice
  Line Plan Business   → AT_SAPProductGroup / AT_SportsCategoryEN
  Category             → AT_ProductHierarchyL2
  Product Line         → AT_ProductHierarchyL1
  Factory Name         → AT_VendorName
  Segment              → AT_MerchandiseCategory
─────────────────────────────────────────────────────────────────

FIX LOG:
  v1.1  - batches_parent: brand.replace(' ', '') → CLH_NewBalanceBatches
  v1.2  - _build_generic_code: brand_code prefix removed → UFFBLV4BB65
    v1.3  - AT_Generic: generic_code
    v1.4  - KEY_Article: generic_code
  v1.5  - _build_sap_style_code: color suffix removed → max 9 chars
  v1.6  - AT_BrandGroup: id_val=b_label (brand name, not code)
  v1.7  - AT_Color: NRF color name → MAA Color Code LOV mapping
  v1.8  - _build_variant_code: brand prefix removed → max 17 chars
  v1.9  - else block (no sizes): pure_generic used, ParentID removed
  v1.10 - AT_Generic in variant: brand prefix removed
  v1.11 - FIXED: broken cr_by.set syntax error removed, duplicate cr_sap
          removed, all 4 ClassificationReferences have valid Type attributes,
          
           MODE line removed for production
  v1.12 - CPL_SAPHierarchy target fixed: CLH_F → CLH_F_TempProductHierarchy
  v1.13 - Variant generation fully disabled — generic article only
          Test slice (mapped_skus[:1]) removed
    v1.14 - _build_generic_code: brand_code prefix restored → NEWUFFBLV4BB650
"""

import re
import os
import sys
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Base directory — overridden by Lambda via LAMBDA_TMP_DIR env var
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR     = BASE_DIR / "input"
PRICELIST_DIR = INPUT_DIR / "linelist_footwear"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PRICELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_footwear_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}
LOV_BY_AGE = {
    "ADULT": "Adult",     "ADULTS": "Adult",   "AD": "Adult",
    "JUNIOR": "Junior",   "YOUTH": "Youth",    "KIDS": "Kids",
    "CHILDREN": "Children", "CHILD": "Child",  "INFANT": "Infant",
    "ALL AGES": "All Ages", "ALL": "All Ages", "CH": "Children",
    "AA": "All Ages",     "UNISEX": "Adult",   "MENS": "Adult",
    "WOMENS": "Adult",    "BOYS": "Children",  "GIRLS": "Children",
    "INFANTS": "Infant",
}
LOV_BY_AGE_ID_MAP: dict[str, str] = {
    "AD": "ADULT",
    "CH": "KIDS",
    "AA": "ALL AGES",
    "IN": "INFANT",
    "JR": "KIDS",
}
LOV_SAP_ARTICLE_CATEGORY = {
    "1": "Generic", "0": "Single", "10": "Sell set (Hampers)",
}
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
    "Licensed": "License", "INLINE": "Inline", "TEAM": "Inline",
}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter", "AL": "All Season",
    "SS26": "Spring-Summer", "SS25": "Spring-Summer",
    "FW26": "Fall-Winter",  "FW25": "Fall-Winter",
}
LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
    "FW": "Footwear",
}
LOV_COO_NAME_TO_CODE: dict[str, str] = {
    "AFGHANISTAN": "AF", "ALBANIA": "AL", "ALGERIA": "DZ",
    "ANGOLA": "AO", "ARGENTINA": "AR", "ARMENIA": "AM",
    "AUSTRALIA": "AU", "AUSTRIA": "AT", "AZERBAIJAN": "AZ",
    "BAHRAIN": "BH", "BANGLADESH": "BD", "BELARUS": "BY",
    "BELGIUM": "BE", "BOLIVIA": "BO", "BRAZIL": "BR",
    "BRUNEI": "BN", "BULGARIA": "BG", "CAMBODIA": "KH",
    "CAMEROON": "CM", "CANADA": "CA", "CHILE": "CL",
    "CHINA": "CN", "COLOMBIA": "CO", "COSTA RICA": "CR",
    "CROATIA": "HR", "CUBA": "CU", "CZECH REPUBLIC": "CZ",
    "DENMARK": "DK", "ECUADOR": "EC", "EGYPT": "EG",
    "ETHIOPIA": "ET", "FINLAND": "FI", "FRANCE": "FR",
    "GERMANY": "DE", "GHANA": "GH", "GREECE": "GR",
    "HONG KONG": "HK", "HUNGARY": "HU", "ICELAND": "IS",
    "INDIA": "IN", "INDONESIA": "ID", "IRAN": "IR",
    "IRAQ": "IQ", "IRELAND": "IE", "ISRAEL": "IL",
    "ITALY": "IT", "JAMAICA": "JM", "JAPAN": "JP",
    "JORDAN": "JO", "JORDON": "JO",
    "KAZAKHSTAN": "KZ", "KENYA": "KE", "KUWAIT": "KW",
    "LAOS": "LA", "LATVIA": "LV", "LEBANON": "LB",
    "LITHUANIA": "LT", "LUXEMBOURG": "LU", "MACAU": "MO",
    "MADAGASCAR": "MG", "MALAYSIA": "MY", "MALDIVES": "MV",
    "MALTA": "MT", "MAURITIUS": "MU", "MEXICO": "MX",
    "MOLDOVA": "MD", "MONGOLIA": "MN", "MOROCCO": "MA",
    "MOZAMBIQUE": "MZ", "MYANMAR": "MM", "NAMIBIA": "NA",
    "NEPAL": "NP", "NETHERLANDS": "NL", "NEW ZEALAND": "NZ",
    "NIGERIA": "NG", "NORWAY": "NO", "OMAN": "OM",
    "PAKISTAN": "PK", "PANAMA": "PA", "PARAGUAY": "PY",
    "PERU": "PE", "PHILIPPINES": "PH", "PHILLIPINES": "PH",
    "POLAND": "PL", "PORTUGAL": "PT", "QATAR": "QA",
    "ROMANIA": "RO", "RUSSIA": "RU", "RUSSIAN FED.": "RU",
    "SAUDI ARABIA": "SA", "SENEGAL": "SN", "SINGAPORE": "SG",
    "SLOVAKIA": "SK", "SLOVENIA": "SI", "SOMALIA": "SO",
    "SOUTH AFRICA": "ZA", "SOUTH KOREA": "KR", "SPAIN": "ES",
    "SRI LANKA": "LK", "SUDAN": "SD", "SWEDEN": "SE",
    "SWITZERLAND": "CH", "SYRIA": "SY", "TAIWAN": "TW",
    "TAJIKISTAN": "TJ", "TANZANIA": "TZ", "THAILAND": "TH",
    "TUNISIA": "TN", "TURKEY": "TR", "TURKMENISTAN": "TM",
    "UGANDA": "UG", "UKRAINE": "UA", "UNITED KINGDOMS": "GB",
    "UK": "GB", "URUGUAY": "UY", "USA": "US",
    "UZBEKISTAN": "UZ", "VENEZUELA": "VE", "VIETNAM": "VN",
    "YEMEN": "YE", "ZAMBIA": "ZM", "ZIMBABWE": "ZW",
    "COMBODIA": "KH",  # typo alias
}

# ── NRF Color Name → MAA Color Code LOV ──────────────────────────────────────
NRF_COLOR_NAME_TO_MAA: dict[str, str] = {
    "PINK": "PK",        "BLACK": "005",        "WHITE": "W",
    "NAVY": "NAV",       "NAVY BLUE": "NAV",  "RED": "R",
    "BLUE": "12W",       "GREY": "GRE",       "GRAY": "GRE",
    "GREEN": "G",        "YELLOW": "Y",       "ORANGE": "O",
    "BROWN": "700",      "BEIGE": "18",       "PURPLE": "P",
    "GOLD": "KGO",       "SILVER": "SV",      "MULTI": "MI",
    "MULTICOLOR": "MI",  "CREAM": "CM",       "OLIVE": "OLI",
    "TEAL": "T",         "MAROON": "M",       "LIGHT GREY": "LGY",
    "LIGHT GRAY": "LGY", "DARK BLUE": "DBL",  "ROSE": "ROS",
    "SAND": "SAN",       "MINT": "MNT",       "LILAC": "L",
    "TURQUOISE": "TUQ",  "VIOLET": "VIO",     "MAGENTA": "MAG",
    "KHAKI": "K",        "CHARCOAL": "LPC",   "COBALT": "CB",
    "COPPER": "COP",     "DENIM": "DM",       "FUCHSIA": "FU0",
    "IVORY": "IVO",      "LAVENDER": "LAV",   "LIME": "LM",
    "MUSTARD": "MD",     "NATURAL": "NA",     "PEACH": "PC0",
    "PLUM": "PLU",       "RUST": "RUS",       "SAGE": "SAG",
    "SALMON": "SLM",     "SKY BLUE": "SB0",   "TAN": "TN5",
    "TAUPE": "TAU",      "BURGUNDY": "XM0",   "CAMEL": "CAM",
    "ANTHRACITE": "ANR", "BRONZE": "BZ",      "AQUA": "A",
    "EMERALD": "E",      "GRAPE": "GR",       "INDIGO": "IDG",
    "INK": "INK",        "LEMON": "LEM",      "MAUVE": "MAU",
    "MOCHA": "MOC",      "MUSHROOM": "MS0",   "NO COLOR": "000",
}

# ── Footwear: Category/Line Plan → Sports Category ───────────────────────────
LOV_LINE_PLAN_TO_SPORTS_CAT: dict[str, str] = {
    "RUNNING":               "Running",
    "TRAINING":              "Running",
    "OTHER FOP":             "Other",
    "LIFESTYLE":             "Lifestyle / Casual",
    "BASKETBALL":            "Basketball",
    "KIDS LIFESTYLE":        "Lifestyle / Casual",
    "TENNIS":                "Tennis / Padel",
    "KIDS PERFORMANCE":      "Running",
    "BASKETBALL AND SOFTBALL": "Other",
    "SKATE":                 "Skateboarding",
    "SOCCER":                "Soccer",
    "GOLF":                  "Golf",
    "WALKING":               "Walking",
    "BADMINTON":             "Badminton",
}

LOV_SIZE_PROFILE_TO_GENDER: dict[str, str] = {
    "WOMENS": "F", "MENS": "M", "UNISEX": "U",
    "YOUTH": "U", "KIDS": "U", "GIRLS": "F",
    "BOYS": "M", "INFANTS": "U",
}
LOV_SIZE_PROFILE_TO_AGE: dict[str, str] = {
    "WOMENS": "AD", "MENS": "AD", "UNISEX": "AD",
    "YOUTH": "CH", "KIDS": "CH", "GIRLS": "CH",
    "BOYS": "CH", "INFANTS": "CH",
}
DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E", "FOOTWEAR": "F",
    "APPAREL": "A", "EQUIPMENT": "Q", "TOYS": "T",
    "SHOES": "F",
}

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

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
                self.lovs.setdefault(str(lov_name).strip(), {})[str(val_name).strip()] = (
                    str(val_id).strip() if val_id else str(val_name).strip()
                )

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
    def __init__(self, path: Path, brand: str = "NEW BALANCE"):
        self.path  = path
        self.brand = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = (
            next((s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()), None)
            or next((s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()), wb.sheetnames[0])
        )
        log.info("[AttrList] Using sheet: '%s'", sheet_name)
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"), None
        )
        if hdr_idx is None:
            wb.close()
            return
        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}
        for row in rows[hdr_idx + 1:]:
            attr = row[col.get("Attributes", 1)] if len(row) > col.get("Attributes", 1) else None
            if not attr or str(attr).strip() in ("", "None"):
                continue
            self.attr_map.append({
                "attribute":      str(attr).strip(),
                "cluster":        self._s(row, col, "Cluster"),
                "indicator":      self._s(row, col, "Indicator"),
                "mapping_logic":  self._s(row, col, "Field Name / Mapping Logic"),
                "from_pricelist": bool(self._s(row, col, "Price List")),
            })
        wb.close()
        log.info("[AttrList] %d attributes loaded", len(self.attr_map))

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class FootwearPriceListLoader:
    """
    Loads the New Balance Footwear Price List.
    Header detection is dynamic — scans for the row containing "Item Number".
    """
    SHEET_NAMES = ["All Sourced", "Pricelist", "Price List", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[FootwearPL] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next((s for s in self.SHEET_NAMES if s in wb.sheetnames), None)
        if target is None:
            target = max(wb.sheetnames, key=lambda s: wb[s].max_row or 0)
        log.info("[FootwearPL] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(str(v).strip() == "Item Number" for v in r if v)),
            None,
        )
        if hdr_idx is None:
            hdr_idx = next(
                (i for i, r in enumerate(rows)
                 if sum(1 for v in r[:15] if isinstance(v, str) and v.strip()) >= 8),
                None,
            )
        if hdr_idx is None:
            log.error("[FootwearPL] Cannot find header row in '%s'", target)
            wb.close()
            return

        log.info("[FootwearPL] Header at row index %d (row %d)", hdr_idx, hdr_idx + 1)
        header = [
            str(h).replace("\n", " ").strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        item_col = next(
            (c for c in ["Item Number", "Item No.", "Item No"] if c in df.columns), None
        )
        if item_col:
            df = df[
                df[item_col].notna()
                & (df[item_col] != "")
                & (df[item_col] != 0)
            ]

        self.df    = df.reset_index(drop=True)
        self.sheet = target
        wb.close()
        log.info("[FootwearPL] %d SKU rows loaded from '%s'", len(self.df), self.sheet)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "N/A") else s


def _coo_to_code(coo_name: str) -> str:
    return LOV_COO_NAME_TO_CODE.get(coo_name.upper().strip(), "")


def _size_profile_to_gender(size_profile: str) -> str:
    return LOV_SIZE_PROFILE_TO_GENDER.get(size_profile.upper().strip(), "U")


def _size_profile_to_age(size_profile: str) -> str:
    return LOV_SIZE_PROFILE_TO_AGE.get(size_profile.upper().strip(), "AD")


def _line_plan_to_sports_cat(line_plan: str) -> str:
    return LOV_LINE_PLAN_TO_SPORTS_CAT.get(line_plan.upper().strip(), "Other")


def _build_sap_style_code(product_number: str, color_code: str) -> str:
    
    return re.sub(r"[^A-Z0-9]", "", product_number.upper())[:9]


def _build_generic_code(brand_code: str, product_number: str, color_code: str) -> str:
    brand     = re.sub(r"[^A-Z0-9]", "", brand_code.upper())[:3]
    base      = re.sub(r"[^A-Z0-9]", "", product_number.upper())[:9].ljust(9, "0")
    color_sfx = re.sub(r"[^A-Z0-9]", "", color_code.upper())  # Full color code, no truncation
    return f"{brand}{base}{color_sfx}"


def _build_variant_code(brand_code: str, product_number: str, color_code: str, size: str) -> str:
    # v1.8: brand prefix removed → base(9) + color_sfx(2) + sap_color(3) + size(3) = 17 chars max
    base      = re.sub(r"[^A-Z0-9]", "", product_number.upper())[:9].ljust(9, "0")
    color_sfx = re.sub(r"[^A-Z0-9]", "", color_code.upper())[:2].ljust(2, "0")
    sap_color = re.sub(r"[^A-Z0-9]", "", color_code.upper())[:3].ljust(3, "0")
    size_code = re.sub(r"[^A-Z0-9/]", "", size.upper())[:3].ljust(3, "0")
    return f"{base}{color_sfx}{sap_color}{size_code}"


def _parse_nrf_color(nrf_color: str) -> tuple[str, str]:
    """Parse NRF Color e.g. 'PINK (650)' → code='650', name='PINK'"""
    m = re.match(r"^(.*?)\s*\((\w+)\)\s*$", nrf_color.strip())
    if m:
        return m.group(2).strip(), m.group(1).strip()
    return nrf_color.strip(), nrf_color.strip()


def _nrf_name_to_maa_color(color_name: str) -> str:
    """Map NRF color name to MAA Color Code LOV value."""
    return NRF_COLOR_NAME_TO_MAA.get(color_name.upper().strip(), "")


def _parse_sizes_region(sizes_raw: str) -> list[str]:
    """
    Footwear sizes stored per-width with newlines:
        '2E:04,045,05,...\\nD:04,045,05,...'
    Extract all unique size values across all widths.
    """
    sizes: list[str] = []
    seen:  set[str]  = set()
    for segment in re.split(r"[\n;]+", sizes_raw):
        segment = re.sub(r"^[^:]+:", "", segment).strip()
        for s in segment.split(","):
            s = s.strip()
            if s and s not in seen:
                seen.add(s)
                sizes.append(s)
    return sizes


def map_sku(row: pd.Series, brand_code: str = "NEW") -> dict:
    item_number    = _s(row.get("Item Number"))
    product_number = _s(row.get("Product Number"))
    nrf_color_raw  = _s(row.get("NRF Color", ""))
    size_profile   = _s(row.get("Size Profile", ""))
    product_line   = _s(row.get("Product Line", ""))
    line_plan_biz  = _s(row.get("Line Plan Business", ""))
    category       = _s(row.get("Category", ""))
    segment        = _s(row.get("Segment", ""))
    sizes_raw      = _s(row.get("Sizes - Region", ""))
    coo_raw        = _s(row.get("Country of Origin", ""))
    nbil_price     = row.get("NBIL Price")
    retail_price   = row.get("Retail Price")
    factory_name   = _s(row.get("Factory Name", ""))
    product_name   = _s(row.get("Product Name", ""))
    intro_period   = _s(row.get("Intro Period", ""))

    color_code, color_name = _parse_nrf_color(nrf_color_raw) if nrf_color_raw else ("", "")

    gender_code  = _size_profile_to_gender(size_profile)
    age_code     = _size_profile_to_age(size_profile)
    coo_code     = _coo_to_code(coo_raw) if coo_raw else ""
    sports_cat   = _line_plan_to_sports_cat(line_plan_biz)
    sap_style    = _build_sap_style_code(product_number, color_code)
    generic_code = _build_generic_code(brand_code, product_number, color_code)

    rrp = str(int(retail_price)) if isinstance(retail_price, (int, float)) and retail_price else \
          str(retail_price).strip() if retail_price and str(retail_price).strip() not in ("", "None", "nan") else ""
    fob = str(round(float(nbil_price), 2)) if isinstance(nbil_price, (int, float)) and nbil_price else \
          str(nbil_price).strip() if nbil_price and str(nbil_price).strip() not in ("", "None", "nan") else ""

    sizes    = _parse_sizes_region(sizes_raw) if sizes_raw else []
    division = "Footwear" if product_line.upper() in ("SHOES", "FOOTWEAR") else product_line

    return {
        "item_number":    item_number,
        "product_number": product_number,
        "color_code":     color_code,
        "color_name":     color_name,
        "color_family":   color_name,
        "display_name":   product_name,
        "brand_code":     brand_code,
        "gender_code":    gender_code,
        "gender_raw":     size_profile,
        "age_code":       age_code,
        "age_raw":        size_profile,
        "division":       division,
        "merch_category": segment,
        "line_plan_biz":  line_plan_biz,
        "category":       category,
        "sports_cat":     sports_cat,
        "silhouette":     "",
        "product_name":   product_name,
        "collection":     "",
        "channel_type":   "Inline",
        "article_type":   "Inline",
        "art_category":   "1",
        "coo_raw":        coo_raw,
        "coo":            coo_code,
        "rrp":            rrp,
        "fob":            fob,
        "supplier":       factory_name,
        "technology":     "",
        "fabric":         "",
        "country_size":   "",
        "fit_version":    "",
        "sap_style_code": sap_style,
        "generic_code":   generic_code,
        "sizes":          sizes,
        "carry_over_new": "",
        "smu_type":       "",
        "gbu":            "",
        "lpa_category":   "",
        "intro_period":   intro_period,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    warns = []
    sku   = mapped["item_number"]

    field_to_at = {
        "product_number": "AT_PrincipalStyleCode",
        "gender_code": "AT_Gender",
        "age_code":    "AT_SAPAge",
        "brand_code":  "AT_Brand",
        "coo":         "AT_CountryOrigin",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {})
        if "mandatory" in (meta.get("cardinality") or "").lower():
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{sku}] MISSING mandatory: {at_id}")

    if not mapped.get("rrp"):
        warns.append(f"[{sku}] MISSING price: Retail Price (AT_OriginalPrice)")
    if not mapped.get("fob"):
        warns.append(f"[{sku}] MISSING cost: NBIL Price (AT_FOB)")

    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER HELPERS
# ══════════════════════════════════════════════════════════════════

def _clean_xml_value(raw) -> str:
    text = "" if raw is None else str(raw).strip()
    return "" if text in ("", "None", "nan", "NaT") else text


def _val(parent, attr_id, value="", id_val=""):
    normalized_id = _clean_xml_value(id_val)
    normalized_value = _clean_xml_value(value)
    if not normalized_id and not normalized_value:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if normalized_id:
        el.set("ID", normalized_id)
        return el
    if normalized_value:
        el.text = normalized_value
    return el


def _multival(parent, attr_id, id_val, label=""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)
    return mv


def _lov(code, lookup, default=""):
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def _get_division_code(product_line: str) -> str:
    pl_upper = (product_line or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in pl_upper:
            return code
    return "F"


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC VALUE WRITER
# ══════════════════════════════════════════════════════════════════

GENERIC_ATTR_PLACEHOLDERS = [
    "AT_SPUGrouping", "AT_SPUGroupingName",
    "AT_EComProductNameEN", "AT_EComProductNameID", "AT_EComProductNamePH",
    "AT_EComProductNameMY", "AT_EComProductNameVN", "AT_EComProductNameTH",
    "AT_EComProductNameKH",
    "AT_ShortDescriptionEN", "AT_ShortDescriptionID", "AT_ShortDescriptionPH",
    "AT_ShortDescriptionMY", "AT_ShortDescriptionVN", "AT_ShortDescriptionTH",
    "AT_ShortDescriptionKH",
    "AT_LongDescriptionEN", "AT_LongDescriptionID", "AT_LongDescriptionPH",
    "AT_LongDescriptionMY", "AT_LongDescriptionVN", "AT_LongDescriptionTH",
    "AT_LongDescriptionKH",
    "AT_CareInstructionEN", "AT_CareInstructionID", "AT_CareInstructionPH",
    "AT_CareInstructionMY", "AT_CareInstructionVN", "AT_CareInstructionTH",
    "AT_CareInstructionKH",
    "AT_ColorDescriptionID", "AT_ColorDescriptionPH", "AT_ColorDescriptionMY",
    "AT_ColorDescriptionVN", "AT_ColorDescriptionTH", "AT_ColorDescriptionKH",
    "AT_Interest", "AT_Occasion", "AT_HeelHeight", "AT_HeelType",
    "AT_GolfClubLength", "AT_GolfClubFlex", "AT_GolfClubLoft",
    "AT_Collection2", "AT_CertificateNumber",
    "AT_EstimatedLandedCost", "AT_ExchangeRate",
    "AT_FreightCost", "AT_Royalty", "AT_Commission",
    "AT_SGS", "AT_MarketingFee", "AT_HaddadOfficeCharge",
    "AT_HaddadOfficeChargeDeduction", "AT_Others",
    "AT_PackagingLength", "AT_PackagingWidth", "AT_PackagingHeight", "AT_PackagingWeight",
    "AT_ProductWeight", "AT_ProductLengthWidthHeight",
    "AT_ImagesSource", "AT_LaunchingDate",
    "AT_MainEANIndicator", "AT_EcomGenderDescriptionEN",
    "AT_EComAgesCategory", "AT_EcommConcept", "AT_EcommSalesChannel",
    "AT_Createdon", "AT_Updatedby", "AT_Updatedon",
    "AT_FOBCurrency",
]


def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None):
    written: set[str] = set()

    # Empty placeholder attributes disabled for clean XML output.
    # for a in GENERIC_ATTR_PLACEHOLDERS:
    #     if a not in written:
    #         _val(vals_el, a)
    #         written.add(a)

    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand",      b_label, id_val=b_code)
    # v1.6: AT_BrandGroup uses brand name as ID
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)

    _val(vals_el, "AT_PrincipalStyleCode",        art["item_number"])
    _val(vals_el, "AT_PrincipalStyleDescription", art["display_name"] or art["item_number"])
    _val(vals_el, "AT_SAPStyleCode",              art["sap_style_code"])

    # Color
    _val(vals_el, "AT_PrincipalColorCode", art["color_code"])
    _val(vals_el, "AT_PrincipalColorName", art["color_name"])

    # v1.7: AT_Color — NRF color name → MAA Color Code LOV
    maa_color_code = _nrf_name_to_maa_color(art["color_name"])
    if maa_color_code:
        _val(vals_el, "AT_Color", id_val=maa_color_code)
    else:
        _val(vals_el, "AT_Color")

    # Gender
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _val(vals_el, "AT_Gender",                    g_label, id_val=g_code)
    _val(vals_el, "AT_BYGender",                  g_label, id_val=g_code)
    _val(vals_el, "AT_PrincipalGenderCode",        art["gender_code"])
    # _val(vals_el, "AT_PrincipalGenderDescription", art["gender_raw"])
    _gender_max = 9
    if mdd is not None:
        _meta = mdd.attributes.get("AT_PrincipalGenderDescription", {})
        _gender_max = int(_meta.get("max_chars") or 9)
    _val(vals_el, "AT_PrincipalGenderDescription", art["gender_raw"][:_gender_max])

    # Age
    sap_age_code  = art["age_code"]
    sap_age_label = LOV_AGE.get(sap_age_code, sap_age_code)
    _val(vals_el, "AT_SAPAge", sap_age_label, id_val=sap_age_code)
    by_age_id = LOV_BY_AGE_ID_MAP.get(sap_age_code, "ADULT")
    _val(vals_el, "AT_BYAge", id_val=by_age_id)
    _val(vals_el, "AT_PrincipalAgeCode",        art["age_code"])
    _val(vals_el, "AT_PrincipalAgeDescription", art["age_raw"])

    # COO
    _val(vals_el, "AT_CountryOrigin", art["coo_raw"], id_val=art["coo"])

    # Season
    # Season  ── FIX: handles both SS26 and SS2026 formats; always emits 4-digit year
    sea_raw = art.get("season", "")
    if not sea_raw:
        sea_raw = ""
    # Extract prefix (SS/FW) and raw year digits
    sea_m = re.match(r'^([A-Z]{2})(\d{2,4})$', sea_raw.strip().upper())
    if sea_m:
        sea_prefix = sea_m.group(1)                                           # "SS"
        sea_digits = sea_m.group(2)                                           # "26" or "2026"
        sea_year_4 = sea_digits if len(sea_digits) == 4 else f"20{sea_digits}"  # "2026"
    else:
        sea_prefix = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
        sea_year_4 = ""

    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_prefix, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_prefix)
    _val(vals_el, "AT_SeasonYear", sea_year_4)

    # Article category & type
    cat_code  = art["art_category"]
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, cat_code)
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)
    at_norm  = art["article_type"].capitalize()
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)

    # Indicators
    _val(vals_el, "AT_BYIndicator",   "Yes", id_val="Y")
    _val(vals_el, "AT_SAPIndicator",  "No",  id_val="N")
    _val(vals_el, "AT_EcomIndicator", "No",  id_val="N")
    _val(vals_el, "AT_UOM", "Each", id_val="EA")

    # ── Material Type — ZINA ─────────────────────────────────────
    _val(vals_el, "AT_MaterialType", id_val="ZINA")

    # Pricing
    _val(vals_el, "AT_OriginalPrice",       art["rrp"])
    _val(vals_el, "AT_CurrentPrice",        art["rrp"])
    _val(vals_el, "AT_FOB",                 art["fob"])
    # _val(vals_el, "AT_RetailPriceCurrency")

    # Vendor
    _val(vals_el, "AT_MainVendorIdentification", "1")

    # Hierarchy
    _val(vals_el, "AT_MerchandiseCategory", (art.get("merch_category") or "")[:9])
    # _val(vals_el, "AT_Silhouette",          art["silhouette"])
    _val(vals_el, "AT_SportsCategoryEN",    art["sports_cat"])
    # Principal Merchandise Hierarchy
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art.get("division", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art.get("product_name", ""))

    # Other
    # _val(vals_el, "AT_TechnologyUsed",  art["technology"])
    # _val(vals_el, "AT_Content",         art["fabric"])
    # _val(vals_el, "AT_Material",        art["fabric"])
    # _val(vals_el, "AT_Collection1",     art["collection"])
    # _val(vals_el, "AT_CountrySize",     art["country_size"])
    # v1.3: AT_Generic without brand prefix
    _val(vals_el, "AT_InboundGenericCode",         art["generic_code"])
    _val(vals_el, "AT_NatureOfArticle", "Retail", id_val="RT1")

    country_val = art.get("country_code", "")
    if country_val:
        _val(vals_el, "AT_Country", country_val, id_val=country_val)


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS
# ══════════════════════════════════════════════════════════════════

def build_classifications(brand, brand_code, season_code):
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    # v1.1: spaces removed from brand name
    batches_parent   = f"CLH_{brand.replace(' ', '')}Batches"
    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", batches_parent)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — PRODUCT XML BUILDER
# ══════════════════════════════════════════════════════════════════

def build_product_xml(art, brand, brand_code, comp_code, sbu, season_id):
    item_number = art["item_number"]
    if not item_number:
        return ""

    div_letter   = _get_division_code(art.get("division", "Footwear"))
    parent_id    = f"PPH_{div_letter}-TempSubCat"
    key_generic = art["generic_code"]

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["display_name"] or item_number

    # ── ClassificationReferences ──────────────────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    merch_id = re.sub(r"[^A-Z0-9]", "", (art["merch_category"] or "").upper())[:4]
    # cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{brand_code}_{div_letter}_{merch_id}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # v1.12: CLH_F_TempProductHierarchy (not CLH_F — Division node invalid for CPL_SAPHierarchy)
    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{div_letter}_TempProductHierarchy")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")
    # ─────────────────────────────────────────────────────────────────────────

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu)

    # v1.13: Variant generation disabled — generic article only
    log.debug("[SKU %s] Generic only — variants disabled.", item_number)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_sku(task):
    row, brand_code, mdd = task
    mapped = map_sku(row, brand_code=brand_code)
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Entry point called by first_product_processor.py (or CLI).
    args must have: brand, brand_code, comp_code, sbu, season, seq
    """
    def first(d: Path, ext="*.xlsx"):
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    pl_files = list(PRICELIST_DIR.glob("*.xlsx"))

    for label, val in [
        ("MDD",        mdd_f),
        ("Attributes", attr_f),
        ("Price List", pl_files),
    ]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    mdd    = MDDLoader(mdd_f)
    _      = AttributesListLoader(attr_f, brand=args.brand)

    all_warnings: list[str] = []

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    for pl_path in pl_files:
        log.info("─── Processing Footwear Price List: %s ───", pl_path.name)

        pl = FootwearPriceListLoader(pl_path)
        if pl.df.empty:
            log.warning("[FootwearPL] Empty dataframe — skipping.")
            continue

        item_col = next(
            (c for c in ["Item Number", "Item No.", "Item No"] if c in pl.df.columns), None
        )
        rows = []
        for _, row in pl.df.iterrows():
            item = str(row.get(item_col or "Item Number", "")).strip()
            if item and item not in ("None", "nan", ""):
                rows.append(row)

        # rows = rows[:10] 
        total_rows = len(rows)
        log.info("[FootwearPL] %d valid SKU rows to process", total_rows)

        out_name = f"{pl_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        log.info("Pass 1/2 — parallel map+validate (%d rows) …", total_rows)
        num_workers = min(8, max(1, total_rows))
        task_args   = [(row, args.brand_code, mdd) for row in rows]
        ordered: list[tuple[int, dict, list]] = []

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_sku, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx, (mapped, warns) = futures[fut], fut.result()
                mapped["season"]       = args.season
                mapped["country_code"] = getattr(args, "country_code", "")   # ← ADD THIS
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_skus = [m for _, m, _ in ordered]
        # mapped_skus = mapped_skus[:1]
        del ordered

        total_variants = sum(len(a.get("sizes", [])) for a in mapped_skus)
        log.info("Pass 1 done — mapped=%d  size_rows=%d", len(mapped_skus), total_variants)

        log.info("Pass 2/2 — streaming XML to %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cls_el  = build_classifications(args.brand, args.brand_code, args.season)
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
            for art in mapped_skus:
                if not art.get("item_number"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, args.brand_code,
                    args.comp_code, args.sbu, season_id,
                )
                product_xml = _XMLNS_RE.sub("", product_xml)
                f.write(f"    {product_xml}\n")
                del product_xml
                written_count += 1
            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        del mapped_skus

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%dKB)", out_path, file_kb)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

        print("═══ SKU SUMMARY ════════════════════════════════════", flush=True)
        print(f"  Price List rows     : {total_rows}",     flush=True)
        print(f"  Mapped SKUs         : {written_count}",  flush=True)
        print(f"  Total size variants : {total_variants}", flush=True)
        print(f"  XML file size       : {file_kb}KB",      flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    rpt_path = LOG_DIR / f"validation_nb_footwear_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")

    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All SKUs passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 10 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — New Balance Footwear v1.13")
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0000")
    p.add_argument("--sbu",        default="FW")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
