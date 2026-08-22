"""
╔══════════════════════════════════════════════════════════════════╗
║     STIBO INBOUND XML GENERATOR — NEW BALANCE  v1.0             ║
║     Price List (Apparel & Accessories) → Stibo STEP XML         ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : Fall26 Inline Apparel & Licensed Accessories Price List - APAC
Primary tab  : Pricelist  (header row index 3, data starts row index 4)

Key column → AT_ attribute mappings
─────────────────────────────────────────────────────────────────
  Product Number       → AT_PrincipalStyleCode  (base style: WT41253)
  Product Number       → AT_PrincipalStyleDescription
  Product Number       → AT_SAPStyleCode (9-digit base + 2-digit color suffix)
  Color Code           → AT_PrincipalColorCode  (e.g. BM4)
  Color Family         → AT_PrincipalColorDescription + localized variants
  Color Name           → AT_PrincipalColorName / AT_PrincipalColorNameEN
  Color Code           → AT_Color / AT_ColorCode (SAP color — via LOV)
  Size Profile         → AT_Gender (Womens→F, Mens→M, Unisex→U, Youth→U)
  Size Profile         → AT_SAPAge / AT_BYAge  (Womens/Mens/Unisex→AD, Youth→CH)
  Sizes                → AT_PrincipalSize / AT_Size (comma-split → per variant)
  COO                  → AT_CountryOrigin  (full name → LOV code)
  NBIL Price           → AT_FOB
  Global Retail Price  → AT_OriginalPrice / AT_CurrentPrice
  Product Line         → AT_SAPProductDivision  (Apparel / Accessories)
  Line Plan Business   → AT_SAPProductGroup + AT_SportsCategoryEN (mapped)
  Silhouette           → AT_SAPProductCategory + AT_Silhouette
  Fit Version          → AT_CountrySize  (Western Fit→US, Asia Fit→ASIA)
  Merchandise Category → AT_MerchandiseCategory
  Collection           → AT_Collection1
  Supplier             → AT_VendorName
  Technologies         → AT_TechnologyUsed
  Primary Fabric Content → AT_Content / AT_Material
  Channel Type         → AT_BYArticleType (Inline/Team→Inline, Licensed→License)
─────────────────────────────────────────────────────────────────
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

INPUT_DIR    = BASE_DIR / "input"
PRICELIST_DIR = INPUT_DIR / "linelist_apparel"   # price list lands in "linelist" folder
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PRICELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
    "ADULT": "Adult", "ADULTS": "Adult", "AD": "Adult",
    "JUNIOR": "Junior", "YOUTH": "Youth", "KIDS": "Kids",
    "CHILDREN": "Children", "CHILD": "Child", "INFANT": "Infant",
    "ALL AGES": "All Ages", "ALL": "All Ages", "CH": "Children",
    "AA": "All Ages",
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
}

# Full country name (as in Price List COO column) → ISO-2 code
# Covers all values seen in the data including misspellings
LOV_COO_NAME_TO_CODE: dict[str, str] = {
    "THAILAND":     "TH",
    "VIETNAM":      "VN",
    "INDONESIA":    "ID",
    "CHINA":        "CN",
    "JORDON":       "JO",   # misspelling of Jordan in source data
    "JORDAN":       "JO",
    "SRI LANKA":    "LK",
    "PAKISTAN":     "PK",
    "USA":          "US",
    "COMBODIA":     "KH",   # misspelling of Cambodia in source data
    "CAMBODIA":     "KH",
    "JAPAN":        "JP",
    "COSTA RICA":   "CR",
    "INDIA":        "IN",
    "BANGLADESH":   "BD",
    "MALAYSIA":     "MY",
    "PHILIPPINES":  "PH",
    "TURKEY":       "TR",
    "GERMANY":      "DE",
    "ITALY":        "IT",
    "FRANCE":       "FR",
    "TAIWAN":       "TW",
    "HONG KONG":    "HK",
    "SINGAPORE":    "SG",
    "AUSTRALIA":    "AU",
}

# Line Plan Business → Sports Category EN (AT_SportsCategoryEN)
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

# Size Profile → SAP Gender code
LOV_SIZE_PROFILE_TO_GENDER: dict[str, str] = {
    "WOMENS":  "F",
    "MENS":    "M",
    "UNISEX":  "U",
    "YOUTH":   "U",
    "KIDS":    "U",
    "GIRLS":   "F",
    "BOYS":    "M",
    "INFANTS": "U",
}

# Size Profile → SAP Age code
LOV_SIZE_PROFILE_TO_AGE: dict[str, str] = {
    "WOMENS":  "AD",
    "MENS":    "AD",
    "UNISEX":  "AD",
    "YOUTH":   "CH",
    "KIDS":    "CH",
    "GIRLS":   "CH",
    "BOYS":    "CH",
    "INFANTS": "CH",
}


# Age code → LOV_BYAge ID (as registered in STIBO)
LOV_AGE_CODE_TO_BY_LOV_ID: dict[str, str] = {
    "AD": "ADULT",
    "CH": "KIDS",
    "AA": "ALL AGES",
    "IN": "INFANT",
    "JR": "KIDS",
}

# Product Line → Division letter (for PPH parent and division code)
DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E",
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "EQUIPMENT":   "Q",
    "TOYS":        "T",
}

# Fit Version → AT_CountrySize
LOV_FIT_VERSION_TO_COUNTRY_SIZE: dict[str, str] = {
    "WESTERN FIT": "US",
    "ASIA FIT":    "ASIA",
    "N/A":         "",
}


# Silhouette raw value → STIBO LOV_Silhouette ID
LOV_SILHOUETTE_MAP: dict[str, str] = {
    "S/S TOP":                  "ST",
    "L/S TOP":                  "LT",
    "PANT":                     "PT",
    "SHORT":                    "SH",
    "TIGHT":                    "TG",
    "DRESS":                    "DR",
    "HOODIES & SWEATSHIRTS":    "HS",
    "JACKET":                   "JK",
    "TANKS/SLEEVELESS/SINGLET": "TK",
    "SOCKS":                    "SK",
    "HEADWEAR":                 "HW",
    "BAG":                      "BG",
    "VEST":                     "VT",
    "GLOVE":                    "GL",
}



# Stibo XML namespace constants
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """
    Loads AT_ attribute metadata and LOVs from the Master Data Dictionary.
    Identical to the Adidas version — MDD is brand-agnostic.
    """
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
    """
    Reads the NEW BALANCE tab from the Attributes List Excel file.
    Captures attribute name, cluster, indicator, and mapping logic.
    """
    def __init__(self, path: Path, brand: str = "NEW BALANCE"):
        self.path  = path
        self.brand = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Find the right sheet — prefer exact match, then partial, skip *v4* sheets
        sheet_name = (
            next((s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()), None)
            or next((s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()), wb.sheetnames[0])
        )
        log.info("[AttrList] Using sheet: '%s'", sheet_name)

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Header row has "Attributes" in col B (index 1)
        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"), None
        )
        if hdr_idx is None:
            log.error("[AttrList] Cannot find header row in '%s'", sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}

        for row in rows[hdr_idx + 1:]:
            attr = row[col.get("Attributes", 1)] if len(row) > col.get("Attributes", 1) else None
            if not attr or str(attr).strip() in ("", "None"):
                continue
            self.attr_map.append({
                "attribute":     str(attr).strip(),
                "cluster":       self._s(row, col, "Cluster"),
                "indicator":     self._s(row, col, "Indicator"),
                "mapping_logic": self._s(row, col, "Field Name / Mapping Logic"),
                "from_pricelist": bool(self._s(row, col, "Price List")),
            })

        wb.close()
        log.info("[AttrList] %d attributes loaded from '%s'", len(self.attr_map), sheet_name)

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class PriceListLoader:
    """
    Loads the New Balance Apparel & Accessories Price List (Pricelist tab).

    Header is at row index 3 (0-based). Each row is one SKU (Item Number).
    The key identifier is Item Number = Product Number + '_' + Color Code.

    Returns:
        self.df    — pandas DataFrame with all rows, header-cleaned
        self.sheet — sheet name used
    """
    SHEET_NAMES = ["Pricelist", "Price List", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[PriceList] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Find the right sheet
        target = next((s for s in self.SHEET_NAMES if s in wb.sheetnames), None)
        if target is None:
            # Fall back to the largest sheet
            target = max(wb.sheetnames, key=lambda s: wb[s].max_row or 0)
        log.info("[PriceList] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Find header — look for "Item Number" in row
        hdr_idx = next(
            (i for i, r in enumerate(rows) if any(str(v).strip() == "Item Number" for v in r if v)),
            None,
        )
        if hdr_idx is None:
            # Fallback: first row with enough non-empty string cells
            hdr_idx = next(
                (i for i, r in enumerate(rows) if sum(1 for v in r[:15] if isinstance(v, str) and v.strip()) >= 8),
                None,
            )
        if hdr_idx is None:
            log.error("[PriceList] Cannot find header row in '%s'", target)
            wb.close()
            return

        log.info("[PriceList] Header at row %d (0-indexed)", hdr_idx)
        header = [
            str(h).replace("\n", " ").strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Filter to rows that have a valid Item Number
        item_col = next((c for c in ["Item Number", "Item No.", "Item No"] if c in df.columns), None)
        if item_col:
            df = df[df[item_col].notna() & (df[item_col] != "") & (df[item_col] != 0)]

        self.df    = df.reset_index(drop=True)
        self.sheet = target
        wb.close()
        log.info("[PriceList] %d SKU rows loaded from '%s'", len(self.df), self.sheet)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safe string — returns '' for None / 'None' / 'nan' / 0."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "N/A") else s


def _coo_to_code(coo_name: str) -> str:
    """
    Convert full country name (as in Price List COO column) to ISO-2 code.
    Falls back to first 2 chars uppercased if not found.
    """
    return LOV_COO_NAME_TO_CODE.get(coo_name.upper().strip(), coo_name[:2].upper())


def _size_profile_to_gender(size_profile: str) -> str:
    """Derive SAP gender code from Size Profile field."""
    return LOV_SIZE_PROFILE_TO_GENDER.get(size_profile.upper().strip(), "U")


def _size_profile_to_age(size_profile: str) -> str:
    """Derive SAP age code from Size Profile field."""
    return LOV_SIZE_PROFILE_TO_AGE.get(size_profile.upper().strip(), "AD")


def _channel_to_article_type(channel_type: str) -> str:
    """
    Map Channel Type → BY Article Type.
    Inline / Team → 'Inline'
    Licensed      → 'License'
    """
    ct = channel_type.upper().strip()
    if "LICENSE" in ct:
        return "License"
    return "Inline"


def _line_plan_to_sports_cat(line_plan: str) -> str:
    """Map Line Plan Business value to Sports Category EN LOV value."""
    return LOV_LINE_PLAN_TO_SPORTS_CAT.get(line_plan.upper().strip(), "Other")


def _fit_version_to_country_size(fit_version: str) -> str:
    """Map Fit Version to Country Size value."""
    return LOV_FIT_VERSION_TO_COUNTRY_SIZE.get(fit_version.upper().strip(), "")


def _build_sap_style_code(product_number: str, color_code: str) -> str:
    """
    SAP Style Code for Apparel/Accessories:
      9-digit Product Number + 2-digit MD-mapped color suffix
    We use the first 2 chars of color code as the suffix (MD would normally provide mapping).
    """
    base  = re.sub(r"[^A-Z0-9]", "", product_number.upper())[:9].ljust(9, "0")
    color_sfx = re.sub(r"[^A-Z0-9]", "", color_code.upper())[:2].ljust(2, "0")
    return f"{base}{color_sfx}"


# AFTER
def _build_generic_code(brand_code: str, product_number: str, color_code: str) -> str:
    """
    AT_Generic for Apparel/Accessories:
      3-digit brand code + Product Number + Color Code
    """
    base = re.sub(r"[^A-Z0-9]", "", product_number.upper())
    color = re.sub(r"[^A-Z0-9]", "", color_code.upper())
    return f"{brand_code}{base}{color}"



def map_sku(row: pd.Series, brand_code: str = "NEW") -> dict:
    """
    Map one Price List row to a normalized article dict.

    Each row in the Price List is one SKU (style + color).
    Sizes are a comma-separated string that will be split into
    individual variant rows by the XML builder.
    """
    # ── Raw field extraction ──────────────────────────────────────────────
    item_number      = _s(row.get("Item Number"))          # WT41253_BM4
    product_number   = _s(row.get("Product Number"))       # WT41253
    color_code       = _s(row.get("Color Code"))           # BM4
    gbu              = _s(row.get("GBU"))
    line_plan_biz    = _s(row.get("Line Plan Business"))   # Running / Training / ...
    category         = _s(row.get("Category"))
    channel_type     = _s(row.get("Channel Type"))         # Inline / Licensed / Team
    lpa_category     = _s(row.get("LPA Category"))
    merch_category   = _s(row.get("Merchandise Category")) # Running / Accessories
    smu_type         = _s(row.get("SMU Type"))
    size_profile     = _s(row.get("Size Profile"))         # Womens / Mens / Unisex / Youth
    global_intro_dt  = row.get("Global Intro Date")
    region_intro_dt  = row.get("Region Intro Date")
    phraseout_dt     = row.get("Phraseout Date")
    display_name     = _s(row.get("Product Display Name")) # Athletics T-Shirt
    carry_over_new   = _s(row.get("Item - CarryOver/New"))
    color_name       = _s(row.get("Color Name"))           # GREEN (400)
    color_family     = _s(row.get("Color Family"))         # GREEN
    product_line     = _s(row.get("Product Line"))         # Apparel / Accessories
    collection       = _s(row.get("Collection"))           # Sport Essentials
    silhouette       = _s(row.get("Silhouette"))           # S/S Top / Pant / Short
    sizes_raw        = _s(row.get("Sizes"))                # XS,S,M,L,XL,2XL
    fit_version      = _s(row.get("Fit Version"))          # Western Fit / Asia Fit
    technologies     = _s(row.get("Technologies"))         # NB IceX
    primary_fabric   = _s(row.get("Primary Fabric Content"))
    global_retail    = row.get("Global Retail Price")
    nbil_price       = row.get("NBIL Price")
    supplier         = _s(row.get("Supplier"))
    coo_raw          = _s(row.get("COO"))                  # Thailand / Vietnam / ...

    # ── Derived fields ────────────────────────────────────────────────────
    gender_code   = _size_profile_to_gender(size_profile)
    age_code      = _size_profile_to_age(size_profile)
    coo_code      = _coo_to_code(coo_raw) if coo_raw else ""
    article_type  = _channel_to_article_type(channel_type)
    sports_cat    = _line_plan_to_sports_cat(line_plan_biz)
    country_size  = _fit_version_to_country_size(fit_version)
    sap_style     = _build_sap_style_code(product_number, color_code)
    generic_code  = _build_generic_code(brand_code, product_number, color_code)

    # Price: cast to string for XML emission
    rrp = str(int(global_retail)) if isinstance(global_retail, (int, float)) and global_retail else ""
    fob = str(round(float(nbil_price), 2)) if isinstance(nbil_price, (int, float)) and nbil_price else ""

    # Sizes: split comma-separated string → list of stripped size strings
    sizes = [s.strip() for s in sizes_raw.split(",") if s.strip()] if sizes_raw else []

    # Article category: Generic (1) unless single-item accessory — default Generic
    art_category = "1"

    return {
        "item_number":    item_number,      # full SKU key  e.g. WT41253_BM4
        "product_number": product_number,   # base style    e.g. WT41253
        "color_code":     color_code,       # principal color code
        "color_name":     color_name,       # e.g. GREEN (400)
        "color_family":   color_family,     # e.g. GREEN
        "display_name":   display_name,     # product display name
        "brand_code":     brand_code,
        "gender_code":    gender_code,      # M / F / U
        "gender_raw":     size_profile,     # raw Size Profile value
        "age_code":       age_code,         # AD / CH
        "age_raw":        size_profile,
        "division":       product_line,     # Apparel / Accessories
        "merch_category": merch_category,
        "line_plan_biz":  line_plan_biz,
        "sports_cat":     sports_cat,
        "silhouette":     silhouette,
        "collection":     collection,
        "channel_type":   channel_type,
        "article_type":   article_type,     # Inline / License
        "art_category":   art_category,
        "coo_raw":        coo_raw,
        "coo":            coo_code,
        "rrp":            rrp,              # Global Retail Price as string
        "fob":            fob,              # NBIL Price as string
        "supplier":       supplier,
        "technology":     technologies,
        "fabric":         primary_fabric,
        "country_size":   country_size,     # US / ASIA / ""
        "fit_version":    fit_version,
        "sap_style_code": sap_style,
        "generic_code":   generic_code,
        "sizes":          sizes,            # list of size strings
        "carry_over_new": carry_over_new,
        "smu_type":       smu_type,
        "gbu":            gbu,
        "lpa_category":   lpa_category,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """
    Check mandatory AT_ attributes are populated.
    Returns a list of warning strings (empty = all good).
    """
    warns = []
    sku   = mapped["item_number"]

    field_to_at = {
        "product_number": "AT_PrincipalStyleCode",
        "gender_code":  "AT_Gender",
        "age_code":     "AT_SAPAge",
        "brand_code":   "AT_Brand",
        "coo":          "AT_CountryOrigin",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {})
        if "mandatory" in (meta.get("cardinality") or "").lower():
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{sku}] MISSING mandatory: {at_id}")

    if not mapped.get("rrp"):
        warns.append(f"[{sku}] MISSING price: Global Retail Price (AT_OriginalPrice)")
    if not mapped.get("fob"):
        warns.append(f"[{sku}] MISSING cost: NBIL Price (AT_FOB)")

    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element:
    """Emit a <Value AttributeID="..."> element. If id_val set, emits ID= attribute."""
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id = str(id_val).strip() if id_val else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
        return el
    clean_val = str(value).strip() if value else ""
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = "") -> ET.Element:
    """Emit a <MultiValue><Value ID="..."> block for LOV multi-value attributes."""
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)
    return mv


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    """Lookup code in a dict → (code, label). Falls back to code as label."""
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def _get_division_code(product_line: str) -> str:
    """Map Product Line to single-letter division code for PPH parent ID."""
    pl_upper = (product_line or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in pl_upper:
            return code
    return "X"


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC & VARIANT VALUE WRITERS
# ══════════════════════════════════════════════════════════════════

# Placeholder AT_ IDs emitted as empty values on the Generic product.
# These are all attributes defined in the MDD that are not sourced from
# the price list and will be populated downstream (e.g. by MD, AI, API).
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



def _add_generic_values(
    vals_el: ET.Element,
    art: dict,
    brand_name: str,
    comp_code: str,
    sbu: str,
    mdd=None
) -> None:
    """
    Write all <Value> elements into the Generic product's <Values> block.
    Order: placeholders first (empty), then data-sourced attributes
    (so downstream tools can detect which fields are populated).
    """
    written: set[str] = set()

    # ── 1. Emit placeholders (empty) first ───────────────────────────────
    # ── 1. Placeholders removed — only send attributes with values ────────
    written.update(GENERIC_ATTR_PLACEHOLDERS)  # mark as written so they are skipped

    # ── 2. Multi-value LOV attributes ────────────────────────────────────
    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    # ── 3. Brand ──────────────────────────────────────────────────────────
    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand", b_label, id_val=b_code)
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)
    written |= {"AT_Brand", "AT_BrandGroup"}

    # ── 4. Style / Item identification ───────────────────────────────────
    # AFTER — send product_number only "WT41253" (≤9 chars, matches linelist_main.py)
    _val(vals_el, "AT_PrincipalStyleCode", art["product_number"])
    # Per mapping logic: description = item number only (not display name)
    _val(vals_el, "AT_PrincipalStyleDescription", art["product_number"])
    _val(vals_el, "AT_SAPStyleCode", art["product_number"])
    written |= {"AT_PrincipalStyleCode", "AT_PrincipalStyleDescription", "AT_SAPStyleCode"}

    # ── 5. Color ──────────────────────────────────────────────────────────
    _val(vals_el, "AT_PrincipalColorCode", art["color_code"])
    # _val(vals_el, "AT_PrincipalColorDescription", art["color_family"])
    _val(vals_el, "AT_PrincipalColorName", art["color_name"])
    # _val(vals_el, "AT_PrincipalColorNameEN", art["color_name"])

    # SAP Color: use Color Code as proxy (MD provides final mapping)
   
    # AFTER:
    color_family_upper = art["color_family"].strip().upper()
    color_lov = (mdd.lovs.get("Color Code") or {}) if mdd else {}

    # 1. Exact match on description (e.g. "GREEN" → "EE0")
    colour_id = color_lov.get(color_family_upper, "")

    # 2. Fuzzy match — strip trailing punctuation/spaces from LOV descriptions
    #    handles cases like "RED ." → code "R"
    if not colour_id:
        for desc, code in color_lov.items():
            if desc.rstrip(". ").upper() == color_family_upper:
                colour_id = code
                break

    # 3. Reverse lookup — in case color_family is already a LOV code (e.g. "EE0")
    if not colour_id:
        rev = {v.upper(): v for v in color_lov.values()}
        colour_id = rev.get(color_family_upper, "")

    # 4. If still not found — send empty, do NOT fall back to raw string
    if not colour_id:
        log.warning(
            "[Color LOV] color_family='%s' NOT found in MDD Color Code LOV — sending empty ID",
            color_family_upper
        )

    log.info("[Color LOV] color_family='%s' → LOV ID='%s'", color_family_upper, colour_id or "(empty)")
    _val(vals_el, "AT_Color", art["color_code"].upper(), id_val=colour_id)
    written |= {
        "AT_PrincipalColorCode", "AT_PrincipalColorDescription",
        "AT_PrincipalColorName", "AT_PrincipalColorNameEN", "AT_Color",
    }

    # ── 6. Gender ─────────────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _val(vals_el, "AT_Gender", g_label, id_val=g_code)
    _val(vals_el, "AT_BYGender", g_label, id_val=g_code)
    # _val(vals_el, "AT_PrincipalGender", art["gender_raw"])
    _val(vals_el, "AT_PrincipalGenderCode", art["gender_code"])
    _val(vals_el, "AT_PrincipalGenderDescription", art["gender_raw"])
    written |= {"AT_Gender", "AT_BYGender", "AT_PrincipalGender",
                "AT_PrincipalGenderCode", "AT_PrincipalGenderDescription"}

    # ── 7. Age ────────────────────────────────────────────────────────────
    sap_age_code  = art["age_code"]
    sap_age_label = LOV_AGE.get(sap_age_code, sap_age_code)
    _val(vals_el, "AT_SAPAge", sap_age_label, id_val=sap_age_code)

    by_age_lov_id = LOV_AGE_CODE_TO_BY_LOV_ID.get(sap_age_code, "ADULT")
    by_age_label  = LOV_BY_AGE.get(sap_age_code, sap_age_code)
    _val(vals_el, "AT_BYAge", by_age_label, id_val=by_age_lov_id)  # sends ID="ADULT" ← CORRECT

    # _val(vals_el, "AT_PrincipalAge", art["age_raw"])
    _val(vals_el, "AT_PrincipalAgeCode", art["age_code"])
    _val(vals_el, "AT_PrincipalAgeDescription", art["age_raw"])
    written |= {"AT_SAPAge", "AT_BYAge", "AT_PrincipalAge",
                "AT_PrincipalAgeCode", "AT_PrincipalAgeDescription"}

    # ── 8. Country of Origin ─────────────────────────────────────────────
    _val(vals_el, "AT_CountryOrigin", art["coo_raw"], id_val=art["coo"])
    written.add("AT_CountryOrigin")

    # ── 9. Season (injected from args, stored in art dict) ───────────────
    sea_raw    = art.get("season", "")
    sea_code   = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label  = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_code)
    written.add("AT_Season")

    # Derive 4-digit year from 2-digit or 4-digit season suffix
    year_raw = sea_raw[2:] if len(sea_raw) > 2 else ""
    year_val = f"20{year_raw}" if len(year_raw) == 2 else year_raw  # FW26→2026, FW2026→2026
    _val(vals_el, "AT_SeasonYear", year_val)
    written.add("AT_SeasonYear")

    # ── 10. Article category & type ───────────────────────────────────────
    cat_code  = art["art_category"]
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, cat_code)
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)
    written.add("AT_SAPArticleCategory")

    at_norm  = art["article_type"].capitalize()
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)
    written.add("AT_BYArticleType")

    # ── 11. Indicators ────────────────────────────────────────────────────
    _val(vals_el, "AT_BYIndicator",   "Yes", id_val="Y")
    _val(vals_el, "AT_SAPIndicator",  "No",  id_val="N")
    _val(vals_el, "AT_EcomIndicator", "No",  id_val="N")
    _val(vals_el, "AT_UOM", "Each", id_val="EA")
    written |= {"AT_BYIndicator", "AT_SAPIndicator", "AT_EcomIndicator", "AT_UOM"}

    # ── Material Type — ZINA ─────────────────────────────────────
    _val(vals_el, "AT_MaterialType", id_val="ZINA")
    written |= {"AT_MaterialType"}

    # ── 12. Pricing ───────────────────────────────────────────────────────
    _val(vals_el, "AT_OriginalPrice", art["rrp"])
    _val(vals_el, "AT_CurrentPrice",  art["rrp"])
    _val(vals_el, "AT_FOB",           art["fob"])
    _val(vals_el, "AT_RetailPriceCurrency")   # blank — set per country
    written |= {"AT_OriginalPrice", "AT_CurrentPrice", "AT_FOB", "AT_RetailPriceCurrency"}

    # ── 13. Vendor ────────────────────────────────────────────────────────
    # _val(vals_el, "AT_VendorName", art["supplier"])
    _val(vals_el, "AT_MainVendorIdentification", "1")
    written |= {"AT_VendorName", "AT_MainVendorIdentification"}

    # ── 14. Merchandise / hierarchy ───────────────────────────────────────
    # _val(vals_el, "AT_MerchandiseCategory", art["merch_category"])
    # _val(vals_el, "AT_ProductHierarchyL1",  art["division"])
    # _val(vals_el, "AT_ProductHierarchyL2",  art["line_plan_biz"])
    # _val(vals_el, "AT_ProductHierarchyL3",  art["silhouette"])
    # _val(vals_el, "AT_SAPProductGroup",     art["line_plan_biz"])
     # ── Principal Merchandise Hierarchy ──────────────────────────────────
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art.get("division", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art.get("silhouette", ""))
    written |= {
        "AT_PrincipalMerchandiseHierarchyL1",
        "AT_PrincipalMerchandiseHierarchyL2",
        "AT_PrincipalMerchandiseHierarchyL3",
    }
    if art.get("silhouette"):
        sil_val = art["silhouette"].strip()
        # Fetch LOV ID from MDD Silhouette LOV tab at runtime
        silhouette_lov = (mdd.lovs.get("Silhouette") or {}) if mdd else {}
        silhouette_id = (
            silhouette_lov.get(sil_val)                  # exact match
            or silhouette_lov.get(sil_val.title())       # e.g. "Pants"
            or silhouette_lov.get(sil_val.upper())       # e.g. "PANTS"
            or ""
        )
        log.info(
            "[Silhouette] '%s' → LOV ID '%s' (from MDD)",
            sil_val, silhouette_id or "NOT FOUND"
        )
        _val(vals_el, "AT_Silhouette", sil_val, id_val=silhouette_id)
    written |= {
         "AT_ProductHierarchyL1",
        "AT_ProductHierarchyL2",  "AT_ProductHierarchyL3",
        "AT_SAPProductGroup",     "AT_Silhouette",
    }

    # ── 15. Sports category ───────────────────────────────────────────────
    _val(vals_el, "AT_SportsCategoryEN", art["sports_cat"])
    written.add("AT_SportsCategoryEN")

    # ── 16. Technology / fabric / collection ──────────────────────────────
    _val(vals_el, "AT_TechnologyUsed", art["technology"])
    # _val(vals_el, "AT_Content",        art["fabric"])
    # _val(vals_el, "AT_Material",       art["fabric"])
    _val(vals_el, "AT_Collection1",    art["collection"])
    written |= {"AT_TechnologyUsed", "AT_Content", "AT_Material", "AT_Collection1"}

    # ── 17. Country size / fit ────────────────────────────────────────────
    # Only write if value exists (skip when blank/N/A)
    if art["country_size"]:
        _val(vals_el, "AT_CountrySize", art["country_size"])
        written.add("AT_CountrySize")

    # ── 18. Generic code ──────────────────────────────────────────────────
    _val(vals_el, "AT_InboundGenericCode", art["generic_code"])
    written.add("AT_InboundGenericCode")

    # ── 19. Nature of Article — default Retail ────────────────────────────
    _val(vals_el, "AT_NatureOfArticle", "Retail", id_val="RT1")
    written.add("AT_NatureOfArticle")

    # AT_Country — from filename country token (e.g. MY, ID)
    # Only write if not already written
    if "AT_Country" not in written:
        country_val = art.get("country_code", "")
        if country_val:
            _val(vals_el, "AT_Country", country_val, id_val=country_val)
        written.add("AT_Country")





# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS BUILDER
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str,
    brand_code: str,
    season_code: str,
) -> ET.Element:
    """
    Build the <Classifications> block.
    Season ID format: CLH_{brand_code}_{SeasonPrefix}{Year}  e.g. CLH_NEW_FW2026
    """
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"   # e.g. CLH_NEW_FW2026
    batches_parent = f"CLH_{brand.replace(' ', '')}Batches"

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

def build_product_xml(
    art: dict,
    brand: str,
    brand_code: str,
    comp_code: str,
    sbu: str,
    season_id: str,
    mdd=None
) -> str:
    """
    Build one Generic + N Variant product fragment for a single SKU.

    Structure:
      <Product UserTypeID="PRD_GenericArticle" ParentID="PPH_A-TempSubCat">
        <KeyValue KeyID="KEY_InboundArticle">NEW{generic_code}</KeyValue>
        <Name>...</Name>
        <ClassificationReference .../>  × 3
        <Values> ... </Values>
        <Product UserTypeID="PRD_VariantArticle">  ← one per size
          <KeyValue KeyID="KEY_Variant">...</KeyValue>
          <Values> ... </Values>
        </Product>
      </Product>
    """
    item_number = art["item_number"]
    if not item_number:
        return ""

    b_code      = art.get("brand_code", brand_code) or brand_code
    div_letter  = _get_division_code(art.get("division", ""))
    parent_id   = f"PPH_{div_letter}-TempSubCat"

    # Generic product element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = art['generic_code']

    # Product display name as <Name>
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["display_name"] or item_number

    # Classification references
    # 1. Merchandiser
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # 2. BY Hierarchy — built from division + merch_category
    div_id  = re.sub(r"[^A-Z0-9]", "", (art["division"] or "").upper())[:2]
    merch_id = re.sub(r"[^A-Z0-9]", "", (art["merch_category"] or "").upper())[:4]
    # cr_by   = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{brand_code}_{div_id}_{merch_id}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # 3. SAP Hierarchy
    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{div_id}")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

    # 4. Unconfirmed season
    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # Generic values
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu,mdd=mdd)

  

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_sku(task):
    """Thread worker: map + validate one Price List row."""
    row, brand_code, mdd = task
    mapped = map_sku(row, brand_code=brand_code)
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — ORCHESTRATOR  (run)
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point called by lambda_function.py (or CLI).

    args must have:
        brand      — e.g. "New Balance"
        brand_code — e.g. "NEW"
        comp_code  — e.g. "0888"
        sbu        — e.g. "SP"
        season     — e.g. "FW26"
        seq        — e.g. 1
    """
    def first(d: Path, ext="*.xlsx"):
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    pl_files = list(PRICELIST_DIR.glob("*.xlsx"))

    # Validate required inputs
    for label, val in [
        ("MDD",        mdd_f),
        ("Attributes", attr_f),
        ("Price List", pl_files),
    ]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    # Load support files
    mdd    = MDDLoader(mdd_f)
    _      = AttributesListLoader(attr_f, brand=args.brand)

    all_warnings: list[str] = []

    # Compute season_id  e.g. CLH_NEW_FW2026
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    for pl_path in pl_files:
        log.info("─── Processing Price List: %s ───", pl_path.name)

        pl = PriceListLoader(pl_path)
        if pl.df.empty:
            log.warning("[PriceList] Empty dataframe — skipping.")
            continue

        # Filter rows with a valid Item Number
        item_col = next(
            (c for c in ["Item Number", "Item No.", "Item No"] if c in pl.df.columns), None
        )
        rows = []
        for _, row in pl.df.iterrows():
            item = str(row.get(item_col or "Item Number", "")).strip()
            if item and item not in ("None", "nan", ""):
                rows.append(row)
        # testing - 1 product 
        # rows = rows[:5]
        total_rows = len(rows)
        log.info("[PriceList] %d valid SKU rows to process", total_rows)

        # Build output filename
        _file_type = getattr(args, "file_type", None) or "Line List"

        # Extract Multi/Mono from price list filename (e.g. "...Multi..." or "...Mono...")
        pl_stem = pl_path.stem  # filename without extension
        multi_mono_match = re.search(r'\b(Multi|Mono)\b', pl_stem, re.IGNORECASE)
        multi_mono = multi_mono_match.group(1).capitalize() if multi_mono_match else "Multi"

        # Output filename = input filename stem + .xml  (same-to-same)
        out_name = f"{pl_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────────────
        log.info("Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
                 total_rows, min(8, total_rows))

        num_workers = min(8, max(1, total_rows))
        # Inject season into each row so map_sku can access it
        task_args = [(row, args.brand_code, mdd) for row in rows]
        ordered: list[tuple[int, dict, list]] = []

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_sku, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx, (mapped, warns) = futures[fut], fut.result()
                # Inject season from args into mapped dict (not in price list)
                mapped["season"] = args.season
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_skus = [m for _, m, _ in ordered]
        del ordered

        # Inject filename-level fields into every mapped article
        _country = getattr(args, "country_code", "")
        for art in mapped_skus:
            if not art.get("country_code"):
                art["country_code"] = _country

        # total_variants = sum(len(a.get("sizes", [])) for a in mapped_skus)
        log.info("Pass 1 done — mapped=%d", len(mapped_skus))

        # ── Pass 2: stream XML ────────────────────────────────────────────
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
                    args.comp_code, args.sbu, season_id, mdd=mdd,
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
        print(f"  Price List rows     : {total_rows}",   flush=True)
        print(f"  Mapped SKUs         : {written_count}", flush=True)
        # print(f"  Total size variants : {total_variants}", flush=True)
        print(f"  XML file size       : {file_kb}KB",    flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    # ── Validation report ─────────────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_nb_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — New Balance v1.0")
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="FW26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
