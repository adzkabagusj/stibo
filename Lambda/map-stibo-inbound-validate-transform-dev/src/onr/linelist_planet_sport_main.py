"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — ON Running  v1.0            ║
║  Line List (Planet Sport) → Stibo STEP XML                     ║
╚══════════════════════════════════════════════════════════════════╝

ON Running (ONR) Line List (Planet Sport) integration.

Linesheet columns → Stibo attribute mapping:
  • Item Number       → AT_PrincipalStyleCode  (principal article code)
  • Style Name        → AT_PrincipalStyleDescription
  • Color             → AT_PrincipalColorDescription  (e.g. "Ash | Cinder")
  • Gender            → AT_Gender / AT_BYGender / AT_PrincipalGender
  • Style Code        → AT_SAPStyleCode
  • Retail Price      → AT_OriginalPrice / AT_CurrentPrice
  • Launch Date       → AT_LaunchingDate
  • Selection         → AT_SAPProductDivision  (Shoes → F Footwear)
  • Vertical          → AT_SAPProductGroup / AT_SAPProductCategory
  • Vertical          → AT_PrincipalMerchHierarchyL2
  • Selection         → AT_PrincipalMerchHierarchyL1

  One Generic per (Item Number + Color) row.
  Article Category = Generic (1) with Variant sub-products for sizes.
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

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
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
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK", "ONR": "ON RUNNING",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex", "K": "Kids"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter",
}
LOV_COUNTRY_ORIGIN = {
    "CN": "China", "VN": "Vietnam", "ID": "Indonesia",
    "KH": "Cambodia", "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia", "TH": "Thailand",
}
LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
    "SM": "Smiggle",
}

# ── ON Running–specific lookup tables ────────────────────────────

# Gender column → SAP Gender code
ONR_GENDER_MAP: dict[str, str] = {
    "M":      "M",
    "W":      "F",
    "WOMEN":  "F",
    "MEN":    "M",
    "UNISEX": "U",
    "U":      "U",
    "YOUTH":  "K",
    "KIDS":   "K",
    "K":      "K",
}

# Gender code → SAP Age code (ON Running is an adult sports brand by default)
ONR_GENDER_TO_AGE: dict[str, str] = {
    "M":  "AD",   # Men → Adults
    "F":  "AD",   # Women → Adults
    "U":  "AD",   # Unisex → Adults
    "K":  "CH",   # Kids/Youth → Children
}

# Selection column → SAP Product Division
ONR_SELECTION_TO_DIVISION: dict[str, str] = {
    "SHOES":     "F",    # Footwear
    "FOOTWEAR":  "F",
    "APPAREL":   "A",
    "ACCESSORIES": "E",
}

# Vertical column → SAP Product Group code
ONR_VERTICAL_TO_PROD_GROUP: dict[str, str] = {
    "PERFORMANCE ALL DAY":  "PA",
    "PERFORMANCE RUNNING":  "PR",
    "PERFORMANCE OUTDOOR":  "PO",
    "PERFORMANCE TRAINING": "PT",
    "CLASSIC":              "CL",
    "LIFESTYLE":            "LS",
    "RUNNING":              "RU",
    "TRAIL":                "TR",
    "TRAINING":             "TN",
    "ALL DAY":              "AD",
    "OUTDOOR":              "OD",
}

# Vertical → SAP Product Category (same mapping; refine as needed)
ONR_VERTICAL_TO_PROD_CATEGORY: dict[str, str] = {
    "PERFORMANCE ALL DAY":  "RU",   # Running
    "PERFORMANCE RUNNING":  "RU",
    "PERFORMANCE OUTDOOR":  "OD",
    "PERFORMANCE TRAINING": "TN",
    "CLASSIC":              "CS",   # Casual
    "LIFESTYLE":            "CS",
}

# Country of Origin default for ON Running (local vendor in VN/CN)
ONR_DEFAULT_COO = "VN"


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
        wb  = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws  = wb["Core Attributes"]
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
        self._load_age_lov(wb)
        self._load_retail_price_currency_lov(wb)
        self._load_sports_category_lov(wb)
        self._load_brand_group_lov(wb)
        wb.close()
        self._build_country_reverse()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))
        
    def _load_brand_group_lov(self, wb):
        """Load BrandGroup LOV as {brand_code → brand_group_id} — same as TAF load_mdd."""
        if "BrandGroup LOV" not in wb.sheetnames:
            # Fallback: same as TAF fallback
            brand_lov = self.lovs.get("Brand", {})
            self.lovs["BrandGroupByCode"] = {
                code: ("ON" if code == "ONR" else code)
                for code in brand_lov.values()
            }
            log.warning("[MDD] BrandGroup LOV sheet not found — using fallback")
            return
        brand_group_by_code: dict[str, str] = {}
        for row in wb["BrandGroup LOV"].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            brand_code = str(row[0]).strip()                                        # col A = brand code
            group_id   = str(row[1]).strip() if len(row) > 1 and row[1] else brand_code  # col B = group ID
            if brand_code:
                brand_group_by_code[brand_code] = group_id
        self.lovs["BrandGroupByCode"] = brand_group_by_code
        log.info("[MDD] BrandGroupByCode: %d entries", len(brand_group_by_code))
        
    def _load_sports_category_lov(self, wb):
        """Load Sports Category LOV: Column A = display name, Column B = LOV ID."""
        sc_sheet = next(
            (s for s in wb.sheetnames if "SPORTS CATEGORY" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sc_sheet:
            log.warning("[MDD] Sports Category LOV sheet not found")
            return
        sports_category: dict[str, str] = {}   # {display_name.upper(): lov_id}
        for row in wb[sc_sheet].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()                               # Column A = display name
            code = str(row[1]).strip() if len(row) > 1 and row[1] else ""  # Column B = LOV ID
            if name and code:
                sports_category[name.upper()] = code
        self.lovs["SportsCategory"] = sports_category
        log.info("[MDD] Sports Category LOV: %d entries", len(sports_category))
        
    def _build_country_reverse(self):
        """Build CountryByCode: code → name, from the Country LOV (name→code) already loaded."""
        country_lov = self.lovs.get("Country", {})
        self.lovs["CountryByCode"] = {code: name for name, code in country_lov.items()}
        log.info("[MDD] CountryByCode reverse map: %d entries", len(self.lovs["CountryByCode"]))

    def _load_retail_price_currency_lov(self, wb):
        """Load Retail Price Currency LOV: {country_name_upper -> currency_code}."""
        if "Retail Price Currency LOV" not in wb.sheetnames:
            return
        retail_price_currency_lov: dict[str, str] = {}
        for row in wb["Retail Price Currency LOV"].iter_rows(min_row=2, values_only=True):
            if row and row[0] and len(row) > 1 and row[1]:
                retail_price_currency_lov[str(row[0]).strip().upper()] = str(row[1]).strip()
        self.lovs["RetailPriceCurrency"] = retail_price_currency_lov
        log.info("[MDD] Retail Price Currency LOV: %d entries", len(retail_price_currency_lov))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = str(val_id).strip() if val_id else str(val_name).strip()

    def _load_age_lov(self, wb):
        sheet_name = next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return
        sap_age_map = {}
        by_age_map  = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display  = str(row[0]).strip()
            sap_age_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            by_age_val   = str(row[5]).strip() if len(row) > 5 and row[5] else ""
            if age_display and sap_age_code:
                sap_age_map[age_display] = sap_age_code
            if age_display and by_age_val:
                by_age_map[age_display] = by_age_val
        self.lovs["SAPAge"] = sap_age_map
        self.lovs["BYAge"]  = by_age_map
        log.info("[MDD] Age LOV loaded — SAPAge: %d entries, BYAge: %d entries",
                len(sap_age_map), len(by_age_map))

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
    """Loads the ON RUNNING tab from the Attributes List workbook."""

    def __init__(self, path: Path, brand: str = "ON RUNNING"):
        self.path      = path
        self.brand     = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()),
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
        log.info("[AttrList] %d attributes loaded from '%s'", len(self.attr_map), sheet_name)

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class RNALoader:
    """
    Loads the 'Source Mapping Related RNA' tab from the Attributes List workbook.

    The tab maps (Country Name, Company Code, SBU, Brand Code) → Brand Type + Brand Category.
    These four key columns come from the input filename metadata.
    """

    RNA_SHEET_KEYWORDS = ["RNA", "SOURCE MAPPING"]

    @staticmethod
    def _norm_country(v: str) -> str:
        return (v or "").strip().upper()

    @staticmethod
    def _norm_comp_code(v: str) -> str:
        s = (v or "").strip()
        if not s:
            return ""
        s2 = s.lstrip("0")
        return s2 if s2 else "0"

    def __init__(self, path: Path):
        self.path   = path
        # key: (country_upper, comp_code_upper, sbu_upper, brand_code_upper)
        self.lookup: dict[tuple, dict] = {}
        self._load()

    def _load(self):
        log.info("[RNA] Loading from: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found in %s", self.path.name)
            wb.close()
            return

        log.info("[RNA] Using sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))

        # Find header row: first row that contains a "COUNTRY" column
        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(isinstance(v, str) and "COUNTRY" in v.upper() for v in r if v)),
            None,
        )
        if hdr_idx is None:
            log.warning("[RNA] Cannot find header row in sheet '%s'", sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip().upper(): i for i, h in enumerate(hdr) if h}

        def _find(*candidates) -> int | None:
            for c in candidates:
                if c.upper() in col:
                    return col[c.upper()]
            return None

        c_country = _find("COUNTRY", "COUNTRY NAME")
        c_comp    = _find("COMPANY CODE", "COMP CODE", "COMPCODE", "COMP_CODE")
        c_sbu     = _find("SBU")
        c_bcode   = _find("BRANDCODE", "BRAND CODE", "REPORTING BRAND CODE MAPPED")
        c_btype   = _find("BRANDTYPE_DETAIL", "BRAND TYPE", "AT_BRANDTYPE", "BRANDTYPE")
        c_bcat    = _find("BRANDCATEGORY", "BRAND CATEGORY", "AT_BRANDCATEGORY", "BRANDCATEGORY")

        missing = [nm for nm, idx in [("Country", c_country), ("CompCode", c_comp),
                          ("SBU", c_sbu), ("BrandCode", c_bcode),
                                      ("BrandType", c_btype), ("BrandCategory", c_bcat)]
                   if idx is None]
        if missing:
            log.warning("[RNA] Missing columns in '%s': %s", sheet_name, missing)

        def _cell(row, idx):
            if idx is None or idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()

        count = 0
        for row in rows[hdr_idx + 1:]:
            country   = _cell(row, c_country)
            comp_code = _cell(row, c_comp)
            sbu       = _cell(row, c_sbu)
            brand_code = _cell(row, c_bcode)
            btype     = _cell(row, c_btype)
            bcat      = _cell(row, c_bcat)
            if not country:
                continue
            key = (
                self._norm_country(country),
                self._norm_comp_code(comp_code),
                (sbu or "").strip().upper(),
                (brand_code or "").strip().upper(),
            )
            self.lookup[key] = {"brand_type": btype, "brand_category": bcat}
            count += 1

        wb.close()
        log.info("[RNA] %d entries loaded from '%s'", count, sheet_name)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        """Return {brand_type, brand_category} for the given keys, or empty strings."""
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu          or "").strip().upper(),
            (brand_code   or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": ""})


class ONRLinelistLoader:
    """
    Loads the ON Running linesheet workbook (Planet Sports).

    Sheet: PLANET SPORTS (or first sheet).
    Header row: row 7 (0-indexed) containing columns:
        Dealers Account, Selection, Gender, Vertical, Style Code,
        Item Number, Style Name, Color, Retail Price, Launch Date, Youth

    Size columns (US sizes) start at col 11+ and are used to build
    variant-level products with size information.

    Each row = one (Item Number + Color) combination = one Generic article.
    Rows with Youth='Youth' or 'Kids' are treated as Kids gender.
    """

    PREFERRED_SHEETS = ["PLANET SPORTS", "Linelist", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self.size_columns: list[str] = []   # US size column names for variant expansion
        self._load()

    def _load(self):
        log.info("[Linelist-ONR] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in self.PREFERRED_SHEETS if s in wb.sheetnames), None
        )
        if target is None:
            best, best_rows = wb.sheetnames[0], 0
            for sn in wb.sheetnames:
                n = wb[sn].max_row or 0
                if n > best_rows:
                    best, best_rows = sn, n
            target = best

        log.info("[Linelist-ONR] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Auto-detect header row: first row containing "Item Number" or "Style Code"
        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and ("ITEM NUMBER" in v.upper() or "STYLE CODE" in v.upper()) for v in r if v):
                hdr_idx = i
                break

        if hdr_idx is None:
            hdr_idx = next(
                (
                    i for i, r in enumerate(rows)
                    if sum(1 for v in r[:12] if isinstance(v, str) and v.strip()) >= 5
                ),
                None,
            )

        if hdr_idx is None:
            log.error("[Linelist-ONR] Cannot find header row"); wb.close(); return

        log.info("[Linelist-ONR] Header at row %d (0-indexed)", hdr_idx)
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]

        # Identify size columns (numeric headers after "Launch Date" / "Youth")
        # These are US sizes like 7, 7.5, 8, 8.5, etc.
        size_start_idx = None
        for i, h in enumerate(header):
            try:
                float(h)
                if size_start_idx is None:
                    size_start_idx = i
            except (ValueError, TypeError):
                pass
        if size_start_idx:
            self.size_columns = [h for h in header[size_start_idx:] if h and h != f"col_{header.index(h) if h in header else 0}"]
            # Filter only numeric-looking size columns
            self.size_columns = [h for h in self.size_columns
                                 if re.match(r'^\d+\.?\d*$', h)]

        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Drop rows where Item Number is blank
        item_col = next(
            (c for c in df.columns if "ITEM" in c.upper() and "NUMBER" in c.upper()), None
        )
        if item_col:
            df = df[
                df[item_col].notna()
                & (~df[item_col].astype(str).str.strip().isin(["", "None", "nan", "TBC"]))
            ]
            df = df.rename(columns={item_col: "Item Number"})

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info(
            "[Linelist-ONR] %d rows loaded | %d size columns detected",
            len(self.df), len(self.size_columns),
        )


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT") else s


def _fmt_date(v) -> str:
    """Format a date to dd-Mon-YYYY lowercase."""
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw:
        return raw
    # Handle "Carry Over" — no date
    if "carry" in raw.lower():
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def map_article_onr(ll_row: dict, brand_code: str = "ONR") -> dict:
    """
    Map one ON Running linesheet row → unified article dict.

    Linesheet columns → article fields:
      - article_no         = Item Number (e.g. 3MF10074782)
      - style_code         = Style Code  (e.g. 3MF1007)
      - model_name         = Style Name  (e.g. Cloud 6)
      - colour             = Color       (e.g. "Ash | Cinder")
      - gender_code        = Gender      (M/W → M/F)
      - selection          = Selection   (Shoes)
      - vertical           = Vertical    (Performance All Day)
      - retail_price       = Retail Price
      - launch_date        = Launch Date (or "Carry Over")
      - youth              = Youth column (if present, override gender to K)
      - dealer             = Dealers Account
    """
    item_number  = _s(ll_row.get("Item Number"))
    style_code   = _s(ll_row.get("Style Code"))
    style_name   = (_s(ll_row.get("Collection 1")) or  
                    _s(ll_row.get("Collection"))   or
                    _s(ll_row.get("Style Name"))) 
    color_raw    = _s(ll_row.get("Color"))
    gender_raw   = _s(ll_row.get("Gender"))
    selection    = _s(ll_row.get("Selection"))
    vertical     = (_s(ll_row.get("Sports Category EN")) or   
                    _s(ll_row.get("Sports Category"))    or
                    _s(ll_row.get("Vertical")))
    retail_price = _s(ll_row.get("Retail Price"))
    launch_date  = ll_row.get("Launch Date")
    youth_flag   = _s(ll_row.get("Youth"))
    dealer       = _s(ll_row.get("Dealers Account"))
    country_size      = _s(ll_row.get("Country Size")) or "US"                 
    ecom_age_category = (_s(ll_row.get("E-com Ages Category")) or                 
                         _s(ll_row.get("E-Com Ages Category")) or
                         _s(ll_row.get("Ecom Ages Category"))) or "Ages 18+ years"

    # ── Derived fields ──────────────────────────────────────────

    # SAP Gender: from Gender column, override to K if Youth is flagged
    if youth_flag and youth_flag.upper() in ("YOUTH", "Y", "YES", "KIDS"):
        gender_code = "K"
    else:
        gender_code = ONR_GENDER_MAP.get(gender_raw.upper(), "U")

    # SAP Age: Adults by default, Children if youth
    age_code = ONR_GENDER_TO_AGE.get(gender_code, "AD")

    # SAP Product Division from Selection
    div_letter = ONR_SELECTION_TO_DIVISION.get(selection.upper(), "F")

    # SAP Product Group from Vertical
    prod_group = ONR_VERTICAL_TO_PROD_GROUP.get(vertical.upper(), "PA")

    # SAP Product Category from Vertical
    prod_category = ONR_VERTICAL_TO_PROD_CATEGORY.get(vertical.upper(), "RU")

    # Country of Origin — default for ON Running
    coo = ONR_DEFAULT_COO

    # Launch Date formatting
    formatted_date = _fmt_date(launch_date)

    # RRP: extract numeric
    rrp_str = ""
    if retail_price:
        m = re.search(r"[\d.]+", str(retail_price))
        rrp_str = m.group() if m else str(retail_price)

    # Generic code: ONR + Style Code (brand 3 chars + style)
    generic_code = f"{brand_code}{style_code}"

    # Color token: extract from pipe-separated color string (first part, cleaned)
    color_parts = color_raw.split("|")
    color_token = re.sub(r"[^A-Z0-9]", "", color_parts[0].upper().strip())[:3].ljust(3, "0")

    # Variant code: Generic + color_token + size placeholder
    variant_code = f"{generic_code}{color_token}000"

    return {
        # Core identifiers
        "article_no":       item_number,
        "style_code":       style_code,
        "model_name":       style_name,
        "brand_code":       brand_code,

        # Colour
        "colour":           color_raw,
        "colour_token":     color_token,

        # Gender / Age
        "gender_code":      gender_code,
        "gender_raw":       gender_raw,
        "age_code":         age_code,
        "age_raw":          "Adults" if age_code == "AD" else "Children",

        # Classification
        "selection":        selection,
        "vertical":         vertical,
        "division":         selection,
        "div_letter":       div_letter,
        "category":         vertical,
        "prod_group":       prod_group,
        "prod_category":    prod_category,

        # Commercial
        "article_type":     "Inline",        # Default: Inline
        "art_category":     "1",             # Generic (footwear with sizes)
        "franchise":        style_name,      # Style Name as franchise
        "launch_date":      formatted_date,

        # Pricing
        "rrp":              rrp_str,
        "currency":         "IDR",           # Indonesia — Planet Sports
        "fob":              "",              # Not in linesheet

        # Origin
        "coo":              coo,

        # Derived codes
        "generic_code":     generic_code,
        "variant_code":     variant_code,
        "sap_style_code":   style_code,

        # Dealer info
        "dealer":           dealer,

        # Hierarchy
        "merch_hierarchy_l1": selection,
        "merch_hierarchy_l2": vertical,

        # No TDD / Backlog data
        "_tdd_sizes":       [],
        "country_size":      country_size,     
        "ecom_age_category": ecom_age_category, 
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Validate mandatory MDD attributes against the mapped article dict."""
    warns = []
    art   = mapped["article_no"]
    field_checks = {
        "article_no":   "AT_PrincipalStyleCode",
        "gender_code":  "AT_Gender",
        "age_code":     "AT_SAPAge",
        "brand_code":   "AT_Brand",
    }
    for field, at_id in field_checks.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id}")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(
    parent:   ET.Element,
    attr_id:  str,
    value:    str = "",
    id_val:   str = "",
    derived:  bool = False,
) -> ET.Element | None:
    if derived:
        return None
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


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = "") -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def resolve_brand_lov_id(
    mdd: MDDLoader | None,
    brand_name: str,
    brand_code: str = "",
) -> tuple[str, str]:
    """
    Resolve Brand LOV ID dynamically from MDD "Brand" LOV using LOV value text.

    Priority:
      1) Match by input brand name (from filename metadata)
      2) Fallback to mapped label from brand_code (legacy behavior)
      3) Fallback ID to brand_code (or first 3 chars of brand name)
    """
    src_brand_name = (brand_name or "").strip()
    src_code = (brand_code or "").strip().upper()

    # Keep brand label from input brand name as the primary source
    brand_label = src_brand_name or _lov(src_code, LOV_BRAND, src_code)[1]

    # Fallback ID when MDD LOV lookup does not find a match
    fallback_id = src_code or (re.sub(r"[^A-Z]", "", src_brand_name.upper())[:3])
    brand_lov_id = fallback_id

    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})

        def _norm(s: str) -> str:
            return re.sub(r"\s+", " ", (s or "").strip()).upper()

        target_names = [brand_label]

        # Legacy fallback: if code exists in static table, also try that display label
        if src_code:
            mapped_label = _lov(src_code, LOV_BRAND, src_code)[1]
            if mapped_label and _norm(mapped_label) != _norm(brand_label):
                target_names.append(mapped_label)

        for target in target_names:
            t_norm = _norm(target)
            if not t_norm:
                continue
            for display_name, lov_id in brand_lov.items():
                if _norm(str(display_name)) == t_norm:
                    brand_lov_id = str(lov_id).strip()
                    return brand_lov_id, str(display_name).strip()

    return brand_lov_id, brand_label


def _resolve_retail_price_currency(country_code: str, mdd) -> str:
    """
    Resolve retail price currency from country code using MDD.
    
    Steps:
    1. Look up country_code (e.g., "ID") in CountryByCode → get country name (e.g., "Indonesia")
    2. Look up country name in RetailPriceCurrency LOV → get currency (e.g., "IDR")
    """
    if not country_code or not mdd:
        return ""
    
    code_upper = country_code.strip().upper()
    
    # Step 1: Get country name from country ID
    country_by_code = mdd.lovs.get("CountryByCode", {})
    country_name = country_by_code.get(code_upper, "")
    
    if not country_name:
        return ""
    
    # Step 2: Look up currency from country name (partial match)
    retail_currency_lov = mdd.lovs.get("RetailPriceCurrency", {})
    country_name_upper = country_name.strip().upper()
    
    # Try exact match first
    if country_name_upper in retail_currency_lov:
        return retail_currency_lov[country_name_upper]
    
    # Try partial match (country name contains or is contained in LOV key)
    for lov_country, currency in retail_currency_lov.items():
        if country_name_upper in lov_country or lov_country in country_name_upper:
            return currency
    
    return ""

_SPORTS_CATEGORY_MAP: dict[tuple, str] = {
    ("Shoes",   "Performance All Day"):  "Lifestyle / Casual",
    ("Shoes",   "Performance Running"):  "Running",
    ("Shoes",   "Performance Training"): "Fitness / Training",
    ("Shoes",   "Performance Outdoor"):  "Outdoor / Trail / Hiking",
    ("Shoes",   "Performance Tennis"):   "Tennis / Padel",
    ("Apparel", "Performance All Day"):  "Lifestyle / Casual",
    ("Apparel", "Performance Running"):  "Running",
    ("Apparel", "Performance Training"): "Fitness / Training",
    ("Apparel", "Performance Tennis"):   "Tennis / Padel",
    (None, "Performance All Day"):       "Lifestyle / Casual",
    (None, "Performance Outdoor"):       "Outdoor / Trail / Hiking",
    (None, "Performance Running"):       "Running",
    (None, "Performance Training"):      "Fitness / Training",
    (None, "Performance Tennis"):        "Tennis / Padel",
}


def resolve_sports_category_id(selection: str, vertical: str, mdd) -> tuple[str, str]:
    """Translate Selection+Vertical → (label, LOV ID) via MDD Sports Category LOV."""
    label = (
        _SPORTS_CATEGORY_MAP.get((selection, vertical))
        or _SPORTS_CATEGORY_MAP.get((None, vertical))
    )
    if not label:
        return "", ""
    sc_lov = mdd.lovs.get("SportsCategory", {}) if mdd else {}
    sc_id  = sc_lov.get(label.upper(), "")
    # Pad single-digit IDs: "7" → "07"
    if sc_id and sc_id.isdigit() and len(sc_id) == 1:
        sc_id = sc_id.zfill(2)
    return label, sc_id

def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
) -> None:
    """
    Write all Generic-level <Value> elements for an ON Running article.
    Mapping based on Attributes List (ON RUNNING tab) and MDD Core Attributes.
    """
    written: set[str] = set()

    def _w(attr_id: str, value: str = "", id_val: str = "", derived: bool = False):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val, derived=derived)
            written.add(attr_id)

    def _mw(attr_id: str, id_val: str):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    sbu_code, _ = _lov(sbu, LOV_SBU, sbu)
    _mw("AT_SBU", sbu_code)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    brand_lov_id = art["brand_code"]
    _w("AT_Brand",      id_val=brand_lov_id)
    
    bg_by_code = mdd.lovs.get("BrandGroupByCode", {}) if mdd else {}
    brand_group_id = bg_by_code.get(brand_lov_id, brand_lov_id)
    _w("AT_BrandGroup", id_val=brand_group_id)

    # ── Principal identifiers ────────────────────────────────────
    # Item Number → AT_PrincipalStyleCode (from Attributes List: "FW App / Line sheet : Item Number")
    article_transformed = (
        art['article_no'][1:2] + art['article_no'][3:]
        if len(art['article_no']) > 3
        else art['article_no']
    )
    _w("AT_PrincipalStyleCode", art["article_no"])

    # Style Name → AT_PrincipalStyleDescription (from Attributes List: "FW App / Line sheet : Style Name")
    _w("AT_PrincipalStyleDescription", art["model_name"])

    # Color → AT_PrincipalColorName (from Attributes List: "FW App / Line sheet : Color")
    if art.get("colour"):
        _w("AT_PrincipalColorName", art["colour"])

    # AT_SAPStyleCode = Principal Item Number (omitting 1st and 3rd characters)
    # Example: 3ME10051043 → M10051043
    _w("AT_SAPStyleCode", article_transformed)

    # AT_InboundGenericCode = BrandCode + ArticleNo (omitting 1st and 3rd characters)
    # Example: 3ME10051043 → M10051043, then ONR + M10051043 = ONRM10051043
    at_generic_val = f"{brand_lov_id}{article_transformed}"
    art["at_generic_val"] = at_generic_val
    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    # Gender → AT_Gender (from Attributes List: "FW App / Line sheet : Gender")
    g_code = art["gender_code"]
    g_label = LOV_GENDER.get(g_code, g_code)
    _w("AT_Gender",            g_label, id_val=g_code)
    _w("AT_BYGender",          g_label, id_val=g_code)
    _w("AT_PrincipalGenderCode", g_code)  # Gender column

    # ── Age ──────────────────────────────────────────────────────
    # Default: Adults (from Attributes List: "FW/App Default : Adults")
    age_raw_upper = (art["age_raw"] or "").strip()

    sap_age_map = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map  = mdd.lovs.get("BYAge",  {}) if mdd else {}

    matched_age_display = next(
        (k for k in sap_age_map if k.upper() == age_raw_upper.upper()), None
    )

    _w("AT_SAPAge", id_val="AD")
    _w("AT_BYAge", id_val="ADULT")

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    if not year_match:
        short_match = re.search(r"\d{2}$", sea_raw)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    _w("AT_SeasonYear", season_year)

    # ── Article Category — Generic (1) for footwear ──────────────
    _w("AT_SAPArticleCategory", id_val="1")
    _w("AT_MaterialType", id_val="ZHAW")  # Material Type = ZHAW (Trading Goods)
    _w("AT_SAPProductFlag", id_val="5") # SAP Product Flag — default 5 (Direct Local)

    # ── BY Article Type ──────────────────────────────────────────
    # Default: Inline (from Attributes List: "Default : Inline")
    at_code = art["article_type"]
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   "Yes", id_val="Y")
    _w("AT_SAPIndicator",  "No",  id_val="N")

    # ── UOM — default EA (from Attributes List: "Default : EA") ──
    _w("AT_UOM", "Each", id_val="EA")

    # ── Pricing ──────────────────────────────────────────────────
    # Retail Price → AT_OriginalPrice / AT_CurrentPrice
    _w("AT_OriginalPrice",        art["rrp"])
    _w("AT_CurrentPrice",         art["rrp"])
    # AT_RetailPriceCurrency / AT_FOBCurrency: derived from country_code via MDD
    _rpc = _resolve_retail_price_currency(art.get("country_code", ""), mdd)
    if _rpc:
        _w("AT_RetailPriceCurrency", "", id_val=_rpc)
        _w("AT_FOBCurrency",         "", id_val=_rpc)

    # ── Vendor identification ─────────────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Franchise (from Style Name) ──────────────────────────────
    # From Attributes List: "Line Sheet : Style Name"
    _w("AT_Franchise", (art["franchise"] or "").upper())

    # ── Launching / Launch date ──────────────────────────────────
    _w("AT_LaunchingDate", art["launch_date"])

    # ── Principal Merchandise Hierarchy ───────────────────────────
    # L1 = Selection, L2 = Vertical
    _w("AT_PrincipalMerchandiseHierarchyL1", art.get("selection", ""))
    _w("AT_PrincipalMerchandiseHierarchyL2", art.get("vertical", ""))
    _sc_label, _sc_id = resolve_sports_category_id(art.get("selection", ""), art.get("vertical", ""), mdd)
    if _sc_id:
        _w("AT_SportsCategoryEN", id_val=_sc_id)  
    else:
        log.warning("[SportsCategoryEN] No MDD LOV match for vertical='%s' selection='%s'",
                    art.get("vertical", ""), art.get("selection", ""))
    _w("AT_Collection1",      art.get("model_name", ""))

    # ── Width (from Style Name if "Wide" present) ────────────────
    # From Attributes List: "Line sheet : Style name - Consist of 'Wide' so it would be 'W'"
    if "WIDE" in (art["model_name"] or "").upper():
        _w("AT_Width", "W")

    # ── Country ──────────────────────────────────────────────────
    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country",       id_val=country_val)
        _w("AT_CountryOrigin",  id_val=country_val) 

    # ── Brand Type / Brand Category (from RNA tab via filename metadata) ─
    # Lookup key: Country Name + Company Code + SBU + Brand Code
    # Source: Attributes List → Source Mapping Related RNA tab
    if art.get("brand_type"):             
        _w("AT_BrandType", art["brand_type"])
        
    if art.get("brand_category"):      
        _w("AT_BrandCategory", art["brand_category"])         
    
    _w("AT_CountrySize", id_val=art.get("country_size") or "US")  
    _w("AT_EComAgesCategory", id_val="18+Y")


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    brand_token = brand.title().replace(" ", "")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand_token}Batches"

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
    mdd=None
) -> str:
    """
    Build a <Product> XML fragment for one ON Running article.

    ON Running specifics:
      • UserTypeID = PRD_GenericArticle (footwear — Article Category = Generic)
      • ParentID   = PPH_F-TempSubCat  (Footwear division)
      • KeyID      = KEY_InboundArticle, value = {brand_code}{article_no}  (BrandCode + AT_PrincipalStyleCode)
      • ClassificationReferences: Merchandiser, Unconfirmed season
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code      = art.get("brand_code", brand_code) or brand_code
    brand_token = brand.title().replace(" ", "")
    div_letter  = art.get("div_letter", "F")
    parent_id   = f"PPH_{div_letter}-TempSubCat"

    # KEY_Article: BrandCode + AT_PrincipalStyleCode (omitting 1st and 3rd characters)
    # AT_PrincipalStyleCode = art["article_no"] (Item Number from input)
    # Example: 3ME10051043 → M10051043, then ONR + M10051043 = ONRM10051043
    article_transformed = art['article_no'][1:2] + art['article_no'][3:] if len(art['article_no']) > 3 else art['article_no']
    key_article = f"{b_code}{article_transformed}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["model_name"] or article_no
    )

    # ── Classification references ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand_token}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Values ───────────────────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one ON Running row."""
    row, brand_code, mdd, season, country_code, article_type_from_filename, brand_type, brand_category = row_tuple
    mapped = map_article_onr(row, brand_code=brand_code)
    mapped["season"]                    = season
    mapped["country_code"]              = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    mapped["brand_type"]                = brand_type
    mapped["brand_category"]            = brand_category
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

    args must have attributes:
        brand       str   e.g. "On Running"
        brand_code  str   e.g. "ONR"
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
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Linelist", ll_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="ON RUNNING")
    rna    = RNALoader(attr_f)

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    brand_lov_id, brand_lov_label = resolve_brand_lov_id(
        mdd,
        args.brand,
        args.brand_code,
    )
    log.info(
        "Resolved brand LOV → source_brand_code=%s  brand_label=%s  brand_lov_id=%s",
        args.brand_code, brand_lov_label, brand_lov_id
    )

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    # ── Resolve AT_BrandType / AT_BrandCategory via RNA tab ──────
    _country_code  = (getattr(args, "country_code", "") or "").strip().upper() or "ID"
    _country_name  = mdd.lovs.get("CountryByCode", {}).get(_country_code, _country_code)
    _rna_result    = rna.get(_country_name, args.comp_code, args.sbu, args.brand_code)
    _brand_type    = _rna_result.get("brand_type", "")
    _brand_category = _rna_result.get("brand_category", "")
    log.info(
        "RNA lookup → country_code=%s country_name=%s brand_type=%s brand_category=%s",
        _country_code, _country_name, _brand_type, _brand_category,
    )

    all_warnings: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing Linelist: %s ───", ll_path.name)

        ll = ONRLinelistLoader(ll_path)
        if ll.df.empty:
            log.warning("[Linelist-ONR] Empty dataframe — skipping."); continue

        rows     = [row for _, row in ll.df.iterrows()]
        total_rows = len(rows)
        # ── TEST MODE: limit to 5 articles ──────────────────────────
        # rows = rows[:5]
        total_rows = len(rows)
        log.info("Total valid rows: %d", total_rows)

        # ── Parse filename metadata ──────────────────────────────
        _stem  = ll_path.stem
        _parts = re.split(r"\s*-\s*", _stem)

        if len(_parts) >= 7:
            file_type_from_input = "-".join(_parts[3:-3])
        else:
            file_type_from_input = "Linesheet"

        _mm_match = re.search(r'\b(Multi|Mono)\b', _stem, re.IGNORECASE)
        multi_mono_from_input = _mm_match.group(1).capitalize() if _mm_match else "Multi"

        log.info(
            "Parsed from filename → file_type='%s'  multi_mono='%s'",
            file_type_from_input, multi_mono_from_input,
        )

        out_name = ll_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers   = min(8, max(1, total_rows))
        task_args     = [(row, brand_lov_id, mdd, args.season, _country_code, getattr(args, "article_type_from_filename", ""), _brand_type, _brand_category) for row in rows]
        ordered: list[tuple[int, dict, list]] = []

        log.info("Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
                 total_rows, num_workers)

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_article, t): i for i, t in enumerate(task_args)}
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
            args.brand, brand_lov_id, args.comp_code, args.sbu, args.season,
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
                    art, args.brand, brand_lov_id,
                    args.comp_code, args.sbu, season_id, mdd=mdd
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

        print("═══ ARTICLE SUMMARY (ON RUNNING) ═══════════════════", flush=True)
        print(f"  Linelist rows (valid)  : {total_rows}",   flush=True)
        print(f"  Articles written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb}KB",     flush=True)
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
# SECTION 8 — CLI (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — ON Running v1.0")
    p.add_argument("--brand",      default="On Running")
    p.add_argument("--brand-code", default="ONR")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="FW26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
