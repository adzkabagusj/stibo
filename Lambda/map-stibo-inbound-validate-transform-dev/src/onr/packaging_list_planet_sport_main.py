"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — ON Running  v1.0            ║
║  Planet Sport Order Sheet → Stibo STEP XML                     ║
╚══════════════════════════════════════════════════════════════════╝

ON Running (ONR) Planet Sport inbound integration.

Input file: "PLANET SPORT - ON SS26 (FEB 26).xlsx"
Sheet: PLANET SPORT
Header row: row 6 (1-indexed)

Planet Sport columns → Stibo attribute mapping (from Attributes List "ON RUNNING" tab):
  ┌─────────────────────┬──────────────────────────────────────────────┐
  │ Excel Column        │ Stibo Attribute                             │
  ├─────────────────────┼──────────────────────────────────────────────┤
  │ STYLE CODE          │ AT_PrincipalStyleCode                       │
  │ STYLE NAME          │ AT_PrincipalStyleDescription                │
  │ COLOUR              │ AT_PrincipalColorCode / AT_Color            │
  │ GENDER              │ AT_Gender / AT_BYGender / AT_PrincipalGender│
  │ SIZE                │ AT_PrincipalSize / AT_Size  (variant level) │
  │ EAN                 │ AT_PrincipalBarcode         (variant level) │
  │ RETAIL PRICE        │ AT_OriginalPrice / AT_CurrentPrice          │
  │ LAUNCH DATE         │ AT_LaunchingDate                            │
  │ PRODUCT             │ AT_SAPProductDivision                       │
  │                     │ AT_PrincipalMerchHierarchyL1                │
  │ CATEGORY            │ AT_SAPProductGroup                          │
  │                     │ AT_SAPProductCategory                       │
  │                     │ AT_PrincipalMerchHierarchyL2                │
  │                     │ AT_SportsCategoryEN                         │
  │ STYLE NAME          │ AT_Collection1 / AT_Franchise               │
  │ ACCOUNT             │ (metadata — "PLANET SPORT")                 │
  │ SKU                 │ (internal reference)                        │
  │ ALLOCATION MONTH    │ (logistics)                                 │
  │ WAREHOUSE           │ (logistics)                                 │
  │ REMARKS             │ (notes)                                     │
  └─────────────────────┴──────────────────────────────────────────────┘

Grouping:
  Rows are at the SKU level (one row per Style Code + Colour + Size).
  We GROUP by (STYLE CODE + COLOUR) → one Generic article.
  Each SIZE row within a group → one Variant sub-product.

  Article Category = Generic (1) — footwear with size variants.
"""

import re
import sys
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

# ── Country code → MDD COUNTRY name ──────────────────────────────
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
}

# ── ON Running / Planet Sport–specific lookup tables ─────────────

# Gender column → SAP Gender code
PS_GENDER_MAP: dict[str, str] = {
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

# Gender code → SAP Age code
PS_GENDER_TO_AGE: dict[str, str] = {
    "M":  "AD",   # Men → Adults
    "F":  "AD",   # Women → Adults
    "U":  "AD",   # Unisex → Adults
    "K":  "CH",   # Kids/Youth → Children
}

# PRODUCT column → SAP Product Division letter
PS_PRODUCT_TO_DIVISION: dict[str, str] = {
    "FOOTWEAR":    "F",
    "SHOES":       "F",
    "APPAREL":     "A",
    "ACCESSORIES": "E",
}

# CATEGORY column parsing → SAP Product Group code
# CATEGORY values like "Shoes Performance Training", "Shoes Performance Running"
# We extract the vertical part after "Shoes " or the full value.
PS_VERTICAL_TO_PROD_GROUP: dict[str, str] = {
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

# Vertical → SAP Product Category
PS_VERTICAL_TO_PROD_CATEGORY: dict[str, str] = {
    "PERFORMANCE ALL DAY":  "RU",
    "PERFORMANCE RUNNING":  "RU",
    "PERFORMANCE OUTDOOR":  "OD",
    "PERFORMANCE TRAINING": "TN",
    "CLASSIC":              "CS",
    "LIFESTYLE":            "CS",
}

# Country of Origin default for ON Running
PS_DEFAULT_COO = "VN"


def _extract_vertical_from_category(category: str) -> str:
    """
    Extract the vertical part from the CATEGORY column.
    e.g. "Shoes Performance Training" → "Performance Training"
         "Apparel Running" → "Running"
    """
    c = (category or "").strip()
    # Remove leading product type word (Shoes, Apparel, Accessories, etc.)
    for prefix in ("SHOES", "FOOTWEAR", "APPAREL", "ACCESSORIES"):
        if c.upper().startswith(prefix):
            c = c[len(prefix):].strip()
            break
    return c


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
        wb.close()
        self._build_country_reverse()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _build_country_reverse(self):
        country_lov = self.lovs.get("Country", {})
        self.lovs["CountryByCode"] = {code: name for name, code in country_lov.items()}
        log.info("[MDD] CountryByCode reverse map: %d entries", len(self.lovs["CountryByCode"]))

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
            
            # Size Code LOV uses columns F (LOV ID) and G (display name)
            if "SIZE" in sn.upper() and "CODE" in sn.upper():
                for row in rows[1:]:
                    if not row or len(row) < 7:
                        continue
                    code, name = row[5], row[6]  # F=index 5, G=index 6
                    if name:
                        self.lovs.setdefault(display, {})[str(name).strip()] = (
                            str(code).strip() if code else str(name).strip()
                        )
            else:
                # Other LOV sheets use columns A and B
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
    Maps (Country Name, Company Code, SBU, Brand Code) → Brand Type + Brand Category.
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
        if s.endswith(".0"):
            s = s[:-2]
        s2 = s.lstrip("0")
        return s2 if s2 else "0"

    def __init__(self, path: Path):
        self.path   = path
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
            country    = _cell(row, c_country)
            comp_code  = _cell(row, c_comp)
            sbu        = _cell(row, c_sbu)
            brand_code = _cell(row, c_bcode)
            btype      = _cell(row, c_btype)
            bcat       = _cell(row, c_bcat)
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
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu          or "").strip().upper(),
            (brand_code   or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": ""})

    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        """Match on country + brand only — ignores comp_code and sbu."""
        c = self._norm_country(country_name)
        b = (brand_code or "").strip().upper()
        s = (sbu or "").strip().upper()
        # pass 1: country + sbu + brand
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        # pass 2: country + brand only (first match)
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": ""}


class PlanetSportLoader:
    """
    Loads the Planet Sport order sheet for ON Running.

    Sheet: PLANET SPORT (or first sheet).
    Header row: auto-detected (row containing "STYLE CODE" or "EAN").

    Columns:
        ACCOUNT, LAUNCH DATE, ALLOCATION MONTH, EAN, SKU,
        CATEGORY, PRODUCT, STYLE CODE, GENDER, STYLE NAME,
        COLOUR, RETAIL PRICE, SIZE, QTY, TOTAL, WAREHOUSE, REMARKS

    Each row = one SKU (Style Code + Colour + Size).
    Rows are grouped by (STYLE CODE + COLOUR) to form one Generic article,
    with individual SIZE rows becoming Variant sub-products.
    """

    PREFERRED_SHEETS = ["PLANET SPORT", "PLANET SPORTS", "Linelist", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[PlanetSport-ONR] Loading: %s", self.path.name)
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

        log.info("[PlanetSport-ONR] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Auto-detect header row: first row containing "STYLE CODE" or "EAN"
        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and (
                "STYLE CODE" in v.upper() or "EAN" in v.upper()
            ) for v in r if v):
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
            log.error("[PlanetSport-ONR] Cannot find header row"); wb.close(); return

        log.info("[PlanetSport-ONR] Header at row %d (0-indexed)", hdr_idx)
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]

        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Drop rows where STYLE CODE is blank
        style_col = next(
            (c for c in df.columns if "STYLE" in c.upper() and "CODE" in c.upper()), None
        )
        if style_col:
            df = df[
                df[style_col].notna()
                & (~df[style_col].astype(str).str.strip().isin(["", "None", "nan", "TBC"]))
            ]
            if style_col != "STYLE CODE":
                df = df.rename(columns={style_col: "STYLE CODE"})

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info(
            "[PlanetSport-ONR] %d SKU rows loaded", len(self.df)
        )


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _fmt_date(v) -> str:
    """Format a date to dd-Mon-YYYY lowercase."""
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw:
        return raw
    if "carry" in raw.lower():
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _group_rows_by_generic(df: pd.DataFrame) -> dict[str, list[dict]]:
    """
    Group SKU-level rows by (STYLE CODE + COLOUR) to form Generic articles.
    Returns dict: group_key → list of row dicts (each row = one size/variant).
    """
    groups: dict[str, list[dict]] = defaultdict(list)
    for _, row in df.iterrows():
        style_code = _s(row.get("STYLE CODE"))
        colour     = _s(row.get("COLOUR"))
        if not style_code:
            continue
        key = f"{style_code}||{colour}"
        groups[key].append(dict(row))
    return dict(groups)


def map_generic_planet_sport(
    group_key: str,
    sku_rows: list[dict],
    brand_code: str = "ONR",
) -> dict:
    """
    Map a group of Planet Sport SKU rows (same Style Code + Colour)
    into one Generic article dict with embedded variant/size info.

    All rows in the group share the same Style Code, Colour, Gender,
    Style Name, Category, Product, Retail Price, Launch Date.
    Each row adds a SIZE + EAN for variant expansion.
    """
    # Take shared attributes from the first row (all rows share these)
    first = sku_rows[0]

    style_code   = _s(first.get("STYLE CODE"))
    style_name   = _s(first.get("STYLE NAME"))
    colour       = _s(first.get("COLOUR"))
    gender_raw   = _s(first.get("GENDER"))
    product      = _s(first.get("PRODUCT"))
    category     = _s(first.get("CATEGORY"))
    retail_price = _s(first.get("RETAIL PRICE"))
    launch_date  = first.get("LAUNCH DATE")
    account      = _s(first.get("ACCOUNT"))
    warehouse    = _s(first.get("WAREHOUSE"))
    remarks      = _s(first.get("REMARKS"))
    alloc_month  = _s(first.get("ALLOCATION MONTH"))

    # ── Collect variant-level data (SIZE + EAN per row) ──────────
    variants: list[dict] = []
    for row in sku_rows:
        size_val = _s(row.get("SIZE"))
        ean_val  = _s(row.get("EAN"))
        qty_val  = _s(row.get("QTY"))
        sku_val  = _s(row.get("SKU"))
        if size_val:
            variants.append({
                "size":     size_val,
                "ean":      ean_val,
                "qty":      qty_val,
                "sku":      sku_val,
            })

    # ── Derived fields ──────────────────────────────────────────

    # SAP Gender
    gender_code = PS_GENDER_MAP.get(gender_raw.upper(), "U")

    # SAP Age: Adults by default, Children if kids
    age_code = PS_GENDER_TO_AGE.get(gender_code, "AD")

    # Extract vertical from CATEGORY (e.g. "Shoes Performance Training" → "Performance Training")
    vertical = _extract_vertical_from_category(category)

    # SAP Product Division from PRODUCT column
    div_letter = PS_PRODUCT_TO_DIVISION.get(product.upper(), "F")

    # SAP Product Group from vertical
    prod_group = PS_VERTICAL_TO_PROD_GROUP.get(vertical.upper(), "PA")

    # SAP Product Category from vertical
    prod_category = PS_VERTICAL_TO_PROD_CATEGORY.get(vertical.upper(), "RU")

    # Country of Origin — default
    coo = PS_DEFAULT_COO

    # Launch Date formatting
    formatted_date = _fmt_date(launch_date)

    # RRP: extract numeric value
    rrp_str = ""
    if retail_price:
        m = re.search(r"[\d.]+", str(retail_price))
        rrp_str = m.group() if m else str(retail_price)

    # Generic code: ONR + Style Code
    generic_code = f"{brand_code}{style_code}"

    # Color token: first 3 chars of cleaned colour
    color_parts = colour.split("|") if "|" in colour else [colour]
    color_token = re.sub(r"[^A-Z0-9]", "", color_parts[0].upper().strip())[:3].ljust(3, "0")

    return {
        # Core identifiers
        "article_no":       style_code,        # STYLE CODE → AT_PrincipalStyleCode
        "style_code":       style_code,
        "model_name":       style_name,        # STYLE NAME → AT_PrincipalStyleDescription
        "brand_code":       brand_code,

        # Colour
        "colour":           colour,            # COLOUR → AT_PrincipalColorCode
        "colour_token":     color_token,

        # Gender / Age
        "gender_code":      gender_code,       # GENDER → AT_Gender
        "gender_raw":       gender_raw,        # GENDER → AT_PrincipalGender
        "age_code":         age_code,
        "age_raw":          "Adults" if age_code == "AD" else "Children",

        # Classification
        "product":          product,           # PRODUCT → AT_SAPProductDivision
        "category":         category,          # CATEGORY → AT_SAPProductGroup/Category
        "vertical":         vertical,          # Extracted vertical
        "division":         product,
        "div_letter":       div_letter,
        "prod_group":       prod_group,
        "prod_category":    prod_category,

        # Commercial
        "article_type":     "Inline",          # Default: Inline
        "art_category":     "1",               # Generic (footwear with sizes)
        "franchise":        style_name,        # STYLE NAME → AT_Franchise
        "launch_date":      formatted_date,    # LAUNCH DATE → AT_LaunchingDate

        # Pricing
        "rrp":              rrp_str,           # RETAIL PRICE → AT_OriginalPrice / AT_CurrentPrice
        "currency":         "IDR",             # Indonesia (Planet Sports)
        "fob":              "",                # Not in Planet Sport file

        # Origin
        "coo":              coo,

        # Derived codes
        "generic_code":     generic_code,

        # Hierarchy
        "merch_hierarchy_l1": product,         # PRODUCT → AT_PrincipalMerchHierarchyL1
        "merch_hierarchy_l2": vertical,        # CATEGORY vertical → AT_PrincipalMerchHierarchyL2

        # Sports Category (from CATEGORY column)
        "sports_category":  vertical,          # CATEGORY → AT_SportsCategoryEN

        # Metadata
        "account":          account,           # ACCOUNT
        "warehouse":        warehouse,         # WAREHOUSE
        "remarks":          remarks,           # REMARKS
        "alloc_month":      alloc_month,       # ALLOCATION MONTH

        # Variant data (SIZE + EAN per row)
        "variants":         variants,

        # No TDD / Backlog data
        "_tdd_sizes":       [],
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
    Resolve Brand LOV ID dynamically from MDD "Brand" LOV.
    """
    src_brand_name = (brand_name or "").strip()
    src_code = (brand_code or "").strip().upper()

    brand_label = src_brand_name or _lov(src_code, LOV_BRAND, src_code)[1]
    fallback_id = src_code or (re.sub(r"[^A-Z]", "", src_brand_name.upper())[:3])
    brand_lov_id = fallback_id

    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})

        def _norm(s: str) -> str:
            return re.sub(r"\s+", " ", (s or "").strip()).upper()

        target_names = [brand_label]
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


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
) -> None:
    """
    Write all Generic-level <Value> elements for a Planet Sport article.

    Mapping is derived from the Attributes List "ON RUNNING" tab,
    cross-referenced with MDD "Core Attributes" tab.

    Indicator legend (from Attributes List):
      A = Direct from Principal File (Planet Sport columns)
      B = Mapping from the principal data
      C = Defined from Images
      D = Manual Input
      E = Not applicable
      F = Formula in system
    """
    pass
    # _mw("AT_SBU",         sbu_code)
    # _mw("AT_CompanyCode", comp_code)
    # _w("AT_Brand",        id_val=art["brand_code"])
    # _w("AT_BrandGroup",   id_val="ON")
    # _w("AT_PrincipalStyleDescription", art["model_name"])
    # _w("AT_InboundGenericCode",        ...)
    # _w("AT_SAPAge",        id_val="AD")
    # _w("AT_BYAge",         id_val="ADULT")
    # _w("AT_Gender",        id_val=art["gender_code"])
    # _w("AT_BYGender",      id_val=art["gender_code"])
    # _w("AT_PrincipalGender", id_val=art["gender_code"])
    # _w("AT_Season",        id_val=sea_code)
    # _w("AT_SeasonYear",    season_year)
    # _w("AT_SAPArticleCategory", id_val="1")
    # _w("AT_BYArticleType",      id_val=at_code)
    # _w("AT_BYIndicator",   "Yes", id_val="Y")
    # _w("AT_SAPIndicator",  "No",  id_val="N")
    # _w("AT_UOM",           "Each", id_val="EA")
    # _w("AT_OriginalPrice", art["rrp"])
    # _w("AT_CurrentPrice",  art["rrp"])
    # _w("AT_MainVendorIdentification", "1")
    # _w("AT_Franchise",    (art["franchise"] or "").upper())
    # _w("AT_LaunchingDate", art["launch_date"])
    # _w("AT_Collection1",   art["model_name"])
    # _w("AT_Width",         "Wide" / "Regular")
    # _w("AT_Country",       id_val=country_val)
    # _w("AT_BrandType",     art["brand_type"])
    # _w("AT_BrandCategory", art["brand_category"])


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


def _build_variant_xml(
    parent_el:  ET.Element,
    art:        dict,
    variant:    dict,
    brand_code: str,
    mdd=None,
) -> None:
    """
    Build a <Product> sub-element for one Variant (size + EAN).

    Variant specifics:
      • UserTypeID = PRD_VariantArticle
      • KeyID      = KEY_Variant
      • Key value  = {BrandCode}{StyleCode}{ColorToken}{SizeCode}
      • Attributes: AT_Brand, AT_PrincipalStyleCode, AT_Size, AT_PrincipalSize, 
                    AT_Variant, AT_VariantDescription, AT_PrincipalBarcode
    """
    size_raw = variant["size"]
    
    ean_raw = variant["ean"]
    if ean_raw:
        try:
            # Handle float values like "7640006361076.0"
            ean = str(int(float(ean_raw)))
        except (ValueError, TypeError):
            ean = str(ean_raw).strip()
    else:
        ean = ""

    # Clean size: remove trailing .0
    size_str = str(size_raw).strip()
    if size_str.endswith(".0"):
        size_str = size_str[:-2]

    # Size code for key: pad to 3 chars (e.g. "8" → "080", "10" → "100")
    size_clean = re.sub(r"[^0-9]", "", size_str.replace(".", ""))
    size_code  = size_clean[:3].ljust(3, "0")

    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id = color_lov.get(art["colour"], "") \
                or color_lov.get(art["colour"].upper(), "")
    if colour_id and colour_id.isdigit():
        colour_id = colour_id.zfill(3)

    sap_color_for_key = colour_id if colour_id else art["colour_token"]
    # KEY_Variant format: {generic code (KEY_InboundArticle)}{ColorToken}{SizeCode}
    # Use the transformed generic code (at_generic_val) which has 1st and 3rd chars removed
    generic_code_for_variant = art.get("at_generic_val", art['generic_code'])
    key_variant = f"{generic_code_for_variant}{sap_color_for_key}{size_code}"

    variant_desc = f"Size {size_str}"

    v_el = ET.SubElement(parent_el, f"{{{STIBO_NS}}}Product")
    v_el.set("UserTypeID", "PRD_VariantArticle")

    kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundVariant")
    kv.text = key_variant

    ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")

    # ── DataContainers (DC_Barcode) — one per variant EAN ────────
    if ean:
        dc_root = ET.SubElement(v_el, f"{{{STIBO_NS}}}DataContainers")
        mdc     = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
        mdc.set("Type", "DC_Barcode")

        dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
        dc.set("Analyzer", "true")

        dc_vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

        barcode_val = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
        barcode_val.set("AttributeID", "AT_Barcode")
        barcode_val.text = ean

        barcode_type = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
        barcode_type.set("AttributeID", "AT_BarcodeType")
        barcode_type.set("ID", "P")

        main_ean_ind = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
        main_ean_ind.set("AttributeID", "AT_MainEANIndicator")
        main_ean_ind.set("ID", "Y")


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
    Build a <Product> XML fragment for one Planet Sport Generic article,
    including Variant sub-products for each size.

    Planet Sport specifics:
      • UserTypeID = PRD_GenericArticle (footwear — Article Category = Generic)
      • ParentID   = PPH_F-TempSubCat  (Footwear division)
      • KeyID      = KEY_InboundArticle
      • Key value  = {BrandCode}{StyleCode}  (AT_PrincipalStyleCode)
      • Variant sub-products for each SIZE row (with EAN barcodes)
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code      = art.get("brand_code", brand_code) or brand_code
    brand_token = brand.title().replace(" ", "")
    div_letter  = art.get("div_letter", "F")
    parent_id   = f"PPH_{div_letter}-TempSubCat"

    # KEY_Article: BrandCode + StyleCode (omitting 1st and 3rd characters)
    # Example: 3ME10051043 → M10051043, then ONR + M10051043 = ONRM10051043
    style_transformed = art['style_code'][1:2] + art['style_code'][3:] if len(art['style_code']) > 3 else art['style_code']
    key_article = f"{b_code}{style_transformed}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    # Set at_generic_val for use in _build_variant_xml
    art["at_generic_val"] = key_article

    # ── Generic-level Values ─────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    # _val(vals_el, "AT_CountryOrigin", id_val=art.get("coo", ""))

    # ── Variant sub-products (one per SIZE + EAN) ────────────────
    for variant in art.get("variants", []):
        _build_variant_xml(g_el, art, variant, b_code, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_generic(args_tuple):
    """Thread worker: map + validate one Generic group (Style Code + Colour)."""
    group_key, sku_rows, brand_code, mdd, season, country_code, article_type_from_filename, brand_type, brand_category = args_tuple
    mapped = map_generic_planet_sport(group_key, sku_rows, brand_code=brand_code)
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
        season      str   e.g. "SS26"
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Planet Sport file", ll_files)]:
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
    # Full fallback chain — mirrors crocs/linelist_main.py
    _country_code  = (getattr(args, "country_code", "") or "").strip().upper() or "ID"
    _country_name  = _COUNTRY_MAP.get(_country_code, _country_code)

    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        _country_name, args.comp_code, args.sbu, brand_lov_id,
    )

    # Try exact match first
    _rna_result = rna.get(_country_name, args.comp_code, args.sbu, brand_lov_id)
    _brand_type    = _rna_result.get("brand_type", "")
    _brand_category = _rna_result.get("brand_category", "")

    # Fallback to fuzzy match if exact match fails
    if not _brand_type and not _brand_category:
        log.info("[RNA] Exact match failed, trying fuzzy lookup...")
        _rna_result = rna.get_fuzzy(_country_name, args.sbu, brand_lov_id)
        _brand_type    = _rna_result.get("brand_type", "")
        _brand_category = _rna_result.get("brand_category", "")

    log.info(
        "[RNA] Result → brand_type=%r  brand_category=%r",
        _brand_type, _brand_category,
    )

    all_warnings: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing Planet Sport file: %s ───", ll_path.name)

        loader = PlanetSportLoader(ll_path)
        if loader.df.empty:
            log.warning("[PlanetSport-ONR] Empty dataframe — skipping."); continue

        # ── Group SKU rows by (STYLE CODE + COLOUR) → Generic articles ──
        groups = _group_rows_by_generic(loader.df)
        groups = dict(groups.items())
        # ── TEST MODE: limit to 5 articles ──────────────────────
        # groups = dict(list(groups.items())[:5])
        total_generics = len(groups)
        total_skus     = len(loader.df)
        log.info("Grouped %d SKU rows into %d Generic articles", total_skus, total_generics)

        # ── Parse filename metadata ──────────────────────────────
        _stem  = ll_path.stem
        _parts = re.split(r"\s*-\s*", _stem)

        if len(_parts) >= 7:
            file_type_from_input = "-".join(_parts[3:-3])
        else:
            file_type_from_input = "Planet Sport"

        _mm_match = re.search(r'\b(Multi|Mono)\b', _stem, re.IGNORECASE)
        multi_mono_from_input = _mm_match.group(1).capitalize() if _mm_match else "Multi"

        log.info(
            "Parsed from filename → file_type='%s'  multi_mono='%s'",
            file_type_from_input, multi_mono_from_input,
        )

        out_name = ll_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate (one task per Generic group) ──
        num_workers = min(8, max(1, total_generics))
        task_args   = [
            (gk, rows, brand_lov_id, mdd, args.season, _country_code,
             getattr(args, "article_type_from_filename", ""),
             _brand_type, _brand_category)
            for gk, rows in groups.items()
        ]
        ordered: list[tuple[int, dict, list]] = []

        log.info("Pass 1/2 — parallel map+validate (%d generics, %d workers) …",
                 total_generics, num_workers)

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_generic, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx           = futures[fut]
                mapped, warns = fut.result()
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered]
        del ordered

        log.info("Pass 1 done — mapped=%d Generic articles", len(mapped_articles))

        # ── Pass 2: stream XML to file ───────────────────────────
        log.info("Pass 2/2 — streaming XML to %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        written_count  = 0
        variant_count  = 0
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
                variant_count += len(art.get("variants", []))

            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        del mapped_articles

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%dKB)", out_path, file_kb)

        print("═══ ARTICLE SUMMARY (ON RUNNING — PLANET SPORT) ════", flush=True)
        print(f"  SKU rows (input)       : {total_skus}",      flush=True)
        print(f"  Generic articles       : {written_count}",   flush=True)
        print(f"  Variant articles       : {variant_count}",   flush=True)
        print(f"  XML file size          : {file_kb}KB",       flush=True)
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
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — ON Running Planet Sport v1.0")
    p.add_argument("--brand",      default="On Running")
    p.add_argument("--brand-code", default="ONR")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()