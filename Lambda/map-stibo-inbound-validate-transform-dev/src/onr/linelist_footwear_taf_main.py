"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — ON Running (TAF) v1.0           ║
║   Linesheet (The Athlete's Foot) → Stibo STEP XML               ║
╚══════════════════════════════════════════════════════════════════╝

Source linesheet format: "Linesheet Footwear - ON FW26 (TAF).xlsx"
Sheet name             : "TAF"

Key differences vs. onr.linelist_main (Planet Sports format):
  • Header row at row 8 (index 7), data from row 11
  • Size chart embedded in rows 1-9 (gender-specific US/EU pairs)
  • One row per (Style Code + Item Number/Color) combination
  • 18-digit variant ID = BrandCode(3) + StyleCode(7) + Color3D(3) + Size3D(3)
  • Adidas-style XML with KeyValue/KEY_Article pattern

Adapted from the local TAF ETL prototype (linesheet_loader.py).
"""

import logging
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import openpyxl


# ══════════════════════════════════════════════════════════════════════════════
# PATHS — driven by LAMBDA_TMP_DIR (same convention as onr.linelist_main)
# ══════════════════════════════════════════════════════════════════════════════
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

log = logging.getLogger("onr.linelist_footwear_taf_main")
if not log.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


# ══════════════════════════════════════════════════════════════════════════════
# RNA LOADER — Brand Type / Brand Category Lookup
# ══════════════════════════════════════════════════════════════════════════════
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

        # Debug: print all column headers
        log.info("[RNA] Column headers found: %s", list(col.keys()))

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

        missing = [nm for nm, idx in [
            ("Country", c_country), ("CompCode", c_comp), ("SBU", c_sbu),
            ("BrandCode", c_bcode), ("BrandType", c_btype), ("BrandCategory", c_bcat),
        ] if idx is None]
        if missing:
            log.warning("[RNA] Missing columns in '%s': %s", sheet_name, missing)

        def _cell(row, idx):
            if idx is None or idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()

        count = 0
        for row in rows[hdr_idx + 1:]:
            country = _cell(row, c_country)
            if not country:
                continue
            key = (
                self._norm_country(country),
                self._norm_comp_code(_cell(row, c_comp)),
                _cell(row, c_sbu).upper(),
                _cell(row, c_bcode).upper(),
            )
            self.lookup[key] = {
                "brand_type":     _cell(row, c_btype),
                "brand_category": _cell(row, c_bcat),
            }
            count += 1

        wb.close()
        log.info("[RNA] %d entries loaded from '%s'", count, sheet_name)
        
        # Debug: print first 3 entries
        if self.lookup:
            log.info("[RNA] Sample entries (first 3):")
            for i, (key, val) in enumerate(list(self.lookup.items())[:3]):
                log.info("[RNA]   %s → %s", key, val)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        """Return {brand_type, brand_category} for the given keys, or empty strings."""
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu        or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        result = self.lookup.get(key, {"brand_type": "", "brand_category": ""})
        log.info("[RNA] Lookup key=%s → result=%s", key, result)
        return result

    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        """Match on country + brand only — ignores comp_code."""
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


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Article:
    dealers_account: str
    selection:       str
    gender:          str        # M / W / U / K
    vertical:        str
    style_code:      str        # Generic level (e.g. 3MF1007)
    item_number:     str        # Variant level (e.g. 3MF10074782)
    style_name:      str
    color:           str
    retail_price:    Optional[float]
    launch_date:     Optional[str]
    row_num:         int
    country_size:    str 
    ecom_age_category: str 


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

# LOV-aligned IDs: AT_SAPGender → F/M/U ; AT_SAPAge → AD/CH/AA
GENDER_MAP_FALLBACK = {
    "M": {"sap_gender": "M", "sap_gender_desc": "Male",   "sap_age": "AD", "sap_age_desc": "Adults"},
    "W": {"sap_gender": "F", "sap_gender_desc": "Female", "sap_age": "AD", "sap_age_desc": "Adults"},
    "U": {"sap_gender": "U", "sap_gender_desc": "Unisex", "sap_age": "AA", "sap_age_desc": "All Ages"},
    "K": {"sap_gender": "U", "sap_gender_desc": "Unisex", "sap_age": "CH", "sap_age_desc": "Children"},
}


def gender_info(raw: str, mdd: dict) -> dict:
    """Resolve gender token via MDD Gender LOV; enrich age label via MDD Age LOV.

    Hardcoded GENDER_MAP_FALLBACK is only used when MDD has no Gender LOV entry
    for the incoming code — it is a last-resort fallback, not a primary source.
    """
    key = (raw or "").strip().upper()[:1] or "M"
    gender_lov = (mdd or {}).get("gender", {})
    age_lov    = (mdd or {}).get("age", {})
    fb         = GENDER_MAP_FALLBACK.get(key, GENDER_MAP_FALLBACK["M"])

    if key in gender_lov:
        entry  = gender_lov[key]
        sap_g  = entry.get("sap_gender")      or fb["sap_gender"]
        mdd_gd = (entry.get("sap_gender_desc") or "").strip()
        # Prefer a proper label over a single-char/empty MDD entry
        sap_gd = mdd_gd if len(mdd_gd) > 1 else fb["sap_gender_desc"]
        sap_a  = entry.get("sap_age")         or fb["sap_age"]
        sap_ad = entry.get("sap_age_desc")    or age_lov.get(sap_a) or fb["sap_age_desc"]
    else:
        sap_g, sap_gd = fb["sap_gender"], fb["sap_gender_desc"]
        sap_a         = fb["sap_age"]
        sap_ad        = age_lov.get(sap_a, fb["sap_age_desc"])

    return {
        "sap_gender":      sap_g,
        "sap_gender_desc": sap_gd,
        "sap_age":         sap_a,
        "sap_age_desc":    sap_ad,
    }

# EU Size → SAP Size Code fallback (used if MDD lookup misses)
EU_TO_SAP_FALLBACK = {
    "27.5": "27H", "28.5": "28H", "29": "29",   "29.5": "29H",
    "30":   "30",  "31":   "31",  "31.5": "H31", "32":   "32",
    "33":   "33",  "33.5": "33H", "34":   "34",  "35":   "35",
    "35.5": "35H", "36":   "36",  "36.5": "36H", "37":   "37",
    "37.5": "37H", "38":   "38",  "38.5": "38H", "39":   "39",
    "40":   "40",  "40.5": "40H", "41":   "41",  "42":   "42",
    "42.5": "42H", "43":   "43",  "44":   "44",  "44.5": "44H",
    "45":   "45",  "46":   "46",  "47":   "47N", "47.5": "47Z",
    "48":   "48",  "49":   "49",  "50":   "50",
}


AT = {
    "CompanyCode":          "AT_CompanyCode",
    "SBU":                  "AT_SBU",
    "Country":              "AT_Country",
    "CountryOrigin":        "AT_CountryOrigin",
    "BYIndicator":          "AT_BYIndicator",
    "SAPIndicator":         "AT_SAPIndicator",
    "EcomIndicator":        "AT_EcomIndicator",
    "SAPAge":               "AT_SAPAge",
    "GenericDesc":          "AT_InboundGenericCode",
    "VariantDesc":          "AT_InboundGenericCode",
    "Season":               "AT_Season",
    "Year":                 "AT_Year",
    "Gender":               "AT_Gender",
    "Brand":                "AT_Brand",
    "BrandCode":            "AT_BrandCode",

    "StyleCode":            "AT_PrincipalStyleCode",
    "StyleDesc":            "AT_PrincipalStyleDescription",
    "PrincipalGender":      "AT_PrincipalGender",
    "PrincipalAge":         "AT_PrincipalAge",
    "PrincipalSize":        "AT_PrincipalSize",
    "CombinationStyleColor":"AT_CombinationStyleColor",
    "SportsCategoryEN":     "AT_SportsCategoryEN",  
    "Collection1":          "AT_Collection1", 

    # "Generic":              "AT_Generic",
    "Variant":              "AT_Variant",
    "SAPStyleCode":         "AT_SAPStyleCode",

    "PrincipalColorCode":   "AT_PrincipalColorCode",
    "ColorCode":            "AT_ColorCode",
    "Color":                "AT_Color",
    "ColorDescription":     "AT_ColorDescription",
    "MAAColor":             "AT_MAAColor",

    "SizeCode":             "AT_SizeCode",
    "Size":                 "AT_Size",
    "SizeDescription":      "AT_SizeDescription",
    "MAASize":              "AT_MAASize",
    "EUSizeCode":           "AT_EUSizeCode",
    "USSizeCode":           "AT_USSizeCode",
    "SAPSizeCode":          "AT_SAPSize",

    "ArticleType":          "AT_ArticleType",
    "UOM":                  "AT_UOM",
    "MaterialType":         "AT_SAPArticleType",
    "HierarchyL1":          "AT_PrincipalMerchandiseHierarchyL1",
    "HierarchyL2":          "AT_PrincipalMerchandiseHierarchyL2",

    "RetailPrice":          "AT_OriginalPrice",
    "LaunchDate":           "AT_LaunchingDate",

    "InboundGeneric":       "AT_InboundGenericCode",
    "DealersAccount":       "AT_DealersAccount",
}


# ══════════════════════════════════════════════════════════════════════════════
# SMALL HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _parse_date(val) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%d-%b-%Y").lower()          
    if hasattr(val, "strftime"):
        return val.strftime("%d-%b-%Y").lower()
    s = str(val).strip()
    if not s:
        return None
    if "carry" in s.lower():                             # ← handle "Carry Over"
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return s


def color_code_3d(item_number: str) -> str:
    """Last 3 digits of item number = 3-digit color code (3MF10074782 → 782)."""
    digits = "".join(ch for ch in item_number if ch.isdigit())
    return digits[-3:].zfill(3) if digits else "000"


def size_code_3d(eu_size: str) -> str:
    """EU size → 3-digit numeric code (40 → 040, 40.5 → 405)."""
    digits = eu_size.replace(".", "").strip()
    digits = "".join(ch for ch in digits if ch.isdigit())
    return digits.zfill(3)[:3] if digits else "000"


def get_sap_size(eu_size: str, mdd: dict) -> str:
    return mdd["size_code"].get(eu_size) or EU_TO_SAP_FALLBACK.get(eu_size, eu_size)


def get_sizes_for_gender(gender: str, size_chart: dict, mdd: dict) -> list[dict]:
    us_eu = size_chart.get(gender.upper(), size_chart.get("M", {}))
    return [
        {"us": us, "eu": eu, "sap": get_sap_size(eu, mdd)}
        for us, eu in us_eu.items()
    ]


def build_generic_desc(brand_code: str, style_name: str, gender_desc: str, color: str) -> str:
    color_part = color.split("|")[0].strip() if "|" in color else color
    return f"{brand_code} {style_name} {gender_desc} {color_part}"[:40]


def build_variant_desc(brand_code: str, style_name: str, gender_desc: str, color: str, size: str) -> str:
    color_part = color.split("|")[0].strip() if "|" in color else color
    return f"{brand_code} {style_name} {gender_desc} {color_part} {size}"[:40]


def color_first_part(color: str) -> str:
    return color.split("|")[0].strip() if "|" in color else color


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER — minimal: only Brand / Size Code / Company / SBU LOVs needed
# ══════════════════════════════════════════════════════════════════════════════
def load_mdd(path: Path) -> dict:
    log.info("[MDD] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    def _ms(v):
        return "" if v is None else str(v).strip()

    def _kv_sheet(name, key_col=0, val_col=1):
        result = {}
        if name in wb.sheetnames:
            for row in wb[name].iter_rows(min_row=2, values_only=True):
                k, v = _ms(row[key_col]), _ms(row[val_col])
                if k and v:
                    result[k] = v
        return result

    brand       = _kv_sheet("Brand LOV")
    brand_group = _kv_sheet("BrandGroup LOV") if "BrandGroup LOV" in wb.sheetnames else {}
    company     = _kv_sheet("Company Code LOV")
    sbu         = _kv_sheet("SBU LOV")
    country     = _kv_sheet("Country LOV")

    # Fallback BrandGroup mapping: brand_code → LOV_BrandGroup ID (if not in MDD)
    if not brand_group:
        brand_group = {code: (code if code != "ONR" else "ON") for code in brand.keys()}

    # Country ID to Name map: {"ID" -> "Indonesia", ...}  (country LOV is already code → name)
    country_id_to_name: dict[str, str] = country

    # Retail Price Currency LOV: {country_name_upper -> currency_code}
    retail_price_currency_lov: dict[str, str] = {}
    if "Retail Price Currency LOV" in wb.sheetnames:
        for row in wb["Retail Price Currency LOV"].iter_rows(min_row=2, values_only=True):
            if row and row[0] and len(row) > 1 and row[1]:
                retail_price_currency_lov[_ms(row[0]).upper()] = _ms(row[1])
    log.info("[MDD] Retail Price Currency LOV: %d entries | sheets available: %s",
             len(retail_price_currency_lov),
             [s for s in wb.sheetnames if "RETAIL" in s.upper() or "CURRENCY" in s.upper() or "PRICE" in s.upper()])

    # Gender LOV: parse raw_key -> {sap_gender, sap_gender_desc, sap_age, sap_age_desc}
    # Expected columns (best-effort, any LOV sheet layout): code | label | sap_gender | sap_age | age_desc
    gender: dict[str, dict] = {}
    if "Gender LOV" in wb.sheetnames:
        for row in wb["Gender LOV"].iter_rows(min_row=2, values_only=True):
            cells = [_ms(c) for c in (row or ())]
            if not cells or not cells[0]:
                continue
            code   = cells[0].upper()[:1]
            label  = cells[1] if len(cells) > 1 else ""
            sap_g  = cells[2] if len(cells) > 2 else ""
            sap_a  = cells[3] if len(cells) > 3 else ""
            age_d  = cells[4] if len(cells) > 4 else ""
            gender[code] = {
                "sap_gender":      sap_g or code,
                "sap_gender_desc": label,
                "sap_age":         sap_a,
                "sap_age_desc":    age_d,
            }

    size_code: dict[str, str] = {}
    if "Size Code LOV" in wb.sheetnames:
        for row in wb["Size Code LOV"].iter_rows(min_row=2, values_only=True):
            sap  = _ms(row[0])
            desc = _ms(row[1])
            if sap and desc:
                size_code[desc] = sap

    # Season LOV: id (e.g. FW) ↔ label (Fall-Winter); accept either column order
    season: dict[str, str] = {}
    season_sheet = next(
        (s for s in wb.sheetnames if "SEASON" in s.upper() and "LOV" in s.upper()),
        None,
    )
    if season_sheet:
        for row in wb[season_sheet].iter_rows(min_row=2, values_only=True):
            cells = [_ms(c) for c in (row or ())]
            if len(cells) < 2 or not cells[0] or not cells[1]:
                continue
            a, b = cells[0], cells[1]
            if len(a) <= 3 < len(b):
                season[a.upper()] = b
            elif len(b) <= 3 < len(a):
                season[b.upper()] = a
            else:
                season[a.upper()] = b

    # Age LOV: id → label, e.g. {"AD": "Adults"}
    age: dict[str, str] = {}
    age_sheet = next(
        (s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()),
        None,
    )
    if age_sheet:
        for row in wb[age_sheet].iter_rows(min_row=2, values_only=True):
            cells = [_ms(c) for c in (row or ())]
            if not cells or not cells[0]:
                continue
            label  = cells[0]
            sap_id = cells[2] if len(cells) > 2 else (cells[1] if len(cells) > 1 else "")
            if label and sap_id:
                age[sap_id] = label

    # Sports Category LOV: Column A = display name, Column B = LOV ID
    sports_category: dict[str, str] = {}   # {display_name.upper(): lov_id}
    sc_sheet = next(
        (s for s in wb.sheetnames if "SPORTS CATEGORY" in s.upper() and "LOV" in s.upper()),
        None,
    )
    if sc_sheet:
        for row in wb[sc_sheet].iter_rows(min_row=2, values_only=True):
            name = _ms(row[0]) if row else ""
            code = _ms(row[1]) if row and len(row) > 1 else ""
            if name and code:
                sports_category[name.upper()] = code

    wb.close()
    log.info(
        "[MDD] Brand:%d | SizeCodes:%d | Company:%d | SBU:%d | Country:%d | Gender:%d | Season:%d | Age:%d | SportsCategory:%d",
        len(brand), len(size_code), len(company), len(sbu), len(country),
        len(gender), len(season), len(age), len(sports_category),
    )
    return {
        "brand":     brand,
        "brand_group": brand_group,
        "size_code": size_code,
        "company":   company,
        "sbu":       sbu,
        "country":   country,
        "country_id_to_name": country_id_to_name,
        "retail_price_currency_lov": retail_price_currency_lov,
        "gender":    gender,
        "season":    season,
        "age":       age,
        "sports_category": sports_category,
    }


# ══════════════════════════════════════════════════════════════════════════════
# LINESHEET LOADER (TAF format)
# ══════════════════════════════════════════════════════════════════════════════
def load_linesheet(path: Path, sheet_name: str = "TAF") -> tuple[list[Article], dict]:
    log.info("[Linesheet] Loading: %s (sheet=%s)", path.name, sheet_name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
        log.warning("[Linesheet] TAF sheet not found, using '%s'", sheet_name)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    size_chart = _parse_size_chart(rows)

    # Case-insensitive header dict — avoids silent misses due to casing/spaces
    hdr_raw = {str(v).strip(): i for i, v in enumerate(rows[7]) if v is not None}
    hdr     = {k.upper(): v for k, v in hdr_raw.items()}
    log.info("[Linesheet] Header columns detected: %s", list(hdr_raw.keys()))

    articles = []
    for i, row in enumerate(rows[10:], start=11):
        if all(v is None for v in row):
            continue

        def g(*cols):
            """Case-insensitive multi-name column lookup with fallbacks."""
            for col in cols:
                idx = hdr.get(col.strip().upper())
                if idx is not None and idx < len(row):
                    return row[idx]
            return None

        sc = _s(g("Style Code"))
        it = _s(g("Item Number"))
        if not sc or not it:
            continue

        price = None
        try:
            price = float(g("Retail Price")) if g("Retail Price") else None
        except (ValueError, TypeError):
            pass

        articles.append(Article(
            dealers_account   = _s(g("Dealers Account")),
            selection         = _s(g("Selection")),
            gender            = _s(g("Gender")),
            vertical          = _s(g("Sports Category EN", "Sports Category", "Vertical")),
            style_code        = sc,
            item_number       = it,
            style_name        = _s(g("Collection 1", "Collection", "Style Name", "Style")),
            color             = _s(g("Color")),
            retail_price      = price,
            launch_date       = _parse_date(g("Launch Date")),
            row_num           = i,
            country_size      = _s(g("Country Size")) or "US",
            ecom_age_category = _s(g("E-com Ages Category", "E-Com Ages Category",
                                     "Ecom Ages Category")) or "Ages 18+ years",
        ))

    log.info("[Linesheet] %d articles | size chart M:%d W:%d Kids:%d Youth:%d",
             len(articles),
             len(size_chart.get("M", {})), len(size_chart.get("W", {})),
             len(size_chart.get("Kids", {})), len(size_chart.get("Youth", {})))
    return articles, size_chart


def _parse_size_chart(rows: list) -> dict:
    def pairs(us_row, eu_row, start):
        out = {}
        for i in range(start, max(len(us_row), len(eu_row))):
            us = us_row[i] if i < len(us_row) else None
            eu = eu_row[i] if i < len(eu_row) else None
            if us and eu:
                us_s = str(us).strip()
                eu_s = str(eu).strip()
                if us_s.lower() not in ("us size", "eu size", ""):
                    out[us_s] = eu_s
        return out

    sc = {
        "M":     pairs(rows[1], rows[2], 16),
        "W":     pairs(rows[3], rows[4], 12),
        "Kids":  pairs(rows[5], rows[6], 12),
        "Youth": pairs(rows[7], rows[8], 12),
    }
    sc["U"] = sc["M"].copy()
    return sc


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILDER (Adidas-style: KeyValue / PPH_F-TempSubCat parent)
# ══════════════════════════════════════════════════════════════════════════════
def _val(parent, attr_id: str, text: str = "", id_: str = ""):
    el = ET.SubElement(parent, "Value")
    el.set("AttributeID", attr_id)
    if id_:
        el.set("ID", id_)
    if text:
        el.text = text
    return el


def _keyval(parent, key_id: str, value: str):
    el = ET.SubElement(parent, "KeyValue")
    el.set("KeyID", key_id)
    el.text = value
    return el


def _clf(parent, class_id: str, ref_type: str):
    el = ET.SubElement(parent, "ClassificationReference")
    el.set("ClassificationID", class_id)
    el.set("Type", ref_type)


def _indent(elem, level=0):
    pad = "\n" + "  " * level
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = pad + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = pad
        for child in elem:
            _indent(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = pad
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = pad


LOV_SEASON_LABEL_FALLBACK = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter",
}

# (Selection, Vertical) → (MerchL1, MerchL2)
# Maps the raw TAF linelist values to Stibo-valid Merchandise Hierarchy labels.
_ONR_SAP_MAP: dict[tuple[str, str], tuple[str, str]] = {
    ("Shoes", "Performance All Day"):         ("Footwear", "Sports Shoes"),
    ("Shoes", "Performance Outdoor"):         ("Footwear", "Sports Shoes"),
    ("Shoes", "Performance Running"):         ("Footwear", "Sports Shoes"),
    ("Shoes", "Performance Training"):        ("Footwear", "Sports Shoes"),
    ("Shoes", "Performance Tennis"):          ("Footwear", "Sports Shoes"),
    ("Shoes", "Shoes Performance Running"):   ("Footwear", "Sports Shoes"),
    ("Apparel", "Apparel Performance Running"):("Apparel",  "Sports Running Apparel"),
}
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


def resolve_sports_category_id(selection: str, vertical: str, mdd: dict) -> tuple[str, str]:
    """Translate Selection+Vertical → (label, LOV ID) via MDD Sports Category LOV."""
    label = (
        _SPORTS_CATEGORY_MAP.get((selection, vertical))
        or _SPORTS_CATEGORY_MAP.get((None, vertical))
    )
    if not label:
        return "", ""
    sc_lov = (mdd or {}).get("sports_category", {})
    sc_id  = sc_lov.get(label.upper(), "")
    # Pad single-digit IDs: "7" → "07"
    if sc_id and sc_id.isdigit() and len(sc_id) == 1:
        sc_id = sc_id.zfill(2)
    return label, sc_id


def _build_classifications(brand_name: str, brand_code: str,
                            sea_prefix: str, sea_year: str) -> ET.Element:
    """Mirror the working onr/linelist_main.py classification tree."""
    cls_root = ET.Element("Classifications")
    brand_token      = brand_name.title().replace(" ", "")
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent   = f"CLH_{brand_token}Batches"

    season_label_map = {"SS": "Spring Summer", "FW": "Fall Winter",
                        "AW": "Autumn Winter", "HO": "Holiday"}
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand_name} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

    season_cls = ET.SubElement(cls_root, "Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_parent)
    ET.SubElement(season_cls, "Name").text = season_display

    confirmed = ET.SubElement(season_cls, "Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, "Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, "Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, "Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


def _resolve_retail_price_currency(country_code: str, mdd: dict) -> str:
    """
    Resolve retail price currency from country code using MDD.
    
    Steps:
    1. Look up country_code (e.g., "ID") in country_id_to_name → get country name (e.g., "Indonesia")
    2. Look up country name in retail_price_currency_lov → get currency (e.g., "IDR")
    """
    if not country_code or not mdd:
        return ""
    
    code_upper = country_code.strip().upper()
    
    # Step 1: Get country name from country ID
    country_id_to_name = mdd.get("country_id_to_name", {})
    country_name = country_id_to_name.get(code_upper, "")
    
    if not country_name:
        log.warning("[Currency] country_code='%s' NOT found in country_id_to_name (%d entries)",
                    code_upper, len(country_id_to_name))
        return ""
    
    # Step 2: Look up currency from country name (partial match)
    retail_currency_lov = mdd.get("retail_price_currency_lov", {})
    country_name_upper = country_name.strip().upper()
    
    # Try exact match first
    if country_name_upper in retail_currency_lov:
        return retail_currency_lov[country_name_upper]
    
    # Try partial match (country name contains or is contained in LOV key)
    for lov_country, currency in retail_currency_lov.items():
        if country_name_upper in lov_country or lov_country in country_name_upper:
            return currency
    
    log.warning("[Currency] country_name='%s' NOT found in Retail Price Currency LOV (%d entries: %s)",
                country_name_upper, len(retail_currency_lov), list(retail_currency_lov.keys())[:5])
    return ""


def build_xml(
    articles:    list[Article],
    size_chart:  dict,
    mdd:         dict,
    output_path: Path,
    cfg:         dict,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code   = cfg["brand_code"]
    brand_name   = mdd["brand"].get(brand_code, cfg["brand_name"])
    season       = cfg["season"]
    comp_code    = cfg["comp_code"]
    sbu          = cfg["sbu"]
    country_code = cfg["country_code"]
    article_type = cfg["article_type"]
    brand_type     = cfg.get("brand_type", "")
    brand_category = cfg.get("brand_category", "")

    sbu_label  = mdd.get("sbu",     {}).get(sbu,       sbu)
    comp_label = mdd.get("company", {}).get(comp_code, comp_code)

    sea_prefix = season[:2].upper() if len(season) >= 2 else season
    sea_short  = season[2:] if len(season) > 2 else ""
    sea_year   = f"20{sea_short}" if len(sea_short) == 2 else sea_short
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id_full   = f"CLH_{brand_code}_{full_season_code}"
    sea_label        = (mdd.get("season", {}).get(sea_prefix)
                        or LOV_SEASON_LABEL_FALLBACK.get(sea_prefix, sea_prefix))

    brand_token = brand_name.title().replace(" ", "")

    at_norm  = (article_type or "Inline").strip().capitalize()
    at_label = "License" if at_norm.lower().startswith("lic") else "Inline"

    root = ET.Element("STEP-ProductInformation")
    root.set("xmlns:xsi",          STIBO_XSI)
    root.set("xsi:schemaLocation", STIBO_SCHEMA)
    root.set("WorkspaceID",        "Main")
    root.set("ContextID",          "Context1")
    root.append(ET.Comment(
        f" {brand_code} TAF ETL | Season:{season} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} "
    ))

    # ── Classifications (must precede Products) ──────────────────
    root.append(_build_classifications(brand_name, brand_code, sea_prefix, sea_year))

    products_el = ET.SubElement(root, "Products")

    generics: dict[str, list[Article]] = defaultdict(list)
    for art in articles:
        generics[art.item_number].append(art)

    log.info("[XML] Building — %d generics (Generic only, no variants)", len(generics))

    for _item_number, variants in generics.items():
        rep          = variants[0]
        ginfo        = gender_info(rep.gender, mdd)
        # AT_PrincipalStyleCode = Item Number (omitting 1st and 3rd characters)
        # Example: 3MF10074782 → M10074782
        item_number = rep.item_number
        principal_style_code = (item_number[1:2] + item_number[3:]) if len(item_number) > 3 else item_number
        # Generic code = BrandCode + PrincipalStyleCode (max 12 chars)
        # Example: ONR + M10074782 = ONRM10074782
        generic_code = f"{brand_code}{principal_style_code}"

        gen_el = ET.SubElement(products_el, "Product")
        gen_el.set("UserTypeID", "PRD_GenericArticle")
        gen_el.set("ParentID",   "PPH_F-TempSubCat")
        _keyval(gen_el, "KEY_InboundArticle", generic_code)
        ET.SubElement(gen_el, "Name").text = rep.style_name or generic_code

        _clf(gen_el, f"CLH_{brand_token}Articles", "CPL_Merchandiser")
        _clf(gen_el, f"{season_id_full}UA",        "CPL_UnConfirmedForSeason")

        gv = ET.SubElement(gen_el, "Values")
        _val(gv, AT["Brand"],        id_=brand_code) 
        brand_group_id = mdd.get("brand_group", {}).get(brand_code, brand_code)
        _val(gv, "AT_BrandGroup",    id_=brand_group_id)
        if brand_type:
            _val(gv, "AT_BrandType",     brand_type)
        if brand_category:
            _val(gv, "AT_BrandCategory", brand_category)
        _val(gv, AT["Season"],       "",         sea_prefix)   # ID only
        _val(gv, "AT_SeasonYear",    sea_year)
        _val(gv, AT["CompanyCode"],  comp_label, comp_code)
        _val(gv, AT["SBU"],          sbu_label,  sbu)
        if country_code:
            _val(gv, AT["Country"],  "",         country_code)  # ID only
            # AT_CountryOrigin: default is always local vendor in each country (per Attributes List row 54)
            _val(gv, AT["CountryOrigin"], "",        country_code)
        _val(gv, "AT_BYArticleType", at_label,   at_label)
        _val(gv, AT["UOM"],          "Each",     "EA")
        # HierarchyL1/L2: map via SAP map ("Shoes" → "Footwear", etc.)
        _sap     = _ONR_SAP_MAP.get((rep.selection, rep.vertical))
        merch_l1 = _sap[0] if _sap else rep.selection
        merch_l2 = _sap[1] if _sap else rep.vertical
        _val(gv, AT["HierarchyL1"], merch_l1)
        _val(gv, AT["HierarchyL2"], merch_l2)
        # AT_SportsCategoryEN: resolve via MDD LOV — send ID only
        _sc_label, _sc_id = resolve_sports_category_id(rep.selection, rep.vertical, mdd)
        if _sc_id:
            _val(gv, "AT_SportsCategoryEN", "", _sc_id)
        else:
            log.warning("[SportsCategoryEN] No MDD LOV match for vertical='%s' selection='%s'",
                        rep.vertical, rep.selection)
        if rep.style_name:
            _val(gv, AT["Collection1"], rep.style_name)   
        _val(gv, AT["StyleCode"], rep.item_number)
        _val(gv, AT["StyleDesc"], rep.style_name) 
        if rep.color:
            _val(gv, "AT_PrincipalColorName", rep.color)         # Color column
        _val(gv, AT["InboundGeneric"], generic_code)  # defining attr for KEY_InboundArticle — required for new products
        # AT_Gender / AT_BYGender / AT_SAPAge: send LOV ID only (no text label)
        _val(gv, AT["Gender"],            "", ginfo["sap_gender"])
        _val(gv, "AT_BYGender",           "", ginfo["sap_gender"])
        _val(gv, "AT_PrincipalGenderCode", (rep.gender or "").strip())  # Gender column from Excel
        _val(gv, AT["SAPAge"],            "", ginfo["sap_age"])
        _val(gv, "AT_BYAge",              "", "ADULT") 
        if rep.retail_price:
            _val(gv, AT["RetailPrice"], str(int(rep.retail_price)))
        # AT_RetailPriceCurrency / AT_FOBCurrency: derived from country_code via MDD
        _rpc = _resolve_retail_price_currency(country_code, mdd)
        if _rpc:
            _val(gv, "AT_RetailPriceCurrency", "", _rpc)
            _val(gv, "AT_FOBCurrency",         "", _rpc)
        if rep.launch_date:
            _val(gv, AT["LaunchDate"], rep.launch_date)
            
        _val(gv, "AT_MainVendorIdentification", "1") 
        _val(gv, "AT_Franchise", (rep.style_name or "").upper()) 
        _val(gv, AT["BYIndicator"],       "Yes", "Y")
        _val(gv, AT["SAPIndicator"],      "No",  "N")
        _val(gv, "AT_SAPArticleCategory", "",    "1")     # Generic
        _val(gv, "AT_MaterialType",       "",    "ZHAW")  # Trading Goods
        _val(gv, "AT_SAPProductFlag",     "",    "5")     # Direct Local
        _val(gv, "AT_CountrySize", id_=rep.country_size or "US")
        _val(gv, "AT_EComAgesCategory",   id_="18+Y")

    # Pretty print + write
    _indent(root)
    xml_str = ET.tostring(root, encoding="unicode", xml_declaration=False)
    xml_str = re.sub(r'\s+xmlns="[^"]*"', "", xml_str)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(xml_str)
        f.write("\n")

    log.info("[XML] Written → %s  (%.1f KB)",
             output_path, output_path.stat().st_size / 1024)
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# LAMBDA ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run(args, auditor=None):
    """
    Lambda dispatcher entry. Reads from LINELIST_DIR / MDD_DIR, writes XML to XML_OUT_DIR.

    args attributes used:
        brand_code   e.g. "ONR"
        brand        e.g. "On Running"
        comp_code    e.g. "0888"
        sbu          e.g. "SP"
        season       e.g. "FW26"
        country_code (optional) e.g. "ID"
        article_type_from_filename (optional) e.g. "Inline"
    """
    def first(d: Path, ext: str = "*.xlsx") -> Optional[Path]:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    if not mdd_f:
        log.error("No MDD file found in %s — aborting.", MDD_DIR)
        sys.exit(1)
    if not attr_f:
        log.warning("No Attributes List file found in %s — proceeding without RNA lookup.", ATTR_DIR)
    if not ll_files:
        log.error("No linelist file found in %s — aborting.", LINELIST_DIR)
        sys.exit(1)

    mdd = load_mdd(mdd_f)
    rna = RNALoader(attr_f) if attr_f else None

    brand_code   = getattr(args, "brand_code", "ONR") or "ONR"
    brand_name   = mdd["brand"].get(brand_code, getattr(args, "brand", "ON RUNNING") or "ON RUNNING")
    season       = getattr(args, "season", "FW26") or "FW26"
    comp_code    = getattr(args, "comp_code", "0888") or "0888"
    sbu          = getattr(args, "sbu", "SP") or "SP"
    country_code = getattr(args, "country_code", "") or "ID"
    article_type = getattr(args, "article_type_from_filename", "") or "Inline"

    sea_prefix     = season[:2].upper() if len(season) >= 2 else season
    sea_year_short = season[2:] if len(season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"{sea_prefix}{sea_year}"

    # ── Resolve AT_BrandType / AT_BrandCategory ──────────────────
    _brand_type     = ""
    _brand_category = ""
    if rna:
        _country_name = mdd.get("country", {}).get(country_code.upper(), country_code)
        log.info(
            "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
            _country_name, comp_code, sbu, brand_code,
        )
        _rna_result = rna.get(_country_name, comp_code, sbu, brand_code)
        if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
            log.warning("[RNA] Exact match failed — trying fuzzy (country+sbu+brand)")
            _rna_result = rna.get_fuzzy(_country_name, sbu, brand_code)
        _brand_type     = _rna_result.get("brand_type",     "")
        _brand_category = _rna_result.get("brand_category", "")
        log.info(
            "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
            _country_name, _brand_type, _brand_category,
        )

    cfg = {
        "brand_code":     brand_code,
        "brand_name":     brand_name,
        "season":         season,
        "season_id":      season_id,
        "comp_code":      comp_code,
        "sbu":            sbu,
        "country_code":   country_code,
        "article_type":   article_type,
        "brand_type":     _brand_type,
        "brand_category": _brand_category,
    }

    log.info(
        "[TAF] Pipeline args → brand=%s code=%s comp=%s sbu=%s season=%s country=%s type=%s",
        brand_name, brand_code, comp_code, sbu, season, country_code, article_type,
    )

    written: list[str] = []
    for ll_path in ll_files:
        log.info("─── Processing TAF linesheet: %s ───", ll_path.name)
        articles, size_chart = load_linesheet(ll_path)
        if not articles:
            log.warning("[TAF] No articles parsed — skipping %s", ll_path.name)
            continue

        # # ── TEST MODE: keep middle 5 articles (contiguous rows) ───────────
        # _n = len(articles)
        # _start = max(0, (_n - 5) // 2)
        # articles = articles[_start:_start + 5]

        out_name = ll_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name
        build_xml(articles, size_chart, mdd, out_path, cfg)
        written.append(str(out_path))

        print("═══ ARTICLE SUMMARY (ON RUNNING — TAF) ═════════════", flush=True)
        print(f"  Linesheet rows         : {len(articles)}",  flush=True)
        print(f"  Unique generics        : {len({a.item_number for a in articles})}", flush=True)
        print(f"  XML file size          : {out_path.stat().st_size // 1024}KB", flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    if auditor and written:
        auditor.set_xml_uploads(written)


# ── Local CLI ──────────────────────────────────────────────────────────────────
def main():
    import argparse
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — ON Running TAF v1.0")
    p.add_argument("--brand",        default="On Running")
    p.add_argument("--brand-code",   default="ONR")
    p.add_argument("--comp-code",    default="0888")
    p.add_argument("--sbu",          default="SP")
    p.add_argument("--season",       default="FW26")
    p.add_argument("--country-code", default="ID")
    p.add_argument("--article-type-from-filename", default="Inline")
    run(p.parse_args())


if __name__ == "__main__":
    main()
