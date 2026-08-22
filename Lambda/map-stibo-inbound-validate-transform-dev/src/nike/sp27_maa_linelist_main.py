"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Nike SP27 MAA  v1.0            ║
║   SP27_MAA_NIKE_Linelist.xlsx  →  Stibo STEP XML               ║
╚══════════════════════════════════════════════════════════════════╝

Input file : SP27_MAA_NIKE_Linelist.xlsx  (1 sheet: Sheet1)
Header row : Row 1 (0-based index 0) — single header row, no merge

Actual columns in this file:
  Image                  col A  (same as Style_Color — ignored)
  Unnamed: 1             col B  (empty — ignored)
  Season Year            col C  e.g. "SP2027"
  Style Name             col D  e.g. "Nike Initiator"
  Secondary Style Name   col E  e.g. "Men's Shoes"
  Style_Color            col F  e.g. "394055-012"  ← primary key
  Currency Code          col G  e.g. "IDR"
  Wholesale Price        col H  numeric float
  Retail Price           col I  numeric int
  Recommended            col J  "Y" / "N"
  Priority               col K  (often blank)
  Custom Assortment Name col L  (often blank)
  Channel/Tier           col M  e.g. "Added Product"
  door clusters          col N  (often blank)
  Color Code             col O  numeric int (e.g. 12 → "012")
  Color Description      col P  e.g. "DK SMOKE GREY/SMOKE GREY-BLACK"
  Product Lifecycle      col Q  "Active" / "Inactive"
  Delivery               col R  numeric int
  UOM                    col S  always blank — derived from Product Type
  Begin Futures Offer Date col T e.g. "2027-01-01"
  End Futures Offer Date col U  e.g. "2027-03-19"
  Product Type           col V  "Footwear" / "Apparel" / "Equipment"
  Body Type              col W  e.g. "Low Top" / "Top" / "Other" / "Bag"
  Silhouette             col X  e.g. "Not Applicable" / "Remastered"
  Category               col Y  e.g. "Sportswear" / "Running" / "Tennis"
  SubCategory            col Z  e.g. "Nsw Running" / "Fundamentals"
  Franchise              col AA e.g. "Initiator" / "Air Max Ivo"
  Collection             col AB e.g. "Vault" / "Sport Inspired"
  Gender                 col AC "Mens" / "Womens" / "Boys" / "Girls" / "Unisex"
  Age                    col AD "Adult" / "Youth" / "Grade School" / "Other"
  Carryover Status       col AE "New" / "Carryover"
  Key Marketing          col AF "Y" / "N"
  Replen                 col AG "Y" / "N"
  League                 col AH (often blank)
  Team                   col AI  (often blank)
  Sell-Ready Validation  col AJ (often blank)
  Distribution Segment   col AK e.g. "SG"
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

# ── Directory setup ──────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent.parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"
NAMING_DIR   = INPUT_DIR / "naming"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for _d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, NAMING_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"sp27maa_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter", "SP": "Spring",
    "SM": "Summer",
}
LOV_COUNTRY_ORIGIN = {
    "CN": "China", "VN": "Vietnam", "ID": "Indonesia",
    "KH": "Cambodia", "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia", "TH": "Thailand", "SG": "Singapore",
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
    "KH": "CAMBODIA",
}

# ── SP27 MAA Nike-specific lookup tables ─────────────────────────

SP27MAA_GENDER_MAP: dict[str, str] = {
    "MENS":   "M", "MEN":    "M", "MALE":   "M",
    "BOYS":   "M", "BOY":    "M",
    "WOMENS": "F", "WOMEN":  "F", "FEMALE": "F",
    "GIRLS":  "F", "GIRL":   "F",
    "UNISEX": "U",
}

SP27MAA_AGE_MAP: dict[str, str] = {
    "ADULT":        "AD",
    "ADULTS":       "AD",
    "YOUTH":        "CH",
    "GRADE SCHOOL": "CH",
    "KIDS":         "CH",
    "CHILDREN":     "CH",
    "INFANT":       "IN",
    "TODDLER":      "IN",
    "ALL AGES":     "AA",
    "JUNIOR":       "JR",
    "OTHER":        "AD",
}

SP27MAA_PRODUCT_TYPE_TO_DIVISION: dict[str, str] = {
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "EQUIPMENT":   "E",
    "ACCESSORIES": "E",
}

SP27MAA_PRODUCT_TYPE_TO_UOM: dict[str, str] = {
    "FOOTWEAR":    "PR",
    "APPAREL":     "EA",
    "EQUIPMENT":   "EA",
    "ACCESSORIES": "EA",
}


# ══════════════════════════════════════════════════════════════════
# STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace("", STIBO_NS)


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """Loads Core Attributes + LOVs from the shared MDD Excel."""

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
        self._load_age_lov(wb)
        self._load_retail_price_currency(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

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
        sheet_name = next(
            (s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None
        )
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return
        sap_age_map: dict[str, str] = {}
        by_age_map:  dict[str, str] = {}
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
        log.info(
            "[MDD] Age LOV loaded — SAPAge: %d entries, BYAge: %d entries",
            len(sap_age_map), len(by_age_map),
        )

    def _load_named_lov_sheets(self, wb):
        skip = {"Core Attributes", "Simple LOVs"}
        for sn in wb.sheetnames:
            if sn in skip:
                continue
            if "AGE" in sn.upper() and "LOV" in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            sub = self.lovs.setdefault(display, {})
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                v0 = str(row[0]).strip() if row[0] is not None else ""
                v1 = str(row[1]).strip() if row[1] is not None else ""
                if not v0 and not v1:
                    continue
                if v0 and not v1:
                    sub.setdefault(v0, v0)
                    continue
                if v1 and not v0:
                    sub.setdefault(v1, v1)
                    continue
                if len(v0) < len(v1):
                    code = v0
                elif len(v1) < len(v0):
                    code = v1
                elif v0.isupper() and not v1.isupper():
                    code = v0
                elif v1.isupper() and not v0.isupper():
                    code = v1
                else:
                    code = v0
                sub.setdefault(v0, code)
                sub.setdefault(v1, code)

    def _build_country_reverse(self):
        country_lov = self.lovs.get("Country", {})
        self.lovs["CountryByCode"] = {code: name for name, code in country_lov.items()}
        log.info(
            "[MDD] CountryByCode reverse map: %d entries",
            len(self.lovs["CountryByCode"]),
        )
    def _load_retail_price_currency(self, wb):
        """
        Load 'Retail Price Currency LOV' sheet directly.
        Sheet structure: Col A = currency display name (e.g. 'Indonesian Rupiah',
                         'Philippine Peso'), Col B = currency code (e.g. 'IDR', 'PHP').
        Then build a {ISO_country_code → currency_code} map using keyword matching
        against _COUNTRY_MAP country names. Handles both the plural country name
        (e.g. "PHILIPPINES") and its singular/adjective form (e.g. "PHILIPPINE"),
        since currency display names use the adjective form ("Philippine Peso"),
        not the plural country name ("Philippines").
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "RETAIL PRICE CURRENCY" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Retail Price Currency LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
    
        # Raw LOV: {currency_display_name_upper → currency_code}
        raw_lov: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2 or not row[0] or not row[1]:
                continue
            display_name = str(row[0]).strip().upper()  # e.g. "INDONESIAN RUPIAH"
            currency_code = str(row[1]).strip()          # e.g. "IDR"
            if display_name and currency_code:
                raw_lov[display_name] = currency_code
    
        # Build {ISO_country_code → currency_code} by matching country name
        # against currency display names (e.g. "INDONESIA" in "INDONESIAN RUPIAH").
        # Also try the singular/adjective form (strip trailing "S") to catch
        # cases like "PHILIPPINES" → "PHILIPPINE PESO".
        country_currency: dict[str, str] = {}
        for iso_code, country_name in _COUNTRY_MAP.items():
            keyword = country_name.upper()  # e.g. "PHILIPPINES"
            keyword_singular = keyword[:-1] if keyword.endswith("S") else keyword  # "PHILIPPINE"
    
            matched = next(
                (code for display, code in raw_lov.items() if keyword in display),
                "",
            )
            matched_via = "plural" if matched else ""
    
            if not matched and keyword_singular != keyword:
                matched = next(
                    (code for display, code in raw_lov.items() if keyword_singular in display),
                    "",
                )
                matched_via = "singular" if matched else ""
    
            if matched:
                country_currency[iso_code] = matched
                log.info(
                    "[MDD] CountryCurrency: %s → %s (via '%s' form, keyword='%s')",
                    iso_code, matched, matched_via,
                    keyword if matched_via == "plural" else keyword_singular,
                )
            else:
                log.warning(
                    "[MDD] CountryCurrency: no match found for %s (tried '%s' and '%s')",
                    iso_code, keyword, keyword_singular,
                )
    
        self.lovs["CountryCurrency"] = country_currency
        log.info("[MDD] CountryCurrency map: %d entries → %s", len(country_currency), country_currency)

class AttributesListLoader:
    """Loads the NIKE tab from the shared Attributes List workbook."""

    def __init__(self, path: Path, brand: str = "NIKE"):
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
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"), None
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
        # strip .0 suffix (e.g. "888.0" → "888")
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

        print(f"\n[RNA DEBUG] All sheets in {self.path.name}:", flush=True)
        for i, sn in enumerate(wb.sheetnames):
            print(f"  {i+1}. '{sn}'", flush=True)
        print(f"[RNA DEBUG] Looking for keywords: {self.RNA_SHEET_KEYWORDS}\n", flush=True)

        sheet_name = next(
            (s for s in wb.sheetnames
             if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found in %s", self.path.name)
            wb.close()
            return

        print(f"[RNA DEBUG] Found sheet: '{sheet_name}'", flush=True)
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
        c_bcat    = _find("BRANDCATEGORY", "BRAND CATEGORY", "AT_BRANDCATEGORY")

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

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu        or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": ""})

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


class NamingConventionLoader:
    """Loads naming convention rules for output filename building."""

    def __init__(self, path: Path):
        self.path  = path
        self.rules: dict[str, str] = {}
        self._load()

    def _load(self):
        log.info("[Naming] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        for sheet in ["Approved", "Sheet1"]:
            if sheet not in wb.sheetnames:
                continue
            for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
                if not row or not row[0]:
                    continue
                brand   = str(row[0]).strip().upper()
                pattern = row[2] if len(row) > 2 else None
                if pattern and "[" in str(pattern):
                    self.rules[brand] = str(pattern).strip()
            break
        wb.close()
        log.info("[Naming] %d brand rules", len(self.rules))

    def build_filename(self, brand, comp_code, sbu, file_type, multi_mono, season, seq):
        return f"{comp_code}-{sbu}-{brand}-{file_type}-{multi_mono}-{season}-{seq}.xml"


class SP27MAALinelistLoader:
    """Loads the Nike SP27 MAA linelist workbook."""

    PREFERRED_SHEETS = ["Sheet1", "Linelist", "MAA", "MAPI"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[SP27MAA-Linelist] Loading: %s", self.path.name)
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

        log.info("[SP27MAA-Linelist] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and "STYLE_COLOR" in v.upper() for v in r if v):
                hdr_idx = i
                break

        if hdr_idx is None:
            hdr_idx = next(
                (
                    i for i, r in enumerate(rows)
                    if sum(1 for v in r[:15] if isinstance(v, str) and v.strip()) >= 5
                ),
                None,
            )

        if hdr_idx is None:
            log.error("[SP27MAA-Linelist] Cannot find header row")
            wb.close()
            return

        log.info("[SP27MAA-Linelist] Header at row index %d", hdr_idx)

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        style_col = next(
            (c for c in df.columns if c.strip().upper() == "STYLE_COLOR"), None
        )
        if style_col:
            df = df[
                df[style_col].notna()
                & (~df[style_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]
        else:
            log.warning("[SP27MAA-Linelist] 'Style_Color' column not found")

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info("[SP27MAA-Linelist] %d valid rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — HELPERS
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw:
        return raw
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _pad_color(raw) -> str:
    if raw is None:
        return "000"
    try:
        return str(int(float(str(raw)))).zfill(3)
    except (ValueError, TypeError):
        s = str(raw).strip().zfill(3)
        return s[:3] if len(s) >= 3 else s


def _price_str(v) -> str:
    if v is None:
        return ""
    try:
        return str(int(float(str(v))))
    except (ValueError, TypeError):
        m = re.search(r"[\d.]+", str(v))
        return m.group() if m else str(v)


def _fob_str(v) -> str:
    if v is None:
        return ""
    try:
        f = round(float(v), 2)
        return f"{f:.2f}".rstrip("0").rstrip(".")
    except (ValueError, TypeError):
        m = re.search(r"[\d.]+", str(v))
        return m.group() if m else str(v)
    
def _resolve_retail_price_currency(country_code: str, mdd: "MDDLoader | None") -> str:
    """country_code (e.g. 'ID') → Retail Price Currency LOV ID (e.g. 'IDR')."""
    if not country_code or not mdd:
        return ""
    return mdd.lovs.get("CountryCurrency", {}).get(country_code.strip().upper(), "")

def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def resolve_brand_lov_id(mdd: MDDLoader | None, brand_code: str) -> tuple[str, str]:
    src_code    = (brand_code or "").strip().upper()
    _, brand_label = _lov(src_code, LOV_BRAND, src_code)
    brand_lov_id   = src_code
    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})
        for display_name, lov_id in brand_lov.items():
            if (display_name or "").strip().upper() == (brand_label or "").strip().upper():
                brand_lov_id = str(lov_id).strip()
                break
    return brand_lov_id, brand_label


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_article_sp27maa(row: dict, brand_code: str = "NIK") -> dict:
    style_color       = _s(row.get("Style_Color"))
    style_name        = _s(row.get("Style Name"))
    sec_name          = _s(row.get("Secondary Style Name"))
    color_code_raw    = row.get("Color Code")
    color_desc        = _s(row.get("Color Description"))
    gender_raw        = _s(row.get("Gender"))
    age_raw           = _s(row.get("Age"))
    season_year       = _s(row.get("Season Year"))
    product_type      = _s(row.get("Product Type"))
    body_type         = _s(row.get("Body Type"))
    silhouette        = _s(row.get("Silhouette"))
    category          = _s(row.get("Category"))
    sub_category      = _s(row.get("SubCategory"))
    franchise         = _s(row.get("Franchise"))
    collection        = _s(row.get("Collection"))
    retail_price      = row.get("Retail Price")
    wholesale         = row.get("Wholesale Price")
    currency          = _s(row.get("Currency Code"))
    begin_date        = row.get("Begin Futures Offer Date")
    delivery          = _s(row.get("Delivery"))
    carryover         = _s(row.get("Carryover Status"))
    channel_tier      = _s(row.get("Channel/Tier"))
    dist_segment      = _s(row.get("Distribution Segment"))
    product_lifecycle = _s(row.get("Product Lifecycle"))
    key_marketing     = _s(row.get("Key Marketing"))

    sap_style_code  = style_color.rsplit("-", 1)[0] if "-" in style_color else style_color
    sap_style_clean = re.sub(r"[^A-Z0-9]", "", sap_style_code.upper())[:9]
    colour_token    = _pad_color(color_code_raw)
    gender_code     = SP27MAA_GENDER_MAP.get(gender_raw.upper().strip(), "U")
    age_code        = SP27MAA_AGE_MAP.get(age_raw.upper().strip(), "AD")
    div_letter      = SP27MAA_PRODUCT_TYPE_TO_DIVISION.get(product_type.upper(), "A")
    uom_code        = SP27MAA_PRODUCT_TYPE_TO_UOM.get(product_type.upper(), "EA")
    coo             = "SG"
    launch_date     = _fmt_date(begin_date)
    generic_code    = f"{brand_code}{sap_style_clean}"
    rrp_str         = _price_str(retail_price)
    fob_str         = _fob_str(wholesale)

    return {
        "article_no":         style_color,
        "sap_style_code":     sap_style_clean,
        "model_name":         sec_name or style_name,
        "style_name":         style_name,
        "brand_code":         brand_code,
        "colour":             color_desc,
        "colour_code":        colour_token,
        "colour_token":       colour_token,
        "gender_code":        gender_code,
        "gender_raw":         gender_raw,
        "age_code":           age_code,
        "age_raw":            age_raw,
        "product_type":       product_type,
        "category":           category,
        "sub_category":       sub_category,
        "silhouette":         silhouette,
        "franchise":          franchise,
        "div_letter":         div_letter,
        "body_type":          body_type,
        "collection":         collection,
        "article_type":       "Inline",
        "art_category":       "1",
        "carryover":          carryover,
        "channel_tier":       channel_tier,
        "delivery":           delivery,
        "dist_segment":       dist_segment,
        "product_lifecycle":  product_lifecycle,
        "key_marketing":      key_marketing,
        "rrp":                rrp_str,
        "currency":           currency or "IDR",
        "fob":                fob_str,
        "fob_currency":       currency or "IDR",
        "coo":                coo,
        "uom_code":           uom_code,
        "launch_date":        launch_date,
        "generic_code":       generic_code,
        "season_year_raw":    season_year,
        "_tdd_sizes":         [],
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    warns = []
    art   = mapped["article_no"]
    field_checks = {
        "article_no":  "AT_PrincipalStyleCode",
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
    value:   str = "",
    id_val:  str = "",
    derived: bool = False,
) -> ET.Element | None:
    if derived:
        return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id = str(id_val).strip() if id_val else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    clean_val = str(value).strip() if value else ""
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd:        MDDLoader | None = None,
) -> None:
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
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    brand_lov_id = ""
    if mdd:
        bl = mdd.lovs.get("Brand", {})
        for cand in (art["brand_code"], art["brand_code"].upper(), brand_name, brand_name.upper()):
            if cand and cand in bl:
                brand_lov_id = str(bl[cand]).strip()
                break
    if not brand_lov_id:
        brand_lov_id = art["brand_code"]
    _w("AT_Brand", id_val=brand_lov_id)

    # ── Brand Group ──────────────────────────────────────────────
    _w("AT_BrandGroup", id_val=brand_name.upper())

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["article_no"])           # ← col Style_Color (raw)
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["colour_code"])          # ← col Color Code (zero-padded)
    _w("AT_PrincipalColorName",        art["colour"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])

    # ── Colour LOV lookup ─────────────────────────────────────────
    colour_id = ""
    if mdd:
        color_lov = (
            mdd.lovs.get("Color Code", {})
            or mdd.lovs.get("Color", {})
            or mdd.lovs.get("AT_Color", {})
        )
        for cand in (
            art["colour_code"], art["colour"], art["colour"].upper(),
            (art["colour"].split("/")[0].strip().upper() if art["colour"] else ""),
            (art["colour"].split("-")[-1].strip().upper() if art["colour"] else ""),
        ):
            if cand and cand in color_lov:
                colour_id = str(color_lov[cand]).strip()
                break
    if colour_id and colour_id.isdigit():
        colour_id = colour_id.zfill(3)
    if colour_id:
        _w("AT_Color", id_val=colour_id)

    sap_color_code        = art["colour_code"]
    at_generic_val        = f"{brand_lov_id}{art['sap_style_code']}{sap_color_code}"
    art["at_generic_val"] = at_generic_val
    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    g_code  = art["gender_code"]
    g_label = LOV_GENDER.get(g_code, g_code)
    _w("AT_Gender",                     g_label, id_val=g_code)
    _w("AT_BYGender",                   g_label, id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age ──────────────────────────────────────────────────────
    age_raw_upper = (art["age_raw"] or "").strip().upper()
    sap_age_map   = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map    = (
        (mdd.lovs.get("BYAge",    {}) if mdd else {})
        or (mdd.lovs.get("BY Age", {}) if mdd else {})
        or (mdd.lovs.get("AT_BYAge", {}) if mdd else {})
    )

    matched_age_display = next(
        (k for k in sap_age_map if k.upper() == age_raw_upper), None
    )
    sap_age_code  = sap_age_map.get(matched_age_display, art["age_code"]) if matched_age_display else art["age_code"]
    sap_age_label = matched_age_display or LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge", sap_age_label, id_val=sap_age_code)

    by_age_val = ""
    for cand in (art["age_raw"], age_raw_upper, (art["age_raw"] or "").title()):
        if cand and cand in by_age_map:
            by_age_val = str(by_age_map[cand]).strip()
            break
    if by_age_val:
        _w("AT_BYAge", id_val=by_age_val)
    _w("AT_PrincipalAgeDescription", art["age_raw"])                # ← col Age (raw, no fallback)

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    # ── Country of Origin ────────────────────────────────────────
    coo_code  = art["coo"]
    coo_label = LOV_COUNTRY_ORIGIN.get(coo_code, coo_code)
    _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    # ── Article category / type / indicators ─────────────────────
    _w("AT_SAPArticleCategory", id_val="1")

    at_code          = art.get("article_type", "Inline")
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    _w("AT_BYIndicator",  "Yes", id_val="Y")
    _w("AT_SAPIndicator", "No",  id_val="N")

    # ── UOM ──────────────────────────────────────────────────────
    _w("AT_UOM", id_val="EA")   #FIX -> Default to EA

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice", art["rrp"])
    _w("AT_CurrentPrice",  art["rrp"])
    cur_raw = (art.get("currency") or "").strip()
    cur_id  = ""
    if mdd and cur_raw:
        cur_lov = (
            mdd.lovs.get("Currency", {})
            or mdd.lovs.get("RetailPriceCurrency", {})
            or mdd.lovs.get("FOBCurrency", {})
        )
        for cand in (cur_raw, cur_raw.upper()):
            if cand and cand in cur_lov:
                cur_id = str(cur_lov[cand]).strip()
                break
    # ← FIX 3: AT_FOB removed — do not send
    if cur_id:
        _w("AT_FOBCurrency", id_val=cur_id)
        
    rpc_id = _resolve_retail_price_currency(art.get("country_code", ""), mdd)
    if rpc_id:
        _w("AT_RetailPriceCurrency", id_val=rpc_id)
        
    # ── Width ─────────────────────────────────────────────────────
    # Footwear only → default "Regular" | Apparel/Equipment → do NOT send
    if (art.get("product_type") or "").strip().upper() == "FOOTWEAR":
        wi_id = ""
        if mdd:
            wi_lov = mdd.lovs.get("Width", {})
            for cand in ("Regular", "REGULAR", "regular"):
                if cand in wi_lov:
                    wi_id = str(wi_lov[cand]).strip()
                    break
        if wi_id:
            _w("AT_Width", id_val=wi_id)
        else:
            _w("AT_Width", "Regular")

    # ── Nature of Article ────────────────────────────────────────
    noa_id = ""
    if mdd:
        noa_lov = (
            mdd.lovs.get("Nature of Article", {})
            or mdd.lovs.get("NatureOfArticle", {})
        )
        for cand in ("Regular", "REGULAR", "REG"):
            if cand in noa_lov:
                noa_id = str(noa_lov[cand]).strip()
                break
    if noa_id:
        _w("AT_NatureOfArticle", id_val=noa_id)

    # ── SAP Product Flag ─────────────────────────────────────────
    spf_id = ""
    if mdd:
        spf_lov = (
            mdd.lovs.get("SAP Product Flag", {})
            or mdd.lovs.get("SAPProductFlag", {})
        )
        for cand in ("Direct - Local", "DIRECT - LOCAL", "Local Production", "5"):
            if cand in spf_lov:
                spf_id = str(spf_lov[cand]).strip()
                break
    if spf_id:
        _w("AT_SAPProductFlag", id_val=spf_id)

    # ── Material Type ─────────────────────────────────────────────
    _w("AT_MaterialType", "direct", id_val="ZHAW")                      

    # ── Collections ──────────────────────────────────────────────
    _w("AT_Collection1", art.get("franchise", ""))
    # AT_Collection2 not sent per spec

    # ── Franchise ────────────────────────────────────────────────
    fr_raw = (art.get("franchise") or "").strip()
    fr_id  = ""
    if mdd and fr_raw:
        fr_lov = mdd.lovs.get("Franchise", {})
        for cand in (fr_raw, fr_raw.upper(), fr_raw.title()):
            if cand and cand in fr_lov:
                fr_id = str(fr_lov[cand]).strip()
                break
    if fr_id:
        _w("AT_Franchise", id_val=fr_id)

    # ── Silhouette ───────────────────────────────────────────────
    si_raw = (art.get("silhouette") or "").strip()
    si_id  = ""
    if mdd and si_raw:
        si_lov = mdd.lovs.get("Silhouette", {})
        for cand in (si_raw, si_raw.upper(), si_raw.title()):
            if cand and cand in si_lov:
                si_id = str(si_lov[cand]).strip()
                break
    if si_id:
        _w("AT_Silhouette", id_val=si_id)

    # ── Sports Category EN ───────────────────────────────────────
    sc_raw = (art.get("category") or "").strip()
    sc_id  = ""
    if mdd and sc_raw:
        sc_lov = mdd.lovs.get("Sports Category", {}) or mdd.lovs.get("SportsCategory", {})
        for cand in (sc_raw, sc_raw.upper(), sc_raw.title()):
            if cand and cand in sc_lov:
                sc_id = str(sc_lov[cand]).strip()
                break
    if sc_id:
        _w("AT_SportsCategoryEN", id_val=sc_id)

    # ── Interest = SubCategory ───────────────────────────────────
    in_raw = (art.get("sub_category") or "").strip()
    in_id  = ""
    if mdd and in_raw:
        in_lov = mdd.lovs.get("Interest", {})
        for cand in (in_raw, in_raw.upper(), in_raw.title()):
            if cand and cand in in_lov:
                in_id = str(in_lov[cand]).strip()
                break
    if in_id:
        _w("AT_Interest", id_val=in_id)

    # AT_EComProductNameEN not sent per spec


    # ── Principal Merchandise Hierarchy L1–L5 ────────────────────
    _w("AT_PrincipalMerchandiseHierarchyL1", art.get("product_type",  ""))
    _w("AT_PrincipalMerchandiseHierarchyL2", art.get("category",      ""))
    _w("AT_PrincipalMerchandiseHierarchyL3", art.get("sub_category",  ""))
    _w("AT_PrincipalMerchandiseHierarchyL4", art.get("silhouette",    ""))
    _w("AT_PrincipalMerchandiseHierarchyL5", art.get("franchise",     ""))

    # ── Country Size ─────────────────────────────────────────────
    # Footwear / Equipment → "US"  |  Apparel → do NOT send
    if (art.get("product_type") or "").strip().upper() != "APPAREL":
        _w("AT_CountrySize", "US")

    # ── Article Status ───────────────────────────────────────────
    as_id = ""
    if mdd:
        as_lov = mdd.lovs.get("Article Status", {}) or mdd.lovs.get("ArticleStatus", {})
        for cand in ("Active", "ACTIVE", "active"):
            if cand in as_lov:
                as_id = str(as_lov[cand]).strip()
                break
    if as_id:
        _w("AT_ArticleStatus", id_val=as_id)

    # ── Brand Type / Brand Category ───────────────────────────────
    # Populated in run() via full RNA fallback chain, passed through
    # _process_article into art dict
    _w("AT_BrandType",     art.get("brand_type",     ""))
    _w("AT_BrandCategory", art.get("brand_category", ""))


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent   = f"CLH_{brand}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
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
    mdd:        MDDLoader | None = None,
) -> str:
    article_no = art["article_no"]
    if not article_no:
        return ""

    div_letter        = art.get("div_letter", "A")
    parent_id         = f"PPH_{div_letter}-TempSubCat"
    sap_color_for_key = art["colour_code"]
    key_article       = f"{art['brand_code']}{art['sap_style_code']}{sap_color_for_key}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["model_name"] or article_no

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one SP27 MAA Nike row."""
    (row, brand_code, mdd, season, country_code,
     article_type_from_filename, brand_type, brand_category) = row_tuple
    mapped = map_article_sp27maa(row, brand_code=brand_code)
    mapped["season"]                     = season
    mapped["country_code"]               = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    mapped["brand_type"]                 = brand_type
    mapped["brand_category"]             = brand_category
    warns = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    naming_f = first(NAMING_DIR)
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Linelist", ll_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="NIKE")
    rna    = RNALoader(attr_f)
    naming = NamingConventionLoader(naming_f) if naming_f else None

    # ── RNA debug ────────────────────────────────────────────────
    print("\n" + "=" * 70, flush=True)
    print("RNA LOADER DEBUG:", flush=True)
    print(f"  Total RNA entries loaded: {len(rna.lookup)}", flush=True)
    for i, (key, val) in enumerate(list(rna.lookup.items())[:10]):
        print(f"    {i+1}. Key={key} → Value={val}", flush=True)
    nik_entries = {k: v for k, v in rna.lookup.items() if k[3] == "NIK"}
    print(f"  All NIK entries ({len(nik_entries)}):", flush=True)
    for key, val in list(nik_entries.items()):
        print(f"    {key} → {val}", flush=True)
    print("=" * 70 + "\n", flush=True)

    log.info(
        "SP27 MAA pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    brand_lov_id, brand_lov_label = resolve_brand_lov_id(mdd, args.brand_code)
    log.info(
        "Resolved brand LOV → source=%s  label=%s  lov_id=%s",
        args.brand_code, brand_lov_label, brand_lov_id,
    )

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    # ── Resolve AT_BrandType / AT_BrandCategory ──────────────────
    # Full fallback chain — mirrors crocs/linelist_main.py
    _country_code = (getattr(args, "country_code", "") or "").strip().upper() or "ID"
    _country_name = _COUNTRY_MAP.get(_country_code, _country_code)

    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        _country_name, args.comp_code, args.sbu, brand_lov_id,
    )

    _rna_result = rna.get(_country_name, args.comp_code, args.sbu, brand_lov_id)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, args.comp_code, args.sbu, args.brand_code)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, str(int(args.comp_code)), args.sbu, brand_lov_id)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, str(int(args.comp_code)), args.sbu, args.brand_code)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        log.warning("[RNA] Exact match failed — trying fuzzy (country+sbu+brand)")
        _rna_result = rna.get_fuzzy(_country_name, args.sbu, brand_lov_id)

    _brand_type     = _rna_result.get("brand_type",     "")
    _brand_category = _rna_result.get("brand_category", "")

    # ── RNA lookup debug ─────────────────────────────────────────
    print("\n" + "=" * 70, flush=True)
    print("RNA LOOKUP DEBUG INFO:", flush=True)
    print(f"  Input country_code    : '{_country_code}'",   flush=True)
    print(f"  Resolved country_name : '{_country_name}'",   flush=True)
    print(f"  Input comp_code       : '{args.comp_code}'",  flush=True)
    print(f"  Input sbu             : '{args.sbu}'",        flush=True)
    print(f"  Input brand_lov_id    : '{brand_lov_id}'",    flush=True)
    print(f"  ───────────────────────────────────────",      flush=True)
    print(f"  RESULT brand_type     : '{_brand_type}'",      flush=True)
    print(f"  RESULT brand_category : '{_brand_category}'",  flush=True)
    print(f"  Full result dict      : {_rna_result}",        flush=True)
    print("=" * 70 + "\n", flush=True)

    log.info(
        "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
        _country_name, _brand_type, _brand_category,
    )

    all_warnings: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing SP27 MAA Linelist: %s ───", ll_path.name)

        ll = SP27MAALinelistLoader(ll_path)
        if ll.df.empty:
            log.warning("[SP27MAA-Linelist] Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in ll.df.iterrows()]
        # rows = rows[:5]  # ← TEST MODE: uncomment to limit to 5 articles
        total_rows = len(rows)
        log.info("Total valid rows: %d", total_rows)

        out_name = ll_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (
                row,
                args.brand_code,
                mdd,
                args.season,
                getattr(args, "country_code", ""),
                getattr(args, "article_type_from_filename", ""),
                _brand_type,       # ← passed to every article
                _brand_category,   # ← passed to every article
            )
            for row in rows
        ]
        ordered: list[tuple[int, dict, list]] = []

        log.info(
            "Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
            total_rows, num_workers,
        )

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

        print("═══ ARTICLE SUMMARY (NIKE SP27 MAA) ═══════════════", flush=True)
        print(f"  Linelist file          : {ll_path.name}",      flush=True)
        print(f"  Linelist rows (valid)  : {total_rows}",        flush=True)
        print(f"  Articles written       : {written_count}",     flush=True)
        print(f"  XML file size          : {file_kb}KB",         flush=True)
        print(f"  Brand Type in XML      : '{_brand_type}'",     flush=True)
        print(f"  Brand Category in XML  : '{_brand_category}'", flush=True)
        print("════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"sp27maa_validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — Nike SP27 MAA v1.0"
    )
    p.add_argument("--brand",      default="Nike")
    p.add_argument("--brand-code", default="NIK")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SP27")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()