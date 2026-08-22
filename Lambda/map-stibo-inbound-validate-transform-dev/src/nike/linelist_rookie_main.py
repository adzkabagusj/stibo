"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — Nike  v1.0                   ║
║  Linelist (SP27 ROOKIE) → Stibo STEP XML                       ║
╚══════════════════════════════════════════════════════════════════╝

Nike-specific details vs Smiggle:
  • Single input file — the SP27 ROOKIE Nike Linelist
  • Header is at row 1 (0-based row 0) — no multi-row header
  • Principal Style Code  = Style_Color column (e.g. "IR6678-480")
  • Principal Style Desc  = Secondary Style Name column
  • Style Name            = Style Name column (franchise / model name)
  • Color Code            = Color Code column (3-digit, e.g. "480")
  • Color Description     = Color Description column
  • Gender                = Gender column (Men/Women/Unisex/Boys/Girls)
  • Age                   = Age column (Adult/Young Athletes)
  • Season Year           = Season Year column (e.g. "SP2027")
  • Product Type          = Product Type column (Apparel/Footwear/Equipment)
  • Body Type             = Body Type column (Top/Other/…)
  • Silhouette            = Silhouette column
  • Category              = Category column
  • SubCategory           = SubCategory column
  • Franchise             = Franchise column
  • Collection            = Collection column
  • Retail Price          = Retail Price column
  • Wholesale Price       = Wholesale Price column
  • Currency Code         = Currency Code column (IDR)
  • UOM                   = UOM column
  • Begin Futures Date    = Begin Futures Offer Date
  • Delivery              = Delivery column
  • Carryover Status      = Carryover Status column

  One Generic per (Style_Color) row.
  SAP Style Code = style part of Style_Color (before the dash).
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
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
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

# ── Nike-specific lookup tables ──────────────────────────────────

NIKE_GENDER_MAP: dict[str, str] = {
    "MEN":    "M", "MALE":   "M", "BOYS":  "M", "BOY":    "M",
    "WOMEN":  "F", "FEMALE": "F", "GIRLS": "F", "GIRL":   "F",
    "UNISEX": "U",
}

NIKE_AGE_MAP: dict[str, str] = {
    "ADULT":          "AD", "ADULTS":         "AD",
    "YOUNG ATHLETES": "CH", "YOUTH":          "CH",
    "KIDS":           "CH", "CHILDREN":       "CH",
    "INFANT":         "IN", "TODDLER":        "IN",
    "ALL AGES":       "AA", "JUNIOR":         "JR",
}

NIKE_PRODUCT_TYPE_TO_DIVISION: dict[str, str] = {
    "FOOTWEAR":    "F", "APPAREL":     "A",
    "EQUIPMENT":   "E", "ACCESSORIES": "E",
}

NIKE_PRODUCT_TYPE_TO_UOM: dict[str, str] = {
    "FOOTWEAR":    "PR", "APPAREL":     "EA",
    "EQUIPMENT":   "EA", "ACCESSORIES": "EA",
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

        self._load_named_lov_sheets(wb)
        self._load_age_lov(wb)
        self._load_retail_price_currency(wb)
        wb.close()
        self._build_country_reverse()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

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
        log.info(
            "[MDD] Age LOV loaded — SAPAge: %d entries, BYAge: %d entries",
            len(sap_age_map), len(by_age_map),
        )

    def _build_country_reverse(self):
        country_lov = self.lovs.get("Country", {})
        self.lovs["CountryByCode"] = {code: name for name, code in country_lov.items()}
        log.info(
            "[MDD] CountryByCode reverse map: %d entries",
            len(self.lovs["CountryByCode"]),
        )

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            import re as _re
            display = _re.sub(r"(?i)\s*lov\s*", "", sn).strip()
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                code, name = row[0], row[1]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )
    def _load_retail_price_currency(self, wb):
        """
        Load 'Retail Price Currency LOV' sheet dynamically starting from row 3.
        Builds a {ISO_country_code → currency_code} map.
        """
        # Step 1: Load Country LOV from MDD if present
        country_id_to_name: dict[str, str] = dict(_COUNTRY_MAP)
        country_sheet = next(
            (s for s in wb.sheetnames 
             if "COUNTRY" in s.upper() and "LOV" in s.upper()
             and "SIZE" not in s.upper() and "ORIGIN" not in s.upper()),
            None,
        )
        if country_sheet:
            for row in wb[country_sheet].iter_rows(min_row=2, values_only=True):
                if row and row[0] and len(row) > 1 and row[1]:
                    v0 = str(row[0]).strip()
                    v1 = str(row[1]).strip()
                    if v0 and v1:
                        country_id_to_name[v0.upper()] = v1.upper()
                        country_id_to_name[v1.upper()] = v0.upper()

        sheet_name = next(
            (s for s in wb.sheetnames if "RETAIL PRICE CURRENCY" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Retail Price Currency LOV sheet not found")
            return

        # Data starts at row 3 in "Retail Price Currency LOV" sheet
        raw_lov: dict[str, str] = {}
        for row in wb[sheet_name].iter_rows(min_row=3, values_only=True):
            if not row or len(row) < 2 or not row[0] or not row[1]:
                continue
            display_name = str(row[0]).strip().upper()  # e.g. "INDONESIAN RUPIAH"
            currency_code = str(row[1]).strip()          # e.g. "IDR"
            if display_name and currency_code:
                raw_lov[display_name] = currency_code

        country_currency: dict[str, str] = {
            "ID": "IDR", "PH": "PHP", "TH": "THB",
            "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD"
        }
        for code, name in country_id_to_name.items():
            if len(code) <= 3 and code.isalpha():
                iso_code = code.upper()
                keyword = name.upper()
                if iso_code == "KH" or "CAMBODIA" in keyword:
                    country_currency[iso_code] = "USD"
                    continue
                keyword_singular = keyword[:-1] if keyword.endswith("S") else keyword
                matched = next(
                    (c for display, c in raw_lov.items() if keyword in display or keyword_singular in display),
                    "",
                )
                if matched:
                    country_currency[iso_code] = matched

        self.lovs["CountryCurrency"] = country_currency
        log.info("[MDD] CountryCurrency map: %d entries → %s", len(country_currency), country_currency)


class AttributesListLoader:
    """Loads the NIKE tab from the Attributes List workbook."""

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
        # ── FIX: strip .0 suffix (e.g. "888.0" → "888") ─────────
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
            (s for s in wb.sheetnames if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found in %s", self.path.name)
            wb.close()
            return
        else:
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

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        """Return {brand_type, brand_category} for the given keys, or empty strings."""
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


class NikeLinelistLoader:
    """Loads the Nike linelist workbook (SP27 ROOKIE format)."""

    PREFERRED_SHEETS = ["Sheet1", "Linelist", "MAPI"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[Linelist-NIKE] Loading: %s", self.path.name)
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

        log.info("[Linelist-NIKE] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and "STYLE" in v.upper() for v in r if v):
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
            log.error("[Linelist-NIKE] Cannot find header row")
            wb.close()
            return

        log.info("[Linelist-NIKE] Header at row %d (0-indexed)", hdr_idx)
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

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info("[Linelist-NIKE] %d rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT") else s


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

_COUNTRY_NAME_TO_CODE: dict[str, str] = {
    "INDONESIA": "ID", "ID": "ID",
    "PHILIPPINES": "PH", "PHILIPPINE": "PH", "PH": "PH",
    "THAILAND": "TH", "TH": "TH",
    "SINGAPORE": "SG", "SG": "SG",
    "MALAYSIA": "MY", "MY": "MY",
    "VIETNAM": "VN", "VN": "VN",
    "CAMBODIA": "KH", "CAMBODIAN": "KH", "KH": "KH",
}

def _parse_country_code_from_file_or_args(filename: str, cli_country: str = "") -> str:
    stem_upper = Path(filename).stem.upper().replace("_", " ").replace("-", " ")
    tokens = stem_upper.split()
    for token in tokens:
        if token in _COUNTRY_NAME_TO_CODE and token not in ("SP", "FF", "NIK"):
            return _COUNTRY_NAME_TO_CODE[token]
    cli_c = (cli_country or "").strip().upper()
    if cli_c:
        return cli_c
    return "ID"

def _resolve_retail_price_currency(country_code: str, mdd: "MDDLoader | None") -> str:
    """country_code (e.g. 'ID', 'KH', 'PH') → Retail Price Currency LOV ID (e.g. 'IDR', 'USD', 'PHP')."""
    fallback_map = {
        "ID": "IDR", "PH": "PHP", "TH": "THB",
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD"
    }
    code_upper = (country_code or "").strip().upper()
    if not code_upper:
        return ""
    if mdd:
        res = mdd.lovs.get("CountryCurrency", {}).get(code_upper, "")
        if res:
            return res
    return fallback_map.get(code_upper, "")
    
def map_article_nike(ll_row: dict, brand_code: str = "NIK") -> dict:
    style_color  = _s(ll_row.get("Style_Color"))
    style_name   = _s(ll_row.get("Style Name"))
    sec_name     = _s(ll_row.get("Secondary Style Name"))
    color_code   = _s(ll_row.get("Color Code"))
    color_desc   = _s(ll_row.get("Color Description"))
    gender_raw   = _s(ll_row.get("Gender"))
    age_raw      = _s(ll_row.get("Age"))
    season_year  = _s(ll_row.get("Season Year"))
    product_type = _s(ll_row.get("Product Type"))
    body_type    = _s(ll_row.get("Body Type"))
    silhouette   = _s(ll_row.get("Silhouette"))
    category     = _s(ll_row.get("Category"))
    sub_category = _s(ll_row.get("SubCategory"))
    franchise    = _s(ll_row.get("Franchise"))
    collection   = _s(ll_row.get("Collection"))
    retail_price = _s(ll_row.get("Retail Price"))
    wholesale    = _s(ll_row.get("Wholesale Price"))
    currency     = _s(ll_row.get("Currency Code"))
    uom_raw      = _s(ll_row.get("UOM"))
    begin_date   = ll_row.get("Begin Futures Offer Date")
    delivery     = _s(ll_row.get("Delivery"))
    carryover    = _s(ll_row.get("Carryover Status"))
    channel_tier = _s(ll_row.get("Channel/Tier"))

    sap_style_code = re.sub(r"-", "", style_color)[:9] if style_color else ""
    gender_code    = NIKE_GENDER_MAP.get(gender_raw.upper(), "U")
    age_code       = NIKE_AGE_MAP.get(age_raw.upper(), "AD")
    div_letter     = NIKE_PRODUCT_TYPE_TO_DIVISION.get(product_type.upper(), "A")
    uom_code       = uom_raw.upper() if uom_raw else NIKE_PRODUCT_TYPE_TO_UOM.get(product_type.upper(), "EA")
    coo            = "SG"
    launch_date    = _fmt_date(begin_date)

    sap_style_clean = re.sub(r"[^A-Z0-9]", "", sap_style_code.upper())[:9]
    generic_code    = f"{brand_code}{sap_style_clean}"
    colour_token    = color_code.strip().zfill(3) if color_code else "000"
    variant_code    = f"{generic_code}{colour_token}000"

    rrp_str = ""
    if retail_price:
        m = re.search(r"[\d.]+", str(retail_price))
        rrp_str = m.group() if m else str(retail_price)

    fob_str = ""
    if wholesale:
        m = re.search(r"[\d.]+", str(wholesale))
        fob_str = m.group() if m else str(wholesale)

    article_no_clean = re.sub(r"-", "", style_color) if style_color else ""

    return {
        "article_no":       style_color,
        "article_no_clean": article_no_clean,
        "sap_style_code":   sap_style_clean,
        "model_name":       sec_name or style_name,
        "style_name":       style_name,
        "brand_code":       brand_code,
        "colour":           color_desc,
        "colour_code":      colour_token,
        "colour_token":     colour_token,
        "gender_code":      gender_code,
        "gender_raw":       gender_raw,
        "age_code":         age_code,
        "age_raw":          age_raw,
        "product_type":     product_type,
        "division":         product_type,
        "div_letter":       div_letter,
        "body_type":        body_type,
        "silhouette":       silhouette,
        "category":         category,
        "sub_category":     sub_category,
        "article_type":     "Inline",
        "art_category":     "1",
        "franchise":        franchise,
        "collection":       collection,
        "carryover":        carryover,
        "channel_tier":     channel_tier,
        "delivery":         delivery,
        "rrp":              rrp_str,
        "currency":         currency or "IDR",
        "fob":              fob_str,
        "fob_currency":     currency or "IDR",
        "coo":              coo,
        "uom_code":         uom_code,
        "launch_date":      launch_date,
        "generic_code":     generic_code,
        "variant_code":     variant_code,
        "season_year_raw":  season_year,
        "_tdd_sizes":       [],
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
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
# SECTION 4 — XML HELPERS
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
    mdd: MDDLoader | None, brand_name: str, fallback_code: str = ""
) -> tuple[str, str]:
    brand_upper  = (brand_name or "").strip().upper()
    brand_label  = brand_upper
    brand_lov_id = fallback_code or brand_upper[:3]

    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})
        log.info(
            "[resolve_brand_lov_id] MDD Brand LOV (%d entries): %s",
            len(brand_lov), dict(list(brand_lov.items())[:20]),
        )

        for key, val in brand_lov.items():
            key_up = (key or "").strip().upper()
            val_up = str(val).strip().upper() if val else ""
            if key_up == brand_upper and val_up != brand_upper:
                brand_lov_id = str(val).strip()
                brand_label  = brand_upper
                log.info(
                    "[resolve_brand_lov_id] MATCH → key='%s'  val='%s'  → brand_lov_id='%s'",
                    key, val, brand_lov_id,
                )
                break
        else:
            for key, val in brand_lov.items():
                val_up = str(val).strip().upper() if val else ""
                key_up = (key or "").strip().upper()
                if val_up == brand_upper and key_up != brand_upper:
                    brand_lov_id = str(key).strip()
                    brand_label  = brand_upper
                    log.info(
                        "[resolve_brand_lov_id] MATCH (reverse) → key='%s'  val='%s'  → brand_lov_id='%s'",
                        key, val, brand_lov_id,
                    )
                    break
            else:
                log.warning(
                    "[resolve_brand_lov_id] NO MATCH for '%s' — falling back to '%s'",
                    brand_upper, brand_lov_id,
                )

    log.info(
        "[resolve_brand_lov_id] RESULT → brand_lov_id='%s'  brand_label='%s'",
        brand_lov_id, brand_label,
    )
    return brand_lov_id, brand_label


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
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
    brand_lov_id = art["brand_code"]
    _w("AT_Brand",      id_val=brand_lov_id)
    _w("AT_BrandGroup", id_val=brand_name.upper())

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["article_no"])       
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["colour_code"])      
    _w("AT_PrincipalColorName",        art["colour"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])

    # ── AT_InboundGenericCode ────────────────────────────────────
    at_generic_val        = f"{brand_lov_id}{art['article_no_clean']}"
    art["at_generic_val"] = at_generic_val
    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    g_code  = art["gender_code"]
    g_label = LOV_GENDER.get(g_code, g_code)
    _w("AT_Gender",                     g_label, id_val=g_code)
    _w("AT_BYGender",                   g_label, id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age ──────────────────────────────────────────────────────
    age_raw_upper = (art["age_raw"] or "").strip()
    sap_age_map   = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map    = mdd.lovs.get("BYAge",  {}) if mdd else {}

    matched_age_display = next(
        (k for k in sap_age_map if k.upper() == age_raw_upper.upper()), None
    )
    sap_age_code  = sap_age_map.get(matched_age_display, art["age_code"]) if matched_age_display else art["age_code"]
    sap_age_label = matched_age_display or LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge", sap_age_label, id_val=sap_age_code)

    by_age_val = by_age_map.get(matched_age_display, "") if matched_age_display else ""
    if not by_age_val:
        by_age_val = next(
            (v for k, v in by_age_map.items() if k.upper() == age_raw_upper.upper()), ""
        )
        if not by_age_val:
            if art["age_code"] == "CH":
                by_age_val = "KIDS"
            elif art["age_code"] == "AD":
                by_age_val = "ADULT"
            else:
                by_age_val = age_raw_upper or "ADULT"
    _w("AT_BYAge", by_age_val.upper(), id_val=by_age_val.upper())
    _w("AT_PrincipalAgeDescription", art["age_raw"])              

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

    # ── Country of Origin ────────────────────────────────────────
    coo_code        = art["coo"]
    country_by_code = mdd.lovs.get("CountryByCode", {}) if mdd else {}
    coo_label       = country_by_code.get(coo_code) or LOV_COUNTRY_ORIGIN.get(coo_code, coo_code)
    _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    # ── Article category / type / indicators ─────────────────────
    _w("AT_SAPArticleCategory", id_val="1")

    at_code              = art.get("article_type", "Inline")
    at_from_filename     = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    _w("AT_BYIndicator",   "Yes", id_val="Y")
    _w("AT_SAPIndicator",  "No",  id_val="N")
    # _w("AT_EcomIndicator", "")

    # ── UOM ──────────────────────────────────────────────────────
    # FIX 4: always default to EA
    _w("AT_UOM", id_val="EA")

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice", art["rrp"])
    _w("AT_CurrentPrice",  art["rrp"])
    # FIX 3: AT_FOB removed — do not send
    _w("AT_FOBCurrency",   art["fob_currency"], id_val=art["fob_currency"])
    
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
        
    # ── SAP Product Flag ─────────────────────────────────────────
    _w("AT_SAPProductFlag", "Local Production", id_val="5")

    # ── Material Type ─────────────────────────────────────────────
    _w("AT_MaterialType", "direct", id_val="ZHAW")                          

    # AT_EComProductNameEN not sent per spec                  


    # ── Principal Merchandise Hierarchy (L1–L5) ──────────────────
    _w("AT_PrincipalMerchandiseHierarchyL1", art.get("product_type",  ""))
    _w("AT_PrincipalMerchandiseHierarchyL2", art.get("category",      ""))
    _w("AT_PrincipalMerchandiseHierarchyL3", art.get("sub_category",  ""))
    _w("AT_PrincipalMerchandiseHierarchyL4", art.get("silhouette",    ""))
    _w("AT_PrincipalMerchandiseHierarchyL5", art.get("franchise",     ""))

    # ── Collection ───────────────────────────────────────────────
    _w("AT_Collection1", art.get("franchise", ""))

    # ── Country ──────────────────────────────────────────────────
    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)

    # ── Brand Type / Brand Category ───────────────────────────────
    _w("AT_BrandType",     art.get("brand_type",     ""))
    _w("AT_BrandCategory", art.get("brand_category", ""))

    # ── Country Size ─────────────────────────────────────────────
    # Footwear / Equipment → "US"  |  Apparel → do NOT send       
    if (art.get("product_type") or "").strip().upper() != "APPAREL":
        _w("AT_CountrySize", "US")


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
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
    mdd=None,
) -> str:
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = art.get("brand_code", brand_code) or brand_code
    div_letter = art.get("div_letter", "A")
    parent_id  = f"PPH_{div_letter}-TempSubCat"
    key_article = f"{b_code}{art['article_no_clean']}"

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
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    (row, brand_code, mdd, season, country_code,
     article_type_from_filename, brand_type, brand_category) = row_tuple
    mapped = map_article_nike(row, brand_code=brand_code)
    mapped["season"]                    = season
    mapped["country_code"]              = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    mapped["brand_type"]                = brand_type
    mapped["brand_category"]            = brand_category
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
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

    mdd = MDDLoader(mdd_f)
    _al = AttributesListLoader(attr_f, brand="NIKE")
    rna = RNALoader(attr_f)

    # ── RNA debug ────────────────────────────────────────────────
    print("\n" + "=" * 70, flush=True)
    print("RNA LOADER DEBUG:", flush=True)
    print(f"  Total RNA entries loaded: {len(rna.lookup)}", flush=True)
    for i, (key, val) in enumerate(list(rna.lookup.items())[:10]):
        print(f"    {i+1}. Key={key} → Value={val}", flush=True)
    nik_entries = {k: v for k, v in rna.lookup.items() if k[3] == "NIK"}
    print(f"  All NIK entries ({len(nik_entries)}):", flush=True)
    for key, val in list(nik_entries.items())[:5]:
        print(f"    {key} → {val}", flush=True)
    print("=" * 70 + "\n", flush=True)

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    brand_lov_id, brand_lov_label = resolve_brand_lov_id(
        mdd, brand_name=args.brand, fallback_code=args.brand_code,
    )
    log.info(
        "Resolved brand LOV → brand_name='%s'  brand_label='%s'  brand_lov_id='%s'",
        args.brand, brand_lov_label, brand_lov_id,
    )

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    # ── Resolve AT_BrandType / AT_BrandCategory ──────────────────
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
    print(f"  Input country_code    : '{_country_code}'",  flush=True)
    print(f"  Resolved country_name : '{_country_name}'",  flush=True)
    print(f"  Input comp_code       : '{args.comp_code}'", flush=True)
    print(f"  Input sbu             : '{args.sbu}'",       flush=True)
    print(f"  Input brand_lov_id    : '{brand_lov_id}'",   flush=True)
    print(f"  ───────────────────────────────────────",     flush=True)
    print(f"  RESULT brand_type     : '{_brand_type}'",     flush=True)
    print(f"  RESULT brand_category : '{_brand_category}'", flush=True)
    print(f"  Full result dict      : {_rna_result}",       flush=True)
    print("=" * 70 + "\n", flush=True)

    log.info(
        "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
        _country_name, _brand_type, _brand_category,
    )

    all_warnings: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing Linelist: %s ───", ll_path.name)

        ll = NikeLinelistLoader(ll_path)
        if ll.df.empty:
            log.warning("[Linelist-NIKE] Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in ll.df.iterrows()]
        # rows = rows[:5]  # <- TEST MODE: uncomment to limit to 5 articles
        # -- DEV LIMIT: set to 0 or remove to process all products --
        # DEV_PRODUCT_LIMIT = 5
        # if DEV_PRODUCT_LIMIT:
        #     rows = rows[:DEV_PRODUCT_LIMIT]
        #     log.warning("[DEV] Product limit active — processing %d article(s) only.", DEV_PRODUCT_LIMIT)
        total_rows = len(rows)
        log.info("Total valid rows: %d", total_rows)

        _stem  = ll_path.stem
        _parts = re.split(r"\s*-\s*", _stem)
        file_type_from_input  = "-".join(_parts[3:-3]) if len(_parts) >= 7 else "Linelist"
        _mm_match             = re.search(r'\b(Multi|Mono)\b', _stem, re.IGNORECASE)
        multi_mono_from_input = _mm_match.group(1).capitalize() if _mm_match else "Multi"

        log.info(
            "Parsed from filename → file_type='%s'  multi_mono='%s'",
            file_type_from_input, multi_mono_from_input,
        )

        out_path = XML_OUT_DIR / (ll_path.stem + ".xml")

        file_country_code = _parse_country_code_from_file_or_args(
            ll_path.name, getattr(args, "country_code", "")
        )
        log.info("[Linelist-Nike] Country code resolved: '%s' for file: %s", file_country_code, ll_path.name)

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (
                row, brand_lov_id, mdd, args.season,
                file_country_code,
                getattr(args, "article_type_from_filename", ""),
                _brand_type,
                _brand_category,
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
        log.info("Pass 2/2 — streaming XML to %s …", out_path.name)
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

        print("═══ ARTICLE SUMMARY (NIKE) ═════════════════════════", flush=True)
        print(f"  Linelist rows (valid)  : {total_rows}",    flush=True)
        print(f"  Articles written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb}KB",     flush=True)
        print(f"  Brand Type in XML      : '{_brand_type}'", flush=True)
        print(f"  Brand Category in XML  : '{_brand_category}'", flush=True)
        print("════════════════════════════════════════════════════", flush=True)

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
# SECTION 8 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — Nike v1.0")
    p.add_argument("--brand",      default="Nike")
    p.add_argument("--brand-code", default="NIK")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SP27")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()