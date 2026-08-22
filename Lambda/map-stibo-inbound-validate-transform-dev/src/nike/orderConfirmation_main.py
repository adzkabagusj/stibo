"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — Nike  v1.0                   ║
║  Order Confirmation Report → Stibo STEP XML                    ║
╚══════════════════════════════════════════════════════════════════╝

Nike Order Confirmation details:
  • Input   : Nike Order Confirmation report (Excel / TSV / CSV)
  • Each row = one size line for a Style-Color combination
  • Groups rows by Prod Cd → one Generic article per Style-Color

  Column → XML attribute mappings:
    Prod Cd                      → AT_PrincipalStyleCode  (e.g. "IF1746-202")
                                   AT_SAPStyleCode        (before dash  "IF1746")
                                   AT_PrincipalColorCode  (after dash   "202")
    Matl Desc                    → Name, AT_PrincipalStyleDescription,
                                   AT_EComProductNameEN
    Gndr Desc                    → AT_Gender, AT_BYGender,
                                   AT_PrincipalGenderDescription
    Gndr Age Desc                → AT_SAPAge, AT_BYAge,
                                   AT_PrincipalAgeDescription
    Silh Desc                    → AT_Silhouette, AT_PrincipalMerchandiseHierarchyL4
    Colr Short Desc              → AT_PrincipalColorName
    Gbl Cat Core Focs Desc       → AT_Collection1, AT_Franchise,
                                   AT_PrincipalMerchandiseHierarchyL2/L5
    Cat Desc                     → AT_Interest,
                                   AT_PrincipalMerchandiseHierarchyL3
    Sprt Acty Desc               → AT_SportsCategoryEN
    Div Cd                       → AT_ProductDivisionID,
                                   AT_PrincipalMerchandiseHierarchyL1
    SO Doc Hdr Nbr               → AT_SODocumentNumber
    Cust PO Nbr                  → AT_CustomerPONumber
    Cust Sold To Nm              → AT_SoldToCustomer
    CRD Dt                       → AT_CRDDate
    CCD Dt Bus Seasn Yr Cd       → AT_Season + AT_Year (+ Classification IDs)
    CCD Dt Bus Mo Yr (Mon-YYYY)  → AT_SeasonMonthYear
    Lnch Dt / Valid From Date    → AT_LaunchingDate
    Trans MSRP                   → AT_OriginalPrice, AT_CurrentPrice
    Whlsl Trans Prc              → AT_FOB
    TC Cd                        → AT_RetailPriceCurrency, AT_FOBCurrency
    EAN UPC Cd                   → AT_Barcode (first size row)
    Cnfrmd Qty      (SUM sizes)  → AT_ConfirmedQty
    Cnfrmd Net Amt  (SUM sizes)  → AT_ConfirmedNetAmount
    Opn Qty         (SUM sizes)  → AT_OpenQty
"""

from __future__ import annotations

import logging
import os
import re
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# ── Directory setup ──────────────────────────────────────────────
BASE_DIR    = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR   = BASE_DIR / "input"
ORDER_DIR   = INPUT_DIR / "linelist"      # files arrive in linelist/ via variant dispatch
MDD_DIR     = INPUT_DIR / "mdd"
ATTR_DIR    = INPUT_DIR / "attributes"
OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for _d in [ORDER_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_orderconfirm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# LOV / LOOKUP TABLES
# ══════════════════════════════════════════════════════════════════

LOV_BRAND: dict[str, str] = {
    "ADI": "ADIDAS",  "NIK": "NIKE",    "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO",    "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}

# Country code → MDD COUNTRY name (for RNA lookup)
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
    "KH": "CAMBODIA",
}

# Nike Div Cd → (division_letter, division_label, PPH_parent, default_uom)
NIKE_DIV_CD_MAP: dict[str, tuple] = {
    "10": ("A", "Apparel",   "PPH_A-TempSubCat", "EA"),
    "20": ("F", "Footwear",  "PPH_F-TempSubCat", "PAA"),
    "30": ("E", "Equipment", "PPH_E-TempSubCat", "EA"),
    "40": ("E", "Equipment", "PPH_E-TempSubCat", "EA"),
}

# Nike Gndr Desc → SAP gender code
NIKE_GNDR_DESC_MAP: dict[str, str] = {
    "UNISEX": "U", "MEN":    "M", "MENS":   "M",
    "MALE":   "M", "BOYS":   "M", "BOY":    "M",
    "WOMEN":  "F", "WOMENS": "F", "FEMALE": "F",
    "GIRLS":  "F", "GIRL":   "F",
}

# Gndr Age Desc substring rules → (SAP Age code, BY Age code)
# BY Age code must match Stibo LOV_BYAge IDs: ADULT, ALL AGES, GRADE SCHOOL, INFANT, KIDS, PRESCHOOL
# LOV_Age (AT_SAPAge) only contains AD, AA, CH in all supported environments.
# IN (Infants) and JR (Youth/Junior) are not valid LOV IDs — map to CH.
NIKE_AGE_DESC_RULES: list[tuple] = [
    ("GRD SCHOOL",   "CH", "GRADE SCHOOL"),
    ("GRADE SCHOOL", "CH", "GRADE SCHOOL"),
    ("PRE SCHOOL",   "CH", "PRESCHOOL"),
    ("LITTLE KIDS",  "CH", "KIDS"),
    ("YOUTH",        "CH", "KIDS"),
    ("JUNIOR",       "CH", "KIDS"),
    ("TODDLER",      "CH", "INFANT"),
    ("INFANT",       "CH", "INFANT"),
    ("BABY",         "CH", "INFANT"),
    ("MEN",          "AD", "ADULT"),
    ("WOMEN",        "AD", "ADULT"),
    ("ADULT",        "AD", "ADULT"),
]

# Season code → display name for Classification <Name>
LOV_SEASON_DISPLAY: dict[str, str] = {
    "FA": "Fall Winter",   "FW": "Fall Winter",
    "SS": "Spring Summer", "SP": "Spring",
    "HO": "Holiday",       "SM": "Summer",
    "AL": "All Season",    "AW": "Autumn Winter",
}

# SAP age code → display label (matches Stibo LOV_Age)
LOV_AGE: dict[str, str] = {
    "AD": "Adults",
    "CH": "Children",
    "IN": "Infants",
    "JR": "Youth",
    "AA": "All Ages",
}

# Gender code → display label
_GENDER_LABEL: dict[str, str] = {
    "M": "Male",
    "F": "Female",
    "U": "Unisex",
}

# MAA 3-character size code mapping  (KEY_Variant = inbound_code + 3-char size)
#
# Nike size sources (from "Sz Desc" column in Order Confirmation):
#   Adult footwear (US numeric): 5, 6, 6.5 … 12  →  float × 10 → 3 digits
#   Kids Years  (Y): 1Y, 1.5Y, 2Y … 6.5Y          →  explicit map
#   Toddler C   (C): 3C … 13C                       →  explicit map
#   Apparel alpha: S, M, L, XL, 2XL, M/L, L/XL    →  explicit map
#   Special: 1SIZE, MISC                            →  explicit map
MAA_SIZE_CODE_MAP: dict[str, str] = {
    # ── Kids Years (Y) ──────────────────────────────────────────────────────
    "1Y":   "01Y",  "1.5Y": "1HY",
    "2Y":   "02Y",  "2.5Y": "2HY",
    "3Y":   "03Y",  "3.5Y": "3HY",
    "4Y":   "04Y",  "4.5Y": "4HY",
    "5Y":   "05Y",  "5.5Y": "5HY",
    "6Y":   "06Y",  "6.5Y": "6HY",
    "7Y":   "07Y",  "7.5Y": "7HY",
    "8Y":   "08Y",  "9Y":   "09Y",
    # ── Toddler C ───────────────────────────────────────────────────────────
    "3C":  "03C", "4C":  "04C", "5C":  "05C", "6C":  "06C",
    "7C":  "07C", "8C":  "08C", "9C":  "09C", "10C": "10C",
    "11C": "11C", "12C": "12C", "13C": "13C",
    # ── Apparel / alpha ─────────────────────────────────────────────────────
    "XXXS": "XXS", "XXS": "XXS", "XS": "XSS",
    "S":    "SMM", "M":   "MMM", "L":  "LGG", "XL":  "XLL",
    "2XL":  "2XL", "3XL": "3XL", "4XL": "4XL", "5XL": "5XL",
    "XS/S": "XSS", "S/M": "SMM", "M/L": "MLL", "L/XL": "LXL",
    # ── Special ─────────────────────────────────────────────────────────────
    "OS": "OSS", "O/S": "OSS", "ONE SIZE": "OSS", "1SIZE": "OSS",
    "MISC": "MSC",
}


def _maa_size_code(size_val: str) -> str:
    """Return a 3-character MAA size code for use in KEY_Variant.

    Priority:
      1. Direct lookup in MAA_SIZE_CODE_MAP (Y/C/alpha/special sizes).
      2. Numeric (adult footwear): parse as float → int(val × 10) → 3-digit str.
         e.g. "5" → 050,  "6.5" → 065,  "10" → 100,  "10.5" → 105
         This avoids collisions between e.g. "10" (adult) and "10C" (toddler).
      3. Fallback: keep A-Z0-9, truncate/pad to 3 chars.
    """
    s = size_val.strip().upper()
    if s in MAA_SIZE_CODE_MAP:
        return MAA_SIZE_CODE_MAP[s]
    # Numeric adult path: only if string is all-numeric (digits + optional dot)
    if re.match(r'^\d+(\.\d+)?$', s):
        try:
            code = str(int(round(float(s) * 10))).zfill(3)
            return code[-3:]  # cap at 3 chars
        except ValueError:
            pass
    # Alpha fallback: keep A-Z0-9, pad/truncate to 3
    alnum = re.sub(r"[^A-Z0-9]", "", s)
    return alnum[:3].ljust(3, "X") if alnum else "XXX"


# Sold-to customer code → company code
SOLD_TO_COMPANY_MAP: dict[str, str] = {
    "0005064337": "0888",
    "0005064338": "0886",
    "0005064339": "0882",
}

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
ET.register_namespace("xsi", STIBO_XSI)


# ══════════════════════════════════════════════════════════════════
# SECTION 0 — RNA LOADER  (Brand Type + Brand Category)
# ══════════════════════════════════════════════════════════════════

class RNALoader:
    """
    Loads 'Source Mapping Related RNA' tab from the MDD / Attributes List workbook.
    Maps (Country Name, Company Code, SBU, Brand Code) → Brand Type + Brand Category.
    """

    RNA_SHEET_REQUIRED = ("SOURCE MAPPING", "RNA")

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

    def _load(self) -> None:
        log.info("[RNA] Loading from: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.warning("[RNA] Cannot open workbook: %s", e)
            return

        sheet_name = next(
            (s for s in wb.sheetnames
             if all(kw in s.upper() for kw in self.RNA_SHEET_REQUIRED)),
            None,
        )
        if not sheet_name:
            log.warning(
                "[RNA] No sheet matching 'Source Mapping … RNA' in %s. Sheets: %s",
                self.path.name, wb.sheetnames,
            )
            wb.close()
            return

        log.info("[RNA] Using sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        wb.close()

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(isinstance(v, str) and "COUNTRY" in v.upper() for v in r if v)),
            None,
        )
        if hdr_idx is None:
            log.warning("[RNA] Cannot find header row in sheet '%s'", sheet_name)
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip().upper(): i for i, h in enumerate(hdr) if h}

        def _find(*candidates) -> int | None:
            for c in candidates:
                if c.upper() in col:
                    return col[c.upper()]
            return None

        c_country = _find("COUNTRY", "COUNTRY NAME")
        c_comp    = _find("COMPCODE", "COMPANY CODE", "COMP CODE", "COMP_CODE")
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
        """Fuzzy lookup ignoring company code — returns first match."""
        norm_country = self._norm_country(country_name)
        norm_sbu     = (sbu or "").strip().upper()
        norm_brand   = (brand_code or "").strip().upper()

        for key, val in self.lookup.items():
            if key[0] == norm_country and key[2] == norm_sbu and key[3] == norm_brand:
                return val
        return {"brand_type": "", "brand_category": ""}


def _find_rna_source_file() -> Path | None:
    """Return the newest workbook in MDD_DIR / ATTR_DIR that contains the RNA sheet."""
    candidates: list[Path] = []
    for d in (ATTR_DIR, MDD_DIR):
        if d.exists():
            candidates += list(d.glob("*.xlsx")) + list(d.glob("*.xlsm"))
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for p in candidates:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            has_rna = any(
                all(kw in s.upper() for kw in RNALoader.RNA_SHEET_REQUIRED)
                for s in wb.sheetnames
            )
            wb.close()
            if has_rna:
                log.info("[RNA] Selected workbook: %s", p.name)
                return p
            log.info("[RNA] Skipping %s — no RNA sheet", p.name)
        except Exception as e:
            log.warning("[RNA] Cannot inspect %s: %s", p.name, e)
    return None


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADER
# ══════════════════════════════════════════════════════════════════

class NikeOrderConfirmLoader:
    """
    Loads the Nike Order Confirmation report.

    Supports: .xlsx, .xls, .csv, .txt (tab / comma / semicolon separated).

    Each unique Prod Cd (Style_Color) becomes one aggregated record:
      - Numeric fields (Cnfrmd Qty, Cnfrmd Net Amt, Opn Qty) are SUMMED.
      - All other fields take the first non-empty value across size rows.
      - All EANs are collected into _eans list.
    """

    PREFERRED_SHEETS = ["Sheet1", "Order Confirmation", "Data", "Report", "MAPI"]

    SUM_COLUMNS = ["Cnfrmd Qty", "Cnfrmd Net Amt", "Opn Qty"]

    def __init__(self, path: Path):
        self.path    = Path(path)
        self.raw_df: pd.DataFrame     = pd.DataFrame()
        self.grouped: dict[str, dict] = {}
        self._load()

    def _load(self):
        suffix = self.path.suffix.lower()
        if suffix in (".xlsx", ".xls"):
            self.raw_df = self._load_excel()
        elif suffix in (".csv", ".txt", ".tsv"):
            self.raw_df = self._load_delimited()
        else:
            raise ValueError(f"Unsupported file type: {suffix!r}")

        self.raw_df.columns = [str(c).strip() for c in self.raw_df.columns]
        log.info(f"Loaded {len(self.raw_df):,} rows from {self.path.name}")
        log.debug(f"Columns: {list(self.raw_df.columns)}")
        self._group_by_prod_cd()

    def _load_excel(self) -> pd.DataFrame:
        xl    = pd.ExcelFile(self.path)
        sheet = next(
            (s for s in self.PREFERRED_SHEETS if s in xl.sheet_names),
            xl.sheet_names[0],
        )
        log.info(f"Using sheet: '{sheet}'")
        return pd.read_excel(self.path, sheet_name=sheet, dtype=str).fillna("")

    def _load_delimited(self) -> pd.DataFrame:
        for sep in ("\t", ",", ";"):
            try:
                df = pd.read_csv(self.path, sep=sep, dtype=str).fillna("")
                if len(df.columns) > 5:
                    return df
            except Exception:
                continue
        raise ValueError(f"Could not parse file: {self.path}")

    def _group_by_prod_cd(self):
        bucket: dict[str, list[dict]] = defaultdict(list)

        for row in self.raw_df.to_dict(orient="records"):
            prod_cd = _s(row.get("Prod Cd", ""))
            if not prod_cd:
                continue
            bucket[prod_cd].append(row)

        for prod_cd, rows in bucket.items():
            first  = dict(rows[0])
            totals: dict[str, float] = {col: 0.0 for col in self.SUM_COLUMNS}
            eans:   list[str]        = []

            for r in rows:
                for col in self.SUM_COLUMNS:
                    raw = _s(r.get(col, ""))
                    if raw:
                        try:
                            totals[col] += float(re.sub(r"[,\s]", "", raw))
                        except ValueError:
                            pass
                ean = _s(r.get("EAN UPC Cd", ""))
                if ean and ean not in eans:
                    eans.append(ean)

            first.update(totals)
            first["_eans"]      = eans
            first["_size_rows"] = rows
            self.grouped[prod_cd] = first

        log.info(f"Grouped into {len(self.grouped):,} unique Style-Color articles")


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — SHARED HELPERS
# ══════════════════════════════════════════════════════════════════
def _build_variant_name(art: dict, size_val: str) -> str:
    brand  = art.get("brand_code", "")[:3].upper()
    style  = art.get("name", "").strip()
    color  = art.get("color_name", "").strip()
    size   = size_val.strip() if size_val else ""
    g_char = {"M": "M", "F": "W", "U": "U"}.get(art.get("gender_code", "U"), "U")
    gender = f"(A/{g_char})"

    brand_display = LOV_BRAND.get(brand.upper(), brand)
    style_clean   = re.sub(
        rf"^{re.escape(brand_display)}\s+", "", style, flags=re.IGNORECASE
    ).strip()

    # Full formula: Brand + Style + Gender + Color + Size
    parts = [p for p in [brand, style_clean, gender, color, size] if p]
    name  = " ".join(parts)

    if len(name) <= 40:
        return name

    # Over 40 — shorten style name, keep gender + color + size
    suffix = f" {gender} {color} {size}".strip()
    avail  = 40 - len(brand) - 1 - len(suffix)
    if avail > 3:
        style_short = style_clean[:avail]
        if len(style_clean) > avail and " " in style_short:
            style_short = style_short[:style_short.rfind(" ")].rstrip()
        candidate = f"{brand} {style_short} {suffix}".strip()
        if len(candidate) <= 40:
            return candidate
        # candidate still over 40 — drop color
        parts_no_color = [p for p in [brand, style_short, gender, size] if p]
        return " ".join(parts_no_color)[:40]
    
    # Last resort: drop color, keep gender + size
    parts_no_color = [p for p in [brand, style_clean, gender, size] if p]
    return " ".join(parts_no_color)[:40]

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "*UNK*") else s


def _clean_int(v) -> str:
    raw = _s(v)
    if not raw:
        return ""
    try:
        return str(int(float(re.sub(r"[,\s]", "", raw))))
    except ValueError:
        return raw


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y")
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y")
    raw = _s(v)
    if not raw:
        return ""
    if re.search(r"[A-Za-z]", raw):
        return raw
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y")
        except ValueError:
            pass
    return raw

def _parse_season(season_yr_cd: str) -> tuple[str, str]:
    m = re.match(r"([A-Za-z]+)(\d{2,4})", season_yr_cd.strip())
    if not m:
        return "", season_yr_cd
    code = m.group(1).upper()
    yr   = m.group(2)
    if len(yr) == 2:
        yr = "20" + yr
    return code, yr


def _resolve_gender(gndr_desc: str) -> str:
    return NIKE_GNDR_DESC_MAP.get(gndr_desc.strip().upper(), "U")


def _resolve_age(gndr_age_desc: str) -> tuple[str, str]:
    upper = gndr_age_desc.upper()
    for keyword, sap_code, by_code in NIKE_AGE_DESC_RULES:
        if keyword in upper:
            return sap_code, by_code
    return "AD", "ADULT"


def _resolve_division(div_cd: str) -> tuple[str, str, str, str]:
    return NIKE_DIV_CD_MAP.get(div_cd.strip(), ("F", "Footwear", "PPH_F-TempSubCat", "PAA"))


def _season_cls_id(brand_code: str, season_code: str, year: str) -> str:
    return f"CLH_{brand_code}_{season_code}{year}"


# Stibo LOV_SizeCode: display value → LOV ID
# FIX: Corrected mapping to use actual Stibo LOV IDs (without leading zeros for C sizes)
_STIBO_SIZE_LOV_MAP: dict[str, str] = {
    # ── Toddler C — LOV IDs from Stibo Size Code sheet ────────────
    "1C":  "01C",  "2C":  "02C",  "3C":  "03C",  "4C":  "04C",
    "5C":  "05C",  "6C":  "06C",  "7C":  "07C",  "8C":  "08C",
    "9C":  "09C",  "10C": "10C",  "11C": "11C",  "12C": "12C",
    "13C": "13C",
    # ── Kids Years (Y) — LOV IDs from Stibo Size Code sheet ───────
    "1Y":  "01Y",  "2Y":  "02Y",  "3Y":  "03Y",  "4Y":  "04Y",
    "5Y":  "05Y",  "6Y":  "06Y",  "7Y":  "07Y",  "8Y":  "08Y",
    "9Y":  "09Y",
    # ── Half sizes (Y) — LOV IDs map via 0xH pattern ──────────────
    # Stibo: 01H=1.5, 02H=2.5, 03H=3.5, 04H=4.5, 05H=5.5, 06H=6.5, 07H=7.5, 08H=8.5, 09H=9.5
    "1.5Y": "1HY",  "2.5Y": "2HY",  "3.5Y": "3HY",  "4.5Y": "4HY",
    "5.5Y": "5HY",  "6.5Y": "6HY",  "7.5Y": "7HY",
    # ── Apparel alpha ─────────────────────────────────────────────
    "XXS":  "XXS",  "XS":   "XSS",
    "S":    "SML",
    "L":    "LGE",
    "2XL":  "2XL",  "3XL":  "3XL",  "4XL":  "4XL",  "5XL":  "5XL",
    "XXL":  "XXL",
    # ── Special ───────────────────────────────────────────────────
    "OS":   "0OS",  "O/S":  "0OS",  "ONE SIZE": "0TU",  "1SIZE": "0TU",
}

def _stibo_size_lov(size_val: str) -> str:
    """
    Derive Stibo LOV_SizeCode ID from a size string.

    Priority:
      1. Direct lookup (C/Y kids, apparel alpha: S→SML, L→LGE, XS→XSS, etc.).
      2. Numeric adult footwear: whole → '5','6','10'; half → '06H','09H','10H'.
      3. Sizes with no LOV entry (M, XL, etc.) → '' — AT_Size not written.
    """
    s = size_val.strip().upper()
    # Direct lookup for C / Y sizes and apparel
    direct = _STIBO_SIZE_LOV_MAP.get(s)
    if direct:
        return direct
    # Numeric adult footwear
    try:
        f = float(s)
        whole = int(f)
        if f == whole:
            return str(whole)         # e.g. '5', '6', '10'
        else:
            return f"{whole:02d}H"    # e.g. '06H', '09H', '10H'
    except (ValueError, TypeError):
        return ""


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_order_confirm_nike(
    prod_cd:    str,
    agg_row:    dict,
    brand_code: str = "NIK",
    season_from_filename: str = "",
    year_from_filename: str = "",
    country_from_filename: str = "SG",
    brand_type: str = "",
    brand_category: str = "",
) -> dict:
    # ── Raw fields ───────────────────────────────────────────────
    matl_desc    = _s(agg_row.get("Matl Desc"))
    gndr_desc    = _s(agg_row.get("Gndr Desc"))
    gndr_age     = _s(agg_row.get("Gndr Age Desc"))
    silh_desc    = _s(agg_row.get("Silh Desc"))
    colr_desc    = _s(agg_row.get("Colr Short Desc"))
    gbl_core     = _s(agg_row.get("Gbl Cat Core Focs Desc"))
    gbl_sum      = _s(agg_row.get("Gbl Cat Sum Desc"))
    cat_desc     = _s(agg_row.get("Cat Desc"))
    sprt_acty    = _s(agg_row.get("Sprt Acty Desc"))
    div_cd       = _s(agg_row.get("Div Cd"))
    so_doc       = _s(agg_row.get("SO Doc Hdr Nbr"))
    cust_po      = _s(agg_row.get("Cust PO Nbr"))
    sold_to_nm   = _s(agg_row.get("Cust Sold To Nm"))
    sold_to_cd   = _s(agg_row.get("Cust Sold To Cd"))
    crd_dt       = _s(agg_row.get("CRD Dt"))
    season_yr_cd = _s(agg_row.get("CCD Dt Bus Seasn Yr Cd"))
    season_mo_yr = _s(agg_row.get("CCD Dt Bus Mo Yr (Mon-YYYY)"))
    lnch_dt      = _s(agg_row.get("Lnch Dt")) or _s(agg_row.get("Valid From Date"))
    trans_msrp   = _s(agg_row.get("Trans MSRP"))
    whlsl_prc    = _s(agg_row.get("Whlsl Trans Prc"))
    tc_cd        = _s(agg_row.get("TC Cd"))
    eans: list   = agg_row.get("_eans", [])

    cnfrmd_qty     = agg_row.get("Cnfrmd Qty",     0)
    cnfrmd_net_amt = agg_row.get("Cnfrmd Net Amt", 0.0)
    opn_qty        = agg_row.get("Opn Qty",        0)

    # ── Derived ──────────────────────────────────────────────────
    if "-" in prod_cd:
        sap_style_code, color_raw = prod_cd.rsplit("-", 1)
    else:
        sap_style_code = prod_cd
        color_raw      = "000"
    # MAA colour code — always exactly 3 digits (zero-pad short, take last 3 if >3)
    # Nike prod_cd colours are standard 3-digit codes e.g. "202", "001"
    _cr = re.sub(r"[^0-9A-Z]", "", color_raw.strip().upper())
    colour_token = _cr.zfill(3)[-3:]   # pad to 3 → keep last 3 chars

    style_clean  = re.sub(r"[^A-Z0-9]", "", sap_style_code.upper())
    inbound_code = f"{brand_code}{style_clean}{colour_token}"

    gender_code          = _resolve_gender(gndr_desc)
    sap_age_code, by_age = _resolve_age(gndr_age)

    div_letter, div_label, parent_id, default_uom = _resolve_division(div_cd)

    # Use season from filename if provided, otherwise parse from data
    if season_from_filename and year_from_filename:
        season_code = season_from_filename
        year_str = year_from_filename
    else:
        season_code, year_str = _parse_season(season_yr_cd) if season_yr_cd else ("", "")

    # Normalize season codes — FA→FW (FA not in Stibo LOV_Season)
    _SEASON_NORM = {"FA": "FW", "AW": "FW"}
    season_code = _SEASON_NORM.get(season_code.upper(), season_code.upper()) if season_code else season_code

    comp_code     = SOLD_TO_COMPANY_MAP.get(sold_to_cd.strip(), "0888")
    rrp_str       = _clean_int(trans_msrp)
    fob_str       = _clean_int(whlsl_prc)
    currency      = tc_cd.strip().upper() if tc_cd else "IDR"
    brand_display = LOV_BRAND.get(brand_code.upper(), brand_code.capitalize())

    # Strip leading brand name from matl_desc to avoid duplication (e.g. "NIKE NIKE Air Max")
    matl_desc_clean = re.sub(
        rf"^{re.escape(brand_display)}\s+", "", matl_desc, flags=re.IGNORECASE
    ).strip()
    ecom_name = f"{brand_display} {matl_desc_clean}".strip()

    return {
        # Identifiers
        "inbound_code":   inbound_code,
        "prod_cd":        prod_cd,
        "prod_cd_no_dash": prod_cd.replace("-", ""),
        "sap_style_code": sap_style_code,
        "color_code":     colour_token,
        "brand_code":     brand_code,
        "comp_code":      comp_code,
        "parent_id":      parent_id,
        "country_code":   country_from_filename,
        # Names
        "name":           matl_desc,
        "ecom_name":      ecom_name,
        # Gender / Age
        "gender_code":    gender_code,
        "gndr_desc":      gndr_desc,
        "sap_age_code":   sap_age_code,
        "by_age_code":    by_age,
        "gndr_age_desc":  gndr_age,
        # Division / UOM
        "div_letter":     div_letter,
        "div_label":      div_label,
        "uom":            default_uom,
        # Season
        "season_code":    season_code,
        "year":           year_str,
        "season_yr_cd":   season_yr_cd,
        "season_mo_yr":   season_mo_yr,
        # Prices
        "rrp":            rrp_str,
        "fob":            fob_str,
        "currency":       currency,
        # Order confirmation fields
        "so_doc_nbr":     so_doc,
        "cust_po":        cust_po,
        "sold_to_nm":     sold_to_nm,
        "crd_date":       _fmt_date(crd_dt),
        "launch_date":    _fmt_date(lnch_dt),
        "cnfrmd_qty":     _clean_int(cnfrmd_qty),
        "cnfrmd_net_amt": _clean_int(cnfrmd_net_amt),
        "opn_qty":        _clean_int(opn_qty),
        # Descriptors
        "color_name":      colr_desc,
        "silhouette":      silh_desc,
        "franchise":       gbl_core,
        "collection1":     gbl_core,
        "sports_category": sprt_acty,
        "interest":        cat_desc,
        # Merchandise hierarchy
        "mh_l1": div_label,
        "mh_l2": gbl_core,
        "mh_l3": cat_desc,
        "mh_l4": silh_desc,
        "mh_l5": gbl_sum,
        # EAN / size rows (used to build PRD_VariantArticle with DC_Barcode)
        "ean":        eans[0] if eans else "",
        "size_rows":  agg_row.get("_size_rows", []),
        # Country of Origin (ISO-2 code if present in source)
        "coo": _s(agg_row.get("Cntry of Orig Cd") or agg_row.get("COO") or agg_row.get("Country of Origin", "")),
        # Brand Type / Brand Category (passed from run() after RNA lookup)
        "brand_type":     brand_type,
        "brand_category": brand_category,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

_REQUIRED = [
    "inbound_code", "prod_cd", "sap_style_code", "color_code",
    "name", "gender_code", "season_code", "year", "div_letter",
]

def validate(art: dict) -> list[str]:
    errors: list[str] = []

    for field in _REQUIRED:
        if not art.get(field):
            errors.append(f"Missing required field: '{field}'")

    code = art.get("inbound_code", "")
    if code and not re.match(r"^[A-Z]{3}[A-Z0-9]{3,12}$", code):
        errors.append(f"Unexpected InboundCode format: {code!r}")

    if art.get("gender_code") not in ("M", "F", "U"):
        errors.append(f"Unknown gender code: {art.get('gender_code')!r}")

    if art.get("sap_age_code") not in ("AD", "CH", "AA"):
        errors.append(f"Unknown SAP age code: {art.get('sap_age_code')!r}")

    return errors


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — XML ELEMENT HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(
    parent:  ET.Element,
    attr_id: str,
    value:   str = "",
    id_val:  str = "",
) -> ET.Element | None:
    clean_id  = str(id_val).strip() if id_val  else ""
    clean_val = str(value).strip()  if value   else ""
    if clean_id  in ("", "None", "nan"): clean_id  = ""
    if clean_val in ("", "None", "nan"): clean_val = ""
    if not clean_id and not clean_val:
        return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if clean_id:  el.set("ID", clean_id)
    if clean_val: el.text = clean_val
    return el


def _val_always(parent: ET.Element, attr_id: str, value: str = "") -> ET.Element:
    """Write <Value> even when value is empty (required for schema compliance)."""
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if value:
        el.text = str(value)
    return el


def _multival(parent: ET.Element, attr_id: str, id_list: list[str]) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    for id_v in id_list:
        if id_v:
            v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
            v.set("ID", id_v)


def _add_product_values(
    vals_el:    ET.Element,
    art:        dict,
    sbu:        str,
    brand_type: str = "",
    brand_cat:  str = "",
) -> None:
    pass
    # _multival(vals_el, "AT_SBU",         [sbu] if sbu else [])
    # _multival(vals_el, "AT_CompanyCode", [art["comp_code"]] if art["comp_code"] else [])
    # _val(vals_el, "AT_Brand",            id_val=art["brand_code"])
    # _val(vals_el, "AT_BrandGroup",       id_val=art["brand"].upper())
    # bt = art.get("brand_type", "") or brand_type
    # bc = art.get("brand_category", "") or brand_cat
    # _val(vals_el, "AT_BrandType",        bt)
    # _val(vals_el, "AT_BrandCategory",    bc)
    # _val(vals_el, "AT_PrincipalStyleCode",        art["prod_cd_no_dash"])
    # _val(vals_el, "AT_PrincipalStyleDescription", art["name"])
    # _val(vals_el, "AT_PrincipalColorName",        art["color_name"])
    # _val(vals_el, "AT_PrincipalColorCode",        art["color_code"])
    # _val(vals_el, "AT_SAPStyleCode",              art["sap_style_code"])
    # _val(vals_el, "AT_InboundGenericCode",        art["inbound_code"])
    # _val(vals_el, "AT_Gender",   id_val=art["gender_code"])
    # _val(vals_el, "AT_BYGender", id_val=art["gender_code"])
    # _val(vals_el, "AT_SAPAge",   id_val=art["sap_age_code"])
    # _val(vals_el, "AT_BYAge",    id_val=art["by_age_code"])
    # _val(vals_el, "AT_Season",    LOV_SEASON_DISPLAY.get(art["season_code"], art["season_code"]), id_val=art["season_code"])
    # _val(vals_el, "AT_SeasonYear", art["year"])
    # _val(vals_el, "AT_CountryOrigin",      id_val=art.get("coo", ""))
    # _val(vals_el, "AT_Country",            id_val=art["country_code"])
    # _val(vals_el, "AT_SAPArticleCategory", id_val="1")
    # _val(vals_el, "AT_BYArticleType",      id_val="Inline")
    # _val(vals_el, "AT_BYIndicator",  "Yes", id_val="Y")
    # _val(vals_el, "AT_SAPIndicator", "No",  id_val="N")
    # _val(vals_el, "AT_UOM",                id_val=art["uom"])
    # _val(vals_el, "AT_OriginalPrice",      art["rrp"])
    # _val(vals_el, "AT_CurrentPrice",       art["rrp"])
    # _val(vals_el, "AT_FOB",                art["fob"])
    # _val(vals_el, "AT_FOBCurrency",        id_val=art["currency"])
    # _val(vals_el, "AT_NatureOfArticle",    id_val="REG")
    # _val(vals_el, "AT_SAPProductFlag",     id_val="5")
    # _val(vals_el, "AT_EComProductNameEN",  art["ecom_name"])
    # _val(vals_el, "AT_LaunchingDate",      art.get("launch_date", ""))
    # _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art["mh_l1"])
    # _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art["mh_l2"])
    # _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art["mh_l3"])
    # _val(vals_el, "AT_PrincipalMerchandiseHierarchyL4", art["mh_l4"])
    # _val(vals_el, "AT_PrincipalMerchandiseHierarchyL5", art["mh_l5"])
    # _val(vals_el, "AT_CountrySize",        "US")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand_code:  str,
    season_code: str,
    year:        str,
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    if not sea_year and year:
        sea_year = year
    elif not sea_year:
        sea_year = year or ""

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    brand_name     = LOV_BRAND.get(brand_code.upper(), brand_code)
    brand_token    = brand_name.title().replace(" ", "")
    batches_parent = f"CLH_{brand_token}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
        "FA": "Fall Winter",   "AL": "All Season",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand_name.title()} {sea_name} {sea_year}".strip()
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
    season_id:  str,
    sbu:        str,
    brand_type: str = "",
    brand_cat:  str = "",
) -> ET.Element:
    brand_code  = art["brand_code"]
    brand_name  = LOV_BRAND.get(brand_code.upper(), brand_code)
    brand_token = brand_name.title().replace(" ", "")
    parent_id   = art["parent_id"]

    prod_el = ET.Element(f"{{{STIBO_NS}}}Product")
    prod_el.set("UserTypeID", "PRD_GenericArticle")
    prod_el.set("ParentID",   parent_id)

    kv = ET.SubElement(prod_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = art["inbound_code"]

    # ── Generic product values ───────────────────────────────────
    vals_el = ET.SubElement(prod_el, f"{{{STIBO_NS}}}Values")
    _val(vals_el, "AT_MaterialType", "direct", id_val="ZHAW")
    

    # ── Variant articles — one PRD_VariantArticle per size row ───
    # "Sz Desc" is the primary size column in Nike Order Confirmation
    _SIZE_COLS = [
        "Sz Desc",
        "SZ DESC",
        "Sz desc",
        "Size Desc",
        "Matl Sz Desc",
        "Matl Desc Siz Dim Val",
        "Sz Cd",
        "Size",
        "Sz Dim Val",
        "Fwd Ord Siz Dim Val",
        "Siz Dim Val",
    ]

    seen_eans: set[str] = set()
    for size_row in art.get("size_rows", []):
        ean = _s(size_row.get("EAN UPC Cd", ""))
        if not ean or ean in seen_eans:
            continue
        seen_eans.add(ean)

        size_val    = next((_s(size_row.get(c)) for c in _SIZE_COLS if _s(size_row.get(c))), "")
        size_lov_id = _stibo_size_lov(size_val) if size_val else ""

        if not size_val:
            log.warning(
                "[SIZE] prod_cd=%s ean=%s — Sz Desc column not found",
                art["prod_cd"], ean,
            )
            size_code = "000"
        else:
            size_code = _maa_size_code(size_val)

        var_id = f"{art['inbound_code']}{size_code}"

        var_el = ET.SubElement(prod_el, f"{{{STIBO_NS}}}Product")
        var_el.set("UserTypeID", "PRD_VariantArticle")
        var_kv = ET.SubElement(var_el, f"{{{STIBO_NS}}}KeyValue")
        var_kv.set("KeyID", "KEY_InboundVariant")
        var_kv.text = var_id

        # ── Variant values — AT_PrincipalSizeCode from Sz Desc ── 
        var_vals_el = ET.SubElement(var_el, f"{{{STIBO_NS}}}Values")
        if size_val:
            _val(var_vals_el, "AT_PrincipalSizeCode", size_val)

        # ── Barcode DataContainer ────────────────────────────────
        dcs_el = ET.SubElement(var_el, f"{{{STIBO_NS}}}DataContainers")
        mdc_el = ET.SubElement(dcs_el, f"{{{STIBO_NS}}}MultiDataContainer")
        mdc_el.set("Type", "DC_Barcode")
        dc_el  = ET.SubElement(mdc_el, f"{{{STIBO_NS}}}DataContainer")
        dc_el.set("Analyzer", "true")
        dcv_el = ET.SubElement(dc_el,  f"{{{STIBO_NS}}}Values")
        _val(dcv_el, "AT_Barcode",          ean)
        _val(dcv_el, "AT_BarcodeType",      id_val="P")
        _val(dcv_el, "AT_MainEANIndicator", id_val="Y")

    return prod_el


def build_full_xml(
    articles:    list[dict],
    brand_code:  str,
    season_code: str,
    year:        str,
    sbu:         str,
    brand_type:  str = "",
    brand_cat:   str = "",
) -> str:
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    root = ET.Element(f"{{{STIBO_NS}}}STEP-ProductInformation")
    root.set(f"{{{STIBO_XSI}}}schemaLocation", STIBO_SCHEMA)
    root.set("ExportTime",       now)
    root.set("ExportContext",    "Context1")
    root.set("ContextID",        "Context1")
    root.set("WorkspaceID",      "Main")
    root.set("UseContextLocale", "false")

    season_id   = _season_cls_id(brand_code, season_code, year)
    products_el = ET.SubElement(root, f"{{{STIBO_NS}}}Products")

    for art in articles:
        products_el.append(
            build_product_xml(art, season_id, sbu,
                              brand_type=brand_type, brand_cat=brand_cat)
        )

    ET.indent(root, space="  ")
    xml_body = ET.tostring(root, encoding="unicode", xml_declaration=False)
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_body


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(task: tuple) -> tuple[str | None, dict | None, list[str]]:
    prod_cd, agg_row, brand_code, season_from_filename, year_from_filename, country_from_filename, brand_type, brand_category = task
    try:
        art    = map_order_confirm_nike(
            prod_cd, agg_row, brand_code,
            season_from_filename, year_from_filename, country_from_filename,
            brand_type=brand_type, brand_category=brand_category
        )
        errors = validate(art)
        if errors:
            log.warning(f"[{prod_cd}] Validation: {errors}")
        return art["inbound_code"], art, errors
    except Exception as exc:
        log.error(f"[{prod_cd}] Processing failed: {exc}", exc_info=True)
        return None, None, [str(exc)]


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

    Expected args attributes:
      args.order_confirm_path : Path to order confirmation file (required)
      args.brand_code         : e.g. "NIK"            (default: "NIK")
      args.brand              : e.g. "Nike"           (default: "Nike")
      args.sbu                : e.g. "SP"             (default: "SP")
      args.comp_code          : e.g. "0888"           (default: "0888")
      args.seq                : e.g. 1                (default: 1)
      args.multi_mono         : e.g. "Multi"          (default: "Multi")
      args.country_code       : e.g. "ID"             (default: "")
      args.brand_type         : override brand type   (default: RNA lookup)
      args.brand_category     : override brand cat    (default: RNA lookup)
      args.output_path        : Override output path  (default: auto)
      args.workers            : Thread pool size      (default: 4)
    """
    brand_code   = getattr(args, "brand_code",   "NIK")
    sbu          = getattr(args, "sbu",          "SP")
    comp_code    = getattr(args, "comp_code",    "0888")
    workers      = getattr(args, "workers",      4)
    country_code = getattr(args, "country_code", "")

    # ── Brand Type / Brand Category ──────────────────────────────
    brand_type = getattr(args, "brand_type",     "") or ""
    brand_cat  = getattr(args, "brand_category", "") or ""

    # ── Resolve input path first to extract country code ─────────
    input_path_raw = getattr(args, "order_confirm_path", None)
    if not input_path_raw:
        candidates = [f for f in ORDER_DIR.iterdir() if f.is_file()]
        if not candidates:
            log.error(f"No input file found in {ORDER_DIR}")
            return None
        input_path_raw = str(candidates[0])
    input_path = Path(input_path_raw)

    # ── Extract country code from filename ───────────────────────
    filename_upper = input_path.stem.upper()
    country_match = re.search(r'\b(ID|MY|PH|SG|TH|VN)\b', filename_upper)
    
    if country_match:
        country_code_from_file = country_match.group(1)
        log.info(f"Extracted country from filename → {country_code_from_file}")
    else:
        country_code_from_file = country_code or "SG"  # Use args or default fallback
        log.warning(f"Could not extract country from filename, using: {country_code_from_file}")

    # ── RNA Lookup with full fallback chain ──────────────────────
    if not brand_type or not brand_cat:
        rna_path = _find_rna_source_file()
        if rna_path:
            try:
                rna          = RNALoader(rna_path)
                country_name = _COUNTRY_MAP.get(country_code_from_file, country_code_from_file)

                # Try exact match first: Nike brand code with company code
                result = rna.get(country_name, comp_code, sbu, "NIK")
                log.info(
                    "[RNA] Exact lookup → country=%s comp=%s sbu=%s brand=NIK → "
                    "brand_type='%s'  brand_category='%s'",
                    country_name, comp_code, sbu,
                    result.get("brand_type", ""), result.get("brand_category", ""),
                )

                # Fallback 1: Try with args.brand_code if different from NIK
                if (not result.get("brand_type") and not result.get("brand_category")) and brand_code != "NIK":
                    result = rna.get(country_name, comp_code, sbu, brand_code)
                    log.info(
                        "[RNA] Fallback1 → brand=%s → brand_type='%s'  brand_category='%s'",
                        brand_code, result.get("brand_type", ""), result.get("brand_category", ""),
                    )

                # Fallback 2: Fuzzy match (ignore company code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, "NIK")
                    log.info(
                        "[RNA] Fuzzy lookup → country=%s sbu=%s brand=NIK → "
                        "brand_type='%s'  brand_category='%s'",
                        country_name, sbu,
                        result.get("brand_type", ""), result.get("brand_category", ""),
                    )

                brand_type = brand_type or result.get("brand_type",     "")
                brand_cat  = brand_cat  or result.get("brand_category", "")

            except Exception as e:
                log.warning("[RNA] lookup failed: %s", e)
        else:
            log.warning(
                "[RNA] No RNA source workbook found in %s or %s — brand metadata will be empty.",
                MDD_DIR, ATTR_DIR,
            )

    log.info(
        "[OrderConfirm] brand_type='%s'  brand_category='%s'",
        brand_type, brand_cat,
    )

    log.info("=" * 60)
    log.info(f"Order Confirmation XML Generator — Brand: {brand_code}")
    log.info(f"Input  : {input_path}")
    log.info("=" * 60)

    # ── Extract season from filename ─────────────────────────────
    # Look for pattern like SP27, FW27, SS2027, etc. in filename
    season_match = re.search(r'(SP|FW|SS|FA|HO|SM|AW|AL)(\d{2,4})', filename_upper)
    
    if season_match:
        season_code = season_match.group(1)
        year_raw = season_match.group(2)
        # Convert 2-digit year to 4-digit
        if len(year_raw) == 2:
            year = f"20{year_raw}"
        else:
            year = year_raw
        log.info(f"Extracted from filename → season={season_code}  year={year}")
    else:
        # Fallback to data if not in filename
        season_code = ""
        year = ""
        log.warning("Could not extract season from filename, will use data from file")

    # ── Step 1 : Load & group ────────────────────────────────────
    loader = NikeOrderConfirmLoader(input_path)
    if not loader.grouped:
        log.error("No valid articles found in input. Aborting.")
        return None

    # ── TEST MODE: limit to 5 articles ──────────────────────────
    # loader.grouped = dict(list(loader.grouped.items())[:5])

    # ── Step 2 : Map + validate in parallel ─────────────────────
    # FIX: Pass brand_type and brand_category to each article task
    tasks = [
        (k, v, brand_code, season_code, year, country_code_from_file, brand_type, brand_cat)
        for k, v in loader.grouped.items()
    ]
    articles: list[dict]           = []
    errors:   dict[str, list[str]] = {}

    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {pool.submit(_process_article, t): t[0] for t in tasks}
        for future in as_completed(future_map):
            prod_cd = future_map[future]
            _, art, errs = future.result()
            if art:
                articles.append(art)
            if errs:
                errors[prod_cd] = errs

    log.info(f"Processed {len(articles):,} articles  |  {len(errors)} with issues")

    if not articles:
        log.error("No valid articles to write. Aborting.")
        return None

    # ── Step 3 : Use season from filename or determine from data ─────────────────────
    if not season_code or not year:
        # Fallback: determine from data if not extracted from filename
        season_counter = Counter((a["season_code"], a["year"]) for a in articles)
        season_code, year = season_counter.most_common(1)[0][0]
        log.info(f"Season from data: {season_code}{year}  (from {season_counter.most_common(1)[0][1]} articles)")
    else:
        log.info(f"Using season from filename: {season_code}{year}")

    # ── Resolve output path ──────────────────────────────────────
    if getattr(args, "output_path", None):
        output_path = Path(args.output_path)
    else:
        # Use input filename stem + .xml (matching rookie linelist approach)
        output_path = XML_OUT_DIR / (input_path.stem + ".xml")

    log.info(f"Output : {output_path}")

    # ── Step 4 : Build XML ───────────────────────────────────────
    xml_str = build_full_xml(
        articles, brand_code, season_code, year, sbu,
        brand_type=brand_type,
        brand_cat=brand_cat,
    )

    # ── Step 5 : Write output ────────────────────────────────────
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(xml_str, encoding="utf-8")
    log.info(f"XML written → {output_path}  ({len(articles)} products)")

    # ── Step 6 : Audit callback ──────────────────────────────────
    result = {
        "output":  str(output_path),
        "total":   len(loader.grouped),
        "success": len(articles),
        "errors":  errors,
    }
    if auditor:
        try:
            auditor.set_etl_result(
                total=len(loader.grouped),
                success=len(articles),
                errors=len(errors),
                output=str(output_path),
            )
        except AttributeError:
            pass

    return result


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — CLI  (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Nike Order Confirmation Report → Stibo STEP XML"
    )
    parser.add_argument("--input",          "-i", required=True)
    parser.add_argument("--output",         "-o", default=None)
    parser.add_argument("--brand",          "-b", default="Nike")
    parser.add_argument("--brand-code",           default="NIK")
    parser.add_argument("--sbu",                  default="SP")
    parser.add_argument("--comp-code",            default="0888")
    parser.add_argument("--seq",                  type=int, default=1)
    parser.add_argument("--multi-mono",           default="Multi")
    parser.add_argument("--country-code",         default="")
    parser.add_argument("--brand-type",           default="")
    parser.add_argument("--brand-category",       default="")
    parser.add_argument("--workers",        "-w", type=int, default=4)
    parsed = parser.parse_args()

    class _Args:
        order_confirm_path = parsed.input
        brand              = parsed.brand
        brand_code         = parsed.brand_code
        sbu                = parsed.sbu
        comp_code          = parsed.comp_code
        seq                = parsed.seq
        multi_mono         = parsed.multi_mono
        country_code       = parsed.country_code
        brand_type         = parsed.brand_type
        brand_category     = parsed.brand_category
        workers            = parsed.workers
        output_path        = parsed.output

    result = run(_Args())
    if result:
        log.info(
            f"Done ✓  success={result['success']}/{result['total']}"
            + (f"  errors={len(result['errors'])}" if result["errors"] else "")
        )
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()