"""
aldo/oor.py  — v5.0
====================
OOR Size Run processor — generates ONE XML from WSL001 OOR Size Run Report.

FIX LOG v5.0:
  - FIX 24: MDD Dynamic Loading — LOV_COUNTRY, LOV_COMPANY_CODE, LOV_SBU,
            LOV_CURRENCY_MAP, ALDO_SILHOUETTE_MAP, MDD_SIZE_LOV, MDD_SIZE_DESC
            
  - MDD_PATH: Set via Lambda env var, or download from S3 and provide local path.

FIX LOG v4.0:
  - FIX 21: Variant Values block contains ONLY AT_Size + AT_Variant
  - FIX 22: UPC/EAN moved to DataContainers/DC_Barcode
  - FIX 23: _add_barcode_datacontainer helper added

FIX LOG v3.0:
  - KEY_Variant fix, AT_Size fix, MDD_SIZE_LOV, color_id padding fixes
"""

import os
import re
import logging
import pandas as pd
from xml.etree import ElementTree as ET
from datetime import datetime
from pathlib import Path

log = logging.getLogger("aldo_etl")

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
BRAND_CODE   = "AOD"
BRAND_NAME   = "ALDO"

SKIP_CATEGORIES = ["service", "gfs", "shoe care", "shoecare"]

# ─────────────────────────────────────────────────────────────────

# Lambda env: os.environ.get("MDD_PATH", "/tmp/mdd/Master Data Dictionary (MAA).xlsx")
# Local dev:   Path(__file__).parent / "metadata" / "Master Data Dictionary (MAA).xlsx"
# ─────────────────────────────────────────────────────────────────
MDD_PATH = Path(os.environ.get(
    "MDD_PATH",
    str(Path(__file__).parent / "metadata" / "Master Data Dictionary (MAA).xlsx")
))

# ─────────────────────────────────────────────────────────────────
# BUILT-IN FALLBACK DICTS (used when MDD file not found)

# ─────────────────────────────────────────────────────────────────
_FALLBACK_LOV_COUNTRY: dict[str, str] = {
    "VN": "Vietnam",    "CN": "China",       "ID": "Indonesia",
    "KH": "Cambodia",   "BD": "Bangladesh",  "IN": "India",
    "MY": "Malaysia",   "TH": "Thailand",    "PH": "Philippines",
    "PK": "Pakistan",   "TR": "Turkey",      "CA": "Canada",
    "IT": "Italy",      "PT": "Portugal",    "ES": "Spain",
    "CH": "Switzerland",
}

_FALLBACK_LOV_COMPANY_CODE: dict[str, str] = {
    "0191": "PT. Aldo Indonesia Adiprk",
    "0195": "Aldo Singapore",
    "0196": "Aldo Malaysia",
    "1191": "PT. Aldo Ind Adiprk Rtl",
}

_FALLBACK_TRENZASHOP_COMP_TO_STIBO: dict[str, str] = {
    "200782": "0191", "0191": "0191", "1191": "1191",
    "0195":   "0195", "0196": "0196",
}

_FALLBACK_LOV_SBU: dict[str, str] = {
    "DO": "Aldo Ind Imp",          "D6": "Aldo Malaysia",
    "D5": "Aldo Singapore",        "DW": "Aldo Thailand",
    "PF": "Fashion",               "KV": "Fashion Footwear",
    "OF": "Fashion Footwear MAS",  "JD": "Fashion Footwear MY",
    "TA": "Fashion Thailand",      "FQ": "Foot Locker",
    "K2": "Foot Locker Malaysia",  "K0": "Foot Locker Singapore",
    "VO": "Foot Locker Vietnam",   "TF": "Footlocker Thailand",
    "K7": "MAA Cambodia",          "CH": "MAA Children",
    "FF": "MAA Fashion Footwear",  "FT": "MAA Foot Locker",
    "GO": "MAA Golf",              "SP": "MAA Sport",
    "SB": "MFA Badminton",         "GI": "Mitra Gaya Indah",
    "PY": "Payless",               "FY": "PSI Foot Locker",
    "PR": "PSI Retail",            "IF": "Sport Direct Philiphine",
    "IZ": "Sports Direct ID",      "TT": "Toys Thailand",
}

_FALLBACK_LOV_CURRENCY_MAP: dict[str, str] = {
    "US DOLLAR": "USD",  "US DOLLARS": "USD", "UNITED STATES DOLLAR": "USD", "USD": "USD",
    "INDONESIAN RUPIAH": "IDR", "RUPIAH": "IDR", "IDR": "IDR",
    "SINGAPORE DOLLAR": "SGD", "SGD": "SGD",
    "MALAYSIAN RINGGIT": "MYR", "RINGGIT": "MYR", "MYR": "MYR",
    "BRITISH POUND": "GBP", "POUND STERLING": "GBP", "GBP": "GBP",
    "EURO": "EUR", "EUR": "EUR",
    "AUSTRALIAN DOLLAR": "AUD", "AUD": "AUD",
    "CANADIAN DOLLAR": "CAD", "CAD": "CAD",
    "JAPANESE YEN": "JPY", "YEN": "JPY", "JPY": "JPY",
    "THAI BAHT": "THB", "BAHT": "THB", "THB": "THB",
    "VIETNAMESE DONG": "VND", "DONG": "VND", "VND": "VND",
    "CHINESE YUAN": "CNY", "YUAN": "CNY", "CNY": "CNY", "RMB": "RMB",
    "PHILIPPINE PESO": "PHP", "PHP": "PHP",
    "CAMBODIAN RIEL": "KHR", "RIEL": "KHR", "KHR": "KHR",
    "INDIAN RUPEE": "INR", "INR": "INR",
}

_FALLBACK_ALDO_SILHOUETTE_MAP: dict[str, str] = {
    "HEELED SHOES":  "Pumps",            "PUMPS":       "Pumps",
    "FLAT SHOES":    "Slip On",          "SLIP ON":     "Slip On",
    "BOOTS":         "Boot",             "BOOT":        "Boot",
    "CHELSEA":       "Chelsea",          "SANDAL":      "Sandal",
    "SANDALS":       "Sandal",           "STRAP SANDAL":"Strap Sandal",
    "MULES":         "Mules",            "MULE":        "Mules",
    "SNEAKER":       "Sneaker Low Top",  "SNEAKERS":    "Sneaker Low Top",
    "HIGH TOP":      "Sneaker High Top", "LOW TOP":     "Sneaker Low Top",
    "LOAFER":        "Loafer",           "LOAFERS":     "Loafer",
    "BALLERINA":     "Ballerina",        "BALLET":      "Ballerina",
    "MARY JANE":     "Mary Jane",        "MARYJANE":    "Mary Jane",
    "MOCCASIN":      "Moccasins",        "MOCCASINS":   "Moccasins",
    "CLOG":          "Clog",             "CLOGS":       "Clog",
    "FLIP FLOP":     "Flip Flop",        "FLIP-FLOP":   "Flip Flop",
    "SLIDES":        "Slides",           "SLIDE":       "Slides",
    "OPEN TOE":      "Open Toe Pumps",   "THONG":       "Thong",
    "BACKSTRAP":     "Backstrap",        "FISHERMAN":   "Fisherman",
    "TOE RING":      "Toe Ring",         "SPIKE":       "Spike",
    "SPIKELESS":     "Spikeless",        "MONK":        "Monk",
}

# Size fallbacks for sizes NOT in MDD (S, M, L, XL, XS, OS, S/M, M/L, etc.)
_SIZE_FALLBACKS: dict[str, str] = {
    "XS":  "00X", "S":   "00S", "M":   "00M", "L":   "00L",
    "XL":  "00X", "XXL": "00X", "XXXL":"00X",
    "OS":  "000", "ONE SIZE": "000", "0": "000",
    "S/M": "00S", "M/L": "00M", "L/XL": "00L",
    "2T":  "002", "3T": "003", "4T": "004",
    "10C": "10C", "C":  "00C", "F":  "00F", "H":  "00H",
}

# Size desc fallbacks for codes NOT covered by MDD
_SIZE_DESC_FALLBACKS: dict[str, str] = {
    "00X": "XS",  "00S": "S",   "00M": "M",   "00L": "L",
    "00C": "C",   "00F": "F",   "00H": "H",   "000": "OS",
    "10C": "10C",
}

# ─────────────────────────────────────────────────────────────────
# RUNTIME LOV DICTS — MDD xlsx 
# ─────────────────────────────────────────────────────────────────
LOV_COUNTRY:             dict[str, str] = dict(_FALLBACK_LOV_COUNTRY)
LOV_COMPANY_CODE:        dict[str, str] = dict(_FALLBACK_LOV_COMPANY_CODE)
TRENZASHOP_COMP_TO_STIBO: dict[str, str] = dict(_FALLBACK_TRENZASHOP_COMP_TO_STIBO)
LOV_SBU:                 dict[str, str] = dict(_FALLBACK_LOV_SBU)
LOV_CURRENCY_MAP:        dict[str, str] = dict(_FALLBACK_LOV_CURRENCY_MAP)
ALDO_SILHOUETTE_MAP:     dict[str, str] = dict(_FALLBACK_ALDO_SILHOUETTE_MAP)

# MDD_SIZE_LOV: OOR size string -> Stibo AT_Size LOV ID
# MDD_SIZE_DESC: Stibo LOV ID -> human-readable desc (for AT_Size text)
MDD_SIZE_LOV:  dict[str, str] = {}
MDD_SIZE_DESC: dict[str, str] = {}

_mdd_loaded = False


# ─────────────────────────────────────────────────────────────────
# MDD LOADER — Load all LOVs dynamically from the MDD Excel file.
# ─────────────────────────────────────────────────────────────────
def _load_mdd_lovs(mdd_path: Path | None = None) -> bool:
    """
    
    Returns True if loaded successfully, False otherwise.
    """
    global LOV_COUNTRY, LOV_COMPANY_CODE, TRENZASHOP_COMP_TO_STIBO
    global LOV_SBU, LOV_CURRENCY_MAP, ALDO_SILHOUETTE_MAP
    global MDD_SIZE_LOV, MDD_SIZE_DESC, _mdd_loaded

    if _mdd_loaded:
        return True

    path = Path(mdd_path) if mdd_path else MDD_PATH
    if not path.exists():
        log.warning("[MDD] File not found: %s — using fallback dicts", path)
        _build_fallback_size_lovs()
        _mdd_loaded = True
        return False

    log.info("[MDD] Loading LOVs from: %s", path.name)
    try:
        xl = pd.ExcelFile(str(path))
        _load_country_origin_lov(xl)
        _load_company_code_lov(xl)
        _load_sbu_lov(xl)
        _load_currency_lov(xl)
        _load_silhouette_lov(xl)
        _load_size_lovs(xl)
        _mdd_loaded = True
        log.info("[MDD] All LOVs loaded successfully from MDD.")
        return True
    except Exception as exc:
        log.error("[MDD] Failed to load: %s — using fallback dicts", exc)
        _build_fallback_size_lovs()
        _mdd_loaded = True
        return False


def _clean_cell(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("none", "nan", "") else s


def _load_country_origin_lov(xl: pd.ExcelFile):
    """Sheet: 'Country Origin LOV' — Code -> Country Origin Name"""
    global LOV_COUNTRY
    df = pd.read_excel(xl, sheet_name="Country Origin LOV", dtype=str).fillna("")
    loaded: dict[str, str] = {}
    for _, row in df.iterrows():
        code = _clean_cell(row.get("Code", "")).upper()
        name = _clean_cell(row.get("Country Origin Name", ""))
        if code and name and code not in loaded:
            loaded[code] = name
    LOV_COUNTRY = loaded
    log.info("[MDD] LOV_COUNTRY loaded: %d entries", len(LOV_COUNTRY))


def _load_company_code_lov(xl: pd.ExcelFile):
    """Sheet: 'Company Code LOV' — Value ID of LOV -> Values of LOV"""
    global LOV_COMPANY_CODE, TRENZASHOP_COMP_TO_STIBO
    df = pd.read_excel(xl, sheet_name="Company Code LOV", dtype=str).fillna("")
    loaded: dict[str, str] = {}
    comp_map: dict[str, str] = {}
    for _, row in df.iterrows():
        id_raw = _clean_cell(row.get("Value ID of LOV", ""))
        label  = _clean_cell(row.get("Values of LOV", ""))
        if not id_raw or not label:
            continue
        loaded[id_raw] = label
        # Index both zero-padded 4-digit and plain numeric forms
        try:
            n = int(id_raw)
            loaded[str(n)] = label
            padded = f"{n:04d}"
            loaded[padded] = label
            # TRENZASHOP map: both raw and padded -> padded
            comp_map[str(n)]  = padded
            comp_map[padded]  = padded
            comp_map[id_raw]  = padded
        except ValueError:
            comp_map[id_raw] = id_raw
    LOV_COMPANY_CODE = loaded
    TRENZASHOP_COMP_TO_STIBO = comp_map
    log.info("[MDD] LOV_COMPANY_CODE loaded: %d entries", len(LOV_COMPANY_CODE))


def _load_sbu_lov(xl: pd.ExcelFile):
    """Sheet: 'SBU LOV' — Value ID of LOV -> Values of LOV"""
    global LOV_SBU
    df = pd.read_excel(xl, sheet_name="SBU LOV", dtype=str).fillna("")
    loaded: dict[str, str] = {}
    for _, row in df.iterrows():
        code  = _clean_cell(row.get("Value ID of LOV", "")).upper()
        label = _clean_cell(row.get("Values of LOV", ""))
        if code and label:
            loaded[code] = label
    LOV_SBU = loaded
    log.info("[MDD] LOV_SBU loaded: %d entries", len(LOV_SBU))


def _load_currency_lov(xl: pd.ExcelFile):
    """Sheet: 'Retail Price Currency LOV' — full name -> ISO code"""
    global LOV_CURRENCY_MAP
    df = pd.read_excel(xl, sheet_name="Retail Price Currency LOV", dtype=str).fillna("")
    loaded: dict[str, str] = {}
    for _, row in df.iterrows():
        name_raw = _clean_cell(row.iloc[0])
        iso      = _clean_cell(row.iloc[1]).upper()
        if not iso or iso.startswith("VALUE ID"):
            continue
        if name_raw.upper() == "VALUES OF LOV":
            continue
        if name_raw:
            loaded[name_raw.upper()] = iso
        if iso:
            loaded[iso] = iso  # passthrough e.g. "IDR" -> "IDR"

    
    _aliases = {
        "US DOLLAR": "USD", "US DOLLARS": "USD", "DOLLAR": "USD",
        "RUPIAH": "IDR", "RINGGIT": "MYR", "POUND STERLING": "GBP",
        "STERLING": "GBP", "EURO": "EUR", "YEN": "JPY", "BAHT": "THB",
        "DONG": "VND", "YUAN": "CNY", "RENMINBI": "CNY",
        "RIEL": "KHR", "CAMBODIAN RIEL": "KHR",
        "RUPEE": "INR", "INDIAN RUPEE": "INR",
        "FRANC": "CHF", "RIYAL": "SAR", "RAND": "ZAR",
        "WON": "KRW", "PESO": "MXN",
    }
    for k, v in _aliases.items():
        loaded.setdefault(k, v)
    LOV_CURRENCY_MAP = loaded
    log.info("[MDD] LOV_CURRENCY_MAP loaded: %d entries", len(LOV_CURRENCY_MAP))


def _load_silhouette_lov(xl: pd.ExcelFile):
    """Sheet: 'Silhouette LOV' — builds OOR raw -> MDD canonical map"""
    global ALDO_SILHOUETTE_MAP
    df = pd.read_excel(xl, sheet_name="Silhouette LOV", dtype=str).fillna("")
    canonical_names: set[str] = set()
    for _, row in df.iterrows():
        name = _clean_cell(row.get("Silhouette", ""))
        if name:
            canonical_names.add(name)

    # OOR raw strings -> canonical (using MDD canonical names as targets)
    _synonyms: dict[str, str] = {
        "HEELED SHOES":  "Pumps",            "PUMPS":       "Pumps",
        "FLAT SHOES":    "Slip On",          "SLIP ON":     "Slip On",
        "BOOTS":         "Boot",             "BOOT":        "Boot",
        "CHELSEA":       "Chelsea",          "SANDAL":      "Sandal",
        "SANDALS":       "Sandal",           "STRAP SANDAL":"Strap Sandal",
        "MULES":         "Mules",            "MULE":        "Mules",
        "SNEAKER":       "Sneaker Low Top",  "SNEAKERS":    "Sneaker Low Top",
        "HIGH TOP":      "Sneaker High Top", "LOW TOP":     "Sneaker Low Top",
        "LOAFER":        "Loafer",           "LOAFERS":     "Loafer",
        "BALLERINA":     "Ballerina",        "BALLET":      "Ballerina",
        "MARY JANE":     "Mary Jane",        "MARYJANE":    "Mary Jane",
        "MOCCASIN":      "Moccasins",        "MOCCASINS":   "Moccasins",
        "CLOG":          "Clog",             "CLOGS":       "Clog",
        "FLIP FLOP":     "Flip Flop",        "FLIP-FLOP":   "Flip Flop",
        "SLIDES":        "Slides",           "SLIDE":       "Slides",
        "OPEN TOE":      "Open Toe Pumps",   "THONG":       "Thong",
        "BACKSTRAP":     "Backstrap",        "FISHERMAN":   "Fisherman",
        "TOE RING":      "Toe Ring",         "SPIKE":       "Spike",
        "SPIKELESS":     "Spikeless",        "MONK":        "Monk",
    }
    loaded: dict[str, str] = {}
    for raw, target in _synonyms.items():
        
        if target in canonical_names:
            loaded[raw] = target
        else:
            loaded[raw] = target  # keep anyway even if not in MDD this version

    # Self-map canonical names (case-insensitive lookup support)
    for name in canonical_names:
        loaded[name.upper()] = name

    ALDO_SILHOUETTE_MAP = loaded
    log.info("[MDD] ALDO_SILHOUETTE_MAP loaded: %d entries", len(ALDO_SILHOUETTE_MAP))


def _load_size_lovs(xl: pd.ExcelFile):
    """
    Sheet: 'Size Code LOV' — builds MDD_SIZE_LOV and MDD_SIZE_DESC.
    Column group 3: Size Code.2 = Stibo LOV ID, Size Code Description.2 = human desc.
    e.g.  '36' -> '36',  '06H' -> '6.5',  '000' -> 'NO SIZE'
    """
    global MDD_SIZE_LOV, MDD_SIZE_DESC
    df = pd.read_excel(xl, sheet_name="Size Code LOV", dtype=str).fillna("")

    size_code: dict[str, str] = {}  # human_desc -> Stibo_ID
    size_desc: dict[str, str] = {}  # Stibo_ID   -> human_desc

    for _, row in df.iterrows():
        code = _clean_cell(row.get("Size Code", ""))
        desc = _clean_cell(row.get("Size Code Description", ""))
        if not code or not desc:
            continue
        size_desc[code] = desc                    # '06H' -> '6.5'
        size_code.setdefault(desc, code)          # '6.5' -> '06H' (first wins)
        size_code.setdefault(desc.upper(), code)  # '6.5'.upper() -> '06H'

    # Merge fallback (S, M, L, XL, XS, OS, S/M, M/L etc.)
    for k, v in _SIZE_FALLBACKS.items():
        size_code.setdefault(k, v)
        size_code.setdefault(k.upper(), v)

    
    for k, v in _SIZE_DESC_FALLBACKS.items():
        size_desc.setdefault(k, v)

    MDD_SIZE_LOV  = size_code
    MDD_SIZE_DESC = size_desc
    log.info("[MDD] MDD_SIZE_LOV loaded: %d entries, MDD_SIZE_DESC: %d entries",
             len(MDD_SIZE_LOV), len(MDD_SIZE_DESC))


def _build_fallback_size_lovs():
    
    global MDD_SIZE_LOV, MDD_SIZE_DESC

    size_code: dict[str, str] = {}
    size_desc: dict[str, str] = {}

    # Whole numbers 1-150 -> zero-padded 3-digit (e.g. '36' -> '036') [original logic]
    for n in range(1, 151):
        s = str(n)
        padded = f"{n:03d}"
        size_code[s] = s
        size_desc[padded] = s

    # Half sizes [original oor.py fallback]
    _halves = {
        "1.5":"01H","2.5":"02H","3.5":"03H","4.5":"04H",
        "5.5":"00S","6.5":"01-","7.5":"01X","8.5":"02-",
        "9.5":"02X","10.5":"01C","11.5":"01K",
    }
    for desc, code in _halves.items():
        size_code[desc] = code
        size_desc[code] = desc

    # Apparel / special
    size_code.update(_SIZE_FALLBACKS)
    size_desc.update(_SIZE_DESC_FALLBACKS)

    MDD_SIZE_LOV  = size_code
    MDD_SIZE_DESC = size_desc
    log.info("[MDD] Fallback MDD_SIZE_LOV built: %d entries", len(MDD_SIZE_LOV))


_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "#DIV/0!", "N/A", "#",
                       "Not assigned", "not assigned") else s


def _val(parent, attr_id: str, value: str = "", id_val: str = "") -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip()  if value  else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent, attr_id: str, id_val: str, label: str = "") -> ET.Element:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    if label:
        v.text = str(label)
    return mv


def _normalize_currency(raw: str) -> str:
    if not raw:
        return ""
    cleaned = raw.strip().upper()
    if cleaned in LOV_CURRENCY_MAP:
        return LOV_CURRENCY_MAP[cleaned]
    cleaned = cleaned.rstrip(". ")
    if cleaned in LOV_CURRENCY_MAP:
        return LOV_CURRENCY_MAP[cleaned]
    log.warning("[Currency] Unknown currency '%s' — skipping AT_RetailPriceCurrency", raw)
    return ""


def _map_silhouette(raw: str) -> str:
    if not raw:
        return ""
    upper = raw.upper().strip()
    if upper in ALDO_SILHOUETTE_MAP:
        return ALDO_SILHOUETTE_MAP[upper]
    for k, v in ALDO_SILHOUETTE_MAP.items():
        if k in upper:
            return v
    return ""


def _mdd_size_code(size_raw: str) -> str:
    """
    OOR size string -> Stibo AT_Size LOV ID (from MDD).
    Never crashes — falls back to '000' (NO SIZE) with warning.
    """
    s = size_raw.strip()
    # 1. Direct lookup
    if s in MDD_SIZE_LOV:
        return MDD_SIZE_LOV[s]
    # 2. Case-insensitive
    upper = s.upper()
    if upper in MDD_SIZE_LOV:
        return MDD_SIZE_LOV[upper]
    # 3. Float whole-number ("6.0" -> "6")
    try:
        f = float(s)
        if f == int(f):
            whole = str(int(f))
            if whole in MDD_SIZE_LOV:
                return MDD_SIZE_LOV[whole]
    except ValueError:
        pass
    log.warning("[SizeCode] Unknown size '%s' — using '000' (NO SIZE)", size_raw)
    return "000"


def _resolve_season_target(args_season: str) -> str:
    s = args_season.strip().upper()
    m = re.match(r'^([A-Z]{2})(20)?(\d{2})$', s)
    if m:
        return f"{m.group(1)}{m.group(3)}"
    return s


def _season_matcher(target_season: str):
    target = target_season.strip().upper()
    m = re.match(r"^([A-Z]{2})(\d{2})$", target)
    if not m:
        return lambda value: str(value).strip().upper() == target
    season_code, season_yr = m.groups()

    def _matches(value):
        v = str(value).strip().upper()
        if not v:
            return False
        if v == target:
            return True
        return season_code in v and season_yr in v

    return _matches


def _read_oor_sheet(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(str(path))
    sheet = "All" if "All" in xl.sheet_names else xl.sheet_names[0]
    log.info("[OOR] Using sheet: '%s'", sheet)
    return pd.read_excel(xl, sheet_name=sheet, dtype=str).fillna("")


def _parse_filename_metadata(path: Path) -> dict:
    stem = path.stem.upper()
    season_code, season_year = "", ""
    m = re.search(r"(?<![A-Z])([A-Z]{2})(20\d{2}|\d{2})(?![0-9A-Z])", stem)
    if m:
        season_code = m.group(1)
        season_year = m.group(2)
        if len(season_year) == 2:
            season_year = f"20{season_year}"
    article_type = (
        "License" if re.search(r"(?:^|[-_ ])LICEN[SC]E(?:D)?(?:$|[-_ ])", stem)
        else ("Inline" if re.search(r"(?:^|[-_ ])INLINE(?:$|[-_ ])", stem) else "")
    )
    m2 = re.search(r"(?:^|[-_])([A-Z]{2})(?:[-_]\d+)?$", path.stem.upper())
    country = m2.group(1) if m2 else ""
    return {
        "season_code": season_code,
        "season_year": season_year,
        "article_type": article_type,
        "country_code": country,
    }


# ─────────────────────────────────────────────────────────────────
# FIX 22+23: EAN/UPC → DataContainers/DC_Barcode
# ─────────────────────────────────────────────────────────────────
def _add_barcode_datacontainer(parent_el, ean: str, variant_key: str):
    if not ean:
        return
    dc_root = ET.SubElement(parent_el, f"{{{STIBO_NS}}}DataContainers")
    mdc     = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")
    dc      = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc_vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")
    v_bc    = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_bc.set("AttributeID", "AT_Barcode")
    v_bc.text = ean
    v_bt    = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_bt.set("AttributeID", "AT_BarcodeType")
    v_bt.set("ID", "P")
    v_main  = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_main.set("AttributeID", "AT_MainEANIndicator")
    v_main.set("ID", "Y")


# ─────────────────────────────────────────────────────────────────
# CLASSIFICATIONS
# ─────────────────────────────────────────────────────────────────
def _build_classifications(season_code: str, season_year: str):
    cls_root  = ET.Element(f"{{{STIBO_NS}}}Classifications")
    season_id = f"CLH_{BRAND_CODE}_{season_code}{season_year}"
    label_map = {"SS": "Spring Summer", "FW": "Fall Winter"}

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", "CLH_AldoBatches")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = (
        f"{BRAND_NAME} {label_map.get(season_code, season_code)} {season_year}"
    )

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{season_code} {season_year} Confirmed Articles"
    )

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{season_code} {season_year} Unconfirmed Articles"
    )

    return cls_root, season_id


# ─────────────────────────────────────────────────────────────────
# PRODUCT XML BUILDERS
# ─────────────────────────────────────────────────────────────────
def _build_oor_product_xml(art: dict, season_code: str, season_year: str,
                           season_id: str, comp_code: str, sbu: str,
                           article_type: str = "", country_code: str = "") -> str:
    article_no = art["article_no"]
    color_id = art["color_id"]
    key_article = f"{BRAND_CODE}{article_no}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", "PPH_F-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["style_name"] or article_no

    for cls_id, cls_type in [
        ("CLH_AldoArticles", "CPL_Merchandiser"),
        (f"{season_id}UA", "CPL_UnConfirmedForSeason"),
    ]:
        cr = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
        cr.set("ClassificationID", cls_id)
        cr.set("Type", cls_type)

    vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _val(vals, "AT_Brand", BRAND_NAME, id_val=BRAND_CODE)
    _val(vals, "AT_InboundGenericCode", key_article)
    _val(vals, "AT_PrincipalStyleCode", article_no)
    _val(vals, "AT_SAPStyleCode", article_no)
    _val(vals, "AT_PrincipalStyleDescription", art["style_name"])

    if country_code:
        _val(vals, "AT_Country", id_val=country_code)

    coo_code = _s(art.get("coo", ""))
    if coo_code:
        coo_label = LOV_COUNTRY.get(coo_code, coo_code)
        _val(vals, "AT_CountryOrigin", coo_label, id_val=coo_code)

    _val(vals, "AT_Color")
    if color_id:
        _val(vals, "AT_PrincipalColorCode", color_id)
        _val(vals, "AT_PrincipalColorName", art["color_desc"])

    _val(vals, "AT_Silhouette", _map_silhouette(art["silhouette"]))
    _val(vals, "AT_Collection1", art.get("article_collection", ""))
    _val(vals, "AT_MerchandiseCategory", art["merch_category"][:9])
    _val(vals, "AT_OriginalPrice", art.get("us_retail", ""))
    _val(vals, "AT_CurrentPrice", art.get("us_retail", ""))

    curr = _normalize_currency(art.get("currency") or "")
    if curr:
        _val(vals, "AT_RetailPriceCurrency", id_val=curr)

    file_season_code = art.get("file_season_code", "") or season_code
    file_season_year = art.get("file_season_year", "") or season_year
    full_year = f"20{file_season_year}" if len(file_season_year) == 2 else file_season_year
    _val(vals, "AT_Season", id_val=file_season_code)
    _val(vals, "AT_SeasonYear", full_year)
    if article_type:
        _val(vals, "AT_BYArticleType", article_type, id_val=article_type)

    raw_comp = art.get("comp_code") or comp_code
    stibo_comp = TRENZASHOP_COMP_TO_STIBO.get(raw_comp, raw_comp)
    comp_label = LOV_COMPANY_CODE.get(stibo_comp, stibo_comp)
    _multival(vals, "AT_CompanyCode", stibo_comp, comp_label)

    sbu_label = LOV_SBU.get(sbu, sbu)
    _multival(vals, "AT_SBU", sbu, sbu_label)
    _val(vals, "AT_BYIndicator", "Yes", id_val="Y")
    _val(vals, "AT_SAPIndicator", "No", id_val="N")
    _val(vals, "AT_UOM", "Pair", id_val="PAA")

    return ET.tostring(g_el, encoding="unicode")


def _build_oor_sizerun_product_xml(art: dict, sizes: list, season_code: str,
                                   season_year: str, season_id: str,
                                   comp_code: str, sbu: str,
                                   article_type: str = "", country_code: str = "") -> str:
    article_no = art["article_no"]
    color_id = art["color_id"]
    key_article = f"{BRAND_CODE}{article_no}"
    coo_code = (country_code or "").strip().upper()

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", "PPH_F-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["style_name"] or article_no

    for cls_id, cls_type in [
        ("CLH_AldoArticles", "CPL_Merchandiser"),
        (f"{season_id}UA", "CPL_UnConfirmedForSeason"),
    ]:
        cr = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
        cr.set("ClassificationID", cls_id)
        cr.set("Type", cls_type)

    vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _val(vals, "AT_Brand", BRAND_NAME, id_val=BRAND_CODE)
    _val(vals, "AT_InboundGenericCode", key_article)
    _val(vals, "AT_PrincipalStyleCode", article_no)
    _val(vals, "AT_SAPStyleCode", article_no)
    _val(vals, "AT_PrincipalStyleDescription", art["style_name"])
    if coo_code:
        _val(vals, "AT_Country", id_val=coo_code)

    _val(vals, "AT_Color")
    if color_id:
        _val(vals, "AT_PrincipalColorCode", color_id)
        _val(vals, "AT_PrincipalColorName", art["color_desc"])

    _val(vals, "AT_Silhouette", _map_silhouette(art["silhouette"]))
    _val(vals, "AT_Collection1", art.get("article_collection", ""))
    _val(vals, "AT_MerchandiseCategory", art["merch_category"][:9])
    _val(vals, "AT_PrincipalMerchandiseHierarchyL1", art.get("merch_division", ""))
    _val(vals, "AT_OriginalPrice", art.get("us_retail", ""))
    _val(vals, "AT_CurrentPrice", art.get("us_retail", ""))
    _val(vals, "AT_Season", id_val=season_code)
    # _val(vals, "AT_SeasonYear", season_year)
    full_year = f"20{season_year}" if len(season_year) == 2 else season_year
    _val(vals, "AT_SeasonYear", full_year)
    if article_type:
        _val(vals, "AT_BYArticleType", article_type, id_val=article_type)
    # if art.get("plm_colorway_id"):
    #     _val(vals, "AT_PLMColorwayID", art["plm_colorway_id"])

    curr = _normalize_currency(art.get("currency") or "")
    if curr:
        _val(vals, "AT_RetailPriceCurrency", id_val=curr)

    raw_comp = art.get("comp_code") or comp_code
    stibo_comp = TRENZASHOP_COMP_TO_STIBO.get(raw_comp, raw_comp)
    comp_label = LOV_COMPANY_CODE.get(stibo_comp, stibo_comp)
    _multival(vals, "AT_CompanyCode", stibo_comp, comp_label)

    sbu_label = LOV_SBU.get(sbu, sbu)
    _multival(vals, "AT_SBU", sbu, sbu_label)
    _val(vals, "AT_BYIndicator", "Yes", id_val="Y")
    _val(vals, "AT_SAPIndicator", "No", id_val="N")
    _val(vals, "AT_UOM", "Pair", id_val="PAA")

    seen_variants: set[str] = set()
    for size_rec in sizes:
        size_raw = str(size_rec.get("size", "")).strip()
        if not size_raw:
            continue

        size_code = _mdd_size_code(size_raw)
        variant_key = f"{key_article}{color_id}{size_code}"
        if len(variant_key) > 18:
            max_sc = 18 - len(key_article) - len(color_id)
            log.warning(
                "[Variant] Key too long (%d chars): %s — truncating size code to %d chars",
                len(variant_key), variant_key, max_sc,
            )
            size_code = size_code[:max(1, max_sc)]
            variant_key = f"{key_article}{color_id}{size_code}"

        if variant_key in seen_variants:
            continue
        seen_variants.add(variant_key)

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        # v_kv.set("KeyID", "KEY_Variant")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant_key

        # ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = f"Size {size_raw}"

        v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
        # AT_Size: ID = Stibo LOV ID from MDD, text = human-readable desc from MDD
        size_desc_text = MDD_SIZE_DESC.get(size_code, size_raw)
        _val(v_vals, "AT_Size", id_val=size_code)
        # _val(v_vals, "AT_Variant", variant_key)
        _val(v_vals, "AT_InboundVariantCode", variant_key)

        upc = size_rec.get("upc", "")
        if upc:
            # _add_barcode_datacontainer(v_el, upc)
            _add_barcode_datacontainer(v_el, upc, variant_key)


    return ET.tostring(g_el, encoding="unicode")


# ─────────────────────────────────────────────────────────────────
# XML WRITER
# ─────────────────────────────────────────────────────────────────
def _write_xml(out_path: Path, cls_el: ET.Element,
               products_xml_list: list, export_time: str):
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    header = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
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
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(header)
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for pxml in products_xml_list:
            pxml = _XMLNS_RE.sub("", pxml)
            f.write(f"    {pxml}\n")
        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")


# ─────────────────────────────────────────────────────────────────
# OOR LOADERS
# ─────────────────────────────────────────────────────────────────
class OORLoader:
    """WSL001 - OOR Report: one row per article+color, no variants."""

    def __init__(self, path: Path):
        self.path = path
        self.articles: list = []
        self.sizes: dict = {}
        meta = _parse_filename_metadata(path)
        self.file_country_code = meta.get("country_code", "")
        self.file_season_code  = meta.get("season_code",  "")
        self.file_season_year  = meta.get("season_year",  "")
        self._load()

    def _load(self):
        log.info("[OOR] Loading: %s", self.path.name)
        df = _read_oor_sheet(self.path)
        df = df[~df["Product Category Name"].str.lower().apply(
            lambda x: any(k in x for k in SKIP_CATEGORIES)
        )]
        log.info("[OOR] Rows after category filter: %d", len(df))

        seen = set()
        for _, row in df.iterrows():
            art_no = _s(row.get("Generic ID", ""))
            if not art_no:
                continue
            color_id = _s(row.get("Color ID", ""))
            art_key  = f"{art_no}_{color_id}"

            if art_key not in seen:
                seen.add(art_key)
                self.articles.append({
                    "article_no":         art_no,
                    "style_name":         _s(row.get("Style Name", "")),
                    "color_id":           color_id,
                    "color_desc":         _s(row.get("Color Name", "")),
                    "coo":                _s(row.get("COO", "")),
                    "article_collection": _s(row.get("Article Collection", "")),
                    "silhouette":         _s(row.get("Silhouette", "")),
                    "selling_season":     _s(row.get("Selling Season", "")),
                    "us_retail":          _s(row.get("US Sugg. Retail Price", "")),
                    "currency":           _s(row.get("Document currency", "")),
                    "order_qty":          _s(row.get("Order Qty", "")),
                    "open_qty":           _s(row.get("Open Qty", "")),
                    "invoiced_qty":       _s(row.get("Invoiced Qty", "")),
                    "merch_category":     _s(row.get("Merchandise Category Name", "")),
                    "merch_division":     _s(row.get("Merchandise Division Name", "")),
                    "product_category":   _s(row.get("Product Category Name", "")),
                    "comp_code":          _s(row.get("Sold to ID", "")),
                    "milestone":          _s(row.get("Milestone", "")),
                    "plm_colorway_id":    _s(row.get("PLM Colorway ID", "")),
                    "file_season_code":   self.file_season_code,
                    "file_season_year":   self.file_season_year,
                })
                self.sizes[art_key] = []

            size = _s(row.get("Size", ""))
            if size:
                self.sizes[art_key].append({
                    "size":      size,
                    "sku":       _s(row.get("Sku", "")),
                    "upc":       _s(row.get("UPC", "")),
                    "order_qty": _s(row.get("Order Qty", "")),
                    "open_qty":  _s(row.get("Open Qty", "")),
                })


class OORSizeRunLoader:
    """WSL001 - OOR Size Run Report: one row per article+color+size."""

    def __init__(self, path: Path, target_season: str):
        self.path = path
        self.articles: list = []
        self.sizes: dict = {}
        meta = _parse_filename_metadata(path)
        self.file_country_code = meta.get("country_code", "")
        self.file_season_code = meta.get("season_code", "")
        self.file_season_year = meta.get("season_year", "")
        self._load(target_season)

    def _load(self, target_season: str):
        log.info("[OORSizeRun] Loading: %s  (season filter: %s)", self.path.name, target_season)
        df = _read_oor_sheet(self.path)

        before_count = len(df)
        df = df[df["Selling Season"].apply(_season_matcher(target_season))]
        log.info("[OORSizeRun] Rows before: %d, after season filter: %d", before_count, len(df))

        df = df[~df["Product Category Name"].str.lower().apply(
            lambda x: any(k in x for k in SKIP_CATEGORIES)
        )]
        log.info("[OORSizeRun] Rows after category filter: %d", len(df))

        seen: set = set()
        for _, row in df.iterrows():
            art_no = _s(row.get("Generic ID", ""))
            color_id = _s(row.get("Color ID", ""))
            if not art_no:
                continue

            art_key = f"{art_no}_{color_id}"
            size = _s(row.get("Size", ""))
            sku = _s(row.get("Sku", ""))
            upc = _s(row.get("UPC", ""))
            order_qty = _s(row.get("Order Qty", ""))
            open_qty = _s(row.get("Open Qty", ""))

            if art_key not in seen:
                seen.add(art_key)
                self.articles.append({
                    "article_no":         art_no,
                    "style_name":         _s(row.get("Style Name", "")),
                    "color_id":           color_id,
                    "color_desc":         _s(row.get("Color Name", "")),
                    "coo":                _s(row.get("COO", "")),
                    "article_collection": _s(row.get("Article Collection", "")),
                    "merch_category":     _s(row.get("Merchandise Category Name", "")),
                    "merch_division":     _s(row.get("Merchandise Division Name", "")),
                    "product_category":   _s(row.get("Product Category Name", "")),
                    "silhouette":         _s(row.get("Silhouette", "")),
                    "selling_season":     _s(row.get("Selling Season", "")),
                    "us_retail":          _s(row.get("US Sugg. Retail Price", "")),
                    "eu_retail":          _s(row.get("EU Sugg. Retail Price", "")),
                    "gb_retail":          _s(row.get("GB Sugg. Retail Price", "")),
                    "currency":           _s(row.get("Document currency", "")),
                    "plm_colorway_id":    _s(row.get("PLM Colorway ID", "")),
                    "comp_code":          _s(row.get("Sold to ID", "")),
                    "file_season_code":   self.file_season_code,
                    "file_season_year":   self.file_season_year,
                })
                self.sizes[art_key] = []

            if size:
                self.sizes[art_key].append({
                    "size":      size,
                    "sku":       sku,
                    "upc":       upc,
                    "order_qty": order_qty,
                    "open_qty":  open_qty,
                })

        log.info("[OORSizeRun] %d unique (article+color) loaded", len(self.articles))


# ─────────────────────────────────────────────────────────────────
# SUMMARY PRINTER
# ─────────────────────────────────────────────────────────────────
def _print_summary(source_file, total_arts, written, season_code,
                   season_year, comp_code, sbu, out_name, kb):
    print("═══ ALDO OOR SIZERUN SUMMARY ══════════════════════════════════", flush=True)
    print(f"  Source file      : {source_file}",               flush=True)
    print(f"  Total articles   : {total_arts}",                flush=True)
    print(f"  Written to XML   : {written}",                   flush=True)
    print(f"  Season           : {season_code} {season_year}", flush=True)
    print(f"  comp_code        : {comp_code}",                 flush=True)
    print(f"  sbu              : {sbu}",                       flush=True)
    print(f"  XML output       : {out_name}",                  flush=True)
    print(f"  File size        : {kb} KB",                     flush=True)
    print("════════════════════════════════════════════════════", flush=True)


# ─────────────────────────────────────────────────────────────────
# MAIN run() — called from Aldo Lambda
# ─────────────────────────────────────────────────────────────────
def run(args, xml_out_dir: Path, export_time: str,
        season_code: str, season_year: str, season_id: str,
        source_dir: Path | None = None, auditor=None, article_filter=None,
        source_type: str = "oor_size", oor_dir: Path | None = None,
        oor_size_dir: Path | None = None, test_mode: bool = False,
        mdd_path: Path | None = None):
    """
    Process Aldo OOR sources. source_type controls dispatch:
      - "oor":      generic OOR report, no variants
      - "oor_size": OOR size-run report, with variants

    mdd_path: override MDD file location (optional; falls back to MDD_PATH env/constant).
    Returns number of XMLs written (0 or 1).
    """
    # ── MDD load (once per process) ──────────────────────────────
    _load_mdd_lovs(mdd_path)

    source_type = (source_type or "oor_size").strip().lower()
    if source_type not in {"oor", "oor_size"}:
        log.warning("[OOR] Unsupported source_type=%s — skipping", source_type)
        return 0

    if source_type == "oor_size":
        input_dir = oor_size_dir or source_dir
    else:
        input_dir = oor_dir or source_dir

    if input_dir is None:
        log.info("[OOR] No input directory supplied for source_type=%s — skipping", source_type)
        return 0

    source_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xlsm"))
    if not source_files:
        log.info("[OOR] No file found in %s — skipping", input_dir)
        return 0

    preferred = None
    if source_type == "oor_size":
        preferred = next((p for p in source_files if "size" in p.stem.lower()), None)
    elif source_type == "oor":
        preferred = next((p for p in source_files if "size" not in p.stem.lower()), None)
    source_path = preferred or source_files[0]
    log.info("[OOR] source_type=%s selected_file=%s", source_type, source_path.name)

    meta = _parse_filename_metadata(source_path)
    season_code = season_code or meta.get("season_code", "")
    season_year = season_year or meta.get("season_year", "")
    if len(season_year) == 2:
        season_year = f"20{season_year}"
    season_id = f"CLH_{BRAND_CODE}_{season_code}{season_year}"

    if source_type == "oor":
        loader_label = "OOR"
        loader = OORLoader(source_path)
        build_products = lambda: [
            _build_oor_sizerun_product_xml(
                art,
                loader.sizes.get(f"{art['article_no']}_{art['color_id']}", []),
                season_code, season_year, season_id,
                args.comp_code, args.sbu,
                meta.get("article_type", ""),
                meta.get("country_code", ""),
            )
            for art in loader.articles
        ]
    else:
        loader_label = "OORSizeRun"
        target_season = getattr(args, "target_season", None) or f"{season_code}{season_year[-2:]}"
        loader = OORSizeRunLoader(source_path, target_season)
        build_products = lambda: [
            _build_oor_sizerun_product_xml(
                art,
                loader.sizes.get(f"{art['article_no']}_{art['color_id']}", []),
                season_code, season_year, season_id,
                args.comp_code, args.sbu,
                meta.get("article_type", ""),
                meta.get("country_code", ""),
            )
            for art in loader.articles
        ]

    if article_filter:
        before = len(loader.articles)
        loader.articles = [
            a for a in loader.articles if a["article_no"] == article_filter
        ]
        log.info("[%s] article_filter=%s -> %d/%d kept",
                 loader_label, article_filter, len(loader.articles), before)
    # Test
    # loader.articles = loader.articles[:5]
    # log.info("[%s] Capped at %d articles", loader_label, len(loader.articles))

    if not loader.articles:
        log.warning("[%s] No articles after filtering — skipping XML", loader_label)
        return 0

    variants_written = sum(
        len(loader.sizes.get(f"{a['article_no']}_{a['color_id']}", []))
        for a in loader.articles
    )
    barcode_count = sum(
        1
        for a in loader.articles
        for s in loader.sizes.get(f"{a['article_no']}_{a['color_id']}", [])
        if _s(s.get("upc", ""))
    )
    log.info("[%s] variant_count=%d barcode_count=%d", loader_label, variants_written, barcode_count)

    cls_el, _ = _build_classifications(season_code, season_year)
    out_name = f"{source_path.stem}.xml"
    out_path = xml_out_dir / out_name
    products = build_products()

    _write_xml(out_path, cls_el, products, export_time)
    kb = out_path.stat().st_size // 1024

    log.info("OOR XML generated by %s -> %s (%dKB)", loader_label, out_name, kb)
    _print_summary(source_path.name, len(loader.articles), len(loader.articles),
                   season_code, season_year, args.comp_code, args.sbu, out_name, kb)

    if auditor:
        audit_kwargs = {
            "status": "ok",
            "articles": len(loader.articles),
            "filename": source_path.name,
        }
        if source_type == "oor_size":
            audit_kwargs["variants"] = variants_written
        auditor.record_loader(loader_label, **audit_kwargs)
    return 1