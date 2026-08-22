"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — 2XU Order Confirmation          ║
║   order_confirmation_main.py                                    ║
║                                                                  ║
║   Source file  : Order Confirmation (.xlsx)                      ║
║   Sheet        : First sheet (Rep*.tmp or Sheet1)               ║
║                                                                  ║
║   File Layout (Row 1 = header, Row 2+ = data):                   ║
║     DISTRIBUTOR CODE | DISTRIBUTOR NAME | STATUS |              ║
║     ORDER NUMBER | ORDER DESCRIPTION | FOLIO |                  ║
║     BARCODE/EAN | PRODUCT CODE | CONCAT | PRODUCT NAME |        ║
║     COLOUR CODE | COLOUR NAME | SIZE CODE | ORDER QTY |         ║
║     MSRP USD | DISCOUNT % | ACTIVE STATUS | PO NUMBER |         ║
║     FACTORY NAME | SO SKU B2B DUE AMOUNT | COO |                ║
║     HS CODE | FABRIC COMP                                        ║
║                                                                  ║
║   Key Mapping Rules (from 2XU Stibo Mapping + Attributes v6):   ║
║     Generic    : 2XU + 7-digit style code (upper) +             ║
║                  2-char colour mapping                           ║
║                  e.g. 2XUWR7506ABS                              ║
║     Variant    : Generic + 3-char SAP Color + 3-char SAP Size   ║
║                  e.g. 2XUWR7506ABS0050XS                        ║
║     SAP Style  : 7-digit style + 2-char colour mapping          ║
║     Gender     : derived from product code prefix               ║
║                  MA/MR/MT = Male, WA/WR/WT/WQ = Female,         ║
║                  UA/UR/UQ = Unisex                              ║
║     Season     : from ORDER DESCRIPTION e.g. H226CN1 → H2/2026  ║
║     COO        : from COO column → ISO code via MDD             ║
╚══════════════════════════════════════════════════════════════════╝
"""

from __future__ import annotations

import logging
import math
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import openpyxl

# ══════════════════════════════════════════════════════════════════════════════
# STIBO XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')


def strip_xmlns(xml_str: str) -> str:
    return _XMLNS_RE.sub("", xml_str)


def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if id_val:
        clean = str(id_val).strip()
        if clean and clean not in ("None", "nan"):
            el.set("ID", clean)
    if value:
        clean = str(value).strip()
        if clean and clean not in ("None", "nan"):
            el.text = clean
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = "") -> ET.Element:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v  = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    if label:
        v.text = str(label)
    return mv


def _keyval(parent: ET.Element, key_id: str, value: str) -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}KeyValue")
    el.set("KeyID", key_id)
    el.text = value
    return el


def _clf(parent: ET.Element, class_id: str, ref_type: str) -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}ClassificationReference")
    el.set("ClassificationID", class_id)
    el.set("Type", ref_type)
    return el


def open_step_xml(f, export_time: str) -> None:
    f.write(
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


def close_step_xml(f) -> None:
    f.write("</STEP-ProductInformation>\n")


def build_classifications(brand_name: str, brand_code: str, season_code: str) -> ET.Element:
    """Build Season / Confirmed / Unconfirmed Classifications block."""
    sea_prefix, sea_year = derive_season(season_code)
    sid         = f"CLH_{brand_code}_{sea_prefix}{sea_year}"
    label_map   = {
        "H1": "Spring Summer", "H2": "Fall Winter",
        "SS": "Spring Summer", "FW": "Fall Winter",
        "SP": "Spring",        "SM": "Summer",
        "FL": "Fall",          "WN": "Winter",
        "AL": "All Season",    "CO": "Core",
    }
    season_label = label_map.get(sea_prefix, sea_prefix)
    short        = f"{sea_prefix} {sea_year}"

    root       = ET.Element(f"{{{STIBO_NS}}}Classifications")
    season_cls = ET.SubElement(root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         sid)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   f"CLH_{brand_name.replace(' ', '')}Batches")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = (
        f"{brand_name} {season_label} {sea_year}"
    )

    for suffix, type_id, label in (
        ("CA", "CLS_ConfirmedArticles",   f"{short} Confirmed Articles"),
        ("UA", "CLS_UnconfirmedArticles", f"{short} Unconfirmed Articles"),
    ):
        sub = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        sub.set("ID",         f"{sid}{suffix}")
        sub.set("UserTypeID", type_id)
        ET.SubElement(sub, f"{{{STIBO_NS}}}Name").text = label

    return root


# ══════════════════════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
ORDER_DIR    = INPUT_DIR / "orderconfirmation"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [ORDER_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("twoxu.order_confirmation_main")
if not log.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

# Product code prefix → (sap_gender_code, gender_label, gender_raw)
_GENDER_PREFIX_MAP: dict[str, tuple[str, str, str]] = {
    "MA": ("M", "Male",   "Mens"),
    "MR": ("M", "Male",   "Mens"),
    "MT": ("M", "Male",   "Mens"),
    "WA": ("F", "Female", "Womens"),
    "WR": ("F", "Female", "Womens"),
    "WT": ("F", "Female", "Womens"),
    "WQ": ("F", "Female", "Womens"),
    "UA": ("U", "Unisex", "Unisex"),
    "UR": ("U", "Unisex", "Unisex"),
    "UQ": ("U", "Unisex", "Unisex"),
}

# Gender char used in Generic Description: (A/M)=Mens, (A/W)=Womens, (A/U)=Unisex
_DESC_GENDER_CHAR: dict[str, str] = {"M": "M", "F": "W", "U": "U"}

# ── Stibo LOV fallback maps (used when MDD sheet missing/incomplete) ──
# Source: 2XU import errors — these are the ONLY valid LOV IDs accepted.
_SAP_AGE_FALLBACK: dict[str, str] = {
    "ADULTS": "AD", "ADULT": "AD", "AD": "AD", "A": "AD",
    "ALL AGES": "AA", "AA": "AA",
    "CHILDREN": "CH", "CHILD": "CH", "KIDS": "CH", "CH": "CH", "K": "CH",
}
_BY_AGE_FALLBACK: dict[str, str] = {
    "ADULT": "ADULT", "ADULTS": "ADULT", "AD": "ADULT", "A": "ADULT",
    "ALL AGES": "ALL AGES",
    "GRADE SCHOOL": "GRADE SCHOOL", "GS": "GRADE SCHOOL",
    "INFANT": "INFANT", "IN": "INFANT",
    "KIDS": "KIDS", "K": "KIDS", "CHILD": "KIDS", "CHILDREN": "KIDS",
    "PRESCHOOL": "PRESCHOOL", "PS": "PRESCHOOL",
}
# 2XU sends H1/H2 — map to Stibo Season LOV IDs.
_SEASON_FALLBACK: dict[str, str] = {
    "H1": "SS", "H2": "FW",
    "SPRING": "SP", "SUMMER": "SM", "FALL": "FL", "WINTER": "WN",
    "SPRING SUMMER": "SS", "SPRINGSUMMER": "SS",
    "FALL WINTER": "FW", "FALLWINTER": "FW",
    "ALL SEASON": "AL", "CORE": "CO",
    "SS": "SS", "FW": "FW", "SP": "SP", "SM": "SM",
    "FL": "FL", "WN": "WN", "AL": "AL", "CO": "CO",
}
# AT_Country valid LOV (per Stibo): only SEA destinations.
_COUNTRY_FALLBACK: dict[str, str] = {
    "KH": "KH", "CAMBODIA": "KH",
    "ID": "ID", "INDONESIA": "ID",
    "MY": "MY", "MALAYSIA": "MY",
    "PH": "PH", "PHILIPPINES": "PH",
    "SG": "SG", "SINGAPORE": "SG",
    "TH": "TH", "THAILAND": "TH",
    "VN": "VN", "VIETNAM": "VN",
}
# AT_FOBCurrency — common subset; full list comes from MDD Currency LOV.
_CURRENCY_FALLBACK: dict[str, str] = {
    "USD": "USD", "AUD": "AUD", "EUR": "EUR", "GBP": "GBP", "SGD": "SGD",
    "IDR": "IDR", "THB": "THB", "VND": "VND", "MYR": "MYR", "PHP": "PHP",
    "KHR": "KHR", "CNY": "CNY", "RMB": "RMB", "HKD": "HKD", "JPY": "JPY",
    "INR": "INR", "KRW": "KRW", "TWD": "TWD", "NZD": "NZD",
}

# COO name → ISO-2 code (augmented at runtime from MDD)
_COO_FALLBACK: dict[str, str] = {
    "CHINA":        "CN",
    "VIETNAM":      "VN",
    "INDONESIA":    "ID",
    "TAIWAN":       "TW",
    "SRI LANKA":    "LK",
    "HONG KONG":    "HK",
    "MYANMAR":      "MM",
    "BANGLADESH":   "BD",
    "THAILAND":     "TH",
    "MALAYSIA":     "MY",
    "CAMBODIA":     "KH",
    "PHILIPPINES":  "PH",
    "SINGAPORE":    "SG",
}

# Dominant colour (part1) → SAP 3-char Color code.
# Used when full PrincipalColourCode (e.g. BLK/GRY) is missing from
# the Colour Mapping sheet — better than emitting '000' (invalid).
_COLOUR_PART1_SAP_FALLBACK: dict[str, str] = {
    "BLK": "005", "BLA": "005", "BK":  "005",
    "WHT": "W",   "WHI": "W",
    "GRY": "LGY", "GRA": "LGY", "GRE": "LGY",
    "RED": "R",
    "BLU": "B",   "NAV": "NAV", "NVY": "NAV",
    "GRN": "G",
    "YEL": "Y",
    "ORG": "ORG",
    "PNK": "PNK", "PIN": "PNK",
    "PUR": "PUR", "PRP": "PUR",
    "BRN": "BRN", "BRO": "BRN",
    "SLV": "SLV", "SIL": "SLV", "SRF": "SLV",
    "GLD": "GLD", "GOL": "GLD",
}

# Product code prefix → SAP Mapping Range (used when Order Conf missing RANGE col).
_PRODUCT_CODE_RANGE_MAP: dict[str, str] = {
    "MA": "Apparel",     "WA": "Apparel",     "UA": "Apparel",
    "MR": "Compression", "WR": "Compression", "UR": "Compression",
    "MT": "Apparel",     "WT": "Apparel",
    "WQ": "Accessories", "UQ": "Accessories",
}


def derive_sap_lookup_keys(product_code: str, product_name: str) -> tuple[str, str, str]:
    """
    Derive (range, field_of_play, category) when the Order Confirmation file
    does not contain RANGE / FIELD OF PLAY / CATEGORY columns.
      - Range  : from product code prefix (MA/WA/UA → Apparel, etc.)
      - FoP    : default 'Run' (most 2XU SKUs are run/compression)
      - Cat    : '' — lookup will degrade gracefully via fallback if no match
    """
    prefix = product_code[:2].upper() if len(product_code) >= 2 else ""
    range_val = _PRODUCT_CODE_RANGE_MAP.get(prefix, "Apparel")
    fop       = "Run"
    cat       = ""
    return range_val, fop, cat

# SAP Mapping: (Range, Field Of Play, Category) →
#   (sap_div, sap_group, sap_cat, sports_cat)
# Loaded at runtime from 2XU Stibo Mapping sheet (7 columns, NO Gender).
# Key is 3-tuple matching actual sheet layout.
_SAP_MAP_FALLBACK: dict[tuple, tuple] = {
    ("Apparel",     "Run",        "(a) Jacket"):       ("Apparel",     "Sports Running Apparel",  "Jacket",              "Running"),
    ("Apparel",     "Sportswear", "(a) Jacket"):       ("Apparel",     "Sports Training Apparel", "Jacket",              "Fitness"),
    ("Apparel",     "Run",        "(a) Short Sleev"):  ("Apparel",     "Sports Running Apparel",  "Tshirt short Sleeve", "Running"),
    ("Compression", "Run",        "(b) Comp Tights"):  ("Apparel",     "Compression",             "Tights",              "Running"),
    ("Accessories", "Run",        "(e) Socks"):        ("Accessories", "Socks",                   "Running",             "Running"),
    # Fallbacks by Range only — used when Category column missing
    ("Apparel",     "Run",        ""):                 ("Apparel",     "Sports Running Apparel",  "",                    "Running"),
    ("Compression", "Run",        ""):                 ("Apparel",     "Compression",             "",                    "Running"),
    ("Accessories", "Run",        ""):                 ("Accessories", "Sports Running Accessories", "",                 "Running"),
}


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class SizeLine:
    """One data row = one EAN + one size."""
    ean:        str
    size_code:  str   # raw from file e.g. "XS", "M", "L", "OSFA"
    order_qty:  Optional[int]
    row_num:    int


@dataclass
class Article:
    """
    Grouped by (product_code, colour_code) = one PRD_Variant (colour variant).
    Multiple SizeLines = one per EAN/size row.
    """
    distributor_code: str
    order_number:     str
    order_desc:       str          # e.g. "H226CN1"
    product_code:     str          # e.g. "WR7575a"
    product_name:     str          # e.g. "Aero Jacket"
    colour_code:      str          # e.g. "BLK/SRF"
    colour_name:      str          # e.g. "Black/Silver Reflective"
    msrp_usd:         Optional[float]
    coo_raw:          str          # e.g. "TAIWAN"
    hs_code:          str
    fabric_comp:      str
    sizes:            list[SizeLine] = field(default_factory=list)
    # Optional SAP Mapping lookup keys (sourced from order confirmation columns if present)
    range_val:        str = ""
    field_of_play:    str = ""
    category:         str = ""


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _num(val) -> Optional[float]:
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _alnum(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def derive_season(raw: str) -> tuple[str, str]:
    """
    Parse season from ORDER DESCRIPTION or args.season.
    Handles: 'H226CN1' → ('H2', '2026')
             'H226'    → ('H2', '2026')
             'H126'    → ('H1', '2026')
             'H2-2026' → ('H2', '2026')
             'SM26'    → ('SM', '2026')
             'FW2026'  → ('FW', '2026')
             'SS26'    → ('SS', '2026')
    """
    s = (raw or "").strip().upper()

    # Try explicit H1/H2 + 4-digit year e.g. "H2-2026"
    m = re.match(r'(H[12])[^0-9]*(\d{4})', s)
    if m:
        return m.group(1), m.group(2)

    # Try H1/H2 + 2-digit year e.g. "H226"
    m = re.match(r'(H[12])(\d{2})', s)
    if m:
        return m.group(1), f"20{m.group(2)}"

    # Try season code + 4-digit year e.g. "SM2026", "FW2026"
    m = re.match(r'^(SS|FW|SP|SM|FL|WN|AL|CO)(\d{4})$', s)
    if m:
        return m.group(1), m.group(2)

    # Try season code + 2-digit year e.g. "SM26", "FW26"
    m = re.match(r'^(SS|FW|SP|SM|FL|WN|AL|CO)(\d{2})$', s)
    if m:
        return m.group(1), f"20{m.group(2)}"

    # Try bare season code e.g. "SM", "FW", "SS"
    m = re.match(r'^(SS|FW|SP|SM|FL|WN|AL|CO)$', s)
    if m:
        return m.group(1), "2026"

    return "H2", "2026"   # safe fallback


def resolve_gender(product_code: str) -> tuple[str, str, str]:
    """product_code prefix → (sap_code, label, raw_desc)"""
    prefix = product_code[:2].upper() if len(product_code) >= 2 else ""
    return _GENDER_PREFIX_MAP.get(prefix, ("U", "Unisex", "Unisex"))


def resolve_coo(raw: str, coo_map: dict[str, str]) -> str:
    """Country of origin name → ISO-2 code."""
    if not raw:
        return "CN"
    key = raw.strip().upper()
    return coo_map.get(key) or _COO_FALLBACK.get(key, "CN")


def resolve_size_code(size_raw: str, size_lov: dict[str, str]) -> str:
    """
    Raw size → SAP 3-char size code.
    MDD lookup first; fallback: pad/truncate to 3 chars.
    e.g. "XS" → "0XS", "M" → "00M", "OSFA" → "0SA"
    """
    if not size_raw:
        return "000"
    key = size_raw.strip().upper()
    if key in size_lov:
        return size_lov[key]
    clean = _alnum(key)
    if not clean:
        return "000"
    if len(clean) >= 3:
        return clean[:3]
    return clean.zfill(3)


def resolve_colour_mapping(colour_code: str, colour_map: dict[str, str]) -> str:
    """
    Colour Code → 2-char mapping for Generic/SAP Style code.
    Rules from 2XU Mapping sheet:
      1. If EAN exists: first digit + last digit of colour code (before /)
         e.g. BLK/SRF → B + F = BF  (but usually = first+second of first part)
      2. First 2 chars of colour code (before /)
         e.g. BLK/SRF → BL → but example shows BS
    Mapping sheet example: BLK/SRF → first+last = BS
    Use: colour_map from Colour Mapping sheet if available, else derive.
    """
    if colour_code in colour_map:
        return colour_map[colour_code]
    # Derive: first char of part1 + first char of part2 (e.g. BLK/SRF → BS)
    parts = colour_code.split("/")
    if len(parts) >= 2 and parts[1]:
        p1 = _alnum(parts[0])
        p2 = _alnum(parts[1])
        two_char = (p1[0] + p2[0]).upper() if p1 and p2 else (p1[:2].upper() if p1 else "00")
    else:
        p1 = _alnum(colour_code)
        two_char = p1[:2].upper() if len(p1) >= 2 else p1.ljust(2, "0")
    return two_char


def resolve_sap_colour(colour_code: str, sap_colour_map: dict[str, str]) -> str:
    """
    Colour Code → SAP 3-char Color code (103 color guidance).
    e.g. BLK/SRF → 005
    Fallback chain:
      1. Exact match in sap_colour_map (from Colour Mapping sheet)
      2. Dominant colour (part1) lookup in _COLOUR_PART1_SAP_FALLBACK
      3. '000' (last resort — may fail Stibo LOV validation)
    """
    key = colour_code.upper() if colour_code else ""
    if key in sap_colour_map:
        return sap_colour_map[key]
    # Dominant colour fallback — BLK/GRY → BLK → 005
    parts = key.split("/")
    if parts and parts[0]:
        p1 = _alnum(parts[0])[:3]
        if p1 in _COLOUR_PART1_SAP_FALLBACK:
            return _COLOUR_PART1_SAP_FALLBACK[p1]
    return "000"


def build_generic_code(brand_code: str, style_code: str, colour_2d: str) -> str:
    """
    2XU Generic = BrandCode(3) + StyleCode(7, upper) + ColourMapping(2)
    Max 12 chars.
    e.g. 2XU + WR7575A + BS = 2XUWR7575ABS
    """
    sc = style_code.upper()[:7]
    return f"{brand_code}{sc}{colour_2d}"[:12]


def build_sap_style(style_code: str, colour_2d: str) -> str:
    """
    SAP Style = StyleCode(7, upper) + ColourMapping(2)
    Max 9 chars.
    e.g. WR7575A + BS = WR7575ABS
    """
    sc = style_code.upper()[:7]
    return f"{sc}{colour_2d}"[:9]


def build_variant_code(generic_code: str, sap_colour_3d: str, size_3d: str) -> str:
    """
    Variant = Generic(12) + SAPColour(3) + SAPSize(3) = max 18 chars.
    """
    return f"{generic_code}{sap_colour_3d}{size_3d}"[:18]


def build_generic_desc(brand_code: str, product_name: str,
                       gender_code: str, colour_code: str) -> str:
    """
    Generic Description = BrandCode + StyleDesc + (A/Gender) + ColourCode
    Max 40 chars. e.g. "2XU AERO WAFFLE HALF ZIP (A/W) BLK/SRF"
    Normalize: replace special chars like ½ → HALF.
    Gender char in desc: M=Mens, W=Womens, U=Unisex (NOT F for Female).
    Smart truncation: shorten product NAME first to preserve gender tag +
    full colour code suffix instead of cutting mid-word like 'BLK/'.
    """
    MAX = 40
    g_char = _DESC_GENDER_CHAR.get(gender_code, gender_code)
    name   = product_name.replace("½", "HALF").replace("1/2", "HALF").strip().upper()
    suffix = f" (A/{g_char}) {colour_code}"
    # Space reserved for brand + space + suffix; remainder for name.
    reserved = len(brand_code) + 1 + len(suffix)
    avail    = MAX - reserved
    if avail < 1:
        # Suffix alone too long — fall back to plain truncate.
        return f"{brand_code} {name}{suffix}"[:MAX]
    if len(name) > avail:
        name = name[:avail].rstrip()
    return f"{brand_code} {name}{suffix}"[:MAX]


def build_variant_desc(brand_code: str, product_name: str, gender_code: str,
                       sap_colour_desc: str, colour_code: str, size_code: str) -> str:
    """
    Variant Description = BrandCode + StyleDesc + (A/Gender) +
                          SAP Colour Description + Size      (max 40 chars)
    Falls back to PrincipalColorCode when sap_colour_desc empty.
    Smart truncation preserves the colour+size suffix; shortens NAME.
    """
    MAX = 40
    g_char = _DESC_GENDER_CHAR.get(gender_code, gender_code)
    name   = product_name.replace("½", "HALF").replace("1/2", "HALF").strip().upper()
    colour_part = (sap_colour_desc or colour_code or "").strip().upper()
    size_part   = (size_code or "").strip().upper()
    # Suffix is the must-preserve portion.
    suffix = f" (A/{g_char}) {colour_part} {size_part}".rstrip()
    reserved = len(brand_code) + 1 + len(suffix)
    avail    = MAX - reserved
    if avail < 1:
        return f"{brand_code} {name}{suffix}"[:MAX]
    if len(name) > avail:
        name = name[:avail].rstrip()
    return f"{brand_code} {name}{suffix}"[:MAX]


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER
# ══════════════════════════════════════════════════════════════════════════════
def load_mdd(path: Path) -> dict:
    """
    Load from Master Data Dictionary:
      brand_lov       : code → name
      colour_code_lov : code → description  (from Color Code LOV)
      size_lov        : raw_size_upper → sap_3digit_code
      coo_map         : country_name_upper → iso_code
      company_lov     : code → name
      sbu_lov         : code → name
    """
    log.info("[MDD] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    def _ms(v): return "" if v is None else str(v).strip()

    def _kv(sheet_name: str, key_col: int = 0, val_col: int = 1) -> dict[str, str]:
        out: dict[str, str] = {}
        if sheet_name not in wb.sheetnames:
            return out
        for row in wb[sheet_name].iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= max(key_col, val_col):
                continue
            k, v = _ms(row[key_col]), _ms(row[val_col])
            if k and v:
                out[k] = v
        return out

    brand_lov   = _kv("Brand LOV")
    company_lov = _kv("Company Code LOV")
    sbu_lov     = _kv("SBU LOV")

    # Color Code LOV: col0 = code, col1 = description
    colour_code_lov: dict[str, str] = {}
    if "Color Code LOV" in wb.sheetnames:
        for row in wb["Color Code LOV"].iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            code = _ms(row[0]).upper()
            desc = _ms(row[1]).upper() if len(row) > 1 else ""
            if code:
                colour_code_lov[code] = desc

    # Size Code LOV: multiple column groups (col5/6 is main SAP size)
    size_lov: dict[str, str] = {}
    if "Size Code LOV" in wb.sheetnames:
        for row in wb["Size Code LOV"].iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 7:
                continue
            sap  = _ms(row[5])
            desc = _ms(row[6])
            if not sap:
                continue
            sap_clean = _alnum(sap)
            if sap_clean:
                if len(sap_clean) < 3:
                    sap_clean = sap_clean.zfill(3)
                sap_clean = sap_clean[:3]
            if desc:
                size_lov[desc.upper()] = sap_clean
            # Also store raw sap as key
            if sap.upper():
                size_lov[sap.upper()] = sap_clean

    # Country Origin LOV: col0 = ISO code, col1 = name
    coo_map: dict[str, str] = {}
    if "Country Origin LOV" in wb.sheetnames:
        for row in wb["Country Origin LOV"].iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            iso  = _ms(row[0]).upper()
            name = _ms(row[1]).upper()
            if iso and name:
                coo_map[name] = iso

    # ── Generic key/value LOV reader (sheet → {VALUE_UPPER -> ID}) ────────
    # Many MDD sheets follow [ID, Value] layout. Build both directions.
    def _lov_two_way(sheet_name: str, id_col: int = 0, val_col: int = 1) -> dict[str, str]:
        out: dict[str, str] = {}
        if sheet_name not in wb.sheetnames:
            return out
        for row in wb[sheet_name].iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= max(id_col, val_col):
                continue
            code = _ms(row[id_col]).upper()
            name = _ms(row[val_col]).upper()
            if not code:
                continue
            out[code] = code              # ID → ID (validates)
            if name:
                out[name] = code          # description → ID
        return out

    age_lov      = _lov_two_way("Age LOV")
    byage_lov    = _lov_two_way("BY Age LOV")
    season_lov   = _lov_two_way("Season LOV")
    country_lov  = _lov_two_way("Country LOV")
    currency_lov = _lov_two_way("Retail Price Currency LOV")

    wb.close()
    log.info("[MDD] brand=%d colour=%d size=%d coo=%d company=%d sbu=%d "
             "age=%d byage=%d season=%d country=%d currency=%d",
             len(brand_lov), len(colour_code_lov), len(size_lov),
             len(coo_map), len(company_lov), len(sbu_lov),
             len(age_lov), len(byage_lov), len(season_lov),
             len(country_lov), len(currency_lov))
    return {
        "brand_lov":       brand_lov,
        "colour_code_lov": colour_code_lov,
        "size_lov":        size_lov,
        "coo_map":         coo_map,
        "company_lov":     company_lov,
        "sbu_lov":         sbu_lov,
        "age_lov":         age_lov,
        "byage_lov":       byage_lov,
        "season_lov":      season_lov,
        "country_lov":     country_lov,
        "currency_lov":    currency_lov,
    }


# ══════════════════════════════════════════════════════════════════════════════
# LOV RESOLVERS — translate any incoming value to a Stibo-valid LOV ID.
# Returns "" when no mapping found, so caller can skip the attribute write
# instead of triggering IllegalLOVValue errors.
# ══════════════════════════════════════════════════════════════════════════════
def _resolve_lov(raw: str, mdd_lov: dict[str, str], fallback: dict[str, str]) -> str:
    if not raw:
        return ""
    key = str(raw).strip().upper()
    if not key:
        return ""
    return mdd_lov.get(key) or fallback.get(key, "")


def _resolve_lov_fallback_first(raw: str, mdd_lov: dict[str, str], fallback: dict[str, str]) -> str:
    """Prefer hard-coded fallback IDs over MDD (used when MDD sheet columns
    are unreliable or contain descriptions instead of valid Stibo LOV IDs)."""
    if not raw:
        return ""
    key = str(raw).strip().upper()
    if not key:
        return ""
    return fallback.get(key) or mdd_lov.get(key, "")


def resolve_sap_age(raw: str, mdd: dict) -> str:
    return _resolve_lov_fallback_first(raw, mdd.get("age_lov", {}), _SAP_AGE_FALLBACK)


def resolve_by_age(raw: str, mdd: dict) -> str:
    return _resolve_lov_fallback_first(raw, mdd.get("byage_lov", {}), _BY_AGE_FALLBACK)


def resolve_season_id(season_prefix: str, mdd: dict) -> str:
    return _resolve_lov_fallback_first(season_prefix, mdd.get("season_lov", {}), _SEASON_FALLBACK)


def resolve_country(raw: str, mdd: dict) -> str:
    return _resolve_lov_fallback_first(raw, mdd.get("country_lov", {}), _COUNTRY_FALLBACK)


def resolve_currency(raw: str, mdd: dict) -> str:
    return _resolve_lov_fallback_first(raw, mdd.get("currency_lov", {}), _CURRENCY_FALLBACK)


# ══════════════════════════════════════════════════════════════════════════════
# 2XU MAPPING LOADER (Colour Mapping + SAP Mapping sheets)
# ══════════════════════════════════════════════════════════════════════════════
def load_2xu_mapping(path: Path) -> dict:
    """
    Load from 2XU Stibo Mapping file:
      colour_map    : PrincipalColourCode → 2-char mapping code (derived)
      sap_colour_map: PrincipalColourCode → SAP 3-char Color Code
      sap_colour_desc: PrincipalColourCode → SAP Color Description
      sap_map       : (Range, FieldOfPlay, Gender, Category) →
                       (sap_div, sap_group, sap_cat, sports_cat)
    """
    if not path or not path.exists():
        log.warning("[2XU Mapping] File not found: %s — using fallbacks", path)
        return {"colour_map": {}, "sap_colour_map": {}, "sap_colour_desc": {}, "sap_map": _SAP_MAP_FALLBACK}

    log.info("[2XU Mapping] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    def _ms(v): return "" if v is None else str(v).strip()

    # ── Colour Mapping sheet ──────────────────────────────────────────────────
    # Row 0 = header: Principal Colour Code | Principal Colour Description |
    #                 Color Code | Color Description | Notes
    colour_map:     dict[str, str] = {}
    sap_colour_map: dict[str, str] = {}
    sap_colour_desc: dict[str, str] = {}

    if "Colour Mapping" in wb.sheetnames:
        for row in wb["Colour Mapping"].iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 4:
                continue
            principal_code = _ms(row[0]).upper()   # e.g. BLK/SRF
            sap_code       = _ms(row[2]).upper()   # e.g. 005
            sap_desc       = _ms(row[3]).upper()   # e.g. BLACK
            if not principal_code:
                continue
            # Derive 2-char colour mapping: first char of part1 + first char of part2
            parts = principal_code.split("/")
            if len(parts) >= 2 and parts[1]:
                p1 = _alnum(parts[0])
                p2 = _alnum(parts[1])
                two_char = (p1[0] + p2[0]).upper() if p1 and p2 else (p1[:2].upper() if p1 else "00")
            else:
                p1 = _alnum(principal_code)
                two_char = p1[:2].upper() if len(p1) >= 2 else p1.ljust(2, "0")
            colour_map[principal_code]     = two_char
            sap_colour_map[principal_code] = sap_code
            sap_colour_desc[principal_code] = sap_desc

    # ── SAP Mapping sheet ─────────────────────────────────────────────────────
    # Row 1 = header (7 columns, NO Gender):
    #   Range | Field Of Play | Category | SAP Product Division |
    #   SAP Product Group | SAP Product Category | Sports Category
    sap_map: dict[tuple, tuple] = {}
    if "SAP Mapping" in wb.sheetnames:
        rows = list(wb["SAP Mapping"].iter_rows(min_row=2, values_only=True))
        for row in rows:
            if not row or len(row) < 7:
                continue
            range_val  = _ms(row[0])
            fop        = _ms(row[1])
            category   = _ms(row[2])
            sap_div    = _ms(row[3])
            sap_group  = _ms(row[4])
            sap_cat    = _ms(row[5])
            sports_cat = _ms(row[6])
            if range_val and fop:
                sap_map[(range_val, fop, category)] = (
                    sap_div, sap_group, sap_cat, sports_cat
                )

    wb.close()
    log.info("[2XU Mapping] colours=%d sap_entries=%d", len(colour_map), len(sap_map))
    return {
        "colour_map":     colour_map,
        "sap_colour_map": sap_colour_map,
        "sap_colour_desc": sap_colour_desc,
        "sap_map":        sap_map if sap_map else _SAP_MAP_FALLBACK,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ORDER CONFIRMATION LOADER
# ══════════════════════════════════════════════════════════════════════════════
def load_order_confirmation(path: Path) -> list[Article]:
    """
    Load Order Confirmation xlsx.
    Row 1 = header, Row 2+ = data.
    Groups by (PRODUCT CODE, COLOUR CODE) → one Article per colour variant.
    """
    log.info("[OrderConf] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    # Use first sheet
    ws   = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        log.error("[OrderConf] Empty file: %s", path.name)
        return []

    # Build header map
    hdr = {
        str(v).strip().upper(): idx
        for idx, v in enumerate(rows[0])
        if v is not None
    }
    log.info("[OrderConf] Headers: %s", list(hdr.keys()))

    def g(row, col: str):
        idx = hdr.get(col.upper())
        if idx is None:
            return None
        return row[idx] if idx < len(row) else None

    articles: dict[tuple, Article] = {}

    for i, row in enumerate(rows[1:], start=2):
        if all(v is None for v in row):
            continue

        product_code = _s(g(row, "PRODUCT CODE"))
        colour_code  = _s(g(row, "COLOUR CODE"))
        ean          = _s(g(row, "BARCODE/EAN"))

        if not product_code or not colour_code:
            continue

        key = (product_code, colour_code)

        if key not in articles:
            articles[key] = Article(
                distributor_code = _s(g(row, "DISTRIBUTOR CODE")),
                order_number     = _s(g(row, "ORDER NUMBER")),
                order_desc       = _s(g(row, "ORDER DESCRIPTION")),
                product_code     = product_code,
                product_name     = _s(g(row, "PRODUCT NAME")),
                colour_code      = colour_code,
                colour_name      = _s(g(row, "COLOUR NAME")),
                msrp_usd         = _num(g(row, "MSRP USD")),
                coo_raw          = _s(g(row, "COO")),
                hs_code          = _s(g(row, "HS CODE")),
                fabric_comp      = _s(g(row, "FABRIC COMP")),
                range_val        = _s(g(row, "RANGE")),
                field_of_play    = _s(g(row, "FIELD OF PLAY")),
                category         = _s(g(row, "CATEGORY")),
            )

        qty = None
        try:
            q = g(row, "ORDER QTY")
            if q is not None:
                qty = int(float(str(q)))
        except (TypeError, ValueError):
            pass

        articles[key].sizes.append(SizeLine(
            ean       = ean,
            size_code = _s(g(row, "SIZE CODE")),
            order_qty = qty,
            row_num   = i,
        ))

    result = list(articles.values())
    log.info("[OrderConf] %d variants | %d unique styles | %d EAN rows",
             len(result),
             len({a.product_code for a in result}),
             sum(len(a.sizes) for a in result))
    return result


# ══════════════════════════════════════════════════════════════════════════════
# XML INDENT HELPER
# ══════════════════════════════════════════════════════════════════════════════
def _indent(elem: ET.Element, level: int = 0) -> None:
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


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def _build_generic_product(
    style_code:   str,
    variants:     list[Article],
    cfg:          dict,
    mdd:          dict,
    mapping:      dict,
    season_id_cl: str,
) -> list[ET.Element]:
    """
    Build one PRD_GenericArticle plus one PRD_VariantArticle per (colour, size).
    Returns a flat list [generic, variant1, variant2, ...] — variants are
    SIBLINGS of the generic (not nested), each with explicit ParentID pointing
    to the generic's KEY. Stibo's nested-variant structure only imports the
    first variant correctly; siblings fix the 'DataContainer not in this node'
    error on subsequent variants.
    """
    rep          = variants[0]
    brand_code   = cfg["brand_code"]
    brand_name   = cfg["brand_name"]
    sea_prefix   = cfg["season_prefix"]
    sea_year     = cfg["season_year"]
    comp_code    = cfg["comp_code"]
    sbu          = cfg["sbu"]
    country_code = cfg["country_code"]
    article_type = cfg.get("article_type", "Inline")

    colour_map      = mapping.get("colour_map", {})
    sap_colour_map  = mapping.get("sap_colour_map", {})
    sap_colour_desc = mapping.get("sap_colour_desc", {})
    sap_map         = mapping.get("sap_map", {})
    size_lov        = mdd.get("size_lov", {})
    coo_map         = mdd.get("coo_map", {})
    company_lov     = mdd.get("company_lov", {})
    sbu_lov         = mdd.get("sbu_lov", {})

    # ── Resolve rep-level attributes ──────────────────────────────────────────
    gender_code, gender_label, gender_raw = resolve_gender(rep.product_code)
    coo_iso   = resolve_coo(rep.coo_raw, coo_map)
    colour_2d = resolve_colour_mapping(rep.colour_code, colour_map)

    generic_code = build_generic_code(brand_code, style_code, colour_2d)
    sap_style    = build_sap_style(style_code, colour_2d)
    generic_desc = build_generic_desc(brand_code, rep.product_name,
                                      gender_code, rep.colour_code)

    retail_price = str(int(rep.msrp_usd)) if rep.msrp_usd else ""
    # FOB = MSRP * 74% (per Attributes List)
    fob_price = str(round(rep.msrp_usd * 0.74, 2)) if rep.msrp_usd else ""

    # ── Season classification ─────────────────────────────────────────────────
    season_id = f"CLH_{brand_code}_{sea_prefix}{sea_year}"

    # ── Build Generic element ─────────────────────────────────────────────────
    # Identify the generic via the unique key KEY_InboundArticle (= the
    # AT_InboundGenericCode attribute). Do NOT set a system `ID` attribute
    # and do NOT also write AT_InboundGenericCode inside <Values> — Stibo
    # rejects both as either ambiguous lookup or unique-key violation when
    # the value already exists on another system-ID. Variants are nested
    # below as children so they inherit the parent context without needing
    # the parent's (unknown) system ID.
    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_A-TempSubCat")   # A = Apparel/Accessories
    _keyval(gen_el, "KEY_InboundArticle", generic_code)

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")

    def put(attr_id, value="", id_val=""):
        if not (value or id_val):
            return
        _val(vals_el, attr_id, value, id_val=id_val)

    # put("AT_CountryOrigin", id_val=coo_iso)

    # Variant Articles NESTED inside generic; identified by KEY_Variant.
    # Stibo resolves the parent via KEY_InboundArticle on the surrounding
    # <Product> element — no system ID needed for the link.
    for art in variants:
        c2d          = resolve_colour_mapping(art.colour_code, colour_map)
        sap_col_3d   = resolve_sap_colour(art.colour_code, sap_colour_map)
        sap_col_desc = sap_colour_desc.get(art.colour_code.upper(), "")
        g_code       = build_generic_code(brand_code, style_code, c2d)

        for sz in art.sizes:
            size_3d  = resolve_size_code(sz.size_code, size_lov)
            var_id   = build_variant_code(g_code, sap_col_3d, size_3d)
            var_desc = build_variant_desc(brand_code, art.product_name,
                                          gender_code, sap_col_desc,
                                          art.colour_code, sz.size_code)

            var_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Product")
            var_el.set("UserTypeID", "PRD_VariantArticle")
            _keyval(var_el, "KEY_InboundVariant", var_id)

            ET.SubElement(var_el, f"{{{STIBO_NS}}}Values")

            if sz.ean:
                dcs_el = ET.SubElement(var_el, f"{{{STIBO_NS}}}DataContainers")
                mdc_el = ET.SubElement(dcs_el, f"{{{STIBO_NS}}}MultiDataContainer")
                mdc_el.set("Type", "DC_Barcode")
                dc_el  = ET.SubElement(mdc_el, f"{{{STIBO_NS}}}DataContainer")
                dc_el.set("Analyzer", "true")
                dcv_el = ET.SubElement(dc_el, f"{{{STIBO_NS}}}Values")
                _val(dcv_el, "AT_Barcode",          sz.ean)
                _val(dcv_el, "AT_BarcodeType",      id_val="P")
                _val(dcv_el, "AT_MainEANIndicator", id_val="Y")

    return [gen_el]


def build_xml(
    articles:    list[Article],
    output_path: Path,
    cfg:         dict,
    mdd:         dict,
    mapping:     dict,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]

    sea_prefix, sea_year = derive_season(cfg["season"])
    cfg["season_prefix"] = sea_prefix
    cfg["season_year"]   = sea_year
    season_id_cl         = f"CLH_{brand_code}_{sea_prefix}{sea_year}"

    # Group by style_code
    by_style: dict[str, list[Article]] = defaultdict(list)
    for art in articles:
        by_style[art.product_code].append(art)

    log.info("[XML] Building — %d generics | %d variants | %d EAN rows",
             len(by_style), len(articles), sum(len(a.sizes) for a in articles))

    export_time   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written_count = 0

    with open(output_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        open_step_xml(f, export_time)
        f.write("  <Products>\n")

        for style_code, variants in by_style.items():
            elements = _build_generic_product(
                style_code, variants, cfg, mdd, mapping, season_id_cl
            )
            for el in elements:
                _indent(el, level=2)
                xml_str = strip_xmlns(ET.tostring(el, encoding="unicode"))
                f.write(f"    {xml_str}\n")
            written_count += 1

        f.write("  </Products>\n")
        close_step_xml(f)

    log.info("[XML] Written → %s  (%.1f KB) — %d generics",
             output_path, output_path.stat().st_size / 1024, written_count)
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# LAMBDA ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run(args, auditor=None):
    """
    Lambda dispatcher entry point.
    Reads from ORDER_DIR, writes XML to XML_OUT_DIR.

    args attributes (all optional with defaults):
        brand_code   default "2XU"
        brand        default "2XU"
        comp_code    default "0888"
        sbu          default "SP"
        season       default "H226"
        country_code default "ID"
        article_type_from_filename default "Inline"
    """
    def _first_xlsx(d: Path):
        files = list(d.glob("*.xlsx"))
        return files[0] if files else None

    order_f  = _first_xlsx(ORDER_DIR)
    mdd_f    = _first_xlsx(MDD_DIR)

    # 2XU Mapping file — look in attributes dir
    mapping_f = None
    for p in ATTR_DIR.glob("*.xlsx"):
        if "2xu" in p.name.lower() and "mapping" in p.name.lower():
            mapping_f = p
            break
    if not mapping_f:
        # fallback: any xlsx in attributes dir with "mapping"
        for p in ATTR_DIR.glob("*.xlsx"):
            if "mapping" in p.name.lower():
                mapping_f = p
                break

    if not order_f:
        log.error("[2XU OrderConf] No Order Confirmation file in %s", ORDER_DIR)
        return
    if not mdd_f:
        log.warning("[2XU OrderConf] No MDD file — colour/size lookups use fallbacks")

    # Load MDD
    mdd = load_mdd(mdd_f) if mdd_f else {
        "brand_lov": {}, "colour_code_lov": {}, "size_lov": {},
        "coo_map": {}, "company_lov": {}, "sbu_lov": {},
    }

    # Load 2XU Mapping
    mapping = load_2xu_mapping(mapping_f) if mapping_f else {
        "colour_map": {}, "sap_colour_map": {}, "sap_colour_desc": {}, "sap_map": _SAP_MAP_FALLBACK,
    }

    brand_code   = getattr(args, "brand_code",   "2XU")    or "2XU"
    brand_name   = getattr(args, "brand",        "2XU")    or "2XU"
    season       = getattr(args, "season",       "H226")   or "H226"
    comp_code    = getattr(args, "comp_code",    "0888")   or "0888"
    sbu          = getattr(args, "sbu",          "SP")     or "SP"
    country_code = getattr(args, "country_code", "ID")     or "ID"
    article_type = getattr(args, "article_type_from_filename", "Inline") or "Inline"

    # AT_Country = MAA destination country (Indonesia for entity 0888), NOT
    # the COO/season tag from filename. Dispatcher sometimes passes 'CN'
    # parsed from e.g. 'H226-CN-1' — sanitize against the valid LOV set.
    if country_code.upper() not in _COUNTRY_FALLBACK:
        log.warning("[2XU OrderConf] country_code '%s' not in destination LOV — defaulting to 'ID'",
                    country_code)
        country_code = "ID"

    cfg = {
        "brand_code":   brand_code,
        "brand_name":   mdd["brand_lov"].get(brand_code, brand_name),
        "season":       season,
        "comp_code":    comp_code,
        "sbu":          sbu,
        "country_code": country_code,
        "article_type": article_type,
    }

    log.info("[2XU OrderConf] cfg → brand=%s comp=%s sbu=%s country=%s season=%s",
             cfg["brand_name"], cfg["comp_code"], cfg["sbu"],
             cfg["country_code"], cfg["season"])

    # Load order confirmation
    articles = load_order_confirmation(order_f)
    if not articles:
        log.warning("[2XU OrderConf] No articles parsed — skipping %s", order_f.name)
        return

    # ── TEST MODE: limit to 1 article ──────────────────────────
    # articles = articles[:1]

    # Derive season from ORDER DESCRIPTION if not overridden
    if not getattr(args, "season", None):
        first_desc = articles[0].order_desc
        if first_desc:
            derived_pfx, derived_yr = derive_season(first_desc)
            cfg["season"] = f"{derived_pfx}{derived_yr[2:]}"
            log.info("[2XU OrderConf] Season derived from ORDER DESCRIPTION: %s → %s",
                     first_desc, cfg["season"])

    out_path = XML_OUT_DIR / (order_f.stem + ".xml")
    build_xml(articles, out_path, cfg, mdd, mapping)

    if auditor:
        auditor.set_xml_uploads([str(out_path)])

    total_eans = sum(len(a.sizes) for a in articles)
    print("═" * 55, flush=True)
    print("  ARTICLE SUMMARY — 2XU ORDER CONFIRMATION", flush=True)
    print("═" * 55, flush=True)
    print(f"  File           : {order_f.name}", flush=True)
    print(f"  Variants       : {len(articles)}", flush=True)
    print(f"  Unique styles  : {len({a.product_code for a in articles})}", flush=True)
    print(f"  Total EAN rows : {total_eans}", flush=True)
    print(f"  Season         : {cfg['season']}", flush=True)
    print(f"  Country        : {cfg['country_code']}", flush=True)
    print(f"  Company Code   : {cfg['comp_code']}", flush=True)
    print(f"  SBU            : {cfg['sbu']}", flush=True)
    print(f"  XML output     : {out_path.name}  ({out_path.stat().st_size // 1024} KB)", flush=True)
    print("═" * 55, flush=True)


# ══════════════════════════════════════════════════════════════════════════════
# LOCAL CLI
# ══════════════════════════════════════════════════════════════════════════════
def main():
    import argparse
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — 2XU Order Confirmation"
    )
    p.add_argument("--brand-code",   default="2XU")
    p.add_argument("--brand",        default="2XU")
    p.add_argument("--season",       default="H226")
    p.add_argument("--comp-code",    default="0888")
    p.add_argument("--sbu",          default="SP")
    p.add_argument("--country-code", default="ID")
    p.add_argument("--article-type-from-filename", default="Inline")
    run(p.parse_args())


if __name__ == "__main__":
    main()