"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Lotto (LIC FOB OrderForm) v1.0  ║
║   Lotto FW26 LIC FOB OrderForm → Stibo STEP XML                 ║
╚══════════════════════════════════════════════════════════════════╝

Source file format: "FW26 LIC FOB ORDERFORM (1).xlsx"
Sheet name        : "Tabella1"

Layout (rows 1-indexed):
  Rows  2-9   → Size Code reference table (col 21 = code, col 22 = desc,
                cols 23+ = ordered size labels for that code)
  Row   10    → totals row (skipped)
  Row   12    → HEADER  (23 columns, 0..22)
  Rows 13..N  → DATA rows. Each row = one (Style Code + Colour Code) combo.
                Per-row sizes are derived by looking up Size Code (col 21)
                in the reference table and filtering by Size Range (col 19).

Hierarchy emitted (linelist/TAF XML format — fully self-contained):
  PRD_GenericArticle  (style_code)
    └─ PRD_Variant       (style_code + colour_code)
        └─ PRD_SizeVariant  (variant_id + sanitized size, per derived size)

No EAN in source → no AT_Barcode on SizeVariant.
"""

from __future__ import annotations

import logging
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
# INLINE STIBO / TAF XML HELPERS  (self-contained — no cross-brand imports)
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')

LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "JR": "Junior", "AA": "All Ages",
}
LOV_SEASON_NAME = {
    "SS": "Spring Summer", "FW": "Fall Winter",
    "AW": "Autumn Winter", "HO": "Holiday",  "AL": "All Season",
    "SM": "Summer",        "SP": "Spring",
}

LOV_SBU = {"SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear"}
LOV_COUNTRY_ORIGIN = {
    "CN": "China",       "VN": "Vietnam",     "ID": "Indonesia",   "KH": "Cambodia",
    "BD": "Bangladesh",  "IN": "India",        "MY": "Malaysia",    "TH": "Thailand",
    "TN": "Tunisia",     "IT": "Italy",        "PH": "Philippines", "PK": "Pakistan",
}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
    "Licensed": "License", "INLINE": "Inline",
}
LOV_INDICATOR = {"Y": "Yes", "N": "No"}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set"}
LOV_NATURE_OF_ARTICLE = {"REG": "Regular", "WH1": "Wholesale", "PROMO": "Promo"}
LOV_SAP_ARTICLE_TYPE  = {"DIRECT": "Direct", "INDIRECT": "Indirect"}


def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "", derived: bool = False) -> Optional[ET.Element]:
    """Add a <Value AttributeID="..."> child. Either text value or LOV ID."""
    if derived:
        return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id = str(id_val).strip() if id_val else ""
    if clean_id and clean_id not in ("None", "nan"):
        el.set("ID", clean_id)
        return el
    clean_val = str(value).strip() if value else ""
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


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


def strip_xmlns(xml_str: str) -> str:
    """Remove xmlns attributes added by ElementTree serialization."""
    return _XMLNS_RE.sub("", xml_str)


def open_step_xml(f, export_time: str) -> None:
    """Write the STEP-ProductInformation XML preamble and root open tag."""
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    f.write(
        f'<STEP-ProductInformation '
        f'xmlns="{STIBO_NS}" '
        f'xmlns:xsi="{STIBO_XSI}" '
        f'xsi:schemaLocation="{STIBO_SCHEMA}" '
        f'ExportTime="{export_time}" '
        f'ExportContext="Context1" '
        f'WorkspaceID="Main" '
        f'UseContextLocale="false">\n'
    )


def close_step_xml(f) -> None:
    f.write("</STEP-ProductInformation>\n")


def derive_season(season_code: str) -> tuple[str, str]:
    """'FW26' → ('FW', '2026'). Handles bare prefix or already 4-digit year."""
    s = (season_code or "").strip().upper()
    if len(s) < 2:
        return s, ""
    prefix = s[:2]
    tail   = s[2:]
    if len(tail) == 2 and tail.isdigit():
        return prefix, f"20{tail}"
    if len(tail) == 4 and tail.isdigit():
        return prefix, tail
    return prefix, tail


def season_id_full(brand_code: str, season_code: str) -> str:
    """Build classification ID like 'CLH_LOT_FW2026'."""
    pfx, yr = derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"


def derive_franchise(value: str) -> str:
    """Extract a franchise token (first word, uppercased, alphanum-only)."""
    if not value:
        return ""
    first = re.split(r"[\s/\-_,]+", value.strip(), maxsplit=1)[0]
    return re.sub(r"[^A-Za-z0-9]", "", first).upper()


def derive_ecom_name(style_name: str) -> str:
    """Title-case e-commerce display name (collapsed whitespace)."""
    if not style_name:
        return ""
    return re.sub(r"\s+", " ", style_name.strip()).title()


def build_classifications(brand_name: str, brand_code: str, season_code: str) -> ET.Element:
    """Build the <Classifications> root with Season + Confirmed/Unconfirmed children."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    brand_token = brand_name.title().replace(" ", "")

    pfx, yr     = derive_season(season_code)
    season_id   = season_id_full(brand_code, season_code)
    batches_par = f"CLH_{brand_token}Batches"
    sea_name    = LOV_SEASON_NAME.get(pfx, pfx)
    display     = f"{brand_name} {sea_name} {yr}".strip()
    short       = f"{pfx} {yr}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_par)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = display

    for suffix, type_id, label in (
        ("CA", "CLS_ConfirmedArticles",   f"{short} Confirmed Articles"),
        ("UA", "CLS_UnconfirmedArticles", f"{short} Unconfirmed Articles"),
    ):
        sub = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        sub.set("ID",         f"{season_id}{suffix}")
        sub.set("UserTypeID", type_id)
        ET.SubElement(sub, f"{{{STIBO_NS}}}Name").text = label

    return cls_root


def add_generic_values(vals_el: ET.Element, ctx: dict) -> None:
    """Write Generic-level <Value> children from a ``ctx`` dict.

    Only writes a key once and skips empty values, so callers can pass a wide
    dict without worrying about duplicates.
    """
    written: set[str] = set()

    def w(attr_id: str, value: str = "", id_val: str = "") -> None:
        if attr_id in written:
            return
        if not value and not id_val:
            return
        _val(vals_el, attr_id, value=value, id_val=id_val)
        written.add(attr_id)

    w("AT_Brand",       id_val=ctx.get("brand_code", ""))
    w("AT_BrandGroup",  id_val=ctx.get("brand_name", "").upper())
    _sbu = ctx.get("sbu", "")
    w("AT_SBU",         LOV_SBU.get(_sbu, _sbu),           id_val=_sbu)
    w("AT_CompanyCode", id_val=ctx.get("comp_code", ""))
    _coo = ctx.get("coo", "")
    w("AT_CountryOrigin", LOV_COUNTRY_ORIGIN.get(_coo, _coo), id_val=_coo)
    w("AT_PrincipalStyleCode",  ctx.get("style_code", ""))
    w("AT_PrincipalStyleDescription", ctx.get("style_name", ""))
    w("AT_PrincipalColorCode", ctx.get("colour_code", ""))
    w("AT_PrincipalColorName", ctx.get("colour", ""))
    # AT_SAPColor                  — mapping is 'na' for Licensed; left empty. TODO: confirm with Rushi
    w("AT_SAPStyleCode",       ctx.get("style_code", ""))
    # AT_Generic — removed; AT_InboundGenericCode used instead
    w("AT_InboundGenericCode",  ctx.get("generic_code", ""))
    # AT_PrincipalColorCode — 'E' (N/A for Licensed) per mapping
    # AT_Color — colour name not from OrderForm at generic level
    _gc = ctx.get("gender_code", "")
    if _gc in ("M", "F", "U"):
        w("AT_Gender", ctx.get("gender_label", ""), id_val=_gc)
    w("AT_PrincipalGenderDescription", ctx.get("gender_raw", ""))
    # AT_Gender: K (Kids) not in Stibo LOV_Gender — skipped for kids range
    w("AT_SAPAge", ctx.get("age_label", ""), id_val=ctx.get("age_code", ""))
    w("AT_Season",   LOV_SEASON_NAME.get(ctx.get("season_prefix", ""), ctx.get("season_prefix", "")),
                     id_val=ctx.get("season_prefix", ""))
    w("AT_SeasonYear", ctx.get("season_year", ""))
    # AT_SAPArticleCategory, AT_BYArticleType — not in FOB OrderForm mapping
    # AT_SAPArticleType, AT_BYIndicator, AT_SAPIndicator — formula/system fields, not from OrderForm
    w("AT_FOB", ctx.get("fob", ""))
    w("AT_FOBCurrency", ctx.get("fob_currency", ""), id_val=ctx.get("fob_currency", ""))
    w("AT_SAPProductFlag", "A", id_val="A")
    w("AT_PrincipalMerchandiseHierarchyL1", ctx.get("product_type", ""))
    w("AT_PrincipalMerchandiseHierarchyL2", ctx.get("discipline", ""))
    w("AT_MaterialType", "", id_val="ZINA")
    w("AT_CountrySize", "", id_val="EU")
    # AT_RetailPriceCurrency: derived from filename country_code via MDD Country LOV → Retail Price Currency LOV
    _rpc = ctx.get("retail_price_currency", "")
    if _rpc:
        w("AT_RetailPriceCurrency", "", id_val=_rpc)
    # ── Brand Type / Brand Category (from RNA) ──────────────────
    w("AT_BrandType",     ctx.get("brand_type", ""))
    w("AT_BrandCategory", ctx.get("brand_category", ""))
    # AT_FOBCurrency — 'fob currency' column not present in input Excel; left empty. TODO: confirm source column with Rushi
    # AT_BYIndicator, AT_SAPIndicator, AT_EcomIndicator — formula/system fields, not in OrderForm. Noted.
    # AT_PrincipalAgeDescription — not in attribute sheet mapping. Noted.
    # AT_SAPProductGroup, AT_SAPProductCategory — not configured in Stibo for Lotto
    # AT_SAPProductDivision — 'C' (from images), not from OrderForm
    # AT_LaunchingDate, AT_NatureOfArticle — not in FOB OrderForm mapping

# ══════════════════════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR     = BASE_DIR / "input"
ORDERFORM_DIR = INPUT_DIR / "orderform"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in (INPUT_DIR, ORDERFORM_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
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
log = logging.getLogger("lotto.orderform_main")

# ── Country code → MDD COUNTRY name ──────────────────────────────
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
}


# ══════════════════════════════════════════════════════════════════════════════
# RNA LOADER (Brand Type / Brand Category from Attributes List)
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

        def _find(*candidates) -> Optional[int]:
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


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Article:
    row_num:        int
    collection:     str
    gender_raw:     str   # CHILD / JUNIOR / ADULT / M / W / U
    product_type:   str   # Footwear / Apparel / Accessories
    discipline:     str
    universe:       str
    line:           str
    subline:        str
    style_code:     str
    style_name:     str
    colour_code:    str
    colour:         str
    made_in:        str   # CHINA / TUNISIA / ...
    currency:       str   # USD
    whl_price:      Optional[float]
    retail_price:   Optional[float]
    size_range:     str   # e.g. "27-35", "S-XL", "TU-TU"
    size_type:      str   # "EU" / "" / "USA"
    size_code_ref:  str   # "01" .. "15" / "TU"
    size_desc:      str
    sizes:          list[str] = field(default_factory=list)  # resolved


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS  (used only as last-resort fallback when MDD lookup misses)
# ══════════════════════════════════════════════════════════════════════════════
# Minimal fallback maps — primary lookups come from MDD LOV sheets at runtime.
# Mapping sheet: MAN -> Age=Male, WOMAN -> Age=Female, CHILD/JUNIOR -> Age=Children
FALLBACK_GENDER = {
    "M": ("M", "AD"), "MAN": ("M", "AD"), "MEN": ("M", "AD"), "MALE": ("M", "AD"),
    "W": ("F", "AD"),  "WOMAN": ("F", "AD"),  "WOMEN": ("F", "AD"),  "FEMALE": ("F", "AD"), "F": ("F", "AD"),
    "U": ("U", "AD"), "UNISEX": ("U", "AD"), "ADULT": ("U", "AD"), "ADULTS": ("U", "AD"), "SENIOR": ("U", "AD"),
    "K": ("U", "CH"), "KIDS": ("U", "CH"), "CHILD": ("U", "CH"),
    "JUNIOR": ("U", "CH"), "JR": ("U", "CH"),
    "INFANT": ("U", "IN"), "BABY": ("U", "IN"),
}

# Orderform Discipline -> (SAP Product Group, SAP Product Category)
# Source: Mapping Detail Lotto (1).xlsx — SAP Product Group & Category sheet (License section)
SAP_DISCIPLINE_MAP: dict[str, tuple[str, str]] = {
    "BADMINTON":    ("TENNIS",    "TENNIS"),
    "CASUAL":       ("LIFESTYLE", "LIFESTYLE"),
    "FIVE A SIDE":  ("FOOTBALL",  "FOOTBALL"),
    "FUTSAL":       ("FOOTBALL",  "FOOTBALL"),
    "FOOTBALL":     ("FOOTBALL",  "FOOTBALL"),
    "HIKING":       ("OUTDOOR",   "OUTDOOR"),
    "KIDS":         ("LIFESTYLE", "LIFESTYLE"),
    "LIFESTYLE":    ("LIFESTYLE", "LIFESTYLE"),
    "OUTDOOR":      ("OUTDOOR",   "OUTDOOR"),
    "OUTDOOR SHOE": ("OUTDOOR",   "OUTDOOR"),
    "PADEL":        ("TENNIS",    "TENNIS"),
    "PADDLE":       ("TENNIS",    "TENNIS"),
    "POOL":         ("LIFESTYLE", "LIFESTYLE"),
    "RUNNING":      ("RUNNING",   "RUNNING"),
    "SABOTS":       ("LIFESTYLE", "LIFESTYLE"),
    "SANDAL":       ("SANDALS",   "SANDALS"),
    "SANDALS":      ("SANDALS",   "SANDALS"),
    "SOCCER":       ("FOOTBALL",  "FOOTBALL"),
    "TENNIS":       ("TENNIS",    "TENNIS"),
    "PERFORMANCE":  ("FOOTBALL",  "FOOTBALL"),
    "PERFORMANCE ACCESSORIES": ("FOOTBALL", "FOOTBALL"),
}
SAP_UNIVERSE_MAP: dict[str, tuple[str, str]] = {
    "LIFES":      ("LIFESTYLE", "LIFESTYLE"),
    "LIFESTYLE":  ("LIFESTYLE", "LIFESTYLE"),
}

FALLBACK_DIVISION = {
    "FOOTWEAR":    ("F", "Shoes"),
    "SHOES":       ("F", "Shoes"),
    "APPAREL":     ("A", "Apparel"),
    "ACCESSORIES": ("A", "Accessories"),  # PPH_AC-TempSubCat doesn't exist in Stibo; use PPH_A-TempSubCat
}

FALLBACK_COUNTRY = {
    "CHINA": "CN", "VIETNAM": "VN", "INDONESIA": "ID", "CAMBODIA": "KH",
    "BANGLADESH": "BD", "INDIA": "IN", "MALAYSIA": "MY", "THAILAND": "TH",
    "TUNISIA": "TN", "ITALY": "IT", "PHILIPPINES": "PH", "PAKISTAN": "PK",
}


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _num(v) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")

# ── MAA Article Code building blocks (always 3 chars each) ─────────────────
#   Variant ID      = Generic + 3-char Color
#   SizeVariant ID  = Generic + 3-char Color + 3-char Size  (max 18 digits)

def color_3d(code: str, mdd: Optional[dict] = None) -> str:
    """Normalize a colour code to exactly 3 alphanumeric chars.

    If MDD Color Code LOV contains this code, the LOV value is returned.
    Otherwise the raw code is alphanumeric-normalized & padded to 3 chars.
    """
    s = _ALNUM_RE.sub("", code or "").upper()
    if not s:
        return "000"
    if mdd:
        if s in mdd.get("color_code_lov", {}):
            return s if len(s) == 3 else (s[:3] if len(s) > 3 else s.zfill(3))
        # try by description (e.g. raw colour name match)
    if len(s) >= 3:
        return s[:3]
    return s.zfill(3)


def size_3d(label: str, mdd: Optional[dict] = None) -> str:
    """Normalize a size label to exactly 3 alphanumeric chars.

    If MDD Size Code LOV contains this label/code, the SAP 3-digit code is returned.
    Otherwise the raw label is alphanumeric-normalized & padded to 3 chars.

    Examples (no MDD):
        '27'   -> '027'
        '35,5' -> '35H'    (comma = half size, SAP standard)
        'XL'   -> '0XL'
    """
    # Half-size handling: '35,5' / '35.5' -> '35H'
    raw = (label or "").strip()
    if raw:
        m = re.match(r"^(\d+)[,\.]5$", raw)
        if m:
            whole = m.group(1)
            return (whole + "H") if len(whole) >= 2 else (whole.zfill(2) + "H")
    s = _ALNUM_RE.sub("", label or "").upper()
    if not s:
        return "000"
    if mdd:
        sap = mdd.get("size_code_lov", {}).get(s)
        if sap:
            sap = _ALNUM_RE.sub("", sap).upper()
            if sap:
                return sap[:3] if len(sap) >= 3 else sap.zfill(3)
    if len(s) >= 3:
        return s[:3]
    return s.zfill(3)


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER (Master Data Dictionary)
# ══════════════════════════════════════════════════════════════════════════════
def _find_mdd_file() -> Optional[Path]:
    """Locate MDD xlsx in lotto/input/mdd/, fall back to ../onr/input/mdd/ for local dev."""
    if MDD_DIR.exists():
        xs = sorted(MDD_DIR.glob("*.xlsx")) or sorted(MDD_DIR.glob("*.xlsm"))
        if xs:
            return xs[0]
    # local-dev fallback — sibling onr folder
    sibling = Path(__file__).parent.parent / "onr" / "input" / "mdd"
    if sibling.exists():
        xs = sorted(sibling.glob("*.xlsx")) or sorted(sibling.glob("*.xlsm"))
        if xs:
            return xs[0]
    return None


def load_mdd(path: Optional[Path] = None) -> dict:
    """Load MAA Master Data Dictionary.

    Returns a dict containing the explicit Brand/Color/Size LOVs PLUS every
    additional ``*LOV*`` sheet found in the workbook auto-loaded as a generic
    ``{display_name -> code}`` mapping under ``generic_lov``.

    Keys returned:
      brand_lov      : {code -> name}
      color_code_lov : {normalized_code -> description}
      size_code_lov  : {normalized_label -> sap_3digit_code}
      generic_lov    : {sheet_display_name -> {value_name -> code}}
      gender_lov     : {raw_value_upper -> (sap_gender_code, sap_age_code)}
      division_lov   : {raw_value_upper -> (sap_division_letter, label)}
      country_lov    : {country_name_upper -> iso_code}
      country_id_to_name : {iso_code_upper -> country_name}
      retail_price_currency_lov : {country_name_upper -> currency_code}
    """
    if path is None:
        path = _find_mdd_file()
    empty = {
        "brand_lov": {}, "color_code_lov": {}, "size_code_lov": {},
        "generic_lov": {}, "gender_lov": {}, "division_lov": {}, "country_lov": {},
        "country_id_to_name": {}, "retail_price_currency_lov": {},
        "prod_group_lov": {}, "prod_cat_lov": {},
    }
    if path is None or not path.exists():
        log.warning("[MDD] No MDD file found — all lookups will use fallback maps only.")
        return empty

    log.info("[MDD] Loading: %s", path.name)
    wb = _open_wb(path)

    brand_lov: dict[str, str] = {}
    if "Brand LOV" in wb.sheetnames:
        for r in wb["Brand LOV"].iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[1]:
                brand_lov[_s(r[0]).upper()] = _s(r[1])

    color_code_lov: dict[str, str] = {}
    if "Color Code LOV" in wb.sheetnames:
        for r in wb["Color Code LOV"].iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                code = _ALNUM_RE.sub("", _s(r[0])).upper()
                desc = _s(r[1]) if len(r) > 1 else ""
                if code:
                    color_code_lov[code] = desc

    size_code_lov: dict[str, str] = {}
    if "Size Code LOV" in wb.sheetnames:
        for r in wb["Size Code LOV"].iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 7:
                continue
            sap_code = _s(r[5])
            desc     = _s(r[6])
            if not sap_code:
                continue
            sap_norm = _ALNUM_RE.sub("", sap_code).upper()
            if sap_norm and sap_norm not in size_code_lov:
                size_code_lov[sap_norm] = sap_code
            if desc:
                desc_norm = _ALNUM_RE.sub("", desc).upper()
                if desc_norm and desc_norm not in size_code_lov:
                    size_code_lov[desc_norm] = sap_code

    # Generic auto-loader for every other "*LOV*" sheet — first two columns
    # treated as (code, value_name). Stored as {value_name -> code}.
    generic_lov: dict[str, dict[str, str]] = {}
    explicit = {"Brand LOV", "Color Code LOV", "Size Code LOV"}
    for sn in wb.sheetnames:
        if "LOV" not in sn.upper() or sn in explicit:
            continue
        rows = list(wb[sn].iter_rows(values_only=True))
        if len(rows) < 2:
            continue
        display = sn.replace(" LOV", "").replace("LOV ", "").strip() or sn
        bucket: dict[str, str] = {}
        for r in rows[1:]:
            if not r or len(r) < 2:
                continue
            code, name = r[0], r[1]
            if not name and not code:
                continue
            key = _s(name) if name else _s(code)
            val = _s(code) if code else _s(name)
            if key:
                bucket.setdefault(key.upper(), val)
        if bucket:
            generic_lov[display] = bucket

    # ── Gender LOV: read MDD "Gender LOV" sheet directly so SAP Age comes from MDD,
    # not from a hardcoded age_default. Schema: [raw_code, sap_gender_desc, sap_gender, sap_age, sap_age_desc].
    gender_lov: dict[str, tuple[str, str]] = {}
    if "Gender LOV" in wb.sheetnames:
        for r in wb["Gender LOV"].iter_rows(min_row=2, values_only=True):
            if not r or not r[0]:
                continue
            raw_key  = _s(r[0]).upper()
            sap_gen  = (_s(r[2]) if len(r) > 2 else "").upper() or raw_key[:1]
            sap_age  = (_s(r[3]) if len(r) > 3 else "").upper()
            if raw_key and sap_gen:
                gender_lov[raw_key] = (sap_gen, sap_age)
                # also index by full description if present
                desc = _s(r[1]) if len(r) > 1 else ""
                if desc:
                    gender_lov[desc.upper()] = (sap_gen, sap_age)
    # Legacy fallback: derive from generic_lov when MDD has no dedicated sheet.
    if not gender_lov:
        g_src = None
        for cand in ("Gender", "SAPGender", "BYGender", "Principal Gender"):
            if cand in generic_lov:
                g_src = generic_lov[cand]
                break
        if g_src:
            for name_upper, code in g_src.items():
                sap = (code or "").strip().upper()[:1] or name_upper[:1]
                gender_lov[name_upper] = (sap, "")

    # Division LOV
    division_lov: dict[str, tuple[str, str]] = {}
    d_src = None
    for cand in ("Product Division", "SAPProductDivision", "Division"):
        if cand in generic_lov:
            d_src = generic_lov[cand]
            break
    if d_src:
        for name_upper, code in d_src.items():
            letter = (code or "").strip().upper() or name_upper[:1]
            label  = name_upper.title()
            division_lov[name_upper] = (letter, label)

    # ── SAP Product Group / Category LOV: optional MDD sheets keyed by Discipline.
    # Sheet schemas (first row = header):
    #   "SAP Product Group LOV"    : [Discipline, SAP_Group]
    #   "SAP Product Category LOV" : [Discipline, SAP_Category]
    # Universe-based fallback uses the same sheets via the "Universe" column when discipline misses.
    def _two_col_lov(sheet_name: str) -> dict[str, str]:
        out: dict[str, str] = {}
        if sheet_name not in wb.sheetnames:
            return out
        for r in wb[sheet_name].iter_rows(min_row=2, values_only=True):
            if not r or not r[0] or len(r) < 2 or not r[1]:
                continue
            out[_s(r[0]).upper()] = _s(r[1])
        return out

    prod_group_lov = _two_col_lov("SAP Product Group LOV")
    prod_cat_lov   = _two_col_lov("SAP Product Category LOV")

    # Country LOV: {iso_code_upper -> country_name} from "Country LOV" tab (Col A = ID, Col B = name)
    country_id_to_name: dict[str, str] = {}
    if "Country LOV" in wb.sheetnames:
        for r in wb["Country LOV"].iter_rows(min_row=2, values_only=True):
            if r and r[0] and len(r) > 1 and r[1]:
                country_id_to_name[_s(r[0]).upper()] = _s(r[1])

    # Retail Price Currency LOV: {country_name_upper -> currency_code}
    # from "Retail Price Currency LOV" tab (Col A = country name, Col B = currency)
    retail_price_currency_lov: dict[str, str] = {}
    if "Retail Price Currency LOV" in wb.sheetnames:
        for r in wb["Retail Price Currency LOV"].iter_rows(min_row=2, values_only=True):
            if r and r[0] and len(r) > 1 and r[1]:
                retail_price_currency_lov[_s(r[0]).upper()] = _s(r[1])

    # Country LOV: {country_name_upper -> iso_code} (legacy lookup)
    country_lov: dict[str, str] = {}
    c_src = None
    for cand in ("Country", "Country Origin", "Country Of Origin"):
        if cand in generic_lov:
            c_src = generic_lov[cand]
            break
    if c_src:
        for name_upper, code in c_src.items():
            if code:
                country_lov[name_upper] = code.strip().upper()

    wb.close()
    log.info(
        "[MDD] brand=%d color=%d size=%d gender=%d division=%d country=%d generic_sheets=%d",
        len(brand_lov), len(color_code_lov), len(size_code_lov),
        len(gender_lov), len(division_lov), len(country_lov), len(generic_lov),
    )
    return {
        "brand_lov":      brand_lov,
        "color_code_lov": color_code_lov,
        "size_code_lov":  size_code_lov,
        "generic_lov":    generic_lov,
        "gender_lov":     gender_lov,
        "division_lov":   division_lov,
        "country_lov":    country_lov,
        "country_id_to_name": country_id_to_name,
        "retail_price_currency_lov": retail_price_currency_lov,
        "prod_group_lov": prod_group_lov,
        "prod_cat_lov":   prod_cat_lov,
    }


# ══════════════════════════════════════════════════════════════════════════════
# SIZE REFERENCE TABLE
# ══════════════════════════════════════════════════════════════════════════════
def load_size_reference(rows: list[tuple]) -> dict[str, list[str]]:
    """
    Read rows 2..10 (idx 1..9). Each row that has a code in col 21 + sizes in col 23+
    becomes one reference entry.

    Returns: { size_code_id (str) → [size_label, ...] }
    Plus an implicit "TU" → ["TU"] for one-size items.
    """
    ref: dict[str, list[str]] = {}
    for r in rows[1:10]:
        if len(r) < 23:
            continue
        code = _s(r[21])
        if not code or code.lower() == "total":
            continue
        sizes = [_s(c) for c in r[23:] if c is not None and _s(c) != ""]
        if sizes:
            ref[code.upper()] = sizes
    # Add fallbacks
    ref.setdefault("TU", ["TU"])
    log.info("[SizeRef] Loaded %d size code groups: %s", len(ref), sorted(ref.keys()))
    return ref


def _parse_size_range(rng: str) -> Optional[tuple[str, str]]:
    """'27-35' → ('27','35'); 'S-XL' → ('S','XL'); 'TU-TU' → ('TU','TU')."""
    if not rng:
        return None
    parts = rng.replace("–", "-").split("-")
    if len(parts) != 2:
        return None
    a, b = parts[0].strip(), parts[1].strip()
    if not a or not b:
        return None
    return a, b


def derive_sizes_for_article(art: Article, size_ref: dict[str, list[str]]) -> list[str]:
    """Resolve actual sizes for an article from size_code lookup + size_range filter."""
    code = (art.size_code_ref or "").upper().strip()
    full = size_ref.get(code)
    if not full:
        # Fallback — use the desc/range as a single size
        return [art.size_range or art.size_desc or "TU"]

    rng = _parse_size_range(art.size_range)
    if not rng:
        return list(full)

    start, end = rng
    if start not in full or end not in full:
        # range labels don't match the reference list — return the full list
        return list(full)

    i, j = full.index(start), full.index(end)
    if i > j:
        i, j = j, i
    return full[i:j + 1]


# ══════════════════════════════════════════════════════════════════════════════
# ORDERFORM LOADER
# ══════════════════════════════════════════════════════════════════════════════

# Column name aliases: field → list of possible header strings (lowercase)
_COL_ALIASES: dict[str, list[str]] = {
    "style_code":   ["style code", "style no", "art no", "article no", "article code", "code art", "supp art #"],
    "style_name":   ["style name", "style description", "article name", "description", "name"],
    "colour_code":  ["colour code", "color code", "col code", "colour no", "color no"],
    "colour":       ["colour", "color", "colour name", "color name", "colour description"],
    "collection":   ["collection", "collezione"],
    "gender_raw":   ["gender", "genre", "sesso"],
    "product_type": ["product type", "category", "division", "tipo prodotto"],
    "discipline":   ["discipline", "sport", "activity"],
    "universe":     ["universe", "universo"],
    "line":         ["line", "linea"],
    "subline":      ["subline", "sub line", "sub-line"],
    "made_in":      ["made in", "country of origin", "origin", "country"],
    "currency":     ["currency", "cur", "devise"],
    "whl_price":    ["whl price", "wholesale price", "fob price", "fob", "cost price"],
    "retail_price": ["retail price", "rrp", "retail", "retail eur", "retail usd"],
    "size_range":   ["size range", "range tailles", "range"],
    "size_type":    ["size type", "type taille"],
    "size_code_ref":["size code", "size ref", "size group", "codice taglia"],
    "size_desc":    ["size description", "size desc", "taille"],
    "action":       ["action", "status", "remark", "remarks", "note", "notes", "remove"],
}

# Values in the action/status column that mean "skip this row"
_REMOVE_VALUES = frozenset({"remove", "removed", "delete", "deleted", "cancel", "cancelled",
                             "exclude", "excluded", "no", "drop", "dropped"})


def _build_col_map(header_row: tuple) -> dict[str, int]:
    """Build {normalized_header_text → col_index} from a header row tuple."""
    col_map: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        if cell is None:
            continue
        key = re.sub(r"\s+", " ", str(cell).strip().lower())
        if key:
            col_map[key] = i
    return col_map


def _find_col(col_map: dict[str, int], field: str) -> Optional[int]:
    """Return first matching column index for ``field`` using _COL_ALIASES."""
    for alias in _COL_ALIASES.get(field, []):
        if alias in col_map:
            return col_map[alias]
    return None


def _find_data_sheet_and_header(
    wb,
) -> tuple[list[tuple], str, int]:
    """
    Scan every sheet (first 30 rows each) looking for a row that contains
    BOTH a style-code-like column AND a colour-code-like column from
    _COL_ALIASES.  Returns (rows, sheet_name, header_row_idx).

    Fallback: first non-empty sheet, header assumed at row 11.
    """
    style_aliases  = set(_COL_ALIASES["style_code"])
    colour_aliases = set(_COL_ALIASES["colour_code"])

    for sheet_name in wb.sheetnames:
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        log.info("[OrderForm] Sheet '%s' → %d rows", sheet_name, len(rows))
        if not rows:
            continue
        for row_idx, row in enumerate(rows[:30]):
            if not row:
                continue
            cells = {
                re.sub(r"\s+", " ", str(c).strip().lower())
                for c in row if c is not None
            }
            if cells & style_aliases and cells & colour_aliases:
                log.info("[OrderForm] Sheet '%s' matched at row %d", sheet_name, row_idx + 1)
                return rows, sheet_name, row_idx

    # Fallback — return the sheet with the most rows
    log.warning(
        "[OrderForm] No sheet matched style+colour headers. "
        "Falling back to sheet with most rows. Available sheets: %s", wb.sheetnames
    )
    best_sheet: Optional[str] = None
    best_rows:  list[tuple]   = []
    for sheet_name in wb.sheetnames:
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) > len(best_rows):
            best_sheet, best_rows = sheet_name, rows
    if best_sheet:
        log.warning("[OrderForm] Fallback: using sheet '%s' (%d rows) with header row 12",
                    best_sheet, len(best_rows))
        return best_rows, best_sheet, min(11, len(best_rows) - 1)

    return [], (wb.sheetnames[0] if wb.sheetnames else "Sheet1"), 11


def _find_header_row(rows: list[tuple]) -> int:
    """Return 0-based index of header row (the one with 'Style Code' + 'Colour Code')."""
    style_aliases  = set(_COL_ALIASES["style_code"])
    colour_aliases = set(_COL_ALIASES["colour_code"])
    for i, r in enumerate(rows[:30]):
        if not r:
            continue
        cells = {
            re.sub(r"\s+", " ", str(c).strip().lower())
            for c in r if c is not None
        }
        if cells & style_aliases and cells & colour_aliases:
            return i
    return 11  # default to row 12


def _open_wb(path: Path):
    """Open workbook; avoid read_only for xlsm (macro workbooks can be misread)."""
    use_read_only = path.suffix.lower() != ".xlsm"
    return openpyxl.load_workbook(
        str(path),
        read_only=use_read_only,
        data_only=True,
        keep_vba=False,
    )


def _gcell(r: tuple, idx: Optional[int], default: str = "") -> str:
    """Safe cell read with optional fallback index."""
    if idx is None or len(r) <= idx:
        return default
    return _s(r[idx])


def load_orderform(path: Path) -> tuple[list[Article], dict[str, list[str]]]:
    log.info("[OrderForm] Loading: %s", path.name)
    wb = _open_wb(path)
    rows, used_sheet, header_idx = _find_data_sheet_and_header(wb)
    wb.close()
    log.info("[OrderForm] Using sheet '%s' (%d rows), header at row %d",
             used_sheet, len(rows), header_idx + 1)

   
    size_ref = load_size_reference(rows)

    # Guard: if header_idx falls outside the file, nothing to parse
    if header_idx >= len(rows):
        log.warning("[OrderForm] Header idx %d out of range (file has %d rows) — skipping",
                    header_idx, len(rows))
        return [], {}

    # ── Build dynamic column map from header row ──────────────────────────────
    col_map = _build_col_map(rows[header_idx])

    def _ci(field: str, fallback: int) -> int:
        v = _find_col(col_map, field)
        return v if v is not None else fallback

    # Resolve each field's column index (fallback to original hardcoded positions)
    ci_style    = _ci("style_code",    8)
    ci_name     = _ci("style_name",    9)
    ci_colour   = _ci("colour_code",  10)
    ci_col_name = _ci("colour",       11)
    ci_collect  = _ci("collection",    1)
    ci_gender   = _ci("gender_raw",    2)
    ci_prodtype = _ci("product_type",  3)
    ci_disc     = _ci("discipline",    4)
    ci_univ     = _ci("universe",      5)
    ci_line     = _ci("line",          6)
    ci_subline  = _ci("subline",       7)
    ci_madein   = _ci("made_in",      12)
    ci_currency = _ci("currency",     13)
    ci_whl      = _ci("whl_price",    14)
    ci_retail   = _ci("retail_price", 15)
    ci_szrange  = _ci("size_range",   19)
    ci_sztype   = _ci("size_type",    20)
    ci_szcode   = _ci("size_code_ref",21)
    ci_szdesc   = _ci("size_desc",    22)
    ci_action   = _find_col(col_map, "action")  # None if not present

    log.info(
        "[OrderForm] Column map → style=%d name=%d colour=%d col_name=%d "
        "gender=%d made_in=%d whl=%d retail=%d sz_range=%d sz_code=%d action=%s",
        ci_style, ci_name, ci_colour, ci_col_name,
        ci_gender, ci_madein, ci_whl, ci_retail, ci_szrange, ci_szcode,
        ci_action,
    )

    articles: list[Article] = []
    skipped  = 0
    removed  = 0
    for i in range(header_idx + 1, len(rows)):
        r = rows[i]
        if r is None:
            continue
        # All None → stop scanning (sparse rows at bottom of file)
        if all(c is None for c in r):
            continue

        # ── Skip rows marked for removal ──────────────────────────────────
        if ci_action is not None:
            action_val = _gcell(r, ci_action).strip().lower()
            if action_val in _REMOVE_VALUES:
                removed += 1
                continue

        style_code  = _gcell(r, ci_style)
        colour_code = _gcell(r, ci_colour)
        if not style_code or not colour_code:
            skipped += 1
            continue

        row_no  = r[0] if len(r) > 0 else None
        row_num = int(row_no) if isinstance(row_no, (int, float)) else (i + 1)

        art = Article(
            row_num       = row_num,
            collection    = _gcell(r, ci_collect),
            gender_raw    = _gcell(r, ci_gender),
            product_type  = _gcell(r, ci_prodtype),
            discipline    = _gcell(r, ci_disc),
            universe      = _gcell(r, ci_univ),
            line          = _gcell(r, ci_line),
            subline       = _gcell(r, ci_subline),
            style_code    = style_code,
            style_name    = _gcell(r, ci_name),
            colour_code   = colour_code,
            colour        = _gcell(r, ci_col_name),
            made_in       = _gcell(r, ci_madein),
            currency      = _gcell(r, ci_currency) or "USD",
            whl_price     = _num(r[ci_whl])    if ci_whl    is not None and len(r) > ci_whl    else None,
            retail_price  = _num(r[ci_retail]) if ci_retail is not None and len(r) > ci_retail else None,
            size_range    = _gcell(r, ci_szrange),
            size_type     = _gcell(r, ci_sztype),
            size_code_ref = _gcell(r, ci_szcode),
            size_desc     = _gcell(r, ci_szdesc),
        )
        art.sizes = derive_sizes_for_article(art, size_ref)
        articles.append(art)

    unique_styles = {a.style_code for a in articles}
    log.info("[OrderForm] %d variants | %d unique styles | %d skipped | %d removed",
             len(articles), len(unique_styles), skipped, removed)
    return articles, size_ref


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILD
# ══════════════════════════════════════════════════════════════════════════════
def _indent(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            _indent(child, level + 1)
            if not child.tail or not child.tail.strip():
                child.tail = i + "  "
        if not child.tail or not child.tail.strip():
            child.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def _resolve_gender(raw: str, mdd: Optional[dict] = None) -> tuple[str, str, str, str]:
    """raw → (sap_code, sap_label, age_code, age_label).

    Consults MDD ``gender_lov`` first, then falls back to ``FALLBACK_GENDER``.
    """
    key = (raw or "").strip().upper()
    sap_code: Optional[str] = None
    age_code: Optional[str] = None
    if mdd:
        gv = mdd.get("gender_lov", {}).get(key)
        if gv:
            sap_code, age_code = gv
    if not sap_code:
        sap_code, age_code = FALLBACK_GENDER.get(key, ("U", "AD"))
    return sap_code, LOV_GENDER.get(sap_code, sap_code), age_code, LOV_AGE.get(age_code, age_code)


def _resolve_age_gender_from_size_range(size_range: str, gender_raw: str) -> tuple[str, str, str, str]:
    """
    Licensed articles: Both SAP Age AND SAP Gender derived from Size Range.
    Source: Mapping Detail Lotto — SAP Age & Gender sheet (License section)

    Size Range → SAP Age, SAP Gender:
      30-36        → Children, Kids (from gender_raw fallback)
      36-40        → Adults, Female
      39-45        → Adults, Male

    Returns: (age_code, age_label, gender_code, gender_label)
    """
    if not size_range:
        # fallback to gender_raw for non-size-range cases (Apparel TU etc.)
        key = (gender_raw or "").strip().upper()
        gc, gl, ac, al = _resolve_gender(key)
        return ac, al, gc, gl

    parts = size_range.replace("\u2013", "-").split("-")
    if len(parts) != 2:
        key = (gender_raw or "").strip().upper()
        gc, gl, ac, al = _resolve_gender(key)
        return ac, al, gc, gl

    try:
        lo = int(re.sub(r"[^\d]", "", parts[0]))
        hi = int(re.sub(r"[^\d]", "", parts[1]))
    except ValueError:
        key = (gender_raw or "").strip().upper()
        gc, gl, ac, al = _resolve_gender(key)
        return ac, al, gc, gl

    if hi <= 36:
        # Kids range: AT_Gender = U (K not in Stibo LOV_Gender), AT_SAPAge = CH
        return "CH", "Children", "U", "Unisex"
    if lo >= 36 and hi <= 40:
        # Female Adults
        return "AD", "Adults", "F", "Female"
    if lo >= 39:
        # Male Adults
        return "AD", "Adults", "M", "Male"

    # Apparel / TU / non-numeric — fallback to gender_raw
    key = (gender_raw or "").strip().upper()
    gc, gl, ac, al = _resolve_gender(key)
    return ac, al, gc, gl


def _resolve_division(product_type: str, mdd: Optional[dict] = None) -> tuple[str, str]:
    """product_type → (sap_division_letter, label). MDD-first, fallback after."""
    key = (product_type or "").strip().upper()
    if mdd:
        dv = mdd.get("division_lov", {}).get(key)
        if dv:
            return dv
    return FALLBACK_DIVISION.get(key, ("F", product_type or "Shoes"))


def _resolve_country(made_in: str, mdd: Optional[dict] = None) -> str:
    """Country name → ISO-2 code. MDD-first, fallback after, default 'CN'."""
    if not made_in:
        return "CN"
    key = made_in.strip().upper()
    if mdd:
        c = mdd.get("country_lov", {}).get(key)
        if c:
            return c
    return FALLBACK_COUNTRY.get(key, "CN")


def _resolve_retail_price_currency(country_code: str, mdd: Optional[dict] = None) -> str:
    """
    Resolve retail price currency from country code using MDD.
    
    Steps:
    1. Look up country_code (e.g., "ID") in "Country LOV" → get country name (e.g., "Indonesia")
    2. Look up country name in "Retail Price Currency LOV" → get currency (e.g., "IDR")
    
    Uses partial/fuzzy matching for country name lookup.
    """
    if not country_code or not mdd:
        return ""
    
    code_upper = country_code.strip().upper()
    
    # Step 1: Get country name from country ID
    country_id_to_name = mdd.get("country_id_to_name", {})
    country_name = country_id_to_name.get(code_upper, "")
    
    if not country_name:
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
    
    return ""


def _build_generic_product(
    style_code: str,
    variants:   list[Article],
    cfg:        dict,
    season_id_cl: str,
    mdd:        Optional[dict] = None,
) -> ET.Element:
    rep = variants[0]
    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]
    sea_prefix = cfg["season_prefix"]
    sea_year   = cfg["season_year"]

    # Licensed: SAP Age AND SAP Gender both derived from Size Range per Mapping Detail Lotto
    a_code, a_label, g_code, g_label = _resolve_age_gender_from_size_range(rep.size_range, rep.gender_raw)
    div_letter, div_label = _resolve_division(rep.product_type, mdd=mdd)
    coo = _resolve_country(rep.made_in, mdd=mdd)
    # Derive SAP Product Group / Category from Discipline (mapping file: SAP Product Group & Category)
    disc_key = (rep.discipline or "").strip().upper()
    pg, pc = SAP_DISCIPLINE_MAP.get(disc_key) or \
             SAP_UNIVERSE_MAP.get((rep.universe or "").strip().upper()) or \
             ("", "")
    # MAA Article Code: BrandCode + StyleCode + ColourCode  (e.g. LOT123456789)
    generic_code = f"{brand_code}{style_code}{rep.colour_code}"

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   f"PPH_{div_letter}-TempSubCat")
    _keyval(gen_el, "KEY_InboundArticle", generic_code)
    ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name").text = rep.style_name or generic_code

    _clf(gen_el, f"CLH_{cfg['brand_name_for_clh']}Articles", "CPL_Merchandiser")
    _clf(gen_el, f"{season_id_cl}UA", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")
    
    # Resolve retail price currency from country_code via MDD
    retail_price_currency = _resolve_retail_price_currency(cfg.get("country_code", ""), mdd)
    
    ctx = {
        "brand_code":           brand_code,
        "brand_name":           cfg.get("brand_name", ""),
        "sbu":                  cfg["sbu"],
        "comp_code":            cfg["comp_code"],
        "style_code":           style_code,
        "style_name":           rep.style_name,
        "colour_code":          rep.colour_code,
        "colour":               rep.colour,
        "generic_code":         generic_code,
        "gender_code":          g_code,
        "gender_label":         g_label,
        "gender_raw":           rep.gender_raw,
        "age_code":             a_code,
        "age_label":            a_label,
        "season_prefix":        sea_prefix,
        "season_year":          sea_year,
        "coo":                  coo,
        "article_type":  cfg["article_type"],
        "product_type":         rep.product_type,
        "discipline":           rep.discipline,
        "fob":            (str(int(rep.whl_price)) if float(rep.whl_price).is_integer() else str(rep.whl_price)) if rep.whl_price else "",
        "fob_currency":  rep.currency or cfg.get("currency", "USD"),
        "prod_group":    pg,
        "prod_cat":      pc,
        "retail_price_currency": retail_price_currency,
        "brand_type":     cfg.get("brand_type", ""),
        "brand_category": cfg.get("brand_category", ""),
    }
    add_generic_values(vals_el, ctx)

    # ── PRD_VariantArticle — commented out (not sent to XML for now) ─────────
    # brand_clh      = cfg.get("brand_name_for_clh", brand_name)
    # seen_var_keys: set[str] = set()
    #
    # for art in variants:
    #     color3d = color_3d(art.colour_code, mdd=mdd)
    #     v_g, v_g_lbl, v_a, v_a_lbl = _resolve_gender(art.gender_raw, mdd=mdd)
    #
    #     for sz_label in art.sizes:
    #         if not sz_label:
    #             continue
    #         sz3d   = size_3d(sz_label, mdd=mdd)
    #         var_id = f"{generic_code}{color3d}{sz3d}"
    #
    #         if var_id in seen_var_keys:
    #             log.warning("[OrderForm] Duplicate variant '%s' (style=%s colour=%s size=%s) — skipping",
    #                         var_id, style_code, art.colour_code, sz_label)
    #             continue
    #         seen_var_keys.add(var_id)
    #
    #         var_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Product")
    #         var_el.set("UserTypeID", "PRD_VariantArticle")
    #         var_el.set("ParentID",   generic_code)
    #         _keyval(var_el, "KEY_Variant", var_id)
    #         ET.SubElement(var_el, f"{{{STIBO_NS}}}Name").text = (
    #             f"{art.style_name} {art.colour.upper()} {sz_label}".strip()
    #         )
    #
    #         _clf(var_el, f"CLH_{brand_clh}Articles", "CPL_Merchandiser")
    #         _clf(var_el, f"{season_id_cl}UA",        "CPL_UnConfirmedForSeason")
    #
    #         vv = ET.SubElement(var_el, f"{{{STIBO_NS}}}Values")
    #         _val(vv, "AT_PrincipalStyleCode", art.style_code)
    #         _val(vv, "AT_SAPStyleCode",       art.style_code)
    #         _val(vv, "AT_Generic",            generic_code)
    #         _val(vv, "AT_Variant",            var_id)
    #         _val(vv, "AT_PrincipalColorCode", art.colour_code)
    #         _val(vv, "AT_SAPColorCode",       color3d)
    #         if art.colour:
    #             _val(vv, "AT_Color", art.colour.upper())
    #         _val(vv, "AT_Gender",      v_g_lbl, id_val=v_g)
    #         _val(vv, "AT_SAPAge",      v_a_lbl, id_val=v_a)
    #         _val(vv, "AT_EUSizeCode",  sz_label)
    #         _val(vv, "AT_SAPSize",     sz3d)
    #         _val(vv, "AT_PrincipalSize", sz_label)
    #         _val(vv, "AT_UOM",         "Each",  id_val="EA")

    return gen_el


def build_xml(
    articles:    list[Article],
    output_path: Path,
    cfg:         dict,
    mdd:         Optional[dict] = None,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]
    cfg["brand_name_for_clh"] = cfg.get("brand_name_for_clh") or brand_name

    sea_prefix, sea_year = derive_season(cfg["season"])
    cfg["season_prefix"] = sea_prefix
    cfg["season_year"]   = sea_year
    season_id_cl         = season_id_full(brand_code, cfg["season"])

    generics: dict[str, list[Article]] = defaultdict(list)
    for art in articles:
        generics[art.style_code].append(art)

    log.info("[XML] Building — %d generics, %d variants total",
             len(generics), len(articles))

    cls_el  = build_classifications(brand_name, brand_code, cfg["season"])
    cls_str = strip_xmlns(ET.tostring(cls_el, encoding="unicode"))

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written = 0

    with open(output_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        open_step_xml(f, export_time)
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for style_code, variants in generics.items():
            gen_el = _build_generic_product(style_code, variants, cfg, season_id_cl, mdd=mdd)
            _indent(gen_el, level=2)
            xml_str = strip_xmlns(ET.tostring(gen_el, encoding="unicode"))
            f.write(f"    {xml_str}\n")
            written += 1
        f.write("  </Products>\n")
        close_step_xml(f)

    log.info("[XML] Written → %s  (%.1f KB) — %d generics",
             output_path, output_path.stat().st_size / 1024, written)
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# LAMBDA ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def parse_config_from_filename(stem: str) -> dict:
    result: dict = {}
    parts = re.split(r"\s*-\s*", stem)
    if parts and re.match(r"^\d{4}$", parts[0].strip()):
        result["comp_code"] = parts[0].strip()
    if len(parts) > 1 and re.match(r"^[A-Z]{2,3}$", parts[1].strip(), re.IGNORECASE):
        result["sbu"] = parts[1].strip().upper()
    season_idx = None
    for i, p in enumerate(parts):
        if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
            result["season"] = p.strip().upper()
            season_idx = i
            break
    if season_idx is not None:
        for p in parts[season_idx + 1:]:
            p = p.strip()
            if re.match(r"^[A-Z]{2,3}$", p, re.IGNORECASE) and not p.isdigit():
                result["country_code"] = p.upper()
                break
    stem_lower = stem.lower()
    if "inline" in stem_lower:
        result["article_type"] = "Inline"
    elif "license" in stem_lower or " lic " in stem_lower or "lic fob" in stem_lower:
        result["article_type"] = "License"
    elif "sse" in stem_lower:
        result["article_type"] = "SSE"
    return result


def _find_orderform_files() -> list[Path]:
    """Scan ORDERFORM_DIR then INPUT_DIR for *.xlsx / *.xlsm."""
    files: list[Path] = []
    if ORDERFORM_DIR.exists():
        files.extend(ORDERFORM_DIR.glob("*.xlsx"))
        files.extend(ORDERFORM_DIR.glob("*.xlsm"))
    if not files and INPUT_DIR.exists():
        files.extend(p for p in INPUT_DIR.glob("*.xlsx") if p.parent == INPUT_DIR)
        files.extend(p for p in INPUT_DIR.glob("*.xlsm") if p.parent == INPUT_DIR)
    return files


def run(args, auditor=None):
    """
    Lambda dispatcher entry. Reads from ORDERFORM_DIR (lambda) or INPUT_DIR
    (local-dev), writes XML to XML_OUT_DIR.

    args attributes used (all optional with sane defaults):
        brand_code   default "LOT"
        brand        default "Lotto"
        comp_code    default "0888"
        sbu          default "SP"
        season       default "FW26"
        country_code default "ID"
        article_type_from_filename default "License"
    """
    of_files = _find_orderform_files()
    if not of_files:
        log.warning("[Lotto] No OrderForm xlsx files in %s — nothing to do.", ORDERFORM_DIR)
        return

    brand_code    = getattr(args, "brand_code", "LOT")  or "LOT"
    brand_name    = getattr(args, "brand",      "Lotto") or "Lotto"
    base_season   = getattr(args, "season",     "FW26")  or "FW26"
    comp_code     = getattr(args, "comp_code",  "0888")  or "0888"
    sbu           = getattr(args, "sbu",        "SP")    or "SP"
    country_code  = getattr(args, "country_code", "")    or "ID"
    article_type  = getattr(args, "article_type_from_filename", "") or "License"

    base_cfg = {
        "brand_code":           brand_code,
        "brand_name":           brand_name,
        "brand_name_for_clh":   brand_name,
        "comp_code":            comp_code,
        "sbu":                  sbu,
        "currency":             getattr(args, "currency",            "") or "USD",
        "sap_article_category": getattr(args, "sap_article_category", "1")      or "1",
        "sap_article_type":     getattr(args, "sap_article_type",    "DIRECT")  or "DIRECT",
        "by_indicator":         getattr(args, "by_indicator",         "Y")      or "Y",
        "sap_indicator":        getattr(args, "sap_indicator",        "N")      or "N",
        "nature_of_article":    getattr(args, "nature_of_article",    "REG")    or "REG",
        "uom_code":             getattr(args, "uom_code",             "EA")     or "EA",
    }

    mdd = load_mdd()
    if mdd["brand_lov"] and brand_code.upper() not in mdd["brand_lov"]:
        log.warning("[Lotto] brand_code %s not in MDD Brand LOV.", brand_code)
    elif mdd["brand_lov"]:
        log.info("[Lotto] Brand %s validated against MDD → %s",
                 brand_code, mdd["brand_lov"].get(brand_code.upper()))

    # ── Load RNA (Brand Type / Brand Category) from Attributes List ──────────
    rna = None
    attr_files = list(ATTR_DIR.glob("*.xlsx")) + list(ATTR_DIR.glob("*.xlsm"))
    if attr_files:
        rna = RNALoader(attr_files[0])
    else:
        log.warning("[Lotto] No Attributes List file found in %s — AT_BrandType/AT_BrandCategory will be empty.", ATTR_DIR)

    for of_path in of_files:
        log.info("─── Processing OrderForm: %s ───", of_path.name)
        file_cfg = parse_config_from_filename(of_path.stem)
        eff_season  = file_cfg.get("season",       base_season)
        eff_country = file_cfg.get("country_code", country_code)
        eff_type    = file_cfg.get("article_type", article_type)
        eff_comp    = file_cfg.get("comp_code",    comp_code)
        eff_sbu     = file_cfg.get("sbu",          sbu)
        log.info("[Lotto] Filename parsed → season=%s country=%s type=%s comp=%s sbu=%s",
                 eff_season, eff_country, eff_type, eff_comp, eff_sbu)

        # ── Resolve Brand Type / Brand Category from RNA ─────────────────────
        _brand_type     = ""
        _brand_category = ""
        if rna:
            _country_name = _COUNTRY_MAP.get(eff_country.upper(), eff_country.upper())
            _rna_result = rna.get(_country_name, eff_comp, eff_sbu, brand_code)
            if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
                _rna_result = rna.get(_country_name, str(int(eff_comp)) if eff_comp.isdigit() else eff_comp, eff_sbu, brand_code)
            if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
                _rna_result = rna.get_fuzzy(_country_name, eff_sbu, brand_code)
            _brand_type     = _rna_result.get("brand_type", "")
            _brand_category = _rna_result.get("brand_category", "")
            log.info("[RNA] country=%s brand_type='%s' brand_category='%s'",
                     _country_name, _brand_type, _brand_category)

        cfg = {
            **base_cfg,
            "comp_code":      eff_comp,
            "sbu":            eff_sbu,
            "season":         eff_season,
            "country_code":   eff_country,
            "article_type":   eff_type,
            "brand_type":     _brand_type,
            "brand_category": _brand_category,
        }
        articles, _ = load_orderform(of_path)
        # ── TEST MODE: limit to 5 ────────────────────────
        # TEST_PRODUCT_LIMIT = 5
        # if TEST_PRODUCT_LIMIT:
        #     seen_styles: list[str] = []
        #     for a in articles:
        #         if a.style_code not in seen_styles:
        #             if len(seen_styles) >= TEST_PRODUCT_LIMIT:
        #                 break
        #             seen_styles.append(a.style_code)
        #     keep = set(seen_styles)
        #     before = len(articles)
        #     articles = [a for a in articles if a.style_code in keep]
        #     log.info("[Lotto] TEST MODE: limited to %d products / %d rows (was %d rows)",
        #             len(keep), len(articles), before)
        if not articles:
            log.warning("[Lotto] No articles parsed — skipping %s", of_path.name)
            continue
        out_path = XML_OUT_DIR / f"{of_path.stem}.xml"
        build_xml(articles, out_path, cfg, mdd=mdd)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Lotto OrderForm → STEP XML")
    ap.add_argument("--brand",        default="Lotto")
    ap.add_argument("--brand-code",   default="LOT", dest="brand_code")
    ap.add_argument("--season",       default="FW26")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="SP")
    ap.add_argument("--country-code", default="ID",  dest="country_code")
    ap.add_argument("--article-type", default="License", dest="article_type_from_filename")
    run(ap.parse_args())
