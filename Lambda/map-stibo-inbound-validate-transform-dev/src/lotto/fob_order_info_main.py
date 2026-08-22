"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — Lotto FOB Order Info Update v1.0             ║
║  "LOTTO SPA- FW25 FOB order info update DDMMYY.xlsm" → Stibo STEP XML      ║
╚══════════════════════════════════════════════════════════════════════════════╝

Source file format : "LOTTO SPA- FW25 FOB order info update 170725.xlsm"
Sheet name         : "WGSFL1"  (auto-detected as first non-empty sheet)

Layout (Row 1 = headers, Rows 2+ = data — one row per size variant):
  Col  0  : Product code        (style code, numeric, e.g. 222837)
  Col  1  : Description         (style name, e.g. "TACTO 700 ID JR")
  Col  2  : Color code          (3-char, e.g. "DKI")
  Col  3  : Color description   (e.g. "BORROWED BLUE/ALL WHITE/NAVY BLUE")
  Col  4  : SKU                 (style_code + colour_code, e.g. "222837DKI")
  Col  5  : Size                (EU size label; comma for half: "8,5")
  Col  6  : Qty                 (ordered quantity)
  Col  7  : EAN 13              (string — preferred)
  Col  8  : EAN 13              (numeric — fallback when col 7 is empty)
  Col  9  : Gender              (JUNIOR / UNISEX / MALE / FEMALE)
  Col 10  : Made in             (CHINA / VIETNAM / …)
  Col 11  : Composition upper   (English, e.g. "PVC:80.0% - Polyester Textile:20.0%")
  Col 12  : Composition upper   (Indonesian)
  Col 13  : Composition Outsole (English)
  Col 14  : Composition Outsole (Indonesian)
  Col 15  : Composition Lining  (English)
  Col 16  : HS code             (e.g. "64021900")
  Col 17  : Artikel ID          (existing MAA article ID, informational)

Hierarchy emitted (Stibo STEP XML):
  PRD_GenericArticle  (style_code)
    └─ PRD_Variant       (style_code + colour_code)
        └─ PRD_SizeVariant  (variant_id + size_3d — one per source row)
               AT_Barcode     = EAN 13
               AT_BarcodeType = P   (EAN-13)

Key differences vs. orderform_main:
  • Flat format — every row is already a resolved size+EAN pair (no size ref table).
  • AT_Barcode / AT_BarcodeType are emitted on every PRD_SizeVariant.
  • Composition attributes (AT_MaterialUpper, AT_Material, AT_OutsoleComposition)
    are written at the PRD_Variant level (same for all sizes of a colourway).
  • AT_HSCode is written at the PRD_GenericArticle level.
  • Reads .xlsm files via openpyxl (keep_vba=True, data_only=True).
  • Season is extracted from the filename (e.g. "FW25" → season="FW25").
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
# STIBO / STEP XML HELPERS  (self-contained — no cross-brand imports)
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')

LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE    = {
    "AD": "Adults",  "CH": "Children", "IN": "Infant",
    "JR": "Junior",  "AA": "All Ages",
}
LOV_SEASON_NAME = {
    "SS": "Spring Summer",  "FW": "Fall Winter",
    "AW": "Autumn Winter",  "HO": "Holiday",    "AL": "All Season",
}


def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "") -> Optional[ET.Element]:
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
    return _XMLNS_RE.sub("", xml_str)


def open_step_xml(f, export_time: str) -> None:
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
    """'FW25' → ('FW', '2025'). Handles 2-digit or 4-digit year tail."""
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
    pfx, yr = derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"


def build_classifications(brand_name: str, brand_code: str, season_code: str) -> ET.Element:
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


# ══════════════════════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR          = BASE_DIR / "input"
FOB_ORDER_INFO_DIR = INPUT_DIR / "fob_order_info"
MDD_DIR            = INPUT_DIR / "mdd"
ATTR_DIR           = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in (INPUT_DIR, FOB_ORDER_INFO_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"fob_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("lotto.fob_order_info_main")


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS / FALLBACK MAPS
# ══════════════════════════════════════════════════════════════════════════════
FALLBACK_GENDER: dict[str, tuple[str, str]] = {
    "M": ("M", "AD"),    "MAN": ("M", "AD"),   "MEN": ("M", "AD"),
    "MALE": ("M", "AD"),
    "W": ("F", "AD"),     "WOMAN": ("F", "AD"),  "WOMEN": ("F", "AD"),
    "FEMALE": ("F", "AD"), "F": ("F", "AD"),
    "U": ("U", "AD"),     "UNISEX": ("U", "AD"),  "ADULT": ("U", "AD"),
    "ADULTS": ("U", "AD"), "SENIOR": ("U", "AD"),
    "K": ("U", "CH"),     "KIDS": ("U", "CH"),    "CHILD": ("U", "CH"),
    "JUNIOR": ("U", "CH"), "JR": ("U", "CH"),
    "INFANT": ("U", "IN"), "BABY": ("U", "IN"),
}

FALLBACK_COUNTRY: dict[str, str] = {
    "CHINA": "CN",       "VIETNAM": "VN",      "INDONESIA": "ID",
    "CAMBODIA": "KH",    "BANGLADESH": "BD",   "INDIA": "IN",
    "MALAYSIA": "MY",    "THAILAND": "TH",     "TUNISIA": "TN",
    "ITALY": "IT",       "PHILIPPINES": "PH",  "PAKISTAN": "PK",
}

_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(v) -> str:
    """Safe string conversion — strips whitespace, collapses integer floats."""
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


def color_3d(code: str) -> str:
    """Normalize colour code to exactly 3 alphanumeric chars."""
    s = _ALNUM_RE.sub("", code or "").upper()
    if not s:
        return "000"
    return s[:3] if len(s) >= 3 else s.zfill(3)


def size_3d(label: str) -> str:
    """Normalize size label to exactly 3 alphanumeric chars.

    Half-sizes with comma/dot (e.g. '8,5' / '8.5') → '08H'.
    """
    raw = (label or "").strip()
    if raw:
        m = re.match(r"^(\d+)[,\.]5$", raw)
        if m:
            whole = m.group(1)
            return (whole + "H") if len(whole) >= 2 else (whole.zfill(2) + "H")
    s = _ALNUM_RE.sub("", raw).upper()
    if not s:
        return "000"
    return s[:3] if len(s) >= 3 else s.zfill(3)


def _resolve_ean(col7, col8) -> str:
    """Return best EAN 13 string: prefer col7 (string), fallback col8 (numeric)."""
    def _fmt_ean(v) -> str:
        if v is None:
            return ""
        try:
            return str(int(float(str(v).strip())))
        except (ValueError, TypeError):
            return str(v).strip()

    s7 = _fmt_ean(col7)
    if s7 and s7 not in ("None", "nan", "0"):
        return s7

    s8 = _fmt_ean(col8)
    if s8 and s8 not in ("None", "nan", "0"):
        return s8

    return ""


def _resolve_gender(raw: str) -> tuple[str, str, str, str]:
    """raw → (sap_gender_code, sap_gender_label, age_code, age_label)."""
    key = (raw or "").strip().upper()
    sap_code, age_code = FALLBACK_GENDER.get(key, ("U", "AD"))
    return sap_code, LOV_GENDER.get(sap_code, sap_code), age_code, LOV_AGE.get(age_code, age_code)


def _resolve_country(made_in: str) -> str:
    """Country name → ISO-2 code."""
    if not made_in:
        return "CN"
    return FALLBACK_COUNTRY.get(made_in.strip().upper(), "CN")


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class SizeRow:
    """Represents one row from the source file (one size variant with EAN)."""
    style_code:      str
    style_name:      str
    colour_code:     str
    colour_desc:     str
    sku:             str
    size_label:      str
    qty:             int
    ean13:           str          # EAN-13 barcode string
    gender_raw:      str
    made_in:         str
    comp_upper_en:   str          # Composition upper (English)
    comp_upper_id:   str          # Composition upper (Indonesian)
    comp_outsole_en: str          # Composition outsole (English)
    comp_outsole_id: str          # Composition outsole (Indonesian)
    comp_lining:     str          # Composition lining
    hs_code:         str
    artikel_id:      str


@dataclass
class ColourGroup:
    """All size rows for one (style_code, colour_code) combination."""
    style_code:  str
    style_name:  str
    colour_code: str
    colour_desc: str
    gender_raw:  str
    made_in:     str
    comp_upper_en:   str
    comp_outsole_en: str
    comp_lining:     str
    hs_code:         str
    sizes:       list[SizeRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADER
# ══════════════════════════════════════════════════════════════════════════════
def _find_input_files() -> list[Path]:
    """Locate .xlsm / .xlsx files in FOB_ORDER_INFO_DIR, fallback INPUT_DIR root."""
    files: list[Path] = []
    if FOB_ORDER_INFO_DIR.exists():
        files.extend(FOB_ORDER_INFO_DIR.glob("*.xlsm"))
        files.extend(FOB_ORDER_INFO_DIR.glob("*.xlsx"))
    if not files and INPUT_DIR.exists():
        files.extend(p for p in INPUT_DIR.glob("*.xlsm") if p.parent == INPUT_DIR)
        files.extend(p for p in INPUT_DIR.glob("*.xlsx") if p.parent == INPUT_DIR)
    return sorted(files)


def load_fob_order_info(path: Path) -> list[SizeRow]:
    """Parse the FOB Order Info Update file into a list of SizeRow objects.

    Supports .xlsm and .xlsx.  Uses the first sheet (usually 'WGSFL1').
    Row 1 is the header and is skipped.  Rows with no style code are skipped.
    """
    log.info("[FOB] Loading: %s", path.name)

    # openpyxl can read .xlsm directly when keep_vba=True
    wb = openpyxl.load_workbook(str(path), keep_vba=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=2, values_only=True))   # skip header row 1
    wb.close()

    result: list[SizeRow] = []
    skipped = 0

    for i, r in enumerate(rows, start=2):
        if r is None or all(c is None for c in r):
            continue

        style_raw = _s(r[0]) if len(r) > 0 else ""
        if not style_raw or style_raw in ("None", "nan"):
            skipped += 1
            continue

        colour_code = _s(r[2]) if len(r) > 2 else ""
        if not colour_code:
            skipped += 1
            continue

        ean = _resolve_ean(r[7] if len(r) > 7 else None,
                           r[8] if len(r) > 8 else None)

        row = SizeRow(
            style_code      = style_raw,
            style_name      = _s(r[1])  if len(r) > 1  else "",
            colour_code     = colour_code,
            colour_desc     = _s(r[3])  if len(r) > 3  else "",
            sku             = _s(r[4])  if len(r) > 4  else "",
            size_label      = _s(r[5])  if len(r) > 5  else "",
            qty             = int(_num(r[6]) or 0) if len(r) > 6 else 0,
            ean13           = ean,
            gender_raw      = _s(r[9])  if len(r) > 9  else "",
            made_in         = _s(r[10]) if len(r) > 10 else "",
            comp_upper_en   = _s(r[11]) if len(r) > 11 else "",
            comp_upper_id   = _s(r[12]) if len(r) > 12 else "",
            comp_outsole_en = _s(r[13]) if len(r) > 13 else "",
            comp_outsole_id = _s(r[14]) if len(r) > 14 else "",
            comp_lining     = _s(r[15]) if len(r) > 15 else "",
            hs_code         = _s(r[16]) if len(r) > 16 else "",
            artikel_id      = _s(r[17]) if len(r) > 17 else "",
        )
        result.append(row)

    log.info("[FOB] Parsed %d size rows | %d skipped", len(result), skipped)
    return result


def group_by_colour(rows: list[SizeRow]) -> dict[str, dict[str, ColourGroup]]:
    """Group rows: { style_code → { colour_code → ColourGroup } }."""
    groups: dict[str, dict[str, ColourGroup]] = defaultdict(dict)

    for r in rows:
        sc = r.style_code
        cc = r.colour_code
        if cc not in groups[sc]:
            groups[sc][cc] = ColourGroup(
                style_code      = sc,
                style_name      = r.style_name,
                colour_code     = cc,
                colour_desc     = r.colour_desc,
                gender_raw      = r.gender_raw,
                made_in         = r.made_in,
                comp_upper_en   = r.comp_upper_en,
                comp_outsole_en = r.comp_outsole_en,
                comp_lining     = r.comp_lining,
                hs_code         = r.hs_code,
            )
        groups[sc][cc].sizes.append(r)

    return groups


# ══════════════════════════════════════════════════════════════════════════════
# XML INDENT
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


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILD
# ══════════════════════════════════════════════════════════════════════════════
def _build_generic_product(
    style_code:  str,
    color_code:  str,
    cg:          ColourGroup,
    cfg:         dict,
    season_id_cl: str,
) -> ET.Element:
    """Build one PRD_GenericArticle element with its PRD_Variant / PRD_SizeVariant children."""
    brand_code   = cfg["brand_code"]
    brand_name   = cfg["brand_name"]
    sea_prefix   = cfg["season_prefix"]
    sea_year     = cfg["season_year"]
    brand_clh    = cfg.get("brand_name_for_clh", brand_name)

    # Use the first size row of this color as the representative for generic-level attributes
    rep_row = cg.sizes[0] if cg.sizes else None

    g_code, g_label, a_code, a_label = _resolve_gender(
        rep_row.gender_raw if rep_row else ""
    )
    coo = _resolve_country(rep_row.made_in if rep_row else "")

    color3d = color_3d(color_code)
    # MAA Article Code at generic level: BrandCode + StyleCode + ColorCode
    generic_code = f"{brand_code}{style_code}{color3d}"

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_F-TempSubCat")
    _keyval(gen_el, "KEY_InboundArticle", generic_code)
    # ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name").text = cg.style_name or generic_code

    # _clf(gen_el, f"CLH_{brand_clh}Articles",  "CPL_Merchandiser")
    # _clf(gen_el, f"{season_id_cl}UA",         "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")
    # _val(vals_el, "AT_Brand",                     id_val=brand_code)
    # _val(vals_el, "AT_SBU",                       id_val=cfg.get("sbu", "SP"))
    # _val(vals_el, "AT_CompanyCode",               id_val=cfg.get("comp_code", "0888"))
    # _val(vals_el, "AT_CountryOrigin",             id_val=coo)
    # # _val(vals_el, "AT_PrincipalStyleCode",        style_code)
    # # _val(vals_el, "AT_PrincipalStyleDescription", cg.style_name)
    # _val(vals_el, "AT_SAPStyleCode",              style_code)
    # _val(vals_el, "AT_InboundGenericCode",         generic_code)
    # _val(vals_el, "AT_Gender",     g_label, id_val=g_code)
    # _val(vals_el, "AT_SAPAge",     a_label, id_val=a_code)
    # _val(vals_el, "AT_Season",     LOV_SEASON_NAME.get(sea_prefix, sea_prefix),
    #                                id_val=sea_prefix)
    # _val(vals_el, "AT_SeasonYear", sea_year)
    # _val(vals_el, "AT_BYIndicator",         "Yes",     id_val="Y")
    # _val(vals_el, "AT_SAPIndicator",        "No",      id_val="N")
    # _val(vals_el, "AT_SAPProductFlag",      "A",       id_val="A")
    # _val(vals_el, "AT_MaterialType",        "", id_val="ZINA")
    # _val(vals_el, "AT_CountrySize",         "", id_val="EU")
    # _val(vals_el, "AT_UOM",                 "Each",    id_val="EA")
    # # HS Code is consistent for a style — write at generic level
    # if cg.hs_code:
    #     _val(vals_el, "AT_HSCode", cg.hs_code)

    # ── PRD_VariantArticle (one per size — flat 2-level hierarchy) ───
    v_g, v_g_lbl, v_a, v_a_lbl = _resolve_gender(cg.gender_raw)

    seen_sizes: set[str] = set()
    for sz_row in cg.sizes:
        sz_norm = sz_row.size_label
        if not sz_norm:
            continue
        sz3d = size_3d(sz_norm)
        if sz3d in seen_sizes:
            log.warning(
                "[FOB] Duplicate size '%s' (→%s) in %s/%s — skipping",
                sz_norm, sz3d, cg.style_code, color_code,
            )
            continue
        seen_sizes.add(sz3d)

        var_id = f"{generic_code}{sz3d}"

        var_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Product")
        var_el.set("UserTypeID", "PRD_VariantArticle")
        _keyval(var_el, "KEY_InboundVariant", var_id)
        # ET.SubElement(var_el, f"{{{STIBO_NS}}}Name").text = f"Size {sz_norm}"

        vv = ET.SubElement(var_el, f"{{{STIBO_NS}}}Values")
        # _val(vv, "AT_PrincipalStyleDescription", cg.style_name)
        # _val(vv, "AT_SAPStyleCode",              cg.style_code)
        # # AT_InboundGenericCode not valid on PRD_VariantArticle — generic level only
        # # _val(vv, "AT_PrincipalColorCode",        color_code)
        # _val(vv, "AT_Gender",          v_g_lbl, id_val=v_g)
        # _val(vv, "AT_SAPAge",          v_a_lbl, id_val=v_a)
        # _val(vv, "AT_PrincipalSize",   sz_norm)
        # _val(vv, "AT_UOM",             "Each", id_val="EA")
        # if sz_row.ean13:
        #     _val(vv, "AT_PrincipalBarcode", sz_row.ean13)

        if sz_row.ean13:
            dcs_el = ET.SubElement(var_el, f"{{{STIBO_NS}}}DataContainers")
            mdc_el = ET.SubElement(dcs_el, f"{{{STIBO_NS}}}MultiDataContainer")
            mdc_el.set("Type", "DC_Barcode")
            dc_el  = ET.SubElement(mdc_el, f"{{{STIBO_NS}}}DataContainer")
            dc_el.set("Analyzer", "true")
            dcv_el = ET.SubElement(dc_el, f"{{{STIBO_NS}}}Values")
            _val(dcv_el, "AT_Barcode",          sz_row.ean13)
            _val(dcv_el, "AT_BarcodeType",      id_val="P")
            _val(dcv_el, "AT_MainEANIndicator", id_val="Y")

    return gen_el


def build_xml(
    rows:        list[SizeRow],
    output_path: Path,
    cfg:         dict,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]
    cfg["brand_name_for_clh"] = cfg.get("brand_name_for_clh") or brand_name

    sea_prefix, sea_year = derive_season(cfg["season"])
    cfg["season_prefix"] = sea_prefix
    cfg["season_year"]   = sea_year
    season_id_cl         = season_id_full(brand_code, cfg["season"])

    groups = group_by_colour(rows)

    # ── TEST LIMITER (Only 2 generics, and 2 variants per generic) ─────────────────
    # TEST_MODE = True  # Set to False or comment out for production
    # if TEST_MODE:
    #     all_combos = []
    #     for style_code, colour_map in groups.items():
    #         for color_code, cg in colour_map.items():
    #             cg.sizes = cg.sizes[:2]
    #             all_combos.append((style_code, color_code, cg))
        
    #     all_combos = all_combos[:2]
        
    #     groups = {}
    #     for style_code, color_code, cg in all_combos:
    #         if style_code not in groups:
    #             groups[style_code] = {}
    #         groups[style_code][color_code] = cg
    #     log.info("[XML] TEST MODE active: Limited to 2 generic articles and 2 variants per generic")
    # ────────────────────────────────────────────────────────────────────────────────

    log.info(
        "[XML] Building — %d generic articles, %d colour groups, %d size rows",
        len(groups),
        sum(len(v) for v in groups.values()),
        len(rows),
    )

    cls_el  = build_classifications(brand_name, brand_code, cfg["season"])
    cls_str = strip_xmlns(ET.tostring(cls_el, encoding="unicode"))

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written = 0

    with open(output_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        open_step_xml(f, export_time)
        # f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for style_code, colour_map in groups.items():
            for color_code, cg in colour_map.items():
                gen_el = _build_generic_product(style_code, color_code, cg, cfg, season_id_cl)
                _indent(gen_el, level=2)
                xml_str = strip_xmlns(ET.tostring(gen_el, encoding="unicode"))
                f.write(f"    {xml_str}\n")
                written += 1
        f.write("  </Products>\n")
        close_step_xml(f)

    log.info(
        "[XML] Written → %s  (%.1f KB) — %d generic articles",
        output_path,
        output_path.stat().st_size / 1024,
        written,
    )
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# LAMBDA ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def _extract_season_from_filename(filename: str) -> str:
    """Extract season token from filename, e.g. 'FW25' from 'LOTTO SPA- FW25 FOB …'."""
    m = re.search(r"\b([A-Z]{2}\d{2,4})\b", filename, re.IGNORECASE)
    return m.group(1).upper() if m else "FW25"


def run(args, auditor=None):
    """
    Lambda dispatcher entry point for FOB Order Info Update files.

    args attributes used (all optional, with sane defaults):
        brand_code   default "LOT"
        brand        default "Lotto"
        comp_code    default "0888"
        sbu          default "SP"
        season       default auto-detected from filename, else "FW25"
        country_code default "ID"
        article_type_from_filename default "License"
    """
    files = _find_input_files()
    if not files:
        log.warning(
            "[Lotto-FOB] No FOB Order Info .xlsm/.xlsx files in %s — nothing to do.",
            FOB_ORDER_INFO_DIR,
        )
        return

    brand_code   = getattr(args, "brand_code", "LOT") or "LOT"
    brand_name   = getattr(args, "brand",      "Lotto") or "Lotto"
    comp_code    = getattr(args, "comp_code",  "0888")  or "0888"
    sbu          = getattr(args, "sbu",        "SP")    or "SP"
    country_code = getattr(args, "country_code", "ID")  or "ID"

    log.info(
        "[Lotto-FOB] Pipeline → brand=%s code=%s comp=%s sbu=%s country=%s",
        brand_name, brand_code, comp_code, sbu, country_code,
    )

    for file_path in files:
        log.info("─── Processing FOB Order Info: %s ───", file_path.name)

        # Season: args first, then extract from filename
        season = (getattr(args, "season", None) or "").strip()
        if not season or season == "FW26":
            season = _extract_season_from_filename(file_path.stem)

        cfg = {
            "brand_code":         brand_code,
            "brand_name":         brand_name,
            "brand_name_for_clh": brand_name,
            "season":             season,
            "comp_code":          comp_code,
            "sbu":                sbu,
            "country_code":       country_code,
            "article_type":       getattr(args, "article_type_from_filename", "License") or "License",
            "currency":           getattr(args, "currency", "USD") or "USD",
        }
        log.info("[Lotto-FOB] Config: %s", cfg)

        rows = load_fob_order_info(file_path)
        # ── TEST MODE: limit to 5 ────────────────────────
        # rows = rows[:5]
        if not rows:
            log.warning("[Lotto-FOB] No rows parsed — skipping %s", file_path.name)
            continue

        out_path = XML_OUT_DIR / f"{file_path.stem}.xml"
        build_xml(rows, out_path, cfg)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Lotto FOB Order Info Update → STEP XML")
    ap.add_argument("--brand",        default="Lotto")
    ap.add_argument("--brand-code",   default="LOT",  dest="brand_code")
    ap.add_argument("--season",       default="",     help="e.g. FW25 (auto-detected from filename if blank)")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="SP")
    ap.add_argument("--country-code", default="ID",   dest="country_code")
    ap.add_argument("--article-type", default="License", dest="article_type_from_filename")
    run(ap.parse_args())
 