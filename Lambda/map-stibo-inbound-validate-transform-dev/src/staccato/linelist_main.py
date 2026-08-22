"""
╔══════════════════════════════════════════════════════════════════════════════╗
║        STIBO INBOUND XML GENERATOR — Staccato Line List (Inline) v1.0         ║
║  "STACCATO 26Summer linesheet.xlsx" → Stibo STEP XML                         ║
╚══════════════════════════════════════════════════════════════════════════════╝

Source file format : Staccato Line Sheet Excel (.xlsx / .xlsm)
Sheet name         : "linesheet" or first non-empty data sheet

MAPPED ATTRIBUTES (from Staccato(Inline) mapping template):
─────────────────────────────────────────────────────────────────────────────────
Stibo Attribute ID                | Field Name in Brand File / Value
─────────────────────────────────────────────────────────────────────────────────
AT_PrincipalStyleCode             | STYLE NO.
AT_PrincipalStyleDescription      | PRODUCE NAME
AT_PrincipalColorName             | COLOR
AT_PrincipalColorCode             | COLOR
AT_FOB                            | FOB
AT_FOBCurrency                    | USD (Default)
AT_PrincipalMerchandiseHierarchyL1|  MAJOR CATEGORY / MAJOR CATEGORY
AT_PrincipalMerchandiseHierarchyL2| CATEGORY
AT_HeelHeight                     | HEEL LEVEL
AT_Collection1                    | THEME
AT_Gender                         | GENDER (Mapped / Default Female - F)
AT_SAPAge                         | GENDER (Mapped / Default Adults - AD)
AT_BYGender                       | GENDER (Mapped / Default Female)
AT_BYAge                          | GENDER (Mapped / Default Adults)
AT_CountryOrigin                  | CN (Default: China)
AT_SAPProductFlag                 | A (Default: Intercompany)
AT_MaterialType                   | ZINA (Default: Intercompany)
AT_SAPArticleCategory             | 01 (Default: Generic / Variant)
AT_UOM                            | EA (Default: EA)
AT_ArticleType                    | Inline (Default)
AT_CountrySize                    | EUR (Default: Footwear) / NS (Default: Accessories)
AT_BCI                            | Commercial (Default)
AT_ArticleStatus                  | Active (Default — MDD LOV ID)
─────────────────────────────────────────────────────────────────────────────────
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
# STIBO / STEP XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')

_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")


def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "") -> Optional[ET.Element]:
    """Create a Value element under parent. Returns None if both value and id_val are empty/invalid."""
    clean_id = str(id_val).strip() if id_val else ""
    has_id = clean_id and clean_id not in ("None", "nan")
    
    clean_val = str(value).strip() if value else ""
    has_val = clean_val and clean_val not in ("None", "nan")
    
    # Skip entirely if both are empty
    if not has_id and not has_val:
        return None
    
    # Create element only if at least one is valid
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if has_id:
        el.set("ID", clean_id)
    if has_val:
        el.text = clean_val
    return el


def _keyval(parent: ET.Element, key_id: str, value: str) -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}KeyValue")
    el.set("KeyID", key_id)
    el.text = value
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> Optional[ET.Element]:
    """Create a MultiValue element with one Value child using ID."""
    clean_id = str(id_val).strip() if id_val else ""
    if not clean_id or clean_id in ("None", "nan"):
        return None
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", clean_id)
    return mv


def _clf(parent: ET.Element, class_id: str, ref_type: str) -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}ClassificationReference")
    el.set("ClassificationID", class_id)
    el.set("Type", ref_type)
    return el


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


def _get_dirs():
    base = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))
    in_dir   = base / "input"
    ll_dir   = in_dir / "linelist"
    mdd_dir  = in_dir / "mdd"
    attr_dir = in_dir / "attributes"
    out_dir  = base / "output"
    xml_dir  = out_dir / "xml"
    log_dir  = out_dir / "logs"

    for d in (in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir):
        d.mkdir(parents=True, exist_ok=True)
    return base, in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("staccato.linelist_main")


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class StaccatoRow:
    row_num:           int
    brand:             str = ""
    year:              str = ""
    season:            str = ""
    style_no:          str = ""
    produce_name:      str = ""
    principal_style_description: str = ""
    factory_name:      str = ""
    upper_material:    str = ""
    theme:             str = ""
    major_category:    str = ""
    category:          str = ""
    gender:            str = ""
    country:           str = ""
    color:             str = ""
    color_chinese:     str = ""
    materials:         str = ""
    size_run:          str = ""
    heel_height:       str = ""
    heel_level:        str = ""
    product_level:     str = ""
    fob:               Optional[float] = None
    tag_price:         Optional[float] = None
    product_barcode:   str = ""
    photo:             str = ""


@dataclass
class StaccatoGenericGroup:
    style_no:          str
    color:             str
    produce_name:      str = ""
    principal_style_description: str = ""
    major_category:    str = ""
    category:          str = ""
    gender:            str = ""
    country:           str = ""
    season:            str = ""
    year:              str = ""
    size_run:          str = ""
    theme:             str = ""
    heel_level:        str = ""
    fob:               Optional[float] = None
    rows:              list[StaccatoRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER  — reads LOVs from Master Data Dictionary Excel at runtime
# ══════════════════════════════════════════════════════════════════════════════
class MDDLoader:
    """
    Loads LOVs from the shared Master Data Dictionary Excel placed in
    staccato/input/mdd/.

    Relevant sheets (auto-detected by name keyword):
      "Heel Height LOV"   → self.lovs["HEEL HEIGHT"]
      "Country Size LOV"  → self.lovs["COUNTRY SIZE"]
      any sheet with AGE  → self.lovs["BYAGE"] + self.lovs["SAPAGE"]
      "*Gender*LOV*"      → self.lovs["BY GENDER"]
      "Simple LOVs"       → merged into respective self.lovs keys
    """

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict[str, str]] = {}   # KEY → {DISPLAY_UPPER: id}
        self._load()

    # ── public helper ──────────────────────────────────────────────────────
    def lookup(self, lov_key: str, raw: str, fallback: str = "") -> str:
        """Return LOV ID for raw display value; empty string if not found."""
        lov = self.lovs.get(lov_key.upper(), {})
        return lov.get((raw or "").strip().upper(), fallback)

    # ── internal loaders ───────────────────────────────────────────────────
    def _load(self) -> None:
        log.info("[MDD] Loading: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(str(self.path), read_only=True, data_only=True)
        except Exception as exc:
            log.error("[MDD] Cannot open file: %s", exc)
            return
        try:
            for sn in wb.sheetnames:
                snu = sn.upper()
                rows = list(wb[sn].iter_rows(values_only=True))
                if len(rows) < 2:
                    continue
                if snu == "SIMPLE LOVS":
                    self._load_simple_lovs(rows)
                elif "AGE" in snu and "LOV" in snu:
                    self._load_age_lov(rows)
                elif "HEEL HEIGHT" in snu:
                    # Heel Height LOV has Code in col-A, Name in col-B (reversed)
                    self._load_code_first_lov("HEEL HEIGHT", rows)
                else:
                    # Derive key: strip " LOV" / "LOV " suffix/prefix
                    key = snu.replace(" LOV", "").replace("LOV ", "").replace("LOV", "").strip()
                    self._load_two_col(key, rows)
        finally:
            wb.close()
        log.info("[MDD] LOVs loaded — sheets: %s", list(self.lovs.keys()))

    def _load_code_first_lov(self, key: str, rows: list) -> None:
        """Parse sheet with (col-A = LOV code/ID, col-B = display name).
        e.g. Heel Height LOV: F=Flat, H=High, L=Low, M=Medium.
        Stores both code→code and display→code so lookup works either way.
        """
        sub: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            code    = str(row[0]).strip() if row[0] is not None else ""
            display = str(row[1]).strip() if row[1] is not None else ""
            if code and display and code.upper() not in ("NONE", "NAN", "CODE"):
                sub[code.upper()]    = code    # "F" → "F"
                sub[display.upper()] = code    # "FLAT" → "F"
        if sub:
            self.lovs.setdefault(key, {}).update(sub)

    def _load_two_col(self, key: str, rows: list) -> None:
        """Parse sheet with (col-A = display value, col-B = LOV ID)."""
        sub: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] is not None else ""
            id_val  = str(row[1]).strip() if row[1] is not None else ""
            if display and id_val and display.upper() not in ("NONE", "NAN"):
                sub[display.upper()] = id_val   # lookup by display
                sub[id_val.upper()]  = id_val   # lookup by ID itself
        if sub:
            self.lovs.setdefault(key, {}).update(sub)

    def _load_age_lov(self, rows: list) -> None:
        """
        Age LOV multi-column format (mirrors Nike sp27 MDDLoader):
          col A (0) = display (Children / Adults / All Ages)
          col C (2) = SAP Age Code
          col F (5) = BY Age LOV ID  (ADULT / KIDS / ALL AGES …)
        """
        sap_age: dict[str, str] = {}
        by_age:  dict[str, str] = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            display      = str(row[0]).strip()
            sap_code     = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            by_age_id    = str(row[5]).strip() if len(row) > 5 and row[5] else ""
            if display and sap_code:
                sap_age[display.upper()] = sap_code
            if display and by_age_id:
                by_age[display.upper()]   = by_age_id
                by_age[by_age_id.upper()] = by_age_id  # also key by ID
        if sap_age:
            self.lovs.setdefault("SAPAGE", {}).update(sap_age)
        if by_age:
            self.lovs.setdefault("BYAGE", {}).update(by_age)
        log.info("[MDD] Age LOV — SAPAge: %d, BYAge: %d", len(sap_age), len(by_age))

    def _load_simple_lovs(self, rows: list) -> None:
        """Simple LOVs sheet: (lov_name, _, value_display, value_id)."""
        for row in rows[1:]:
            lov_name, _, val_name, val_id = (list(row) + [None] * 4)[:4]
            if not lov_name or not val_name:
                continue
            key     = str(lov_name).strip().upper()
            display = str(val_name).strip()
            id_v    = str(val_id).strip() if val_id else display
            sub = self.lovs.setdefault(key, {})
            sub[display.upper()] = id_v
            sub[id_v.upper()]    = id_v


def _find_mdd_file(mdd_dir: Path) -> Optional[Path]:
    """Return first .xlsx/.xlsm file found in mdd_dir, or None."""
    for suffix in (".xlsx", ".xlsm"):
        files = sorted(mdd_dir.glob(f"*{suffix}"))
        if files:
            return files[0]
    return None


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING
# ══════════════════════════════════════════════════════════════════════════════
COLUMN_ALIASES: dict[str, list[str]] = {
    "brand":          ["brand"],
    "year":           ["year"],
    "season":         ["season"],
    "style_no":       ["style no.", "style no", "style_no", "style"],
    "produce_name":   ["produce name", "product name", "description", "style description"],
    "principal_style_description": ["principal style description"],
    "factory_name":   ["factory name"],
    "upper_material": ["帮面材料", "upper material"],
    "theme":          ["theme", "collection", "collection 1", "collection1",
                       "design theme", "design", "collection/theme", "theme/collection"],
    "major_category": [" major category", "major category", "major_category"],
    "category":       ["category", "sub category"],
    "gender":         ["gender", "sex"],
    "country":        ["country", "origin country", "country of origin"],
    "color":          ["color", "colour"],
    "color_chinese":  ["颜色"],
    "materials":      ["materials", "material"],
    "size_run":       ["size run", "sizes", "size_run"],
    "heel_height":    ["heel/cm", "heel cm", "heel height"],
    "heel_level":     ["heel level", "heel_level"],
    "product_level":  ["product level"],
    "fob":            ["fob", "fob price", "fob(usd)"],
    "tag_price":      ["tag price", "retail price"],
    "product_barcode":["product barcode", "barcode"],
    "photo":          ["photo", "picture", "image"],
}


# ══════════════════════════════════════════════════════════════════════════════
# SEASON LOV MAP  (MDD: Season LOV tab)
# ══════════════════════════════════════════════════════════════════════════════
SEASON_LOV_MAP: dict[str, str] = {
    # Direct LOV codes
    "SP": "SP", "SM": "SM", "FL": "FL", "WN": "WN",
    "CO": "CO", "SS": "SS", "FW": "FW", "AL": "AL",
    # Common label variations
    "SPRING":         "SP",
    "SUMMER":         "SM",
    "SU":             "SM",  # legacy alias
    "FALL":           "FL",
    "AUTUMN":         "FL",
    "AU":             "FL",
    "WINTER":         "WN",
    "WI":             "WN",
    "CORE":           "CO",
    "SPRING-SUMMER":  "SS",
    "SPRING SUMMER":  "SS",
    "FALL-WINTER":    "FW",
    "FALL WINTER":    "FW",
    "ALL SEASON":     "AL",
    "ALL":            "AL",
}

# Keywords that identify accessory / non-footwear categories (Bag, Bag Charm, Parfume, Socks, etc.)
_ACC_KEYWORDS = {"ACCESSORY", "ACCESSORIES", "ACC", "BAG", "BAGS", "CHARM", "CHARMS",
                 "PARFUME", "PERFUME", "PARFUM", "SOCK", "SOCKS", "APPAREL", "CAP", "CAPS", "BELT", "BELTS"}


def _is_accessory(category: str, major_category: str) -> bool:
    """Returns True if category or major_category indicates an accessory item."""
    combined = (f"{category} {major_category}").upper()
    for kw in _ACC_KEYWORDS:
        if kw in combined:
            return True
    return False


def _resolve_country_size(category: str, major_category: str, mdd: Optional[MDDLoader] = None) -> str:
    """
    Returns Country Size LOV ID: EU (footwear) or NS (accessories/no-size).
    LOV IDs are hardcoded from the verified STEP LOV table — MDD is not used
    because MDD data has been observed to carry incorrect IDs (e.g. 'UE' for EU).
    """
    if _is_accessory(category, major_category):
        return "NS"
    return "EU"

# ── AT_Brand LOV  (display name → LOV ID) ────────────────────────────────
BRAND_LOV_MAP: dict[str, str] = {
    "2XU":           "2XU",
    "ADIDAS":        "ADI",
    "AIRWALK":       "AIW",
    "ALDO":          "AOD",
    "ANTA":          "ATA",
    "ASICS":         "ASI",
    "ASTEC":         "ASC",
    "BIRKENSTOCK":   "BCX",  # or BCK
    "CAMPER":        "C41",  # or C4M/C4X/C4P
    "CLARKS":        "CKS",
    "CROCS":         "CCR",
    "DIADORA":       "DIA",
    "DR. MARTENS":   "DRM",
    "DR MARTENS":    "DRM",  # alt format
    "ELLESSE":       "ELL",
    "HEY DUDE":      "HYD",  # or H3Y
    "IMPLUS":        "IPL",
    "K SWISS":       "KSW",
    "LOTTO":         "LOT",
    "NEW BALANCE":   "NEW",
    "NEW ERA":       "NRA",
    "NIKE":          "NIK",
    "ON RUNNING":    "ONR",
    "ONITSUKA TIGER": "ONT",
    "PAZZION":       "PZN",  # or PZZ
    "PTP":           "PTP",
    "REEBOK":        "REE",
    "SMIGGLE":       "IGL",
    "STACCATO":      "SC7",
    "STEVE MADDEN":  "SVM",  # or SVN
    "VIVAIA":        "VVA",
}

# ── AT_BrandType LOV ID Map ───────────────────────────────────────────────
BRAND_TYPE_LOV_MAP: dict[str, str] = {
    "MAA":          "MAA",
    "MAABRAND":     "MAA",
    "NONMAA":       "NONMAA",
    "NONMAABRAND":  "NONMAA",
    "NONSD":        "NONSD",
    "NONSDBRAND":   "NONSD",
    "SD":           "SD",
    "SDBRAND":      "SD",
}


def _resolve_brand_type(raw_val: str) -> str:
    """Resolve raw brand_type label into a valid STEP LOV ID, else empty string."""
    s = (raw_val or "").strip().upper()
    if not s:
        return ""

    # Normalize separators so variants like "NON MAA", "NON-MAA", "NON_MAA" all match.
    compact = re.sub(r"[\s\-_]+", "", s)
    return BRAND_TYPE_LOV_MAP.get(compact, "")


# ── AT_HeelHeight LOV  (STEP: Flat→F, High→H, Low→L, Medium→M) ────────────
HEEL_HEIGHT_LOV_MAP: dict[str, str] = {
    "F":      "F",  "FLAT":   "F",
    "H":      "H",  "HIGH":   "H",
    "L":      "L",  "LOW":    "L",
    "M":      "M",  "MEDIUM": "M",  "MIDDLE": "M",
}


def _resolve_heel_height(heel_raw: str, mdd: Optional[MDDLoader] = None) -> str:
    """Map raw heel height string to AT_HeelHeight LOV ID via MDD, fallback to hardcoded map."""
    if mdd:
        result = mdd.lookup("HEEL HEIGHT", heel_raw)
        if result:
            return result
    # Fallback: hardcoded LOV map (Flat→F, High→H, Low→L, Medium→M)
    return HEEL_HEIGHT_LOV_MAP.get((heel_raw or "").strip().upper(), "")


# ── AT_BYGender LOV IDs ───────────────────────────────────────────────────
# BY Gender LOV: Female / Male / Unisex  (IDs match display values)
# ── AT_BYAge LOV IDs  (STEP: Adult→ADULT, Kids→KIDS, etc.) ──────────────
BY_AGE_LOV_MAP: dict[str, str] = {
    "AD":           "ADULT",
    "ADULT":        "ADULT",
    "ADULTS":       "ADULT",
    "AA":           "ALL AGES",
    "ALL AGES":     "ALL AGES",
    "GS":           "GRADE SCHOOL",
    "GRADE SCHOOL": "GRADE SCHOOL",
    "IN":           "INFANT",
    "INFANT":       "INFANT",
    "CH":           "KIDS",
    "CHILDREN":     "KIDS",
    "K":            "KIDS",
    "KIDS":         "KIDS",
    "PS":           "PRESCHOOL",
    "PRESCHOOL":    "PRESCHOOL",
}


def _resolve_season(season_raw: str, fallback: str = "SM") -> str:
    """Map raw season string from linesheet to Season LOV ID."""
    s = (season_raw or "").strip().upper()
    if s in SEASON_LOV_MAP:
        return SEASON_LOV_MAP[s]
    # Try first word (e.g. "Summer 2026" → "SUMMER")
    first = s.split()[0] if s else ""
    return SEASON_LOV_MAP.get(first, fallback)



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


def color_3d(code: str) -> str:
    s = _ALNUM_RE.sub("", code or "").upper()
    if not s:
        return "000"
    return s[:3] if len(s) >= 3 else s.zfill(3)


def _find_input_files() -> list[Path]:
    base, in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir = _get_dirs()
    files: list[Path] = []
    if ll_dir.exists():
        files.extend(p for p in ll_dir.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm"))
    if not files and in_dir.exists():
        files.extend(p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm"))
    return sorted(files)


def _build_col_map(header_row: tuple) -> dict[str, int]:
    col_map: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        if cell is not None:
            norm = _s(cell).lower().strip()
            if norm:
                col_map[norm] = i
    return col_map


def _find_col(col_map: dict[str, int], field: str) -> Optional[int]:
    for alias in COLUMN_ALIASES.get(field, []):
        if alias in col_map:
            return col_map[alias]
    return None


def load_linesheet(path: Path) -> list[StaccatoRow]:
    log.info("[Staccato] Loading linesheet: %s", path.name)
    wb = openpyxl.load_workbook(str(path), data_only=True)
    
    # Target sheet 'linesheet' or first sheet
    sheet_name = "linesheet" if "linesheet" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        log.warning("[Staccato] No rows found in %s", path.name)
        return []

    # Find header row (usually row 2 or row 1)
    hdr_idx = 0
    for idx, r in enumerate(rows[:10]):
        if r and any(_s(c).lower().strip() in ["style no.", "brand", "product barcode"] for c in r if c):
            hdr_idx = idx
            break

    header_row = rows[hdr_idx]
    col_map = _build_col_map(header_row)

    col_brand          = _find_col(col_map, "brand")
    col_year           = _find_col(col_map, "year")
    col_season         = _find_col(col_map, "season")
    col_style_no       = _find_col(col_map, "style_no")
    col_produce_name   = _find_col(col_map, "produce_name")
    col_principal_desc = _find_col(col_map, "principal_style_description")
    col_factory_name   = _find_col(col_map, "factory_name")
    col_upper_material = _find_col(col_map, "upper_material")
    col_theme          = _find_col(col_map, "theme")
    col_major_cat      = _find_col(col_map, "major_category")
    col_cat            = _find_col(col_map, "category")
    col_gender         = _find_col(col_map, "gender")
    col_color          = _find_col(col_map, "color")
    col_color_zh       = _find_col(col_map, "color_chinese")
    col_materials      = _find_col(col_map, "materials")
    col_size_run       = _find_col(col_map, "size_run")
    col_heel_h         = _find_col(col_map, "heel_height")
    col_heel_l         = _find_col(col_map, "heel_level")
    log.info("[Staccato] Heel columns — heel_height col=%s  heel_level col=%s",
             col_heel_h, col_heel_l)
    log.info("[Staccato] Theme/Collection column — col=%s", col_theme)
    col_prod_level     = _find_col(col_map, "product_level")
    col_fob            = _find_col(col_map, "fob")
    col_tag_price      = _find_col(col_map, "tag_price")
    col_barcode        = _find_col(col_map, "product_barcode")
    col_photo          = _find_col(col_map, "photo")

    def _gcell(r: tuple, idx: Optional[int]) -> str:
        if idx is None or idx >= len(r):
            return ""
        return _s(r[idx])

    result: list[StaccatoRow] = []
    skipped = 0

    for i, r in enumerate(rows[hdr_idx + 1:], start=hdr_idx + 2):
        if r is None or all(c is None for c in r):
            continue

        style = _gcell(r, col_style_no)
        if not style or style.lower() in ("none", "nan", "total", ""):
            skipped += 1
            continue

        row = StaccatoRow(
            row_num         = i,
            brand           = _gcell(r, col_brand),
            year            = _gcell(r, col_year),
            season          = _gcell(r, col_season),
            style_no        = style,
            produce_name    = _gcell(r, col_produce_name),
            principal_style_description = _gcell(r, col_principal_desc),
            factory_name    = _gcell(r, col_factory_name),
            upper_material  = _gcell(r, col_upper_material),
            theme           = _gcell(r, col_theme),
            major_category  = _gcell(r, col_major_cat),
            category        = _gcell(r, col_cat),
            gender          = _gcell(r, col_gender),
            color           = _gcell(r, col_color),
            color_chinese   = _gcell(r, col_color_zh),
            materials       = _gcell(r, col_materials),
            size_run        = _gcell(r, col_size_run),
            heel_height     = _gcell(r, col_heel_h),
            heel_level      = _gcell(r, col_heel_l),
            product_level   = _gcell(r, col_prod_level),
            fob             = _num(r[col_fob]) if col_fob is not None and col_fob < len(r) else None,
            tag_price       = _num(r[col_tag_price]) if col_tag_price is not None and col_tag_price < len(r) else None,
            product_barcode = _gcell(r, col_barcode),
            photo           = _gcell(r, col_photo),
        )
        result.append(row)

    log.info("[Staccato] Parsed %d rows | %d skipped", len(result), skipped)
    return result


def group_generics(rows: list[StaccatoRow]) -> dict[str, dict[str, StaccatoGenericGroup]]:
    """Group rows by { style_no -> { color -> StaccatoGenericGroup } }."""
    groups: dict[str, dict[str, StaccatoGenericGroup]] = defaultdict(dict)
    for r in rows:
        st = r.style_no
        c = r.color or "000"
        if c not in groups[st]:
            groups[st][c] = StaccatoGenericGroup(
                style_no       = st,
                color          = c,
                produce_name   = r.produce_name,
                principal_style_description = r.principal_style_description,
                major_category = r.major_category,
                category       = r.category,
                gender         = r.gender,
                country        = r.country,
                season         = r.season,
                year           = r.year,
                size_run       = r.size_run,
                theme          = r.theme,
                heel_level     = r.heel_level,
                fob            = r.fob,
            )
        groups[st][c].rows.append(r)
    return groups


# ══════════════════════════════════════════════════════════════════════════════
# CLASSIFICATIONS BUILDER
# ══════════════════════════════════════════════════════════════════════════════
_SEASON_LABELS: dict[str, str] = {
    "SP": "Spring",        "SM": "Summer",
    "FL": "Fall",          "WN": "Winter",
    "CO": "Core",          "SS": "Spring Summer",
    "FW": "Fall Winter",   "AL": "All Season",
}


def build_classifications(
    brand_code: str,
    brand_name: str,
    seasons: list[tuple[str, str]],   # [(sea_prefix, sea_year), ...]
) -> ET.Element:
    """
    Build <Classifications> block.

    Structure (parent Batches assumed to exist in STEP):
      CLH_{brand_code}_{sea}{year}        CLS_Season               parent=CLH_{brand}Batches
        CLH_{brand_code}_{sea}{year}CA    CLS_ConfirmedArticles    parent=CLH_{brand_code}_{sea}{year}
        CLH_{brand_code}_{sea}{year}UA    CLS_UnconfirmedArticles  parent=CLH_{brand_code}_{sea}{year}
    """
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    batches_parent = f"CLH_{brand_name}Batches"

    # Emit season classifications
    for sea_prefix, sea_year in seasons:
        season_id      = f"CLH_{brand_code}_{sea_prefix}{sea_year}"
        sea_name       = _SEASON_LABELS.get(sea_prefix, sea_prefix)
        season_display = f"{brand_name} {sea_name} {sea_year}".strip()
        season_short   = f"{sea_prefix} {sea_year}".strip()

        season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
        season_cls.set("ID",         season_id)
        season_cls.set("UserTypeID", "CLS_Season")
        season_cls.set("ParentID",   batches_parent)
        ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

        confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        confirmed.set("ID",         f"{season_id}CA")
        confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
        confirmed.set("ParentID",   season_id)
        ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

        unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        unconfirmed.set("ID",         f"{season_id}UA")
        unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
        unconfirmed.set("ParentID",   season_id)
        ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


# ══════════════════════════════════════════════════════════════════════════════
# RNA LOADER (Brand Type / Brand Category from Attributes mapping workbook)
# ══════════════════════════════════════════════════════════════════════════════
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA", "MY": "MALAYSIA",  "PH": "PHILIPPINES",
    "SG": "SINGAPORE",  "TH": "THAILAND", "VN": "VIETNAM",
}


class RNALoader:
    """Loads 'Source Mapping Related RNA' tab from the Attributes / Brand mapping workbook."""

    RNA_SHEET_KEYWORDS = ["RNA", "SOURCE MAPPING"]

    @staticmethod
    def _norm_country(v: str) -> str:
        return (v or "").strip().upper()

    @staticmethod
    def _norm_comp(v: str) -> str:
        s = (v or "").strip()
        if not s:
            return ""
        if s.endswith(".0"):
            s = s[:-2]
        s2 = s.lstrip("0")
        return s2 if s2 else "0"

    def __init__(self, path: Path):
        self.path = path
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
            (s for s in wb.sheetnames if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] No RNA/SOURCE MAPPING sheet found in %s", self.path.name)
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
                self._norm_comp(_cell(row, c_comp)),
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
            self._norm_comp(comp_code),
            (sbu        or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": ""})

    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        c = self._norm_country(country_name)
        b = (brand_code or "").strip().upper()
        s = (sbu or "").strip().upper()
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": ""}


def _find_rna_source_file(attr_dir: Path) -> Path | None:
    """Locate the Brand mapping / Attributes workbook that contains an RNA sheet."""
    # 1. Exact known filename
    exact = attr_dir / "NEW - Brand mapping files Template.xlsx"
    if exact.exists():
        try:
            wb = openpyxl.load_workbook(exact, read_only=True, data_only=True)
            has_rna = any(any(kw in s.upper() for kw in RNALoader.RNA_SHEET_KEYWORDS) for s in wb.sheetnames)
            wb.close()
            if has_rna:
                return exact
        except Exception as e:
            log.warning("[RNA] Cannot inspect %s: %s", exact.name, e)

    # 2. Keyword search in attr_dir
    keywords = ["brand mapping", "brand_mapping", "mapping template", "brand template"]
    for f in sorted(attr_dir.glob("*.xlsx")):
        if any(kw in f.name.lower() for kw in keywords):
            try:
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                has_rna = any(any(kw in s.upper() for kw in RNALoader.RNA_SHEET_KEYWORDS) for s in wb.sheetnames)
                wb.close()
                if has_rna:
                    return f
            except Exception as e:
                log.warning("[RNA] Cannot inspect %s: %s", f.name, e)
    return None


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILD (ONLY GENERIC PRODUCTS)
# ══════════════════════════════════════════════════════════════════════════════
def _resolve_gender_age(
    gender_raw: str,
    mdd: Optional[MDDLoader] = None,
) -> tuple[str, str, str, str]:
    """
    Returns (AT_Gender id, AT_BYGender id, AT_SAPAge id, AT_BYAge id)

    LOV sources (STEP):
      AT_Gender  (SAP Gender): Female→F, Male→M, Unisex→U
      AT_BYGender             : Female, Male, Unisex
      AT_SAPAge               : AD (Adults), K (Children)
      AT_BYAge                : ADULT, KIDS, etc.  ← from MDD Age LOV col-F
    """
    g = (gender_raw or "").strip().upper()
    if g in ["FEMALE", "FAMALE", "WOMEN", "WOMAN", "F"]:
        sap_gender, by_gender_raw, sap_age, by_age_raw = "F", "Female", "AD", "Adult"
    elif g in ["MALE", "MEN", "MAN", "M"]:
        sap_gender, by_gender_raw, sap_age, by_age_raw = "M", "Male",   "AD", "Adult"
    else:
        sap_gender, by_gender_raw, sap_age, by_age_raw = "F", "Female", "AD", "Adult"  # default

    # AT_BYGender ID — from MDD 'BY Gender' / 'Gender' LOV; fallback to display value ID (F/M/U)
    by_gender_fallback = "F" if sap_gender == "F" else ("M" if sap_gender == "M" else "U")
    if mdd:
        by_gender_id = (
            mdd.lookup("BY GENDER", by_gender_raw)
            or mdd.lookup("GENDER",    by_gender_raw)
            or by_gender_fallback
        )
    else:
        by_gender_id = by_gender_fallback

    # AT_BYAge ID — from MDD Age LOV col-F only; fallback to ADULT (uppercase per STEP LOV)
    by_age_fallback = "ADULT" if by_age_raw == "Adult" else by_age_raw.upper()
    if mdd:
        by_age_id = (
            mdd.lookup("BYAGE",   by_age_raw)
            or mdd.lookup("BY AGE", by_age_raw)
            or by_age_fallback
        )
    else:
        by_age_id = by_age_fallback

    return sap_gender, by_gender_id, sap_age, by_age_id


def _build_generic_product(
    group: StaccatoGenericGroup,
    cfg:   dict,
) -> ET.Element:
    brand_code   = cfg.get("brand_code", "SC7")
    brand_name   = cfg.get("brand_name", "Staccato")
    # Season prefix — cfg (portal) takes priority, fall back to row data
    sea_prefix   = cfg.get("season_prefix") or _resolve_season(group.season, fallback="SM")
    # Season year — cfg (portal) takes priority, fall back to row data then current year
    sea_year     = (cfg.get("season_year") or group.year or str(datetime.now().year))
    season_id    = f"CLH_{brand_code}_{sea_prefix}{sea_year}"

    generic_code = f"{brand_code}{group.style_no}{group.color or ''}"

    sap_gender_code, sap_gender_label, sap_age_code, sap_age_label = _resolve_gender_age(
        group.gender, mdd=cfg.get("mdd")
    )

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    parent_id = "PPH_E-TempSubCat" if _is_accessory(group.category, group.major_category) else "PPH_F-TempSubCat"
    gen_el.set("ParentID",   parent_id)
    _keyval(gen_el, "KEY_InboundArticle", generic_code)
    
    name_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name")
    if group.principal_style_description:
        name_el.text = group.principal_style_description
        
    _clf(gen_el, f"CLH_{brand_name}Articles",  "CPL_Merchandiser")
    _clf(gen_el, f"{season_id}UA",              "CPL_UnConfirmedForSeason")
    
    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")
        

    # ── Mapped & Default Attributes from Staccato(Inline) ──────────────
    # AT_Brand: brand_code IS the AT_Brand LOV ID — use it directly.
    # Do NOT look up from MDD or BRAND_LOV_MAP by brand_name: MDD may store the
    # display name as the ID (e.g. "STACCATO" instead of "SC7").
    _val(vals_el, "AT_Brand",               "", id_val=brand_code)
    _val(vals_el, "AT_BrandGroup",          "", id_val=brand_name.upper())
    _val(vals_el, "AT_BrandStatus",         "", id_val="A")
    # Brand Type resolved from RNA (Source Mapping). Prefer LOV ID, fallback to value text.
    # Some STEP setups resolve BrandType by display value during import.
    raw_brand_type = (cfg.get("brand_type") or "").strip()
    if raw_brand_type:
        bt_id = _resolve_brand_type(raw_brand_type)
        if bt_id:
            _val(vals_el, "AT_BrandType", id_val=bt_id)
        else:
            log.warning(
                "[BrandType] Unmapped value '%s' from RNA. Writing as text value fallback.",
                raw_brand_type,
            )
            _val(vals_el, "AT_BrandType", value=raw_brand_type)
    _multival(vals_el, "AT_SBU",            cfg.get("sbu", "FF"))
    _multival(vals_el, "AT_CompanyCode",    cfg.get("comp_code", "0888"))
    country_code = group.country or cfg.get("country_code", "ID")
    _val(vals_el, "AT_Country",             "", id_val=country_code)
    _RETAIL_CURRENCY_MAP: dict[str, str] = {
        "ID": "IDR", "PH": "PHP", "TH": "THB",
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    retail_currency = _RETAIL_CURRENCY_MAP.get(country_code, "")
    if retail_currency:
        _val(vals_el, "AT_RetailPriceCurrency", "", id_val=retail_currency)
    # AT_InboundArticle is populated automatically by Business Action "Update Attributes Inbound"
    # from KEY_InboundArticle — do not include in import XML (STEP reports it as "not found").
    _val(vals_el, "AT_InboundGenericCode",  generic_code)
    _val(vals_el, "AT_Season",              "", id_val=sea_prefix)
    _val(vals_el, "AT_SeasonYear",          sea_year)

    # Mapped Line Sheet Fields
    _val(vals_el, "AT_PrincipalStyleCode",  group.style_no)
    
    if group.principal_style_description:
        _val(vals_el, "AT_PrincipalStyleDescription", group.principal_style_description)
    if group.color:
        _val(vals_el, "AT_PrincipalColorName", group.color)
    if group.size_run:
        _val(vals_el, "AT_PrincipalSize", group.size_run)
        
    if group.major_category:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", group.major_category)
    if group.category:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", group.category)
    if group.heel_level:
        heel_id = _resolve_heel_height(group.heel_level, mdd=cfg.get("mdd"))
        log.info("[HeelHeight] heel_level='%s' → LOV ID='%s'", group.heel_level, heel_id or "(no match)")
        if heel_id:
            _val(vals_el, "AT_HeelHeight", "", id_val=heel_id)
    else:
        log.info("[HeelHeight] heel_level empty for style=%s color=%s — AT_HeelHeight skipped",
                 group.style_no, group.color)
    if group.theme:
        _val(vals_el, "AT_Collection1", group.theme)
    if group.fob is not None:
        _val(vals_el, "AT_FOB", str(group.fob))

    # Mapped Gender & Age
    _val(vals_el, "AT_Gender",              "", id_val=sap_gender_code)
    _val(vals_el, "AT_SAPAge",              "", id_val=sap_age_code)
    _val(vals_el, "AT_BYGender",            "", id_val=sap_gender_label)
    _val(vals_el, "AT_BYAge",               "", id_val=sap_age_label)

    # Defaults specified in Staccato(Inline) template
    _val(vals_el, "AT_CountryOrigin",       "", id_val="CN")
    _val(vals_el, "AT_FOBCurrency",         "", id_val="USD")
    _val(vals_el, "AT_SAPProductFlag",      "", id_val="A")
    _val(vals_el, "AT_MaterialType",        "", id_val="ZINA")
    _val(vals_el, "AT_SAPArticleCategory",  "", id_val="01")
    _val(vals_el, "AT_UOM",                 "", id_val="EA")
    _val(vals_el, "AT_BYArticleType",         "", id_val=cfg.get("article_type", "Inline"))
    # AT_CountrySize: EUR for footwear, NS (No Size) for accessories — IDs from MDD Country Size LOV
    _val(vals_el, "AT_CountrySize", "", id_val=_resolve_country_size(
        group.category, group.major_category, mdd=cfg.get("mdd")
    ))
    _val(vals_el, "AT_BCI",                 "", id_val="COMMERCIAL")
    # AT_PackDetails — only for Accessories (Bag, Bag Charm, Parfume, Socks, etc.)
    if _is_accessory(group.category, group.major_category):
        _val(vals_el, "AT_PackDetails",     "", id_val="S")          # default: Single
    _val(vals_el, "AT_EComAgesCategory",    "", id_val="18+Y")       # default: Ages 18+ years
    # AT_ArticleStatus: LOV ID 'Active' per MDD (Default when created: Active)
    _val(vals_el, "AT_ArticleStatus",       "", id_val="A")
    _val(vals_el, "AT_PricingDistributionChannel", "", id_val="01")

    return gen_el


def build_xml(rows: list[StaccatoRow], out_xml_path: Path, cfg: dict) -> None:
    groups = group_generics(rows)

    # ──────────────────────────────────────────────────────────────────────────
    # Use ONLY the configured/specified season for Classifications
    # (Do NOT scan rows for multiple seasons — generate exactly 1 classification)
    # ──────────────────────────────────────────────────────────────────────────
    fallback_year   = cfg.get("season_year") or str(datetime.now().year)
    fallback_prefix = cfg.get("season_prefix", "SM")
    # Single season only: use provided season_prefix and season_year from config
    season_list = [(fallback_prefix, fallback_year)]

    # ══════════════════════════════════════════════════════════════════════════
    # TEST LIMITER — Set TEST_MODE = False for production
    # ══════════════════════════════════════════════════════════════════════════
    TEST_MODE = False
    # FOOTWEAR_LIMIT = 2
    # ACCESSORIES_LIMIT = 2

    # if TEST_MODE:
    #     all_pairs = [(sc, cc, cg) for sc, cm in groups.items() for cc, cg in cm.items()]
    #     footwear_pairs = []
    #     accessories_pairs = []
    
    #     for sc, cc, cg in all_pairs:
    #         if _is_accessory(cg.category, cg.major_category):
    #             if len(accessories_pairs) < ACCESSORIES_LIMIT:
    #                 accessories_pairs.append((sc, cc, cg))
    #         else:
    #             if len(footwear_pairs) < FOOTWEAR_LIMIT:
    #                 footwear_pairs.append((sc, cc, cg))
    
    #         if len(footwear_pairs) >= FOOTWEAR_LIMIT and len(accessories_pairs) >= ACCESSORIES_LIMIT:
    #             break
    
    #     selected_pairs = footwear_pairs + accessories_pairs
    #     groups = {}
    #     for sc, cc, cg in selected_pairs:
    #         if sc not in groups:
    #             groups[sc] = {}
    #         groups[sc][cc] = cg
    #     log.info("[XML] TEST MODE: Limited to %d footwear and %d accessories products", len(footwear_pairs), len(accessories_pairs))

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    brand_code = cfg.get("brand_code", "SC7")
    brand_name = cfg.get("brand_name", "Staccato")

    log.info("[Staccato] Writing STEP XML → %s", out_xml_path.name)
    with open(out_xml_path, "w", encoding="utf-8") as f:
        open_step_xml(f, export_time)

        # ── Classifications block (written before Products) ────────────────
        cls_elem = build_classifications(brand_code, brand_name, season_list)
        cls_str  = ET.tostring(cls_elem, encoding="utf-8").decode("utf-8")
        cls_str  = re.sub(r'ns0:', '', cls_str)
        cls_str  = re.sub(r':ns0', '', cls_str)
        f.write(f"  {cls_str}\n")

        f.write("  <Products>\n")

        total_generics = 0
        for style_no, color_dict in groups.items():
            for color, group in color_dict.items():
                gen_elem = _build_generic_product(group, cfg)
                xml_str = ET.tostring(gen_elem, encoding="utf-8").decode("utf-8")
                # Clean namespace prefixes & xmlns declarations from inner Product tags
                xml_str = re.sub(r'ns0:', '', xml_str)
                xml_str = re.sub(r':ns0', '', xml_str)
                xml_str = re.sub(r'\s*xmlns="[^"]+"', '', xml_str)
                f.write(f"    {xml_str}\n")
                total_generics += 1

        f.write("  </Products>\n")
        close_step_xml(f)

    log.info("[Staccato] XML generated successfully: %d Generic product nodes.", total_generics)


def run(args, auditor=None) -> tuple[list[dict], Path | None]:
    log.info("[Staccato-Linelist] ETL started — brand=%s", getattr(args, "brand", "Staccato"))

    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        file_path = Path(input_file_arg)
    else:
        files = _find_input_files()
        if not files:
            log.warning("[Staccato] No linelist .xlsx/.xlsm files found — nothing to process.")
            return [], None
        file_path = max(files, key=lambda p: p.stat().st_mtime)

    log.info("[Staccato] Processing file: %s", file_path.name)

    base, in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir = _get_dirs()

    # ── Load MDD LOVs ────────────────────────────────────────────────────────
    mdd: Optional[MDDLoader] = None
    mdd_path = _find_mdd_file(mdd_dir)
    if mdd_path:
        try:
            mdd = MDDLoader(mdd_path)
        except Exception as exc:
            log.warning("[MDD] Load failed (%s) — using hardcoded fallback maps", exc)
    else:
        log.info("[MDD] No MDD file found in %s — using hardcoded fallback maps", mdd_dir)

    brand_code   = getattr(args, "brand_code",   "SC7")   or "SC7"
    brand_name   = getattr(args, "brand",        "Staccato") or "Staccato"
    comp_code    = getattr(args, "comp_code",    "0888")  or "0888"
    sbu          = getattr(args, "sbu",          "FF")    or "FF"
    country_code = getattr(args, "country_code", "ID")    or "ID"

    # ── Brand Type / Brand Category from RNA ─────────────────────────────────
    brand_type = getattr(args, "brand_type",     "") or ""
    brand_cat  = getattr(args, "brand_category", "") or ""

    if not brand_type or not brand_cat:
        rna_path = _find_rna_source_file(attr_dir)
        if rna_path:
            try:
                rna          = RNALoader(rna_path)
                country_name = _COUNTRY_MAP.get(country_code.upper(), country_code.upper())
                result = rna.get(country_name, comp_code, sbu, brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_code)
                brand_type = brand_type or result.get("brand_type", "")
                brand_cat  = brand_cat  or result.get("brand_category", "")
                log.info(
                    "[Staccato] RNA → country=%s comp=%s sbu=%s brand=%s → type='%s' cat='%s'",
                    country_name, comp_code, sbu, brand_code, brand_type, brand_cat,
                )
            except Exception as e:
                log.warning("[Staccato] RNA lookup failed: %s", e)
        else:
            log.warning("[Staccato] No RNA source workbook found in %s", attr_dir)

    cfg = {
        "brand_code":    brand_code,
        "brand_name":    brand_name,
        "sbu":           sbu,
        "comp_code":     comp_code,
        "country_code":  country_code,
        "season_prefix": getattr(args, "season_prefix", None) or "SM",
        "season_year":   getattr(args, "season_year", None) or None,
        "article_type":  getattr(args, "article_type", "Inline") or "Inline",
        "brand_type":    brand_type,
        "brand_category": brand_cat,
        "mdd":           mdd,    # MDDLoader instance (or None)
    }

    rows = load_linesheet(file_path)
    if not rows:
        log.warning("[Staccato] No valid rows parsed.")
        return [], None

    xml_path = xml_dir / f"{file_path.stem}.xml"
    build_xml(rows, xml_path, cfg)

    if auditor and xml_path:
        auditor.set_xml_uploads([str(xml_path)])

    return [], xml_path


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Staccato Line List → STEP XML")
    ap.add_argument("--brand",        default="Staccato")
    ap.add_argument("--brand-code",   default="SC7", dest="brand_code")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="FF")
    ap.add_argument("--country-code", default="ID", dest="country_code")
    run(ap.parse_args())