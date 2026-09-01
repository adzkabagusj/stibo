"""
STIBO INBOUND XML GENERATOR -- implus Line List (Inline) v1.0
"STACCATO 26Summer linesheet.xlsx" -> Stibo STEP XML

Source file format : implus Line Sheet Excel (.xlsx / .xlsm)
Sheet name         : "linesheet" or first non-empty data sheet

MAPPED ATTRIBUTES (from implus(Inline) mapping template):
---------------------------------------------------------------------------------
Stibo Attribute ID                | Field Name in Brand File / Value
---------------------------------------------------------------------------------
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
AT_ArticleStatus                  | Active (Default -- MDD LOV ID)
---------------------------------------------------------------------------------
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
log = logging.getLogger("implus.linelist_main")


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class implusRow:
    row_num:           int
    color_code:        str = ""
    implus_us:         str = ""
    upc:               str = ""
    style_desc:        str = ""
    color:             str = ""
    size:              str = ""
    country_of_origin: str = ""
    distributor_price: str = ""
    hierarchy_l1:      str = ""

@dataclass
class implusGenericGroup:
    implus_us:         str
    style_desc:        str
    hierarchy_l1:      str
    color:             str = ""
    color_code:        str = ""
    principal_style_description: str = ""
    variants:          list[implusRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER  — reads LOVs from Master Data Dictionary Excel at runtime
# ══════════════════════════════════════════════════════════════════════════════
class MDDLoader:
    """
    Loads LOVs from the shared Master Data Dictionary Excel placed in
    implus/input/mdd/.

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
                elif "BRAND" in snu and "TYPE" in snu and "LOV" in snu:
                    self._load_brand_attr_lov("AT_BrandType", rows)
                elif "BRAND" in snu and "STATUS" in snu and "LOV" in snu:
                    self._load_brand_attr_lov("AT_BrandStatus", rows)
                elif "BRAND" in snu and "CATEGORY" in snu and "LOV" in snu:
                    self._load_brand_attr_lov("AT_BrandCategory", rows)
                elif "BRAND" in snu and "GROUP" in snu and ("LOV" in snu or snu == "BRAND GROUP"):
                    self._load_brand_attr_lov("AT_BrandGroup", rows)
                else:
                    # Derive key: strip " LOV" / "LOV " suffix/prefix
                    key = snu.replace(" LOV", "").replace("LOV ", "").replace("LOV", "").strip()
                    self._load_two_col(key, rows)
        finally:
            wb.close()
        log.info("[MDD] LOVs loaded — sheets: %s", list(self.lovs.keys()))

    def _load_brand_attr_lov(self, attr_key: str, rows: list) -> None:
        """Parse a Brand Type/Status/Category/Group LOV sheet.
        Col A = LOV ID  (e.g. 'NONMAA', 'A', 'ID - SP - NON TOP', 'ASICS')
        Col B = display value  (e.g. 'Non MAA', 'Active', ...)
        Stored as: display.upper() → lov_id  AND  lov_id.upper() → lov_id
        Mirrors the Clarks MDDLoader pattern exactly.
        """
        lov: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] is not None else ""
            display = str(row[1]).strip() if row[1] is not None else ""
            if lov_id and lov_id.upper() not in ("NONE", "NAN"):
                if display:
                    lov[display.upper()] = lov_id
                lov[lov_id.upper()] = lov_id  # also key by ID itself
        if lov:
            self.lovs[attr_key] = lov
            log.info("[MDD] %s loaded — %d entries", attr_key, len(lov))
        else:
            log.warning("[MDD] %s sheet found but empty", attr_key)

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
    "CAMPER":        "C41",  # or C4M/C4X/C4P/C41
    "CLARKS":        "CKS",
    "CROCS":         "CCR",
    "DIADORA":       "DIA",
    "DR. MARTENS":   "DRM",
    "DR MARTENS":    "DRM",  # alt format
    "ELLESSE":       "ELL",
    "HEY DUDE":      "HYD",  # or H3Y
    "IMPLUS":        "IPL",
    "implus":       "IPL",
    "IMP":           "IPL",
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

# ── AT_BrandGroup LOV ID Map ──────────────────────────────────────────────
BRAND_GROUP_LOV_MAP: dict[str, str] = {
    "implus":       "IMPLUS",
    "IMP":           "IMPLUS",
    "IMPLUS":        "IMPLUS",
    "2XU":           "2XU",
    "ADAMS":         "ADAMS",
    "ADIDAS":        "ADIDAS",
    "AIRWALK":       "AIRWALK",
    "ALDO":          "ALDO",
    "ANTA":          "ANTA",
    "ASICS":         "ASICS",
    "ASTEC":         "ASTEC",
    "BIRKENSTOCK":   "BIRKENSTOCK",
    "CAMPER":        "CAMPER",
    "CLARKS":        "CLARKS",
    "CROCS":         "CROCS",
    "DIADORA":       "DIADORA",
    "DR. MARTENS":   "DR. MARTENS",
    "ELLESSE":       "ELLESSE",
    "LOTTO":         "LOTTO",
    "NEW BALANCE":   "NEW BALANCE",
    "NEW ERA":       "NEW ERA",
    "NIKE":          "NIKE",
    "ON":            "ON",
    "ONITSUKA TIGER": "ONITSUKA TIGER",
    "PAZZION":       "PAZZION",
    "PTP":           "PTP",
    "REEBOK":        "REEBOK",
    "SMIGGLE":       "SMIGGLE",
    "STACCATO":      "STACCATO",
}


def _resolve_brand_group(raw_val: str) -> str:
    s = (raw_val or "").strip().upper()
    return BRAND_GROUP_LOV_MAP.get(s, s)


# ── AT_PackDetails LOV ID Map ──────────────────────────────────────────────
PACK_DETAILS_LOV_MAP: dict[str, str] = {
    "MULTI":     "M",
    "PACK OF 2": "2",
    "PACK OF 3": "3",
    "PACK OF 5": "5",
    "PACK OF 6": "6",
    "SINGLE":    "S",
    "M":         "M",
    "2":         "2",
    "3":         "3",
    "5":         "5",
    "6":         "6",
    "S":         "S",
}


def _resolve_pack_details(raw_val: str, mdd: Optional[MDDLoader] = None) -> str:
    s = (raw_val or "").strip().upper()
    if s in PACK_DETAILS_LOV_MAP:
        return PACK_DETAILS_LOV_MAP[s]
    if mdd:
        res = mdd.lookup("PACK DETAILS", raw_val) or mdd.lookup("LOV_PACKDETAILS", raw_val)
        if res:
            return PACK_DETAILS_LOV_MAP.get(res.upper(), res)
    return "S"


# ── AT_BrandType LOV ID Map ───────────────────────────────────────────────
BRAND_TYPE_LOV_MAP: dict[str, str] = {
    "MAA":              "MAA",
    "MAA BRAND":        "MAA",
    "NONMAA":           "NONMAA",
    "NON-MAA BRAND":    "NONMAA",
    "NON-MAA":          "NONMAA",
    "NONSD":            "NONSD",
    "NON-SD BRAND":     "NONSD",
    "NON-SD":           "NONSD",
    "SD":               "SD",
    "SD BRAND":         "SD",
}


def _resolve_brand_type(raw_val: str) -> str:
    """Resolves raw brand_type string (or label from RNA) to the exact STEP LOV ID."""
    s = (raw_val or "").strip().upper()
    return BRAND_TYPE_LOV_MAP.get(s, s)


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


def _is_black_fill(cell, log_ctx: str = "") -> bool:
    """
    True if a cell's background fill is black/near-black.

    Same mechanism as implus/linelist_main_source_sofsole.py's
    _is_black_fill(): the top-level Category (L1) banner row is
    black-filled; sub-category banners use other colors. Handles all 3
    ways Excel can store a fill color: explicit RGB, a theme color, or a
    legacy indexed color.
    """
    try:
        fill = cell.fill
        if fill is None or fill.patternType is None:
            log.info("[Balega][fill]%s no pattern fill -> not black", log_ctx)
            return False

        fg      = fill.fgColor
        c_type  = getattr(fg, "type", None)
        rgb     = getattr(fg, "rgb", None)
        theme   = getattr(fg, "theme", None)
        tint    = getattr(fg, "tint", None)
        indexed = getattr(fg, "indexed", None)
        log.info(
            "[Balega][fill]%s pattern=%s type=%s rgb=%s theme=%s tint=%s indexed=%s",
            log_ctx, fill.patternType, c_type, rgb, theme, tint, indexed,
        )

        is_black = False

        # 1. Explicit RGB (e.g. "FF000000")
        if isinstance(rgb, str) and len(rgb) >= 6:
            hex6 = rgb[-6:]
            try:
                r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
                is_black = r < 40 and g < 40 and b < 40
            except ValueError:
                pass

        # 2. Theme color -- theme index 1 = "Black, Text 1" in the default
        #    Office theme (0=Background1/white, 1=Text1/black, ...).
        if not is_black and c_type == "theme" and theme == 1 and (not tint or abs(tint) < 0.1):
            is_black = True

        # 3. Legacy indexed palette -- 8 = black, 64 = automatic (often black text/fill)
        if not is_black and indexed in (8, 64):
            is_black = True

        log.info("[Balega][fill]%s -> is_black=%s", log_ctx, is_black)
        return is_black
    except Exception as exc:
        log.info("[Balega][fill]%s error checking fill: %s", log_ctx, exc)
        return False


def load_linesheet(path: Path) -> list[implusRow]:
    log.info("[implus] Loading excel file: %s", path.name)
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        log.error("Failed to open excel: %s", e)
        return []

    ws = wb["Balega Asia Gen"] if "Balega Asia Gen" in wb.sheetnames else wb.active

    def _clean_str(val) -> str:
        if val is None: return ""
        s = str(val).strip()
        if s in ("None", "nan"): return ""
        return s

    rows = []
    current_l1 = ""
    header_row_idx = -1
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        cell_0 = _clean_str(row[0])
        cell_1 = _clean_str(row[1])
        if cell_0.lower() == "color code" or cell_1.lower() == "implus us":
            header_row_idx = row_idx
            break

    if header_row_idx == -1:
        log.warning("Could not find header row. Defaulting to row 20.")
        header_row_idx = 20

    # Some files place the section banner (e.g. "ULTRA LIGHT NO SHOW")
    # directly above the header row rather than interspersed among the
    # data rows below it — the main loop below only starts scanning after
    # the header, so that banner would otherwise never be seen and L1
    # would stay blank for the whole sheet. Pre-scan rows above the
    # header for the last black-filled banner row and seed current_l1.
    for row_idx in range(1, header_row_idx):
        row = [c.value for c in next(ws.iter_rows(min_row=row_idx, max_row=row_idx))]
        cell_0 = _clean_str(row[0]) if len(row) > 0 else ""
        cell_1 = _clean_str(row[1]) if len(row) > 1 else ""
        cell_2 = _clean_str(row[2]) if len(row) > 2 else ""
        cell_4 = _clean_str(row[4]) if len(row) > 4 else ""
        if (cell_1 or cell_0) and not cell_2 and not cell_4:
            header_title = cell_1 or cell_0
            if "TOTAL" not in header_title.upper() and "COLOR CODE" not in header_title.upper():
                banner_cell = ws.cell(row=row_idx, column=2)
                log_ctx = f" row={row_idx} text='{header_title}' (pre-header scan)"
                if _is_black_fill(banner_cell, log_ctx):
                    current_l1 = header_title
                    log.info("[Balega] L1 seeded from pre-header banner -> '%s' (row %d)", current_l1, row_idx)

    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, values_only=True), start=header_row_idx + 1):
        cell_0 = _clean_str(row[0])
        cell_1 = _clean_str(row[1])
        cell_2 = _clean_str(row[2]) if len(row) > 2 else ""
        cell_4 = _clean_str(row[4]) if len(row) > 4 else ""

        # Section Header Row: Column B (or A) contains category title (e.g. HIDDEN COMFORT NO SHOW).
        # Only promote it to L1 if Column B's cell is actually black-filled —
        # same mechanism as implus/linelist_main_source_sofsole.py's
        # _is_black_fill() check, so non-black banner rows (if any) don't
        # overwrite the current L1.
        if (cell_1 or cell_0) and not cell_2 and not cell_4:
            header_title = cell_1 or cell_0
            if "TOTAL" not in header_title.upper() and "COLOR CODE" not in header_title.upper():
                banner_cell = ws.cell(row=row_idx, column=2)  # Column B
                log_ctx = f" row={row_idx} text='{header_title}'"
                if _is_black_fill(banner_cell, log_ctx):
                    if header_title != current_l1:
                        current_l1 = header_title
                        log.info("[Balega] L1 set -> '%s' (row %d)", current_l1, row_idx)
                else:
                    log.info(
                        "[Balega] Banner row not black-filled -> not treated as L1: "
                        "'%s' (row %d)", header_title, row_idx,
                    )
            continue

        if not cell_0 and not cell_1:
            continue

        color_code = _clean_str(row[0])
        implus_us = _clean_str(row[1])
        upc = _clean_str(row[2])
        style_desc = _clean_str(row[4])
        color = _clean_str(row[5])
        size = _clean_str(row[6])
        coo = _clean_str(row[7])
        price = _clean_str(row[12])

        rows.append(implusRow(
            row_num=row_idx,
            color_code=color_code,
            implus_us=implus_us,
            upc=upc,
            style_desc=style_desc,
            color=color,
            size=size,
            country_of_origin=coo,
            distributor_price=price,
            hierarchy_l1=current_l1
        ))

    log.info("[implus] Parsed %d valid rows", len(rows))
    return rows


def group_generics(rows: list[implusRow]) -> dict[str, dict[str, implusGenericGroup]]:
    """Group rows by { implus_us -> { color -> implusGenericGroup } }.
    
    For a linelist, each generic represents ONE style + ONE color combination,
    with multiple size variants (following Staccato linelist pattern).
    """
    groups: dict[str, dict[str, implusGenericGroup]] = defaultdict(dict)
    for r in rows:
        style_code = r.implus_us or f"UNKNOWN_{r.row_num}"
        color = r.color or "000"
        
        if color not in groups[style_code]:
            groups[style_code][color] = implusGenericGroup(
                implus_us=style_code,
                style_desc=r.style_desc,
                principal_style_description=f"{r.style_desc} {r.size}".strip(),
                hierarchy_l1=r.hierarchy_l1,
                color=color,
                color_code=r.color_code
            )
        groups[style_code][color].variants.append(r)
    
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

    # Normalize "implus" / "implus" -> "Implus" (without 'e') for STEP classification parent/names
    b_name = "Implus" if (brand_name or "").strip().upper() in ("implus", "IMP", "IMPLUS") else brand_name.strip()
    b_code = "IPL" if (brand_code or "").strip().upper() in ("implus", "IMP", "IPL") else brand_code.strip()

    batches_parent = f"CLH_{b_name}Batches"

    # Emit season classifications
    for sea_prefix, sea_year in seasons:
        season_id      = f"CLH_{b_code}_{sea_prefix}{sea_year}"
        sea_name       = _SEASON_LABELS.get(sea_prefix, sea_prefix)
        season_display = f"{b_name} {sea_name} {sea_year}".strip()
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
        c_bgroup  = _find("BRANDGROUP", "BRAND GROUP", "AT_BRANDGROUP")
        c_bstatus = _find("BRAND STATUS", "BRANDSTATUS", "AT_BRANDSTATUS")

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
                "brand_group":    _cell(row, c_bgroup),
                "brand_status":   _cell(row, c_bstatus),
            }
            count += 1
        log.info("[RNA] %d entries loaded from '%s'", count, sheet_name)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        b_codes = {(brand_code or "").strip().upper(), "IPL", "IMP", "IMPLUS", "BAL"}
        for b in b_codes:
            key = (
                self._norm_country(country_name),
                self._norm_comp(comp_code),
                (sbu        or "").strip().upper(),
                b,
            )
            if key in self.lookup:
                return self.lookup[key]
        return self.get_fuzzy(country_name, sbu, brand_code)

    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        c = self._norm_country(country_name)
        b_codes = {(brand_code or "").strip().upper(), "IPL", "IMP", "IMPLUS", "BAL"}
        s = (sbu or "").strip().upper()
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb in b_codes:
                return val
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb in b_codes:
                return val
        for (kc, _, ks, kb), val in self.lookup.items():
            if kb in b_codes:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


def _find_rna_source_file(attr_dir: Path) -> Path | None:
    """Locate the Brand mapping / Attributes workbook that contains an RNA sheet."""
    search_dirs = [
        attr_dir,
        attr_dir.parent,
        attr_dir.parent.parent,
        Path.cwd(),
        Path.cwd().parent,
    ]
    seen = set()
    for search_dir in search_dirs:
        if not search_dir or not search_dir.exists() or search_dir in seen:
            continue
        seen.add(search_dir)
        for f in sorted(search_dir.glob("*.xlsx")):
            if "brand mapping" in f.name.lower() or "brand_mapping" in f.name.lower() or "template" in f.name.lower():
                try:
                    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                    has_rna = any(any(kw in s.upper() for kw in RNALoader.RNA_SHEET_KEYWORDS) for s in wb.sheetnames)
                    wb.close()
                    if has_rna:
                        return f
                except Exception as e:
                    log.warning("[RNA] Cannot inspect %s: %s", f.name, e)
    return None


# ── Country of Origin text → ISO LOV Code map ────────────────────────────
# Source: MDD "Country Origin LOV" tab
_COO_TEXT_MAP: dict[str, str] = {
    "CHINA":        "CN", "CN": "CN",
    "VIETNAM":      "VN", "VN": "VN",
    "SINGAPORE":    "SG", "SG": "SG",
    "HONG KONG":    "HK", "HK": "HK",
    "MALAYSIA":     "MY", "MY": "MY",
    "USA":          "US", "US":  "US", "UNITED STATES": "US",
    "THAILAND":     "TH", "TH": "TH",
    "AUSTRALIA":    "AU", "AU": "AU",
    "INDIA":        "IN", "IN": "IN",
    "CAMBODIA":     "KH", "KH": "KH",
    "ITALY":        "IT", "IT": "IT",
    "UNITED KINGDOM": "GB", "GB": "GB", "UK": "GB",
    "JAPAN":        "JP", "JP": "JP",
    "SRI LANKA":    "LK", "LK": "LK",
    "PORTUGAL":     "PT", "PT": "PT",
    "PHILIPPINES":  "PH", "PH": "PH",
    "SLOVENIA":     "SI", "SI": "SI",
    "GERMANY":      "DE", "DE": "DE",
    "BRAZIL":       "BR", "BR": "BR",
    "FRANCE":       "FR", "FR": "FR",
    "MEXICO":       "MX", "MX": "MX",
    "TURKEY":       "TR", "TR": "TR",
    "INDONESIA":    "ID", "ID": "ID",
    "SOUTH AFRICA": "ZA", "ZA": "ZA",
    "BANGLADESH":   "BD", "BD": "BD",
    "PAKISTAN":     "PK", "PK": "PK",
    "TAIWAN":       "TW", "TW": "TW",
    "HONDURAS":     "HN", "HN": "HN",
    "AFGHANISTAN":  "AF", "AF": "AF",
    "HONGKONG":     "HK",
}


def _resolve_coo(raw: str) -> str:
    """Map Excel country name to MDD Country Origin LOV ID. Default: HK."""
    if not raw:
        return "HK"
    s = (raw or "").strip().upper()
    val = _COO_TEXT_MAP.get(s, "")
    if val:
        return val
    if len(s) == 2:
        return s
    return "HK"


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILD (ONLY GENERIC PRODUCTS)
# ══════════════════════════════════════════════════════════════════════════════

def _build_generic_product(group: implusGenericGroup, cfg: dict) -> ET.Element:
    brand_code   = cfg.get("brand_code", "IPL")
    brand_name   = cfg.get("brand_name", "IMPLUS")
    sea_prefix   = cfg.get("season_prefix", "SM")
    sea_year     = cfg.get("season_year") or str(datetime.now().year)
    country_code = cfg.get("country_code", "ID")
    mdd: Optional[MDDLoader] = cfg.get("mdd")

    generic_key = f"{brand_code}{group.implus_us}"

    # First variant data (for FOB, CountryOrigin)
    first_var = group.variants[0] if group.variants else None

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_E-TempSubCat")  # implus = accessories/socks

    _keyval(gen_el, "KEY_InboundArticle", generic_key)

    name_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name")
    name_el.text = f"{brand_code} {group.style_desc}".strip()
    
    # ── Classification References ─────────────────────────────────────────────
    # Must mirror the same brand normalization used in build_classifications()
    b_name = "Implus" if (brand_name or "").strip().upper() in ("implus", "IMP", "IMPLUS") else brand_name.strip()
    b_code = "IPL" if (brand_code or "").strip().upper() in ("implus", "IMP", "IPL") else brand_code.strip()

    _clf(gen_el, f"CLH_{b_name}Articles", "CPL_Merchandiser")
    _clf(gen_el, f"CLH_{b_code}_{sea_prefix}{sea_year}UA", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")

    # ── Brand Attributes ──────────────────────────────────────────────────────
    # MDD: Brand LOV — brand_code IS the LOV ID (e.g. IPL for Implus)
    b_lov_id = BRAND_LOV_MAP.get(brand_code.upper(), BRAND_LOV_MAP.get(brand_name.upper(), brand_code))
    _val(vals_el, "AT_Brand",        "", id_val=b_lov_id)
    # MDD: Brand Group — brand_name upper (e.g. IMPLUS / BALEGA SOCKS)
    b_grp_id = _resolve_brand_group(brand_name)
    _val(vals_el, "AT_BrandGroup",   "", id_val=b_grp_id)
    # MDD: Brand Status LOV — A=ACTIVE
    _val(vals_el, "AT_BrandStatus",  "", id_val="A")

    raw_brand_type = (cfg.get("brand_type") or "MAA").strip()
    bt_id = _resolve_brand_type(raw_brand_type)
    if bt_id:
        _val(vals_el, "AT_BrandType", id_val=bt_id)
    else:
        _val(vals_el, "AT_BrandType", value=raw_brand_type)

    # AT_BrandCategory should NOT be populated
    # raw_brand_cat = (cfg.get("brand_category") or f"{country_code} - SD - NON TOP").strip()
    # if raw_brand_cat:
    #     _val(vals_el, "AT_BrandCategory", value=raw_brand_cat)

    # ── RNA / Portal Attributes ───────────────────────────────────────────────
    _multival(vals_el, "AT_SBU",         cfg.get("sbu", "FF"))
    _multival(vals_el, "AT_CompanyCode", cfg.get("comp_code", "0888"))
    # MDD: Country LOV — country code (ID, MY, PH etc.)
    _val(vals_el, "AT_Country",          "", id_val=country_code)

    # MDD: Retail Price Currency LOV — IDR/PHP/THB/SGD/MYR/VND/USD
    _RETAIL_CURRENCY_MAP: dict[str, str] = {
        "ID": "IDR", "PH": "PHP", "TH": "THB",
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    retail_currency = _RETAIL_CURRENCY_MAP.get(country_code, "")
    if retail_currency:
        _val(vals_el, "AT_RetailPriceCurrency", "", id_val=retail_currency)

    # ── Season ────────────────────────────────────────────────────────────────
    # MDD: Season LOV — SP/SM/FL/WN/CO/SS/FW/AL
    _val(vals_el, "AT_Season",     "", id_val=sea_prefix)
    # MDD: SeasonYear — number (text value, not LOV)
    _val(vals_el, "AT_SeasonYear", sea_year)

    # ── Inbound Code ─────────────────────────────────────────────────────────
    _val(vals_el, "AT_InboundGenericCode", generic_key)

    # ── Principal Attributes ─────────────────────────────────────────────────
    _val(vals_el, "AT_PrincipalStyleCode",        group.implus_us)
    _val(vals_el, "AT_PrincipalStyleDescription", group.principal_style_description)
    if group.color_code:
        _val(vals_el, "AT_PrincipalColorCode", group.color_code)
    if group.color:
        _val(vals_el, "AT_PrincipalColorName", group.color)
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", group.hierarchy_l1)

    # ── From First Variant ────────────────────────────────────────────────────
    # AT_CountryOrigin: Always HK (Hong Kong) for implus
    _val(vals_el, "AT_CountryOrigin", "", id_val="HK")

    if first_var and first_var.size:
        _val(vals_el, "AT_PrincipalSize", first_var.size)

    if first_var and first_var.distributor_price:
        _val(vals_el, "AT_FOB", first_var.distributor_price)

    # ── Defaults (from MDD) ───────────────────────────────────────────────────
    # MDD: FOB Currency LOV — USD
    _val(vals_el, "AT_FOBCurrency",    "", id_val="USD")
    # MDD: SAP Product Flag LOV — A=Intercompany
    _val(vals_el, "AT_SAPProductFlag", "", id_val="A")
    # MDD: SAP Article Type LOV — ZINA=Intercompany Articles
    _val(vals_el, "AT_MaterialType",   "", id_val="ZINA")
    # MDD: Article Category LOV — 01=Generic (STIBO format)
    _val(vals_el, "AT_SAPArticleCategory", "", id_val="01")
    # MDD: UOM LOV — EA=Each
    _val(vals_el, "AT_UOM",            "", id_val="EA")

    # MDD: Gender LOV — U=Unisex (SAP Gender code)
    _val(vals_el, "AT_Gender",  "", id_val="U")
    # MDD: Age LOV — AD=Adults (SAP Age code)
    _val(vals_el, "AT_SAPAge",  "", id_val="AD")

    # MDD: BY Gender LOV — F=Female, M=Male, U=Unisex
    by_gender_raw = cfg.get("by_gender", "Unisex")
    by_gender_id = "F" if "FEMALE" in by_gender_raw.upper() else ("M" if "MALE" in by_gender_raw.upper() else "U")
    _val(vals_el, "AT_BYGender", "", id_val=by_gender_id)

    # MDD: Age LOV BY Age column — via MDD lookup, fallback ADULT
    by_age_id = "ADULT"
    if mdd:
        by_age_id = (mdd.lookup("BYAGE", "Adult")
                     or mdd.lookup("BY AGE", "Adult")
                     or mdd.lookup("AGE", "Adults")
                     or "ADULT")
    _val(vals_el, "AT_BYAge",    "", id_val=by_age_id)

    # MDD: BY Article Type LOV — Inline/License/SSE
    _val(vals_el, "AT_BYArticleType", "", id_val=cfg.get("article_type", "Inline"))

    # MDD: Pack Details LOV — S=Single (or dynamic via MDD/PACK_DETAILS_LOV_MAP)
    pack_details_raw = cfg.get("pack_details", "Single")
    pack_details_id = _resolve_pack_details(pack_details_raw, mdd)
    _val(vals_el, "AT_PackDetails", "", id_val=pack_details_id)

    # AT_CountrySize should NOT be sent
    # _val(vals_el, "AT_CountrySize", "", id_val="NS")

    # MDD: BCI LOV — Best Commercial Image (COMMERCIAL is the standard ID)
    _val(vals_el, "AT_BCI", "", id_val="COMMERCIAL")

    # AT_EComAgesCategory should NOT be sent
    # _val(vals_el, "AT_EComAgesCategory", "", id_val="18+Y")

    # MDD: Article Status — A=Active (same pattern as Brand Status A=ACTIVE)
    _val(vals_el, "AT_ArticleStatus", "", id_val="A")

    # MDD: Pricing Distribution Channel LOV — 01=Retail
    _val(vals_el, "AT_PricingDistributionChannel", "", id_val="01")
    

    return gen_el

def build_xml(rows: list[implusRow], out_xml_path: Path, cfg: dict) -> None:
    groups = group_generics(rows)

    # ──────────────────────────────────────────────────────────────────────────
    # Use ONLY the configured/specified season for Classifications
    # (Do NOT scan rows for multiple seasons — generate exactly 1 classification)
    # ──────────────────────────────────────────────────────────────────────────
    fallback_year   = cfg.get("season_year") or str(datetime.now().year)
    fallback_prefix = cfg.get("season_prefix", "SM")
    # Single season only: use provided season_prefix and season_year from config
    season_list = [(fallback_prefix, fallback_year)]

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    brand_code = cfg.get("brand_code", "BAL")
    brand_name = cfg.get("brand_name", "implus")

    log.info("[implus] Writing STEP XML → %s", out_xml_path.name)
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

    log.info("[implus] XML generated successfully: %d Generic product nodes.", total_generics)


def run(args, auditor=None) -> tuple[list[dict], Path | None]:
    log.info("[implus-Linelist] ETL started — brand=%s", getattr(args, "brand", "implus"))

    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        file_path = Path(input_file_arg)
    else:
        files = _find_input_files()
        if not files:
            log.warning("[implus] No linelist .xlsx/.xlsm files found — nothing to process.")
            return [], None
        file_path = max(files, key=lambda p: p.stat().st_mtime)

    log.info("[implus] Processing file: %s", file_path.name)

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

    brand_code   = getattr(args, "brand_code",   "IPL")    or "IPL"
    brand_name   = getattr(args, "brand",        "IMPLUS") or "IMPLUS"
    comp_code    = getattr(args, "comp_code",    "0888")   or "0888"
    sbu          = getattr(args, "sbu",          "FF")     or "FF"
    country_code = getattr(args, "country_code", "ID")     or "ID"

    # ── Resolve brand_lov_id from input filename → MDD BrandLOV ─────────────
    # Pattern: parts[2] of dash-split stem = brand token (e.g. "IMPLUS" / "BALEGA").
    # MDD BrandLOV: display name → LOV ID (e.g. "IMPLUS" → "IPL").
    # Fallback chain: MDD BrandLOV → hardcoded BRAND_LOV_MAP → brand_code default.
    brand_from_file = brand_name
    filename_stem   = file_path.stem
    stem_parts      = filename_stem.split("-")
    if len(stem_parts) >= 3:
        brand_token = stem_parts[2].strip().upper()
        if brand_token:
            brand_from_file = brand_token.title()  # e.g. "Implus" / "Balega"
            log.info("[implus] Brand token from filename: '%s' → '%s'", brand_token, brand_from_file)

    brand_lov_id = brand_code  # fallback
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            looked_up    = brand_lov.get(brand_from_file.upper())
            brand_lov_id = looked_up if looked_up else BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.info(
                "[implus] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_from_file.upper(), brand_lov_id,
                "(FOUND in MDD)" if looked_up else "(NOT FOUND in MDD — using fallback)",
            )
        else:
            brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.warning("[implus] BrandLOV sheet not in MDD — using BRAND_LOV_MAP fallback: '%s'", brand_lov_id)
    else:
        brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
        log.warning("[implus] No MDD loaded — brand_lov_id='%s' from BRAND_LOV_MAP", brand_lov_id)

    # Also derive country_id from filename (2nd-to-last dash token, e.g. "...SP2029-ID-1")
    filename_country_id = ""
    if len(stem_parts) >= 2:
        candidate = stem_parts[-2].strip().upper()
        if re.match(r'^[A-Z]{2}$', candidate):
            filename_country_id = candidate
            log.info("[implus] Country ID from filename: '%s'", filename_country_id)
    if filename_country_id:
        country_code = filename_country_id

    # ── Brand Type / Brand Category from RNA ─────────────────────────────────
    brand_type   = getattr(args, "brand_type",     "") or ""
    brand_cat    = getattr(args, "brand_category", "") or ""
    brand_group  = getattr(args, "brand_group",    "") or ""
    brand_status = getattr(args, "brand_status",   "") or ""

    if not brand_type or not brand_cat:
        rna_path = _find_rna_source_file(attr_dir)
        if rna_path:
            try:
                rna          = RNALoader(rna_path)
                country_name = _COUNTRY_MAP.get(country_code.upper(), country_code.upper())
                result = rna.get(country_name, comp_code, sbu, brand_lov_id)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_lov_id)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_code)
                brand_type   = brand_type   or result.get("brand_type",     "")
                brand_cat    = brand_cat    or result.get("brand_category", "")
                brand_group  = brand_group  or result.get("brand_group",    "")
                brand_status = brand_status or result.get("brand_status",   "")
                log.info(
                    "[Impulse] RNA → country=%s comp=%s sbu=%s brand=%s → type='%s' cat='%s'",
                    country_name, comp_code, sbu, brand_code, brand_type, brand_cat,
                )
            except Exception as e:
                log.warning("[implus] RNA lookup failed: %s", e)
        else:
            log.warning("[implus] No RNA source workbook found in %s", attr_dir)

    cfg = {
        "brand_code":    brand_lov_id,    # resolved LOV ID from MDD BrandLOV
        "brand_name":    brand_from_file,  # display name from filename
        "sbu":           sbu,
        "comp_code":     comp_code,
        "country_code":  country_code,
        "season_prefix": getattr(args, "season_prefix", None) or "SM",
        "season_year":   getattr(args, "season_year", None) or None,
        "article_type":  getattr(args, "article_type", "Inline") or "Inline",
        "brand_type":    brand_type,
        "brand_category": brand_cat,
        "brand_group":   brand_group,
        "brand_status":  brand_status,
        "mdd":           mdd,    # MDDLoader instance (or None)
    }

    rows = load_linesheet(file_path)
    if not rows:
        log.warning("[implus] No valid rows parsed.")
        return [], None

    xml_path = xml_dir / f"{file_path.stem}.xml"
    build_xml(rows, xml_path, cfg)

    if auditor and xml_path:
        auditor.set_xml_uploads([str(xml_path)])

    return [], xml_path


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="implus Line List → STEP XML")
    ap.add_argument("--brand",        default="IMPLUS")
    ap.add_argument("--brand-code",   default="IPL", dest="brand_code")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="FF")
    ap.add_argument("--country-code", default="ID", dest="country_code")
    run(ap.parse_args())