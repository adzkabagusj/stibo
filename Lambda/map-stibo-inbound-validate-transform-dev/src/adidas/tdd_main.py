"""
tdd_main.py
===========
Standalone TDD (ROS) pipeline for Adidas.

FIX LOG (v16):
    - FIX 57: Deduplication — Variants are now deduplicated by v_key to prevent redundant tags in XML.
    - FIX 58: Validation — Article numbers shorter than 2 characters (like "v") are ignored.
    - FIX 59: Lazy Generation — No empty tags written for manual attributes (NOA, Ecom, etc.).
    - FIX 60: Robust MDD — Full color aliases and gender/age mapping restored.
"""

import logging
import os
import re
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

import openpyxl
import pandas as pd

log = logging.getLogger("tdd_main")

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace('', STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')

DIVISION_PARENT_MAP = {
    "ACCESSORIES": "E", "FOOTWEAR": "F", "APPAREL": "A",
    "EQUIPMENT": "Q", "HARDWARE": "Q", "TOYS": "T",
}

_BY_AGE_ID_MAP = {
    "AD": "ADULT", "AA": "ALL", "CH": "CHILDREN",
    "IN": "INFANT", "JR": "JUNIOR",
}

_STIBO_GENDER_VALID = {"F", "M", "U"}

_GENDER_FALLBACK = {
    "W": "F", "WOMEN": "F", "WOMAN": "F", "FEMALE": "F", "GIRLS": "F", "GIRL": "F",
    "M": "M", "MEN": "M", "MAN": "M", "MALE": "M", "BOYS": "M", "BOY": "M",
    "U": "U", "UNISEX": "U", "MIXED": "U", "F": "F",
}

TDD_SHEET  = "Spreadsheet"
HEADER_ROW = 4   
DATA_START = 6   

# =============================================================================
# HELPERS
# =============================================================================

def _s(v) -> str:
    if v is None: return ""
    try:
        if pd.isnull(v): return ""
    except: pass
    if isinstance(v, datetime): return v.strftime("%d.%m.%Y")
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "0") else s

def _norm(text: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())

def _get_division_code(division: str) -> str:
    div_upper = (division or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in div_upper: return code
    return "X"

def _parse_season(season_raw: str) -> tuple:
    s = (season_raw or "").strip().upper()
    if not s: return "SS", "26", "2026"
    match = re.search(r"\((SM|FW|AW|AL)\)", s)
    sea_pre = match.group(1) if match else ("FW" if any(x in s for x in ["FW", "FALL", "WINTER"]) else "SS")
    if sea_pre == "SM": sea_pre = "SS"
    y_match = re.search(r"(\d{4})", s)
    f_y = y_match.group(1) if y_match else "2026"
    return sea_pre, f_y[2:], f_y

def _size_to_3digit(size_val: str, tech_size: str = "") -> str:
    raw = _normalize_tdd_sap_size(size_val) or _normalize_tdd_sap_size(tech_size)
    clean = re.sub(r"[^A-Z0-9\-/]", "", raw.upper())
    return clean[-3:].ljust(3, "0") if clean else "000"

def _normalize_tdd_sap_size(raw: str) -> str:
    val = (raw or "").strip().upper()
    if not val:
        return ""
    half = re.match(r"^0*(\d+)(?:\s*-|[.,]5)$", val)
    if half:
        return f"{int(half.group(1))}H"
    return val

def _is_numeric_only(s: str) -> bool:
    stripped = s.strip()
    if not stripped: return False
    if len(stripped) > 1 and stripped[0] == '0': return False
    return bool(re.match(r"^\d+$", stripped))

# =============================================================================
# MDD LOADER (Robust Version)
# =============================================================================

class MDDLoader:
    _ADIDAS_COLOR_ALIASES = {
        "COREBLACK": "005", "COREWHITE": "W", "FTWRWHITE": "W", "OFFWHITE": "W",
        "CLOUDWHITE": "W", "WONDERWHITE": "W", "COLLEGIATENAVY": "NAV", "LEGENDINK": "INK",
        "BETTERSCARLET": "R", "LUCIDBLUE": "12W",
    }

    def __init__(self, path: Path):
        self.lovs = {}
        self._color_desc_to_id = {}
        self._color_id_to_desc = {}
        self._color_3digit = {}
        self._gender = {}
        self._age = {}
        self._season = {}
        self._size_lov = {}
        if path: self._load(path)

    def _load(self, path):
        log.info(f"[MDD] Loading: {path.name}")
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        
        if "Color Code LOV" in wb.sheetnames:
            for row in wb["Color Code LOV"].iter_rows(min_row=2, values_only=True):
                if row[0]:
                    lov_id, desc = str(row[0]).strip(), str(row[1]).upper() if row[1] else ""
                    self._color_id_to_desc[_norm(lov_id)] = desc
                    if desc: self._color_desc_to_id[_norm(desc)] = lov_id
                    alpha = re.sub(r"[^A-Z]", "", desc)
                    self._color_3digit[_norm(lov_id)] = alpha[:3] if alpha else "XXX"

        # Load Gender LOV dynamically (Col A → Col C mapping)
        self._load_gender_lov(wb)

        # Load Age LOV dynamically (Col A → Col B mapping)
        self._load_age_lov(wb)

        if "Gender LOV" in wb.sheetnames:
            for row in wb["Gender LOV"].iter_rows(min_row=3, values_only=True):
                lbl, sid = _s(row[0]), _s(row[1])
                if sid in _STIBO_GENDER_VALID: self._gender[_norm(lbl)] = (sid, lbl)

        if "Age LOV" in wb.sheetnames:
            for row in wb["Age LOV"].iter_rows(min_row=2, values_only=True):
                lbl, code = _s(row[0]), _s(row[1])
                if code: self._age[_norm(lbl)] = (code, lbl)

        if "Size Code LOV" in wb.sheetnames:
            for row in wb["Size Code LOV"].iter_rows(values_only=True):
                if len(row) > 5:
                    c3, c5 = _s(row[3]), _s(row[5])
                    if c5:
                        if c3: self._size_lov[c3.upper()] = c5
                        self._size_lov[c5.upper()] = c5
        wb.close()

    def _load_gender_lov(self, wb):
        """
        Load Gender LOV: Col A = display value (e.g. 'Male'),
                         Col C = LOV ID (e.g. 'M').
        Also registers common aliases (W, M, U, MEN, WOMEN, FEMALE, etc.).
        """
        if "Gender LOV" not in wb.sheetnames:
            log.warning("[MDD] Gender LOV sheet not found")
            return
        gender_lov = {}
        for row in wb["Gender LOV"].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            display_value = str(row[0]).strip()           # Col A = display value
            lov_id        = str(row[2]).strip() if len(row) > 2 and row[2] else ""  # Col C = LOV ID
            if display_value and lov_id:
                gender_lov[display_value.upper()] = lov_id
        # Register common aliases so short codes / alternate spellings resolve too
        for alias, canonical in [
            ("W",     "FEMALE"),
            ("F",     "FEMALE"),
            ("M",     "MALE"),
            ("MEN",   "MALE"),
            ("WOMEN", "FEMALE"),
            ("U",     "UNISEX"),
        ]:
            if canonical in gender_lov:
                gender_lov.setdefault(alias, gender_lov[canonical])
        self.lovs["GenderLOV"] = gender_lov
        log.info("[MDD] Gender LOV loaded: %d entries", len(gender_lov))

    def _load_age_lov(self, wb):
        """
        Load Age LOV: Col A = display value (e.g. 'Children'),
                      Col B = LOV ID (e.g. 'K').
        """
        if "Age LOV" not in wb.sheetnames:
            log.warning("[MDD] Age LOV sheet not found")
            return
        age_lov = {}
        for row in wb["Age LOV"].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            display_value = str(row[0]).strip()           # Col A = display value
            lov_id        = str(row[1]).strip() if len(row) > 1 and row[1] else ""  # Col B = LOV ID
            if display_value and lov_id:
                age_lov[display_value.upper()] = lov_id
        self.lovs["AgeLOV"] = age_lov
        log.info("[MDD] Age LOV loaded: %d entries", len(age_lov))

    def color_to_lov_id(self, raw: str) -> str:
        if not raw: return ""
        norm = _norm(raw)
        if norm in self._ADIDAS_COLOR_ALIASES: return self._ADIDAS_COLOR_ALIASES[norm]
        return self._color_desc_to_id.get(norm, "")

    def color_3digit(self, lov_id: str) -> str:
        return self._color_3digit.get(_norm(lov_id), "XXX")

    def gender(self, raw: str) -> tuple:
        if not raw: return ("U", "Unisex")
        sid = _GENDER_FALLBACK.get(raw.strip().upper(), "U")
        return (sid, "Male" if sid=="M" else "Female" if sid=="F" else "Unisex")

    def age(self, raw: str) -> tuple:
        if not raw: return ("AD", "Adults")
        return self._age.get(_norm(raw), ("AD", "Adults"))

    def size_to_lov_id(self, raw: str) -> str:
        clean = _normalize_tdd_sap_size(raw)
        if not clean:
            return ""
        return self._size_lov.get(clean, clean)

    def resolve_size_lov_id(self, size_raw: str, tech_size: str) -> str:
        if (size_raw or "").strip():
            return self.size_to_lov_id(size_raw)
        return self.size_to_lov_id(tech_size)

# =============================================================================
# TDD LOADER
# =============================================================================

class TDDLoader:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.article_meta = {}
        self.by_article = {}
        self._load()

    def _load(self):
        log.info(f"[TDDLoader] Reading file: {self.path.name}")
        wb = openpyxl.load_workbook(str(self.path), read_only=True, data_only=True)
        target = "Spreadsheet" if "Spreadsheet" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if len(rows) <= HEADER_ROW: return
        hdr = rows[HEADER_ROW]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}

        for row in rows[DATA_START:]:
            art_no = _s(row[col.get("Article No", 0)])
            # FIX 58: Validation — Skip noise like "v"
            if not art_no or len(art_no) < 2: continue

            if art_no not in self.article_meta:
                self.article_meta[art_no] = {
                    "article_no": art_no, 
                    "material_desc": _s(row[col.get("IAM Material description", 0)]),
                    "color_description": _s(row[col.get("Color/ Description 1", 0)]),
                    "product_division": _s(row[col.get("Product Division", 0)]),
                    "rrp": _s(row[col.get("RRP FINAL", 0)]),
                    "currency": _s(row[col.get("Currency", 0)]), 
                    "gender": _s(row[col.get("Gender", 0)]),
                    "age_group": _s(row[col.get("Age Group", 0)]),
                }
            
            self.by_article.setdefault(art_no, []).append({
                "tech_size": _s(row[col.get("Tech.Size", 0)]),
                "size": _s(row[col.get("Size", 0)]), 
                "ean": _s(row[col.get("EAN/UPC", 0)])
            })

    @property
    def all_articles(self): return list(self.article_meta.keys())
    def get_meta(self, art_no): return self.article_meta.get(art_no)
    def get_sizes(self, art_no): return self.by_article.get(art_no, [])

# =============================================================================
# XML WRITER
# =============================================================================

def _val(parent, attr_id: str, value: str = "", id_val: str = ""):
    c_id, c_val = str(id_val).strip(), str(value).strip()
    if not c_id and not c_val: return None
    if c_id in ("None", "nan") and not c_val: return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value", AttributeID=attr_id)
    if c_id and c_id not in ("None", "nan"): el.set("ID", c_id)
    else: el.text = c_val
    return el

def _multival(parent, attr_id: str, id_val: str):
    if not id_val or str(id_val).strip() in ("None", "nan", ""): return
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue", AttributeID=attr_id)
    ET.SubElement(mv, f"{{{STIBO_NS}}}Value", ID=str(id_val).strip())

def _add_generic_values(vals_el, art, brand_code, comp_code, sbu, mdd, args):
    sea_pre, _, full_yr = _parse_season(args.season)
    _multival(vals_el, "AT_SBU", sbu or "SP")
    _multival(vals_el, "AT_CompanyCode", comp_code or "0888")
    
    if hasattr(args, "country_code") and args.country_code:
        _val(vals_el, "AT_Country", id_val=args.country_code)
    
    _val(vals_el, "AT_Brand", id_val=brand_code)
    _val(vals_el, "AT_PrincipalStyleCode", art["article_no"])
    _val(vals_el, "AT_SAPStyleCode", art["article_no"])
    _val(vals_el, "AT_PrincipalStyleDescription", art["material_desc"])
    # _val(vals_el, "AT_PrincipalColorCode", art["color_description"])
    
    c_lov = mdd.color_to_lov_id(art["color_description"])
    if c_lov: _val(vals_el, "AT_Color", id_val=c_lov)
    
    # ── Gender ── (dynamic lookup from MDD Gender LOV: Col A → Col C) ──────────
    gender_lov    = mdd.lovs.get("GenderLOV") or {}
    gender_raw_up = (art.get("gender") or "").strip().upper()
    # Fallback using old gender method for backward compatibility
    gender_code, _ = mdd.gender(art["gender"])
    g_lov_id      = (
        gender_lov.get(gender_raw_up)
        or gender_lov.get(gender_code.upper())
        or gender_code
    )
    _val(vals_el, "AT_Gender", id_val=g_lov_id)
    
    # ── Age ── (dynamic lookup from MDD Age LOV: Col A → Col B) ──────────────
    age_lov       = mdd.lovs.get("AgeLOV") or {}
    age_raw_up    = (art.get("age_group") or "").strip().upper()
    # Fallback using old age method for backward compatibility
    age_code, _ = mdd.age(art["age_group"])
    sap_age_id    = (
        age_lov.get(age_raw_up)
        or age_lov.get(age_code.upper())
        or age_code
    )
    _val(vals_el, "AT_SAPAge", id_val=sap_age_id)
    
    _val(vals_el, "AT_Season", id_val=sea_pre)
    _val(vals_el, "AT_SeasonYear", full_yr)
    _val(vals_el, "AT_SAPArticleCategory", id_val="1")
    _val(vals_el, "AT_BYArticleType", id_val=getattr(args, "article_type", "Inline"))
    _val(vals_el, "AT_UOM", id_val="EA")
    _val(vals_el, "AT_OriginalPrice", art["rrp"])
    _val(vals_el, "AT_RetailPriceCurrency", id_val=art["currency"])
    _val(vals_el, "AT_InboundGenericCode", f"{brand_code}{art['article_no']}")

def _build_product_xml(art, brand, brand_code, comp_code, sbu, season_id, mdd, args) -> str:
    div = _get_division_code(art["product_division"])
    g_el = ET.Element(f"{{{STIBO_NS}}}Product", UserTypeID="PRD_GenericArticle", ParentID=f"PPH_{div}-TempSubCat", update="true")
    ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue", KeyID="KEY_InboundArticle").text = f"{brand_code}{art['article_no']}"
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["material_desc"] or art["article_no"]
    
    ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference", ClassificationID=f"CLH_{brand.capitalize()}Articles", Type="CPL_Merchandiser")
    ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference", ClassificationID=f"{season_id}UA", Type="CPL_UnConfirmedForSeason")
    
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand_code, comp_code, sbu, mdd, args)
    
    # FIX 57: Deduplicate Variants
    c_lov = mdd.color_to_lov_id(art["color_description"])
    seen_variants = set()
    
    for sz in art.get("_tdd_sizes", []):
        # Resolve size LOV ID from MDD
        s_lov = mdd.resolve_size_lov_id(sz["size"], sz["tech_size"])
        
        # KEY_InboundVariant = KEY_InboundArticle + Size LOV ID
        inbound_article = f"{brand_code}{art['article_no']}"
        v_key = f"{inbound_article}{s_lov}" if s_lov else f"{inbound_article}{_size_to_3digit(sz['size'], sz['tech_size'])}"
        
        if v_key in seen_variants: continue
        seen_variants.add(v_key)
        
        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product", UserTypeID="PRD_VariantArticle", update="true")
        ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue", KeyID="KEY_InboundVariant").text = v_key
        vv = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
        if s_lov: _val(vv, "AT_Size", id_val=s_lov)
        _val(vv, "AT_PrincipalSize", sz["size"] or sz["tech_size"])
        _val(vv, "AT_InboundVariantCode", v_key)
        
    return ET.tostring(g_el, encoding="unicode")

# =============================================================================
# MAIN RUNNER
# =============================================================================

def run_tdd(args, auditor=None, tmp_dir=None):
    base = Path(tmp_dir) if tmp_dir else Path("/Users/ducomac12/.gemini/antigravity/scratch")
    tdd_dir, mdd_dir, xml_out = base/"input/tdd", base/"input/mdd", base/"output/xml"
    xml_out.mkdir(parents=True, exist_ok=True)
    
    tdd_files = list(tdd_dir.glob("*.xlsx"))
    if not tdd_files: return
    
    mdd_file = list(mdd_dir.glob("*.xlsx"))[0] if list(mdd_dir.glob("*.xlsx")) else None
    mdd = MDDLoader(mdd_file)
    loader = TDDLoader(tdd_files[0])
    
    sea_pre, _, full_yr = _parse_season(args.season)
    season_id = f"CLH_{args.brand_code}_{sea_pre}{full_yr}"
    sbu_name = getattr(args, "sbu_name", "MAA Sport")
    art_type = getattr(args, "article_type", "Inline")
    multi_mono = getattr(args, "multi_mono", "Multi")
    country_code = getattr(args, "country_code", "")
    
    out_name = f"{args.comp_code}-{args.sbu}-ADIDAS-TDD & Price List ({sbu_name}) {art_type}-{multi_mono}-{sea_pre}{full_yr}-{country_code}-1.xml"
    out_path = xml_out / out_name
    
    with open(out_path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(f'<STEP-ProductInformation xmlns="{STIBO_NS}" xmlns:xsi="{STIBO_XSI}" xsi:schemaLocation="{STIBO_SCHEMA}" ExportTime="{datetime.now()}" update="true">\n')
        f.write(f'  <Classifications><Classification ID="{season_id}" UserTypeID="CLS_Season" ParentID="CLH_{args.brand.capitalize()}Batches" update="true"><Name>{args.brand} {sea_pre} {full_yr}</Name><Classification ID="{season_id}UA" UserTypeID="CLS_UnconfirmedArticles" update="true"><Name>{sea_pre} {full_yr} Unconfirmed</Name></Classification></Classification></Classifications>\n')
        f.write("  <Products>\n")
        # ── TEST MODE: limit to 5 articles ──────────────────────────
        # loader.article_meta = dict(list(loader.article_meta.items())[:5])
        for art_no in loader.all_articles:
            art = loader.get_meta(art_no)
            art["_tdd_sizes"] = loader.get_sizes(art_no)
            pxml = _XMLNS_RE.sub("", _build_product_xml(art, args.brand, args.brand_code, args.comp_code, args.sbu, season_id, mdd, args))
            f.write(f"    {pxml}\n")
        f.write("  </Products>\n</STEP-ProductInformation>\n")
    
    log.info(f"✓ TDD XML written: {out_path.name}")
