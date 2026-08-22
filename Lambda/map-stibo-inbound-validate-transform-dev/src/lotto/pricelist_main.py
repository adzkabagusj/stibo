"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — Lotto Price List (Inline + Licensed) v1.0    ║
║  "Format_FW26 Price List USD Licensee.xlsx" → Stibo STEP XML                ║
╚══════════════════════════════════════════════════════════════════════════════╝

Source file format : "Format_FW26 Price List USD Licensee.xlsx"
Sheet name         : "Sheet 1" or first non-empty sheet

═══════════════════════════════════════════════════════════════════════════════
BRAND MAPPING TEMPLATE: "Lotto Inline (New Format)" Sheet
═══════════════════════════════════════════════════════════════════════════════

MAPPED ATTRIBUTES (from Price List - Sheet 1):
─────────────────────────────────────────────────────────────────────────────────
Stibo Attribute ID                | Field Name in Brand File
─────────────────────────────────────────────────────────────────────────────────
AT_PrincipalStyleCode             | Style
AT_PrincipalStyleDescription      | Style description
AT_PrincipalColorCode             | Color Code
AT_PrincipalColorName             | Color description
AT_PrincipalGenderDescription     | Gender
AT_PrincipalSize                  | Size range
AT_SAPStyleCode                   | Style
AT_Generic                        | Style (formula: Brand Code + Style)
AT_GenericDescription             | Style description
AT_Variant                        | Style (formula: Generic + Color 3D)
AT_VariantDescription             | Style description
AT_Color                          | Pic (AI-based from images)
AT_SAPAge                         | Gender (derived mapping)
AT_Gender                         | Gender (derived mapping)
AT_BYAge                          | Gender (derived mapping)
AT_BYGender                       | Gender (derived mapping)
AT_CountryOrigin                  | MADE IN
AT_FOB                            | LICENSEE PRICE DIRECT BMOQ
AT_FOBCurrency                    | CURRENCY
AT_PrincipalMerchandiseHierarchyL1| Product Type
AT_PrincipalMerchandiseHierarchyL2| Function
AT_MaterialUpper                  | PRODUCT COMPOSITION 1
─────────────────────────────────────────────────────────────────────────────────

UNMAPPED ATTRIBUTES (NOT included in XML):
─────────────────────────────────────────────────────────────────────────────────
AT_Country, AT_CompanyCode, AT_SBU, AT_Brand (manual portal input)
AT_BYIndicator, AT_SAPIndicator, AT_EcomIndicator (system calculated)
AT_PrincipalGenderCode, AT_PrincipalAgeCode, AT_PrincipalAgeDescription (not available)
AT_Size (BY feedback)
AT_Createdon, AT_Createdby (system default)
AT_MusicalBoxArticleCode, AT_MBStyleCode, AT_MBAssortmentCode (BY feedback)
AT_MusicalBoxArticleDescription, AT_MusicalBoxComponentsQty, AT_MusicalBoxUOM (BY feedback)
AT_Season, AT_SeasonYear (manual portal input)
AT_Vendor, AT_MainVendorIdentification (manual input)
AT_ExchangeRate (not available)
AT_PricingDistributionChannel (default 01 = Retailer)
AT_RetailPriceCurrency (formula based on country)
AT_OriginalPrice, AT_CurrentPrice (BY feedback)
AT_DiscountBucket, AT_DiscountBucketRounded (formula)
AT_EstimatedLandedCost (BY feedback)
AT_FreightCost, AT_Royalty, AT_RoyaltyDeduction (not available)
AT_SGS, AT_MarketingFee, AT_HaddadOfficeCharge (not available)
AT_HaddadOfficeChargeDeduction, AT_SafeguardDutyIdOnly (not available)
AT_LabelCost, AT_Others (not available)
AT_SAPProductFlag (default A)
AT_PrincipalMerchandiseHierarchyL3/L4/L5 (not available)
AT_MaterialType (default ZINA)
AT_SAPArticleCategory (default Generic/Variant)
AT_InternalBarcode, AT_FGBarcode (not available)
AT_MainEANIndicator, AT_EANCategory (formula)
AT_UOM (default EA)
AT_HSCode (manual input)
AT_BCI, AT_PriceRange (manual/formula)
AT_Silhouette, AT_Width, AT_Fastening (not available)
AT_Content, AT_Fabric, AT_PatternPrint, AT_Fit (not available)
AT_Style, AT_Material, AT_PackDetails, AT_StyleType (manual input)
AT_CountrySize (default EU)
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
    """Create a Value element under parent. Skips if both value and id_val are empty."""
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
    """'FW26' → ('FW', '2026'). Handles 2-digit or 4-digit year tail."""
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

INPUT_DIR     = BASE_DIR / "input"
PRICELIST_DIR = INPUT_DIR / "pricelist"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in (INPUT_DIR, PRICELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"pricelist_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("lotto.pricelist_main")


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS / FALLBACK MAPS
# ══════════════════════════════════════════════════════════════════════════════
FALLBACK_GENDER: dict[str, tuple[str, str]] = {
    # (SAP Gender code, SAP Age code)
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

# ── Country code → MDD COUNTRY name ──────────────────────────────
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
}

_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING — Dynamic mapping from Brand Mapping Template
# ══════════════════════════════════════════════════════════════════════════════
# Based on "Lotto Inline (New Format)" sheet in Brand Mapping Template
# Field Name in Brand File → Header aliases (lowercase, for flexible matching)

COLUMN_ALIASES: dict[str, list[str]] = {
    "style":             ["style", "style code", "style no", "art no", "article no"],
    "style_description": ["style description", "style desc", "description", "article name"],
    "color_code":        ["color code", "colour code", "col code", "color"],
    "color_description": ["color description", "colour description", "colour name", "color name"],
    "gender":            ["gender", "sex", "genre"],
    "size_range":        ["size range", "size", "sizes", "taille"],
    "made_in":           ["made in", "country of origin", "origin", "country", "coo"],
    "fob_price":         ["licensee price direct bmoq", "fob", "fob price", "licensee price",
                          "direct bmoq", "cost price", "unit price"],
    "currency":          ["currency", "cur", "fob currency"],
    "product_type":      ["product type", "category", "division", "tipo prodotto"],
    "function":          ["function", "discipline", "sport", "activity", "sub category"],
    "pic":               ["pic", "picture", "image", "thumbnail", "photo"],
    "composition":       ["product composition 1 (highest %)", "product composition 1", "composition",
                          "material composition", "upper composition", "material upper"],
    "ean":               ["ean 13", "ean", "barcode", "ean13"],
}


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


def _find_mdd_file() -> Optional[Path]:
    """Locate MDD xlsx in input/mdd/, lotto/ or base dir, return first match."""
    # 1. Search in MDD_DIR
    if MDD_DIR.exists():
        xs = sorted(MDD_DIR.glob("*.xlsx")) or sorted(MDD_DIR.glob("*.xlsm"))
        if xs:
            return xs[0]
            
    # 2. Search in lotto/ directory (current folder)
    lotto_dir = Path(__file__).parent
    xs = sorted(lotto_dir.glob("*.xlsx")) or sorted(lotto_dir.glob("*.xlsm"))
    xs = [x for x in xs if not x.name.startswith("~$")]
    if xs:
        return xs[0]
        
    # 3. Search in BASE_DIR
    if BASE_DIR.exists():
        xs = sorted(BASE_DIR.glob("*.xlsx")) or sorted(BASE_DIR.glob("*.xlsm"))
        xs = [x for x in xs if not x.name.startswith("~$")]
        if xs:
            return xs[0]
            
    return None


def _open_wb(path: Path):
    """Open openpyxl workbook safely."""
    return openpyxl.load_workbook(path, data_only=True)


def _s(v) -> str:
    """Safe string conversion."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def load_material_lov(mdd_path: Optional[Path] = None) -> dict[str, str]:
    """Load Material - Upper LOV from MDD.
    
    Returns {display_name -> lov_code_id} mapping.
    Falls back to empty dict if file not found or sheet not found.
    """
    if mdd_path is None:
        mdd_path = _find_mdd_file()
    
    if mdd_path is None or not mdd_path.exists():
        log.warning("[MDD] No MDD file found — using composition strings as-is")
        return {}
    
    log.info("[MDD] Loading Material LOVs from: %s", mdd_path.name)
    try:
        wb = _open_wb(mdd_path)
    except Exception as e:
        log.warning("[MDD] Failed to open MDD: %s", e)
        return {}
    
    material_lov: dict[str, str] = {}
    
    # Look for "Material - Upper LOV" or similar sheet
    lov_sheet_names = [sn for sn in wb.sheetnames if "Material" in sn and "Upper" in sn and "LOV" in sn]
    
    if not lov_sheet_names:
        # Fallback: look for any generic "Material Upper" related LOV sheet
        lov_sheet_names = [sn for sn in wb.sheetnames if "Material" in sn and "LOV" in sn]
    
    if not lov_sheet_names:
        log.warning("[MDD] Material - Upper LOV sheet not found")
        return {}
    
    sheet_name = lov_sheet_names[0]
    log.info("[MDD] Using LOV sheet: '%s'", sheet_name)
    
    try:
        ws = wb[sheet_name]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 2:
                continue
            # Typically: r[0] = code/ID, r[1] = display name
            code = _s(r[0]).strip()
            name = _s(r[1]).strip()
            if code and name:
                # Store both directions for flexible lookup
                material_lov[name.lower()] = code
                material_lov[code.lower()] = code
        
        log.info("[MDD] Material - Upper LOV loaded: %d entries", len(material_lov))
        return material_lov
    except Exception as e:
        log.warning("[MDD] Failed to parse Material - Upper LOV: %s", e)
        return {}


def _get_lov_id(val: str, attr: str) -> str:
    v = (val or "").strip().upper()
    if not v:
        return ""
    if attr in ("AT_Gender", "AT_BYGender"):
        if v in ("MALE", "M", "MAN", "MEN"):
            return "M"
        if v in ("FEMALE", "F", "WOMAN", "WOMEN"):
            return "F"
        if v in ("UNISEX", "U", "CHILDREN", "CHILD", "CH", "KIDS", "KID", "JUNIOR", "JR"):
            return "U"
        return v
    elif attr == "AT_SAPAge":
        if v in ("ADULTS", "ADULT", "AD", "MALE", "FEMALE", "M", "F", "MAN", "WOMAN"):
            return "AD"
        if v in ("CHILDREN", "CHILD", "CH"):
            return "CH"
        if v in ("INFANT", "IN"):
            return "IN"
        if v in ("JUNIOR", "JR"):
            return "JR"
        if v in ("ALL AGES", "ALL_AGES", "AA"):
            return "AA"
        return v
    elif attr == "AT_BYAge":
        if v in ("ADULT", "ADULTS", "AD", "MALE", "FEMALE", "M", "F", "MAN", "WOMAN"):
            return "ADULT"
        if v in ("KIDS", "KID", "K"):
            return "KIDS"
        if v in ("CHILDREN", "CHILD", "CH"):
            return "CHILDREN"
        if v in ("JUNIOR", "JR"):
            return "JUNIOR"
        if v in ("ALL AGES", "AA"):
            return "ALL AGES"
        return v
    return v


def load_lotto_gender_mappings(mdd_path: Optional[Path] = None) -> tuple[dict[str, tuple], dict]:
    """Loads Principle and License gender/age mappings from 'Lotto MD Mappings' sheet."""
    if mdd_path is None:
        mdd_path = _find_mdd_file()
    
    principle_map = {}
    
    if mdd_path is None or not mdd_path.exists():
        log.warning("[MDD] No MDD file found for gender mapping.")
        return principle_map, {}
    
    try:
        wb = _open_wb(mdd_path)
        if "Lotto MD Mappings" in wb.sheetnames:
            ws = wb["Lotto MD Mappings"]
            rows = list(ws.iter_rows(values_only=True))
            # Start from row 9 (index 8)
            if len(rows) > 8:
                for r in rows[8:]:
                    if not r or all(c is None for c in r):
                        continue
                    gender_val = _s(r[0]).strip().upper()
                    if not gender_val:
                        continue
                    
                    sap_gender_raw = _s(r[1]).strip()
                    sap_age_raw    = _s(r[2]).strip()
                    by_gender_raw  = sap_gender_raw
                    by_age_raw     = sap_age_raw
                    
                    principle_map[gender_val] = (
                        _get_lov_id(sap_gender_raw, "AT_Gender"),
                        sap_gender_raw,
                        _get_lov_id(sap_age_raw, "AT_SAPAge"),
                        sap_age_raw,
                        _get_lov_id(by_gender_raw, "AT_BYGender"),
                        by_gender_raw,
                        _get_lov_id(by_age_raw, "AT_BYAge"),
                        by_age_raw
                    )
            log.info("[MDD] Loaded gender mappings from 'Lotto MD Mappings' sheet.")
        wb.close()
    except Exception as e:
        log.warning("[MDD] Failed to parse gender mappings from MDD: %s", e)
        
    return principle_map, {}


def resolve_lotto_gender_age(
    gender_raw: str,
    size_range: str = "",
    principle_map: Optional[dict] = None,
    license_map: Optional[dict] = None,
) -> tuple[str, str, str, str, str, str, str, str]:
    """Resolves Gender and Age using Lotto MD Mappings."""
    g_key = (gender_raw or "").strip().upper()
    
    eff_p_map = principle_map or {}
    
    if g_key in eff_p_map:
        val = eff_p_map[g_key]
        return val
        
    # Check fuzzy/substring matches if exact match fails
    for k, val in eff_p_map.items():
        if g_key == k or g_key in k or k in g_key:
            return val
            
    # Fallback to local resolver using FALLBACK_GENDER
    sap_code, sap_label, age_code, age_label = _resolve_gender(gender_raw)
    return (
        _get_lov_id(sap_code, "AT_Gender"),
        sap_label,
        _get_lov_id(age_code, "AT_SAPAge"),
        age_label,
        _get_lov_id(sap_code, "AT_BYGender"),
        sap_label,
        _get_lov_id(age_code, "AT_BYAge"),
        age_label
    )


def _get_all_material_lov_matches(material_str: str, material_lov: Optional[dict] = None) -> set[str]:
    """Get ALL possible LOV IDs that match the material string.
    
    Returns set of all matching LOV IDs.
    """
    if not material_str:
        return set()
    
    matched_codes = set()
    lower = material_str.lower().strip()
    
    # 1. Exact lookup in MDD material LOV if available.
    # No fuzzy/normalized/alias inference is allowed.
    if material_lov:
        direct = material_lov.get(lower)
        if direct:
            matched_codes.add(direct)

    # 2. Exact lookup in hardcoded mappings (no alias expansion).
    val = material_str.upper().strip()
    if val in MATERIAL_UPPER_LOV.values():
        matched_codes.add(val)
    
    for label, code in MATERIAL_UPPER_LOV.items():
        if lower == label.lower():
            matched_codes.add(code)
    
    return matched_codes


def _has_multiple_materials(raw: str) -> bool:
    if not raw:
        return False
    val = str(raw).strip().upper()
    pct_count = val.count("%")
    if pct_count > 1:
        return True
    if pct_count == 1:
        return False
    if any(delim in val for delim in ["/", ",", "+", "&"]) or " AND " in val:
        return True
    return False


def _extract_highest_percentage_material(raw: str) -> str:
    """Extract material with HIGHEST PERCENTAGE from composition string.
    
    Supports formats:
      - "70% POLYURETHANE, 30% GEL"
      - "PVC:80.0% - Polyester Textile:20.0%"
      - "Leather - Bovine:90.0% - PU:10.0%"
    Falls back to entire string if no percentages found.
    Returns '' if input is empty.
    """
    if not raw:
        return ""
    
    # Extract matches of form "Material Name:Percentage%"
    matches = re.findall(r'([^:]+?):\s*(\d+(?:\.\d+)?)\s*%', raw)
    if matches:
        highest_material = None
        highest_percent = -1.0
        for mat_name, pct_str in matches:
            try:
                pct = float(pct_str)
                mat_clean = mat_name.strip()
                if mat_clean.startswith("-"):
                    mat_clean = mat_clean[1:].strip()
                # Clean up any other potential leading delimiters
                mat_clean = mat_clean.lstrip(",/& ").strip()
                if pct > highest_percent:
                    highest_percent = pct
                    highest_material = mat_clean
            except ValueError:
                pass
        if highest_material:
            return highest_material

    # Split on commas, slashes, or dashes as fallback
    parts = [p.strip() for p in re.split(r'[,/\-]', raw) if p.strip()]
    highest_material = None
    highest_percent = -1.0
    
    for part in parts:
        part_clean = part.strip()
        # Format 1: "70% POLYURETHANE"
        m1 = re.match(r'^(\d+(?:\.\d+)?)\s*%\s*(.+)$', part_clean)
        if m1:
            try:
                percent = float(m1.group(1))
                material = m1.group(2).strip()
                if percent > highest_percent:
                    highest_percent = percent
                    highest_material = material
            except ValueError:
                pass
            continue
            
        # Format 2: "PVC:80.0%"
        m2 = re.match(r'^(.+?)\s*:\s*(\d+(?:\.\d+)?)\s*%$', part_clean)
        if m2:
            try:
                percent = float(m2.group(2))
                material = m2.group(1).strip()
                if percent > highest_percent:
                    highest_percent = percent
                    highest_material = material
            except ValueError:
                pass
            continue

    if highest_material is None:
        highest_material = raw
    
    return highest_material


MATERIAL_UPPER_LOV: dict[str, str] = {
    "EVA": "EVA",
    "Leather": "LEA",
    "Mesh": "MES",
    "Polyester": "PES",
    "PU": "PU",
    "Suede": "SUE",
    "Sustainable": "SUS",
    "Synthetic": "SYN",
    "Textile": "TEX",
}

def _normalize_material_text(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", (value or "").lower())).strip()


MATERIAL_UPPER_EXACT_ALIASES: dict[str, str] = {
    "eva": "EVA",
    "leather": "Leather",
    "mesh": "Mesh",
    "polyester": "Polyester",
    "polyurethan": "PU",
    "polyurethane": "PU",
    "pu": "PU",
    "tpu": "PU",
    "suede": "Suede",
    "sustainable": "Sustainable",
    "synthetic": "Synthetic",
    "textile": "Textile",
}


def _map_material_upper(raw: str, material_lov: Optional[dict] = None) -> str:
    """Map composition text to LOV ID for AT_MaterialUpper.
    
    LOGIC:
    1. Extract the material with HIGHEST PERCENTAGE
    2. Find ALL LOV IDs that match this material string
    3. If exactly 1 LOV ID matches → return it
    4. If 0 or multiple LOV IDs match → return empty
    
    Example 1: "10% SUPIMA, 40% PIMA, 50% LEATHER"
      → Highest: LEATHER (50%)
      → Matches: {LEA} → 1 match → return "LEA"
    
    Example 2: "90% SUEDE LEATHER, 10% SUPIMA"
      → Highest: SUEDE LEATHER (90%)
      → Matches: {SUE, LEA} → 2 matches → return empty
    """
    if not raw:
        return ""
    
    # Step 1: Extract material with highest percentage
    highest_material = _extract_highest_percentage_material(raw)
    if not highest_material:
        return ""
    
    # Step 2: Get ALL LOV IDs that match this material
    matched_lov_ids = _get_all_material_lov_matches(highest_material, material_lov)
    
    # Step 3: Return LOV ID only if exactly 1 match found
    if len(matched_lov_ids) == 1:
        return matched_lov_ids.pop()
    
    # Otherwise return empty (0 matches or multiple matches)
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
class PriceListRow:
    """Represents one row from the Price List file."""
    row_num:           int
    style:             str
    style_description: str
    color_code:        str
    color_description: str
    gender:            str
    size_range:        str
    made_in:           str
    fob_price:         Optional[float]
    currency:          str
    product_type:      str
    function:          str
    pic:               str          # Image URL or reference
    composition:       str          # Material Upper / Composition
    ean:               str          # EAN-13 barcode (if available)


@dataclass
class ColourGroup:
    """All rows for one (style_code, colour_code) combination."""
    style:             str
    style_description: str
    color_code:        str
    color_description: str
    gender:            str
    size_range:        str
    made_in:           str
    fob_price:         Optional[float]
    currency:          str
    product_type:      str
    function:          str
    pic:               str
    composition:       str
    rows:              list[PriceListRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADER
# ══════════════════════════════════════════════════════════════════════════════
def _find_input_files() -> list[Path]:
    """Locate .xlsx / .xlsm files in PRICELIST_DIR, fallback INPUT_DIR root."""
    files: list[Path] = []
    if PRICELIST_DIR.exists():
        files.extend(PRICELIST_DIR.glob("*.xlsx"))
        files.extend(PRICELIST_DIR.glob("*.xlsm"))
    if not files and INPUT_DIR.exists():
        files.extend(p for p in INPUT_DIR.glob("*.xlsx") if p.parent == INPUT_DIR)
        files.extend(p for p in INPUT_DIR.glob("*.xlsm") if p.parent == INPUT_DIR)
    return sorted(files)


def _build_col_map(header_row: tuple) -> dict[str, int]:
    """Build {normalized_header_text → col_index} from a header row tuple."""
    col_map: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        if cell is not None:
            norm = _s(cell).lower().strip()
            if norm:
                col_map[norm] = i
    return col_map


def _find_col(col_map: dict[str, int], field: str) -> Optional[int]:
    """Return first matching column index for ``field`` using COLUMN_ALIASES."""
    for alias in COLUMN_ALIASES.get(field, []):
        if alias in col_map:
            return col_map[alias]
    return None


def _find_data_sheet_and_header(wb) -> tuple[list[tuple], str, int]:
    """
    Scan every sheet (first 30 rows each) looking for a row that contains
    BOTH a style-like column AND a color-code-like column from COLUMN_ALIASES.
    Returns (rows, sheet_name, header_row_idx).
    """
    style_aliases = set(COLUMN_ALIASES["style"])
    color_aliases = set(COLUMN_ALIASES["color_code"])

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        for hdr_idx, row in enumerate(rows[:30]):
            if row is None:
                continue
            norm_cells = {_s(c).lower().strip() for c in row if c}
            has_style = any(a in norm_cells for a in style_aliases)
            has_color = any(a in norm_cells for a in color_aliases)
            if has_style and has_color:
                log.info("[PriceList] Found header at row %d in sheet '%s'", hdr_idx + 1, sheet_name)
                return rows, sheet_name, hdr_idx

    # Fallback: first non-empty sheet, header assumed at row 1
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if rows and any(c for c in rows[0] if c):
            log.warning("[PriceList] Using fallback: sheet '%s', header row 1", sheet_name)
            return rows, sheet_name, 0
    return [], "", 0


def load_pricelist(path: Path) -> list[PriceListRow]:
    """Parse the Price List file into a list of PriceListRow objects.

    Supports .xlsx and .xlsm. Auto-detects the sheet and header row.
    """
    log.info("[PriceList] Loading: %s", path.name)

    wb = openpyxl.load_workbook(str(path), keep_vba=True, data_only=True)
    rows, sheet_name, hdr_idx = _find_data_sheet_and_header(wb)
    wb.close()

    if not rows:
        log.warning("[PriceList] No data found in %s", path.name)
        return []

    header_row = rows[hdr_idx]
    col_map = _build_col_map(header_row)

    # Find column indices for each field
    col_style       = _find_col(col_map, "style")
    col_style_desc  = _find_col(col_map, "style_description")
    col_color_code  = _find_col(col_map, "color_code")
    col_color_desc  = _find_col(col_map, "color_description")
    col_gender      = _find_col(col_map, "gender")
    col_size_range  = _find_col(col_map, "size_range")
    col_made_in     = _find_col(col_map, "made_in")
    col_fob         = _find_col(col_map, "fob_price")
    col_currency    = _find_col(col_map, "currency")
    col_prod_type   = _find_col(col_map, "product_type")
    col_function    = _find_col(col_map, "function")
    col_pic         = _find_col(col_map, "pic")
    col_composition = _find_col(col_map, "composition")
    col_ean         = _find_col(col_map, "ean")

    if col_style is None:
        log.error("[PriceList] Could not find 'Style' column in %s", path.name)
        return []

    log.info(
        "[PriceList] Column mapping: style=%s color=%s gender=%s fob=%s composition=%s",
        col_style, col_color_code, col_gender, col_fob, col_composition
    )

    result: list[PriceListRow] = []
    skipped = 0

    def _gcell(r: tuple, idx: Optional[int], default: str = "") -> str:
        if idx is None or idx >= len(r):
            return default
        return _s(r[idx])

    for i, r in enumerate(rows[hdr_idx + 1:], start=hdr_idx + 2):
        if r is None or all(c is None for c in r):
            continue

        style_raw = _gcell(r, col_style)
        if not style_raw or style_raw.lower() in ("none", "nan", "total", ""):
            skipped += 1
            continue

        row = PriceListRow(
            row_num           = i,
            style             = style_raw,
            style_description = _gcell(r, col_style_desc),
            color_code        = _gcell(r, col_color_code),
            color_description = _gcell(r, col_color_desc),
            gender            = _gcell(r, col_gender),
            size_range        = _gcell(r, col_size_range),
            made_in           = _gcell(r, col_made_in),
            fob_price         = _num(r[col_fob]) if col_fob is not None and col_fob < len(r) else None,
            currency          = _gcell(r, col_currency),
            product_type      = _gcell(r, col_prod_type),
            function          = _gcell(r, col_function),
            pic               = _gcell(r, col_pic),
            composition       = _gcell(r, col_composition),
            ean               = _gcell(r, col_ean),
        )
        result.append(row)

    log.info("[PriceList] Parsed %d rows | %d skipped", len(result), skipped)
    return result


def group_by_style_and_colour(rows: list[PriceListRow]) -> dict[str, dict[str, ColourGroup]]:
    """Group rows: { style → { colour_code → ColourGroup } }."""
    groups: dict[str, dict[str, ColourGroup]] = defaultdict(dict)

    for r in rows:
        sc = r.style
        cc = r.color_code or "000"
        if cc not in groups[sc]:
            groups[sc][cc] = ColourGroup(
                style             = sc,
                style_description = r.style_description,
                color_code        = cc,
                color_description = r.color_description,
                gender            = r.gender,
                size_range        = r.size_range,
                made_in           = r.made_in,
                fob_price         = r.fob_price,
                currency          = r.currency,
                product_type      = r.product_type,
                function          = r.function,
                pic               = r.pic,
                composition       = r.composition,
            )
        groups[sc][cc].rows.append(r)

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
    style_code:    str,
    cg:            ColourGroup,
    cfg:           dict,
    season_id_cl:  str,
    material_lov:  dict = None,
    principle_map: dict = None,
    license_map:   dict = None,
) -> ET.Element:
    """Build one PRD_GenericArticle element for a style+color combination.

    Only MAPPED attributes from Brand Mapping Template are included.
    One product is generated per (style_code, color_code) combination.
    """
    brand_code   = cfg["brand_code"]
    brand_name   = cfg["brand_name"]
    sea_prefix   = cfg["season_prefix"]
    sea_year     = cfg["season_year"]
    brand_clh    = cfg.get("brand_name_for_clh", brand_name)

    # Gender/Age derived mapping (unpacks 8-tuple)
    sap_gender_code, sap_gender_label, sap_age_code, sap_age_label, by_gender_code, by_gender_label, by_age_code, by_age_label = \
        resolve_lotto_gender_age(cg.gender, cg.size_range, principle_map, license_map)
    coo = _resolve_country(cg.made_in)

    # KEY_InboundArticle & Generic Code: BrandCode + StyleCode + ColorCode (style+color level)
    color3d = color_3d(cg.color_code)
    generic_code = f"{brand_code}{style_code}{color3d}"
    inbound_article_code = f"{brand_code}{style_code}{color3d}"

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_F-TempSubCat")
    _keyval(gen_el, "KEY_InboundArticle", inbound_article_code)
    ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name").text = cg.style_description or generic_code

    _clf(gen_el, f"CLH_{brand_clh}Articles",  "CPL_Merchandiser")
    _clf(gen_el, f"{season_id_cl}UA",         "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")

    # ── Org / brand / defaults ───────────────────────────────────────────
    _val(vals_el, "AT_Brand",       "", id_val=brand_code)
    _val(vals_el, "AT_SBU",         "", id_val=cfg.get("sbu", "SP"))
    _val(vals_el, "AT_CompanyCode", "", id_val=cfg.get("comp_code", "0888"))
    _val(vals_el, "AT_Season",      "", id_val=sea_prefix)
    if sea_year:
        _val(vals_el, "AT_SeasonYear", str(sea_year))
    country_code = cfg.get("country_code", "ID")
    _val(vals_el, "AT_Country",      "", id_val=country_code)
    
    currency_map = {
        "ID": "IDR", "PH": "PHP", "TH": "THB", 
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD"
    }
    retail_currency = currency_map.get(country_code, "")
    if retail_currency:
        _val(vals_el, "AT_RetailPriceCurrency", id_val=retail_currency)
        
    # _val(vals_el, "AT_SAPProductFlat", id_val="A")
    _val(vals_el, "AT_SAPProductFlag", id_val="A")

    _val(vals_el, "AT_PricingDistributionChannel", "", id_val="01")

    # AT_CountrySize — only include if Product Type is "Footwear" (case-insensitive)
    if (cg.product_type or "").strip().lower() == "footwear":
        _val(vals_el, "AT_CountrySize",  "", id_val="EU")
    _val(vals_el, "AT_UOM",          "", id_val="EA")
    _val(vals_el, "AT_MaterialType", "", id_val="ZINA")
    _val(vals_el, "AT_SAPArticleCategory", "", id_val="01")

    # AT_BrandGroup & Brand Group Mappings
    _val(vals_el, "AT_BrandGroup", "", id_val=brand_name.upper())
    if cfg.get("brand_type"):
        _val(vals_el, "AT_BrandType", cfg["brand_type"])

    # ═══════════════════════════════════════════════════════════════════════════
    # MAPPED ATTRIBUTES ONLY (from Brand Mapping Template)
    # ═══════════════════════════════════════════════════════════════════════════

    # AT_PrincipalStyleCode → Style
    _val(vals_el, "AT_PrincipalStyleCode", style_code)

    # AT_PrincipalStyleDescription → Style description
    if cg.style_description:
        _val(vals_el, "AT_PrincipalStyleDescription", cg.style_description)

    # AT_SAPStyleCode → Style (Do not send for Lotto Inline)
    # _val(vals_el, "AT_SAPStyleCode", style_code)

    # AT_InboundGenericCode & AT_Generic
    _val(vals_el, "AT_InboundGenericCode", inbound_article_code)
    
    # AT_BYArticleType
    at_code = cfg.get("article_type")
    if at_code:
        _val(vals_el, "AT_BYArticleType", at_code, id_val=at_code)

    # AT_PrincipalGenderDescription → Gender (raw value)
    if cg.gender:
        _val(vals_el, "AT_PrincipalGenderDescription", cg.gender)

    # AT_Gender → Gender (derived: SAP Gender code)
    if sap_gender_code:
        _val(vals_el, "AT_Gender", id_val=sap_gender_code)

    # AT_BYGender → Gender (derived: same as AT_Gender)
    if by_gender_code:
        _val(vals_el, "AT_BYGender", id_val=by_gender_code)

    # AT_SAPAge → Gender (derived: SAP Age code)
    if sap_age_code:
        _val(vals_el, "AT_SAPAge", id_val=sap_age_code)

    # AT_BYAge → Gender (derived: same as AT_SAPAge)
    if by_age_code:
        _val(vals_el, "AT_BYAge", id_val=by_age_code)

    # AT_CountryOrigin → MADE IN
    if coo:
        _val(vals_el, "AT_CountryOrigin", "", id_val=coo)

    # AT_PrincipalColorCode → Color Code
    if cg.color_code:
        _val(vals_el, "AT_PrincipalColorCode", cg.color_code)

    # AT_PrincipalColorName → Color description
    if cg.color_description:
        _val(vals_el, "AT_PrincipalColorName", cg.color_description)

    # AT_PrincipalSize → Size range
    if cg.size_range:
        _val(vals_el, "AT_PrincipalSize", cg.size_range)

    # AT_FOB → LICENSEE PRICE DIRECT BMOQ
    if cg.fob_price is not None:
        _val(vals_el, "AT_FOB", str(cg.fob_price))

    # AT_FOBCurrency → CURRENCY
    if cg.currency:
        _val(vals_el, "AT_FOBCurrency", cg.currency, id_val=cg.currency)

    # AT_PrincipalMerchandiseHierarchyL1 → Product Type
    if cg.product_type:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", cg.product_type)

    # AT_PrincipalMerchandiseHierarchyL2 → Function
    if cg.function:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", cg.function)

    # AT_MaterialUpper → PRODUCT COMPOSITION 1 (LOV ID matched directly)
    if cg.composition:
        _mat_upper_id = _map_material_upper(cg.composition, material_lov)
        if _mat_upper_id:
            _val(vals_el, "AT_MaterialUpper", id_val=_mat_upper_id)
        else:
            _val(vals_el, "AT_MaterialUpper")

    # Note: PRD_VariantArticle NOT generated - only Generic level as per requirement

    return gen_el


def build_xml(
    rows:          list[PriceListRow],
    output_path:   Path,
    cfg:           dict,
    material_lov:  dict = None,
    principle_map: dict = None,
    license_map:   dict = None,
) -> Path:
    """Build the STEP XML from PriceList rows."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]
    cfg["brand_name_for_clh"] = cfg.get("brand_name_for_clh") or brand_name

    sea_prefix, sea_year = derive_season(cfg["season"])
    cfg["season_prefix"] = sea_prefix
    cfg["season_year"]   = sea_year
    season_id_cl         = season_id_full(brand_code, cfg["season"])

    groups = group_by_style_and_colour(rows)

    # ══════════════════════════════════════════════════════════════════════════
    # TEST LIMITER — Set to None for production, or a number to limit products
    # ══════════════════════════════════════════════════════════════════════════
    # TEST_PRODUCT_LIMIT = 25  # Uncomment to limit products for testing
    # # TEST_PRODUCT_LIMIT = None  # Set to None to process all products
    # if TEST_PRODUCT_LIMIT is not None:
    #     # Flatten all (style, color) pairs and limit
    #     all_pairs = [(sc, cc, cg) for sc, cm in groups.items() for cc, cg in cm.items()]
    #     all_pairs = all_pairs[:TEST_PRODUCT_LIMIT]
    #     # Rebuild limited groups
    #     groups = {}
    #     for sc, cc, cg in all_pairs:
    #         if sc not in groups:
    #             groups[sc] = {}
    #         groups[sc][cc] = cg
    #     log.info("[XML] TEST MODE: Limited to first %d products", TEST_PRODUCT_LIMIT)

    total_products = sum(len(v) for v in groups.values())
    log.info(
        "[XML] Building — %d styles, %d products (style+color combos), %d total rows",
        len(groups),
        total_products,
        len(rows),
    )

    cls_el  = build_classifications(brand_name, brand_code, cfg["season"])
    cls_str = strip_xmlns(ET.tostring(cls_el, encoding="unicode"))

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written = 0

    with open(output_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        open_step_xml(f, export_time)
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for style_code, colour_map in groups.items():
            # Generate one product per (style_code, color_code) combination
            for color_code, cg in colour_map.items():
                gen_el = _build_generic_product(
                    style_code, cg, cfg, season_id_cl,
                    material_lov, principle_map, license_map
                )
                _indent(gen_el, level=2)
                xml_str = strip_xmlns(ET.tostring(gen_el, encoding="unicode"))
                f.write(f"    {xml_str}\n")
                written += 1
        f.write("  </Products>\n")
        close_step_xml(f)

    log.info(
        "[XML] Written → %s  (%.1f KB) — %d products (style+color)",
        output_path,
        output_path.stat().st_size / 1024,
        written,
    )
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


def _extract_season_from_filename(filename: str) -> str:
    """Extract season token from filename, e.g. 'FW26' from 'Format_FW26 Price List …'."""
    m = re.search(r"\b([A-Z]{2}\d{2,4})\b", filename, re.IGNORECASE)
    return m.group(1).upper() if m else "FW26"


class RNALoader:
    """
    Loads 'Source Mapping Related RNA' tab from the MDD / Attributes List workbook.
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

    def _load(self) -> None:
        log.info("[RNA] Loading from: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.warning("[RNA] Cannot open workbook: %s", e)
            return

        sheet_name = next(
            (s for s in wb.sheetnames
             if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning(
                "[RNA] No sheet matching 'RNA' or 'SOURCE MAPPING' in %s. Sheets present: %s",
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
        """Match on country + brand only — ignores comp_code."""
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


def _find_rna_source_file() -> Path | None:
    # First, try exact filename
    brand_mapping_file = ATTR_DIR / "NEW - Brand mapping files Template.xlsx"
    if brand_mapping_file.exists():
        try:
            wb = openpyxl.load_workbook(brand_mapping_file, read_only=True, data_only=True)
            has_rna = any(
                any(kw in s.upper() for kw in RNALoader.RNA_SHEET_KEYWORDS)
                for s in wb.sheetnames
            )
            wb.close()
            if has_rna:
                log.info("[RNA] Using NEW - Brand mapping files Template.xlsx")
                return brand_mapping_file
        except Exception as e:
            log.warning("[RNA] Cannot load brand mapping file: %s", e)
    
    # Fallback: search for files with brand mapping keywords
    keywords = ["brand mapping", "brand_mapping", "mapping template", "brand template"]
    candidates: list[Path] = []
    
    if ATTR_DIR.exists():
        for file in ATTR_DIR.glob("*.xlsx"):
            filename_lower = file.name.lower()
            if any(kw in filename_lower for kw in keywords):
                candidates.append(file)
    
    # Check MDD_DIR as fallback location
    if MDD_DIR.exists():
        for file in MDD_DIR.glob("*.xlsx"):
            filename_lower = file.name.lower()
            if any(kw in filename_lower for kw in keywords):
                candidates.append(file)
    
    # Verify candidates have RNA sheet
    for p in candidates:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            has_rna = any(
                any(kw in s.upper() for kw in RNALoader.RNA_SHEET_KEYWORDS)
                for s in wb.sheetnames
            )
            wb.close()
            if has_rna:
                log.info("[RNA] Found by keyword: %s", p.name)
                return p
        except Exception as e:
            log.warning("[RNA] Cannot inspect %s: %s", p.name, e)
    
    return None


def run(args, auditor=None) -> tuple[list[dict], Path | None]:
    """
    Lambda dispatcher entry point for Lotto Price List files.
    """
    log.info(
        "[Lotto-PriceList] ETL started — brand=%s  season=%s  seq=%s",
        args.brand, args.season, getattr(args, "seq", 1),
    )

    files = _find_input_files()
    if not files:
        raise FileNotFoundError("No Price List .xlsx/.xlsm files found")

    file_path = max(files, key=lambda p: p.stat().st_mtime)
    log.info("[Lotto-PriceList] Input file: %s", file_path.name)

    file_cfg = parse_config_from_filename(file_path.stem)
    log.info("[Lotto-PriceList] Filename parsed → comp=%s sbu=%s season=%s country=%s type=%s",
             file_cfg.get("comp_code"), file_cfg.get("sbu"), file_cfg.get("season"),
             file_cfg.get("country_code"), file_cfg.get("article_type"))

    # ── SBU auto-remap ───────────────────────────────────────────
    sbu = file_cfg.get("sbu") or getattr(args, "sbu", "SP") or "SP"
    if (sbu or "").upper() == "SB":
        sbu = "SP"

    # ── Country code ─────────────────────────────────────────────
    country_code = file_cfg.get("country_code") or getattr(args, "country_code", "ID") or "ID"
    brand_code   = getattr(args, "brand_code", "LOT") or "LOT"
    brand_name   = getattr(args, "brand",      "Lotto") or "Lotto"
    comp_code    = file_cfg.get("comp_code") or getattr(args, "comp_code",  "0888")  or "0888"

    # ── Season: args first, then extract from filename ───────────
    season = file_cfg.get("season") or (getattr(args, "season", None) or "").strip()
    if not season or season == "FW26":
        season = _extract_season_from_filename(file_path.stem)

    # ── Brand Type / Brand Category from RNA ─────────────────────
    brand_type = getattr(args, "brand_type",     "") or ""
    brand_cat  = getattr(args, "brand_category", "") or ""

    if not brand_type or not brand_cat:
        rna_path = _find_rna_source_file()
        if rna_path:
            try:
                rna          = RNALoader(rna_path)
                country_name = _COUNTRY_MAP.get(country_code.upper(), country_code.upper())

                result = rna.get(country_name, comp_code, sbu, brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get(country_name, comp_code, getattr(args, "sbu", "SP"), brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get(country_name, comp_code, sbu, brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get(country_name, comp_code, getattr(args, "sbu", "SP"), brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get(country_name, str(int(comp_code)) if comp_code.isdigit() else comp_code, sbu, brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get(country_name, str(int(comp_code)) if comp_code.isdigit() else comp_code, getattr(args, "sbu", "SP"), brand_code)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_code)

                brand_type = brand_type or result.get("brand_type", "")
                brand_cat  = brand_cat  or result.get("brand_category", "")
                log.info(
                    "[Lotto-PriceList] RNA → country=%s comp=%s sbu=%s brand=%s → type='%s' cat='%s'",
                    country_name, comp_code, sbu, brand_code,
                    brand_type, brand_cat,
                )
            except Exception as e:
                log.warning("[Lotto-PriceList] RNA lookup failed: %s", e)
        else:
            log.warning(
                "[Lotto-PriceList] No RNA source workbook in %s or %s — brand metadata empty.",
                MDD_DIR, ATTR_DIR,
            )

    cfg = {
        "brand_code":         brand_code,
        "brand_name":         brand_name,
        "brand_name_for_clh": brand_name,
        "season":             season,
        "comp_code":          comp_code,
        "sbu":                sbu,
        "country_code":       country_code,
        "article_type":       file_cfg.get("article_type") or getattr(args, "article_type_from_filename", None),
        "currency":           getattr(args, "currency", "USD") or "USD",
        "brand_type":         brand_type,
        "brand_category":     brand_cat,
    }
    log.info("[Lotto-PriceList] Config: %s", cfg)

    rows = load_pricelist(file_path)
    if not rows:
        log.warning("[Lotto-PriceList] No rows parsed — skipping %s", file_path.name)
        return [], None

    material_lov = load_material_lov()
    principle_map, license_map = load_lotto_gender_mappings()
    
    xml_path = XML_OUT_DIR / f"{file_path.stem}.xml"
    build_xml(rows, xml_path, cfg, material_lov, principle_map, license_map)

    if auditor and xml_path:
        auditor.set_xml_uploads([str(xml_path)])

    log.info("[Lotto-PriceList] ETL complete.")
    return [], xml_path


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Lotto Price List → STEP XML")
    ap.add_argument("--brand",        default="Lotto")
    ap.add_argument("--brand-code",   default="LOT",  dest="brand_code")
    ap.add_argument("--season",       default="",     help="e.g. FW26 (auto-detected from filename if blank)")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="SP")
    ap.add_argument("--country-code", default="ID",   dest="country_code")
    ap.add_argument("--article-type", default="Inline", dest="article_type_from_filename")
    run(ap.parse_args())
