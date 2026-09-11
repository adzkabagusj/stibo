"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Steve Madden Footwear PO v1.0   ║
║   Footwear PO Form → Stibo STEP XML                             ║
║   (Generic Articles + Colour Variants + Size Variants)          ║
╚══════════════════════════════════════════════════════════════════╝

Input file pattern:
    0888-FF-STEVE MADDEN SVM-FOOTWEAR PO FORM MAA Fashion Footwear Inline-Multi-SP2029-ID-1.xlsx

Sheet:
    "All" or first sheet — header row auto-detected

Column layout (from SVM(Inline) brand mapping):
    STYLE                    → AT_PrincipalStyleCode
    STYLE                    → AT_PrincipalStyleDescription
    COLOR CODE               → AT_PrincipalColorCode
    COLOR CODE DESCRIPTION   → AT_PrincipalColorName
    PRICE                    → AT_FOB
    PRODUCT CATEGORY         → AT_PrincipalMerchandiseHierarchyL1
    BARCODES                 → AT_PrincipalBarcode (EAN)
    (Size columns vary)      → AT_PrincipalSize (will be mapped via size chart)

Article structure:
  One Generic per unique STYLE + COLOR CODE
  One Colour Variant per COLOR CODE
  Multiple Size Variants per size column with barcode

KEY formulas:
    KEY_InboundArticle  =  BrandCode + PrincipalStyleCode + PrincipalColorCode
    AT_InboundGenericCode = KEY_InboundArticle
    KEY_InboundVariant  =  KEY_InboundArticle + ColorCode(3) + "000"
    KEY_SizeVariant     =  KEY_InboundVariant + SizeCode(3)

Default values (from brand mapping):
    AT_PrincipalGenderDescription → "Female"
    AT_PrincipalAgeDescription    → "Adults"
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

log = logging.getLogger(__name__)

# ======================================================================
# PATHS
# ======================================================================
BASE_DIR    = Path(os.environ.get("LAMBDA_TMP_DIR", Path(__file__).resolve().parent))

INPUT_DIR       = BASE_DIR / "input" / "footwear_po"
MDD_DIR         = BASE_DIR / "input" / "mdd"
ATTR_DIR        = BASE_DIR / "input" / "attributes"
XML_OUT_DIR     = BASE_DIR / "output" / "xml"
LOG_DIR         = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME  = "SteveMadden"
BRAND_CODE  = "SVM"          # 3-letter SAP brand code for Steve Madden

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')

# ======================================================================
# DEFAULT VALUES (from brand mapping)
# ======================================================================
DEFAULT_GENDER_DESC = "Female"
DEFAULT_AGE_DESC = "Adults"


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value — returns empty string for None/nan/empty."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT", "#VALUE!") else s


def _pad_principal_color_code(color_code: str) -> str:
    """
    Principal Color Code is sent as-is from the principal file, but a few rows
    carry fewer than 3 characters (e.g. "56", "76", "ZX"). Left-pad any code
    shorter than 3 characters with zeros so it is always 3 characters
    ("76" -> "076", "ZX" -> "0ZX"), regardless of whether it is numeric or
    alphanumeric. Codes already 3+ characters are left untouched.
    """
    cc = _s(color_code)
    return cc.zfill(3) if cc and len(cc) < 3 else cc


def _color_code_3d(color_code: str) -> str:
    """
    Convert color code to 3-character format.
    If numeric, pad/trim to 3 digits. Otherwise take first 3 chars uppercase.
    """
    cc = _s(color_code)
    if not cc:
        return "000"
    
    # If numeric, pad to 3 digits
    digits = re.sub(r"[^0-9]", "", cc)
    if digits:
        return digits[:3].zfill(3)
    
    # Otherwise, take first 3 uppercase alphanumeric
    alpha = re.sub(r"[^A-Za-z0-9]", "", cc).upper()
    return alpha[:3].ljust(3, "0") if alpha else "000"


def _resolve_currency_from_country(country_code: str) -> str:
    """
    Resolve Retail Price Currency LOV ID from country code.
    """
    country_to_currency = {
        "ID": "IDR",  # Indonesia → Indonesian Rupiah
        "PH": "PHP",  # Philippines → Philippine Peso
        "TH": "THB",  # Thailand → Thailand Baht
        "SG": "SGD",  # Singapore → Singapore Dollar
        "MY": "MYR",  # Malaysia → Malaysian Ringgit
        "VN": "VND",  # Vietnam → Vietnamese Dong
        "KH": "USD",  # Cambodia → United States Dollar
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


# ======================================================================
# MDD LOADER
# ======================================================================

class MDDLoader:
    """
    Loads LOV tables from the MDD Excel needed for Steve Madden.
    Loads: Color Code LOV, Country Origin LOV, Material LOV, Gender LOV, Brand LOV, etc.
    """

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        self._load_color_lov(wb)
        self._load_country_origin_lov(wb)
        self._load_material_lov(wb)
        self._load_gender_lov(wb)
        self._load_brand_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_brand_category_lov(wb)
        self._load_brand_group_lov(wb)
        self._load_country_lov(wb)
        self._load_heel_height_lov(wb)

        wb.close()
        log.info("[MDD] %d LOV types loaded", len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = str(val_id).strip() if val_id else str(val_name).strip()

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display_name = sn.replace(" LOV", "").replace("LOV ", "").strip()
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                code, name = row[0], row[1]
                if name:
                    self.lovs.setdefault(display_name, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )

    def _load_color_lov(self, wb):
        """Color Code LOV: display name → LOV ID."""
        sheet = next(
            (s for s in wb.sheetnames if "COLOR" in s.upper() and "CODE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Color Code LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["ColorCodeLOV"] = lov
        self.lovs["AT_Color"]     = lov
        log.info("[MDD] Color Code LOV: %d entries", len(lov))

    def _load_country_origin_lov(self, wb):
        """Country Origin LOV: display name → LOV ID."""
        sheet = next(
            (s for s in wb.sheetnames if "COUNTRY" in s.upper() and "ORIGIN" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Country Origin LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["CountryOriginLOV"]  = lov
        self.lovs["AT_CountryOrigin"]  = lov
        log.info("[MDD] Country Origin LOV: %d entries", len(lov))

    def _load_material_lov(self, wb):
        """Material LOV (non-upper): display name → LOV ID."""
        sheet = next(
            (s for s in wb.sheetnames
             if "MATERIAL" in s.upper() and "LOV" in s.upper() and "UPPER" not in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Material LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["MaterialLOV"] = lov
        self.lovs["AT_Material"] = lov
        log.info("[MDD] Material LOV: %d entries", len(lov))

    def _load_gender_lov(self, wb):
        """Gender LOV: display name → LOV ID."""
        sheet = next(
            (s for s in wb.sheetnames if "GENDER" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Gender LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["GenderLOV"] = lov
        self.lovs["AT_Gender"] = lov
        log.info("[MDD] Gender LOV: %d entries", len(lov))

    def _load_brand_lov(self, wb):
        """Brand LOV: col A = Value ID of LOV, col B = Values of LOV (display name)."""
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()
             and "TYPE" not in s.upper() and "GROUP" not in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Brand LOV sheet not found")
            return
        log.info("[MDD] Loading Brand LOV from sheet: '%s'", sheet)
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandLOV"] = lov
        self.lovs["AT_Brand"] = lov
        log.info("[MDD] Brand LOV: %d entries loaded", len(lov))

    def _load_brand_type_lov(self, wb):
        """Brand Type LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND TYPE LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Type LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandTypeLOV"] = lov
        self.lovs["AT_BrandType"] = lov
        log.info("[MDD] Brand Type LOV: %d entries", len(lov))

    def _load_brand_status_lov(self, wb):
        """Brand Status LOV."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Status LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandStatusLOV"] = lov
        self.lovs["AT_BrandStatus"] = lov
        log.info("[MDD] Brand Status LOV: %d entries", len(lov))

    def _load_brand_category_lov(self, wb):
        """Brand Category LOV."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND CATEGORY LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Category LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandCategoryLOV"] = lov
        self.lovs["AT_BrandCategory"] = lov
        log.info("[MDD] Brand Category LOV: %d entries", len(lov))

    def _load_brand_group_lov(self, wb):
        """Brand Group LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP"),
            next((s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP LOV"), None),
        )
        if not sheet_name:
            sheet_name = next(
                (s for s in wb.sheetnames if s.strip().upper() == "BRANDGROUP LOV"),
                None,
            )
        if not sheet_name:
            # Fallback: search for any sheet with BRAND + GROUP (but exclude BRAND TYPE/STATUS/CATEGORY)
            sheet_name = next(
                (s for s in wb.sheetnames 
                 if "BRAND" in s.upper() and "GROUP" in s.upper()
                 and "TYPE" not in s.upper() and "STATUS" not in s.upper() and "CATEGORY" not in s.upper()),
                None,
            )
        if not sheet_name:
            log.warning("[MDD] Brand Group LOV sheet not found")
            log.warning("[MDD] Available sheets: %s", wb.sheetnames)
            return
        log.info("[MDD] Loading Brand Group LOV from sheet: '%s'", sheet_name)
        lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandGroupLOV"] = lov
        self.lovs["AT_BrandGroup"] = lov
        log.info("[MDD] Brand Group LOV: %d entries loaded", len(lov))

    def _load_country_lov(self, wb):
        """Country LOV: 2-letter code → full country name."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            None,
        )
        if not sheet:
            log.warning("[MDD] Country LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            code    = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if code and display:
                lov[code.upper()] = display
        self.lovs["CountryLOV"] = lov
        log.info("[MDD] Country LOV: %d entries", len(lov))

    def _load_heel_height_lov(self, wb):
        """Heel Height LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "HEEL HEIGHT LOV"),
            next((s for s in wb.sheetnames if "HEEL" in s.upper() and "HEIGHT" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Heel Height LOV sheet not found")
            return
        log.info("[MDD] Loading Heel Height LOV from sheet: '%s'", sheet_name)
        lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["HeelHeightLOV"] = lov
        self.lovs["AT_HeelHeight"] = lov
        log.info("[MDD] Heel Height LOV: %d entries loaded", len(lov))

    def country_id_to_name(self, country_id: str) -> str:
        """Convert 2-letter country code to full name."""
        return self.lovs.get("CountryLOV", {}).get((country_id or "").strip().upper(), "")


# ======================================================================
# SVM MAPPING LOADER (for heel height, etc.)
# ======================================================================

class SVMMappingLoader:
    """
    Loads brand-specific mappings from 'SVM MD Mappings' sheet.
    For heel height: Column A has source values (e.g., '75MM', '7.5CM'),
                     Column C has target values (e.g., 'Flat', 'Low', 'Medium', 'High').
    """
    def __init__(self, path: Path):
        self.path = path
        self.heel_height_map: dict[str, str] = {}
        self._load()

    def _load(self):
        log.info("[SVMMapping] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        # Find SVM MD Mappings sheet
        sheet_name = next(
            (s for s in wb.sheetnames if "SVM" in s.upper() and "MAPPING" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[SVMMapping] 'SVM MD Mappings' sheet not found")
            wb.close()
            return
        
        log.info("[SVMMapping] Loading from sheet: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        
        # Find header row with 'Heel Height' in columns A and C (Steve Madden)
        hdr_idx = None
        for i, row in enumerate(rows[:20]):
            if row and len(row) >= 3:
                col_a = str(row[0] or "").strip().upper() if row[0] else ""
                col_c = str(row[2] or "").strip().upper() if len(row) > 2 and row[2] else ""
                if "HEEL" in col_a and "HEIGHT" in col_a:
                    hdr_idx = i
                    break
        
        if hdr_idx is None:
            log.warning("[SVMMapping] Heel Height header not found")
            wb.close()
            return
        
        # Load mappings: column A (source) -> column C (target)
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 3:
                continue
            source_val = str(row[0]).strip().upper() if row[0] else ""
            # Also check column B for alternate format
            source_val_b = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ""
            target_val = str(row[2]).strip() if row[2] else ""
            
            if source_val and target_val:
                self.heel_height_map[source_val] = target_val
            if source_val_b and target_val:
                self.heel_height_map[source_val_b] = target_val
        
        wb.close()
        log.info("[SVMMapping] Heel Height mappings loaded: %d entries", len(self.heel_height_map))

    def get_heel_height_category(self, heel_value: str) -> str:
        """Map heel height value to category (Flat, Low, Medium, High)."""
        return self.heel_height_map.get((heel_value or "").strip().upper(), "")


# ======================================================================
# RNA LOADER (Brand Attributes from "Source Mapping related RNA" sheet)
# ======================================================================

class RNALoader:
    """
    Loads brand attributes from the RNA mapping sheet in the attributes file.
    Lookup key: (country_name, comp_code, sbu, brand_code) → brand_type, brand_status, brand_category, brand_group
    """
    RNA_SHEET_KEYWORDS = ["SOURCE MAPPING", "RNA", "BRAND MAPPING"]

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

    def __init__(self, path: Path, mdd: MDDLoader | None = None):
        self.path = path
        self.mdd = mdd
        self.lookup: dict[tuple[str, str, str, str], dict[str, str]] = {}
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
        c_bcat    = _find("BRANDCATEGORY", "BRAND CATEGORY", "AT_BRANDCATEGORY")
        c_bgroup  = _find("BRANDGROUP", "BRAND GROUP", "AT_BRANDGROUP")
        c_bstatus = _find("BRAND STATUS", "BRANDSTATUS", "AT_BRANDSTATUS")

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
                "brand_group":    _cell(row, c_bgroup),
                "brand_status":   _cell(row, c_bstatus),
            }
            count += 1

        wb.close()
        log.info("[RNA] %d entries loaded from '%s'", count, sheet_name)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        """Return RNA attributes for given keys, or empty strings."""
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu        or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""})

    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        """Match on country + sbu + brand — ignores comp_code."""
        c = self._norm_country(country_name)
        b = (brand_code or "").strip().upper()
        s = (sbu or "").strip().upper()
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}

    def get_by_brand_only(self, brand_code: str) -> dict:
        """Pass 4: match on brand code only — ignores country, comp_code, sbu."""
        b = (brand_code or "").strip().upper()
        for (_, _, _, kb), val in self.lookup.items():
            if kb == b:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


# ======================================================================
# INPUT FILE LOADER
# ======================================================================

class SteveMaddenFootwearPOLoader:
    """
    Loads Steve Madden Footwear PO Form data.
    Auto-detects header row and sheet.
    """

    # Signal words that identify the header row
    HEADER_SIGNALS = {"style", "color code", "price", "barcode"}

    def __init__(self, path: Path):
        self.path           = path
        self.df             = pd.DataFrame()
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[SteveMadden-FootwearPO] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[SteveMadden-FootwearPO] Available sheets: %s", wb.sheetnames)

        # Try "All" first, then first sheet
        target = "All" if "All" in wb.sheetnames else wb.sheetnames[0]
        log.info("[SteveMadden-FootwearPO] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.warning("[SteveMadden-FootwearPO] Sheet '%s' is empty", target)
            return

        # ── Locate header row ──────────────────────────────────────
        hdr_idx = 0
        for i in range(min(20, len(rows))):
            if rows[i]:
                rv = {str(v).strip().lower().replace(" ", "").replace("_", "")
                      for v in rows[i] if v is not None}
                if rv & self.HEADER_SIGNALS:
                    hdr_idx = i
                    break

        log.info("[SteveMadden-FootwearPO] Header at row %d (0-indexed)", hdr_idx)
        raw_header = rows[hdr_idx]
        header = [
            str(h).strip() if h is not None else f"col_{i}"
            for i, h in enumerate(raw_header)
        ]

        # Deduplicate column names
        seen: dict[str, int] = {}
        for ci, col in enumerate(header):
            if col in seen:
                seen[col] += 1
                header[ci] = f"{col}_{seen[col]}"
            else:
                seen[col] = 0

        self.headers = header
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Filter: keep rows with a valid STYLE
        style_col = self._find_col(header, ["STYLE", "Style", "style"])
        if style_col:
            df = df[
                df[style_col].notna()
                & (~df[style_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        self.df = df.reset_index(drop=True)
        log.info(
            "[SteveMadden-FootwearPO] Loaded %d data rows | headers: %s",
            len(self.df), header[:25],
        )

    @staticmethod
    def _find_col(headers: list[str], candidates: list[str]) -> str | None:
        """Case-insensitive column name search."""
        for c in candidates:
            if c in headers:
                return c
            for h in headers:
                if h.lower() == c.lower():
                    return h
        return None

    def find_col(self, *candidates: str) -> str | None:
        return self._find_col(self.headers, list(candidates))


# ======================================================================
# GROUPER — builds Generic + Variant + Size dict from flat rows
# ======================================================================

def group_rows(
    raw_rows: list[dict],
    loader: SteveMaddenFootwearPOLoader,
    brand_code: str,
    mdd: MDDLoader | None,
    svm_mapping: SVMMappingLoader | None = None,
) -> dict[str, dict]:
    """
    Group flat rows into Generic articles keyed by KEY_InboundArticle.

    KEY_InboundArticle  =  brand_code + PrincipalStyleCode
    KEY_InboundVariant  =  KEY_InboundArticle + color_code(3) + "000"
    KEY_SizeVariant     =  KEY_InboundVariant + size_code(3)
    """
    col_style       = loader.find_col("STYLE", "Style")
    col_style_name  = loader.find_col("STYLE NAME", "Style Name")
    col_color_code  = loader.find_col("COLOR CODE", "Color Code", "COLOR_CODE")
    col_color_desc  = loader.find_col("COLOR CODE DESCRIPTION", "Color Description")
    col_price       = loader.find_col("PRICE", "Price", "FOB")
    col_category    = loader.find_col("PRODUCT CATEGORY", "Category")
    col_barcode     = loader.find_col("BARCODES", "Barcode", "EAN")
    col_coo         = loader.find_col("COO", "Country of Origin", "CountryOfOrigin")
    col_heel_height = loader.find_col("HEEL HEIGHT", "Heel Height", "HeelHeight")

    result: dict[str, dict] = {}

    for row in raw_rows:
        style_code = _s(row.get(col_style) if col_style else "")
        if not style_code:
            continue

        color_code = _s(row.get(col_color_code) if col_color_code else "")
        if color_code.endswith(".0"):
            color_code = color_code[:-2]
        color_code = _pad_principal_color_code(color_code)
        if not color_code:
            continue

        color_desc = _s(row.get(col_color_desc) if col_color_desc else "")

        # Country of Origin LOV lookup
        country_origin_raw = _s(row.get(col_coo) if col_coo else "")
        country_origin_lov = ""
        if country_origin_raw and mdd:
            co_lov = mdd.lovs.get("CountryOriginLOV", {})
            country_origin_lov = co_lov.get(country_origin_raw.upper(), country_origin_raw)
        elif country_origin_raw:
            country_origin_lov = country_origin_raw

        # Heel Height LOV lookup
        heel_height_raw = _s(row.get(col_heel_height) if col_heel_height else "")
        heel_height_lov = ""
        if heel_height_raw and svm_mapping and mdd:
            # Step 1: Map raw value to category (Flat, Low, Medium, High)
            heel_category = svm_mapping.get_heel_height_category(heel_height_raw)
            # Step 2: Look up LOV ID from category
            if heel_category:
                hh_lov = mdd.lovs.get("HeelHeightLOV", {})
                heel_height_lov = hh_lov.get(heel_category.upper(), "")

        # Build keys
        generic_key = f"{brand_code}{style_code}{color_code}"
        color_3d = _color_code_3d(color_code)
        variant_key = f"{generic_key}{color_3d}000"

        # Initialize generic if new
        if generic_key not in result:
            result[generic_key] = {
                "generic_key": generic_key,
                "style_code": style_code,
                "AT_InboundGenericCode": generic_key,
                "AT_PrincipalStyleCode": style_code,
                "AT_PrincipalColorCode": color_code,
                "AT_PrincipalColorName": color_desc,
                "AT_PrincipalStyleDescription": style_code,
                "AT_PrincipalGenderDescription": DEFAULT_GENDER_DESC,
                "AT_PrincipalAgeDescription": DEFAULT_AGE_DESC,
                "AT_CountryOrigin": country_origin_lov,
                "AT_Collection1": style_code,
                "AT_BYArticleType": "Inline",
                "AT_HeelHeight": heel_height_lov,
                "AT_FOB": _s(row.get(col_price) if col_price else ""),
                "AT_PrincipalMerchandiseHierarchyL1": _s(row.get(col_category) if col_category else ""),
                "variants": {},
            }

        # Initialize variant if new
        if variant_key not in result[generic_key]["variants"]:
            result[generic_key]["variants"][variant_key] = {
                "variant_key": variant_key,
                "color_code": color_code,
                "color_3d": color_3d,
                "AT_PrincipalColorCode": color_code,
                "AT_PrincipalColorName": _s(row.get(col_color_desc) if col_color_desc else ""),
                "sizes": {},
            }

        # For now, store barcode at variant level (will handle size variants later)
        barcode = _s(row.get(col_barcode) if col_barcode else "")
        if barcode:
            result[generic_key]["variants"][variant_key].setdefault("barcodes", []).append(barcode)

    log.info(
        "[Grouper] %d rows → %d generics, %d total variants",
        len(raw_rows),
        len(result),
        sum(len(g["variants"]) for g in result.values()),
    )
    return result


# ======================================================================
# XML HELPERS
# ======================================================================

def _val(parent: ET.Element, attr_id: str, text: str = "", lov_id: str = "") -> None:
    """Append a <Value AttributeID="..."> element."""
    if not text and not lov_id:
        return
    v = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    v.set("AttributeID", attr_id)
    if lov_id:
        v.set("ID", lov_id)
    if text:
        v.text = text


def _val_text(parent: ET.Element, attr_id: str, text: str) -> None:
    """Add a text-type Value element (only if non-empty)."""
    t = _s(text)
    if t:
        _val(parent, attr_id, text=t)


def _val_lov(parent: ET.Element, attr_id: str, lov_id: str) -> None:
    """Add a LOV-type Value element (only if non-empty)."""
    lid = _s(lov_id)
    if lid:
        _val(parent, attr_id, lov_id=lid)


def _add_barcode_dc(parent: ET.Element, ean: str) -> None:
    """Add DC_Barcode DataContainer with the EAN barcode."""
    dc_root = ET.SubElement(parent, f"{{{STIBO_NS}}}DataContainers")
    mdc     = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")
    dc      = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc.set("Analyzer", "true")
    dcv     = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

    v1 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v1.set("AttributeID", "AT_Barcode")
    v1.text = ean

    v2 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v2.set("AttributeID", "AT_BarcodeType")
    v2.set("ID", "P")

    v3 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v3.set("AttributeID", "AT_MainEANIndicator")
    v3.set("ID", "Y")


# ======================================================================
# CLASSIFICATIONS BUILDER
# ======================================================================

def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    """
    Build the <Classifications> block (mirrors Clarks pattern):
      CLH_{brand_code}_{season}          CLS_Season         parent=CLH_{brand}Batches
        CLH_{brand_code}_{season}CA      CLS_ConfirmedArticles
        CLH_{brand_code}_{season}UA      CLS_UnconfirmedArticles
    """
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_tail   = season_code[2:]
    sea_year   = f"20{sea_tail}" if len(sea_tail) == 2 else sea_tail

    full_season    = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season}"
    batches_parent = f"CLH_{brand.replace(' ', '')}Batches"

    _SEASON_LABELS = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
    }
    sea_name       = _SEASON_LABELS.get(sea_prefix, sea_prefix)
    season_display = f"{brand.replace(' ', '')} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_parent)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


# ======================================================================
# XML BUILDER
# ======================================================================

def build_product_xml(
    generic: dict,
    brand: str,
    brand_code: str,
    comp_code: str,
    sbu: str,
    season: str,
    season_id: str,
    country_code: str,
    brand_type: str = "",
    brand_status: str = "",
    brand_category: str = "",
    brand_group: str = "",
    mdd: "MDDLoader | None" = None,
) -> str:
    """Build XML for one Generic article (no variants — footwear PO, generic only).

    brand_code  — 3-letter code from filename (e.g. 'SVM'), used for both
                  KEY/classification IDs and the AT_Brand LOV value.
    """
    if not generic.get("style_code"):
        return ""

    # ── Generic Article ──────────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_F-TempSubCat")  # Footwear hierarchy

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_key"]

    # Name
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        generic["AT_PrincipalStyleDescription"] or generic["generic_key"]
    )

    # ── Classification References ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Generic Values ───────────────────────────────────────────
    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # ── Principal attributes (direct from source file) ───────────
    _val_text(gv, "AT_InboundGenericCode",              generic["AT_InboundGenericCode"])
    _val_text(gv, "AT_PrincipalStyleCode",              generic["AT_PrincipalStyleCode"])
    _val_text(gv, "AT_PrincipalStyleDescription",        generic["AT_PrincipalStyleDescription"])
    _val_text(gv, "AT_PrincipalColorCode",              generic.get("AT_PrincipalColorCode", ""))
    _val_text(gv, "AT_PrincipalColorName",              generic.get("AT_PrincipalColorName", ""))
    _val_text(gv, "AT_PrincipalGenderDescription",       generic["AT_PrincipalGenderDescription"])
    _val_text(gv, "AT_PrincipalAgeDescription",          generic["AT_PrincipalAgeDescription"])
    _val_text(gv, "AT_PrincipalMerchandiseHierarchyL1",  generic["AT_PrincipalMerchandiseHierarchyL1"])
    _val_text(gv, "AT_FOB",                              generic["AT_FOB"])
    if generic.get("AT_PrincipalBarcode"):
        _val_text(gv, "AT_PrincipalBarcode",             generic["AT_PrincipalBarcode"])

    # ── Country of Origin (LOV) ────────────────────────────
    if generic.get("AT_CountryOrigin"):
        _val_lov(gv, "AT_CountryOrigin", generic["AT_CountryOrigin"])

    # ── Collection ─────────────────────────────────────────────
    if generic.get("AT_Collection1"):
        _val_text(gv, "AT_Collection1", generic["AT_Collection1"])

    # ── Article Type ───────────────────────────────────────────
    _val_lov(gv, "AT_BYArticleType", generic["AT_BYArticleType"])

    # ── Heel Height (LOV) ──────────────────────────────────────
    if generic.get("AT_HeelHeight"):
        _val_lov(gv, "AT_HeelHeight", generic["AT_HeelHeight"])

    # ── Brand (LOV) ──────────────────────────────────────────────
    # Brand LOV ID comes from the filename (parts[2] trailing 3-letter code).
    _val_lov(gv, "AT_Brand", brand_code)

    # ── Brand Type ───────────────────────────────────────────────
    # From RNA BRANDTYPE_DETAIL → look up LOV ID in Brand Type LOV
    log.info(
        "[XML] Brand attrs — brand_type='%s'  brand_status='%s'  brand_category='%s'  brand_group='%s'",
        brand_type, brand_status, brand_category, brand_group,
    )
    if brand_type:
        bt_lov_id = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type.upper(), "")
        if bt_lov_id:
            _val_lov(gv, "AT_BrandType", bt_lov_id)
        else:
            _val_text(gv, "AT_BrandType", brand_type)

    # ── Brand Status ─────────────────────────────────────────────
    # From RNA BRAND STATUS → look up LOV ID in Brand Status LOV
    if brand_status:
        bs_lov_id = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status.upper(), "")
        if bs_lov_id:
            _val_lov(gv, "AT_BrandStatus", bs_lov_id)
        else:
            _val_text(gv, "AT_BrandStatus", brand_status)

    # ── Brand Category ───────────────────────────────────────────
    # From RNA BRANDCATEGORY → look up LOV ID in Brand Category LOV
    if brand_category:
        bc_lov_id = (mdd.lovs.get("AT_BrandCategory", {}) if mdd else {}).get(brand_category.upper(), "")
        if bc_lov_id:
            _val_lov(gv, "AT_BrandCategory", bc_lov_id)
        else:
            _val_text(gv, "AT_BrandCategory", brand_category)

    # ── Brand Group ──────────────────────────────────────────────
    # From RNA BRANDGROUP → look up LOV ID in Brand Group LOV
    if brand_group:
        bg_lov_id = (mdd.lovs.get("AT_BrandGroup", {}) if mdd else {}).get(brand_group.upper(), "")
        if bg_lov_id:
            _val_lov(gv, "AT_BrandGroup", bg_lov_id)
        else:
            _val_text(gv, "AT_BrandGroup", brand_group)

    # ── Company + SBU (MultiValue LOV pattern) ───────────────────
    mv_comp = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_comp.set("AttributeID", "AT_CompanyCode")
    v_comp = ET.SubElement(mv_comp, f"{{{STIBO_NS}}}Value")
    v_comp.set("ID", comp_code)

    mv_sbu = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_sbu.set("AttributeID", "AT_SBU")
    v_sbu = ET.SubElement(mv_sbu, f"{{{STIBO_NS}}}Value")
    v_sbu.set("ID", sbu)

    # ── Season + Season Year ─────────────────────────────────────
    sea_code = season[:2].upper() if len(season) >= 2 else season
    _val_lov(gv, "AT_Season", sea_code)

    year_match = re.search(r"20\d{2}", season)
    if not year_match:
        short_match = re.search(r"\d{2}$", season)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    if season_year:
        _val_text(gv, "AT_SeasonYear", season_year)

    # ── Country (destination country — LOV) ──────────────────────
    if country_code:
        _val_lov(gv, "AT_Country", country_code.upper())

    # ── Retail Price Currency (derived from country code) ────────
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _val_lov(gv, "AT_RetailPriceCurrency", currency_lov_id)

    # ── FOB Currency (default USD) ───────────────────────────────
    _val_lov(gv, "AT_FOBCurrency", "USD")

    # ── System defaults (Clarks pattern) ─────────────────────────
    _val_lov(gv, "AT_SAPProductFlag",     "A")     # Direct Local
    _val_lov(gv, "AT_MaterialType",       "ZINA")  # Trading Goods
    _val_lov(gv, "AT_SAPArticleCategory", "1")     # Generic
    _val_lov(gv, "AT_UOM",                "EA")    # Each
    _val_lov(gv, "AT_BYAge",              "ADULT")
    _val_lov(gv, "AT_BYGender",           "F")
    _val_lov(gv, "AT_SAPAge",             "AD")
    _val_lov(gv, "AT_Gender",             "F")
    _val_lov(gv, "AT_PricingDistributionChannel", "01")
    _val_lov(gv, "AT_BCI",                "COMMERCIAL")
    _val_lov(gv, "AT_CountrySize",        "US")
    _val_lov(gv, "AT_ArticleStatus",      "A")     # Active
    _val_lov(gv, "AT_EComAgesCategory",   "18+Y")  # E-commerce Age Category

    # Generic only — no variant sub-products
    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# MAIN RUN FUNCTION
# ======================================================================

def run(args, auditor=None):
    """
    Main entry point for Steve Madden Footwear PO ETL.
    """
    log.info("[SteveMadden-FootwearPO] Starting ETL ...")

    # ── Locate input files ───────────────────────────────────────
    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        log.error("[SteveMadden-FootwearPO] No input file found in %s", INPUT_DIR)
        return

    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if not mdd_file:
        log.error("[SteveMadden-FootwearPO] No MDD file found in %s", MDD_DIR)
        return

    attr_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    if not attr_file:
        log.error("[SteveMadden-FootwearPO] No attributes file found in %s", ATTR_DIR)
        return

    # ── Load MDD and RNA ─────────────────────────────────────────
    mdd = MDDLoader(mdd_file)
    rna = RNALoader(attr_file, mdd)
    svm_mapping = SVMMappingLoader(attr_file)  # SVM mappings are in brand mapping file

    # ── Base metadata from args (used as fallback) ───────────────
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SS27")
    country_code = getattr(args, "country_code", "")

    log.info(
        "[SteveMadden-FootwearPO] Starting: brand=%s  code=%s  season=%s",
        brand, brand_code, season,
    )

    # ── Extract metadata from filename ────────────────────────────
    # Pattern: 0888-FF-STEVE MADDEN SVM-FOOTWEAR PO FORM MAA Fashion Footwear Inline-Multi-SP2029-ID-1
    #   parts[0]  = comp code           (e.g. '0888')
    #   parts[1]  = SBU                 (e.g. 'FF')
    #   parts[2]  = brand name + code   (e.g. 'STEVE MADDEN SVM')
    #   season    = regex-located token (e.g. 'SP2029')
    #   parts[-2] = country id          (e.g. 'ID')
    #
    # NOTE: brand_code and brand are already correctly parsed by the lambda
    # from parts[2] using the same regex as Crocs — args values are authoritative.
    filename_stem       = input_file.stem
    filename_comp_code  = comp_code    # fallback to args
    filename_sbu        = sbu          # fallback to args
    filename_country_id = ""           # resolved below

    parts = re.split(r"\s*-\s*", filename_stem)
    if len(parts) >= 1 and parts[0].strip():
        filename_comp_code = parts[0].strip().lstrip("0") or "0"   # '0888' → '888'
        log.info("[SteveMadden-FootwearPO] Comp code from filename: '%s'", filename_comp_code)
    if len(parts) >= 2 and parts[1].strip():
        filename_sbu = parts[1].strip().upper()
        log.info("[SteveMadden-FootwearPO] SBU from filename: '%s'", filename_sbu)

    # ── Locate season token by regex ─────────────────────────────
    season_idx = None
    for i, p in enumerate(parts):
        if re.match(r'^[A-Z]{2}\d{2,4}$', p.strip(), re.IGNORECASE):
            sea_num = p.strip()[2:]
            sea_year_2d = sea_num[-2:] if len(sea_num) == 4 else sea_num
            season = f"{p.strip()[:2].upper()}{sea_year_2d}"
            season_idx = i
            log.info("[SteveMadden-FootwearPO] Season extracted from filename: '%s'", season)
            break
    else:
        # Fallback: regex anywhere in stem
        season_match = re.search(r'([A-Z]{2})(\d{2,4})', filename_stem, re.IGNORECASE)
        if season_match:
            sea_num = season_match.group(2)
            sea_year_2d = sea_num[-2:] if len(sea_num) == 4 else sea_num
            season = f"{season_match.group(1).upper()}{sea_year_2d}"
            log.info("[SteveMadden-FootwearPO] Season extracted (fallback): '%s'", season)

    # ── Country: first 2-letter token after season ────────────────
    if season_idx is not None:
        for p in parts[season_idx + 1:]:
            p = p.strip().upper()
            if re.match(r'^[A-Z]{2,3}$', p) and not p.isdigit():
                filename_country_id = p
                log.info("[SteveMadden-FootwearPO] Country ID from filename: '%s'", filename_country_id)
                break
    if not filename_country_id:
        # Fallback: 2nd-to-last dash token
        candidate = parts[-2].strip().upper() if len(parts) >= 2 else ""
        if re.match(r'^[A-Z]{2}$', candidate):
            filename_country_id = candidate
            log.info("[SteveMadden-FootwearPO] Country ID from filename (fallback): '%s'", filename_country_id)

    log.info(
        "[SteveMadden-FootwearPO] Filename metadata: comp=%s  sbu=%s  brand=%s  brand_code=%s  season=%s  country=%s",
        filename_comp_code, filename_sbu, brand, brand_code, season, filename_country_id,
    )

    # ── Brand LOV ID ─────────────────────────────────────────────
    # Taken directly from the filename (parts[2] trailing 3-letter code).
    # For '0888-FF-STEVE MADDEN SVM-...' → brand_code = 'SVM' → brand_lov_id = 'SVM'
    brand_lov_id = brand_code
    log.info("[SteveMadden-FootwearPO] Brand LOV ID from filename: '%s'", brand_lov_id)

    # ── Resolve country name from country ID via MDD ─────────────
    rna_country = ""
    if filename_country_id and mdd:
        rna_country = mdd.country_id_to_name(filename_country_id)
        log.info(
            "[SteveMadden-FootwearPO] Country: ID='%s' → name='%s' (for RNA lookup)",
            filename_country_id, rna_country,
        )
        if not rna_country:
            log.warning("[SteveMadden-FootwearPO] Country ID '%s' not found in Country LOV", filename_country_id)
    if not rna_country:
        rna_country = country_code   # fallback to args

    # ── RNA lookup: brand_type, brand_status, brand_category ─────
    brand_type = brand_status = brand_category = brand_group = ""
    if rna:
        log.info("[RNA] Total entries in lookup: %d", len(rna.lookup))
        for k, v in list(rna.lookup.items())[:5]:
            log.info("[RNA] Sample key: %s → %s", k, v)
        log.info(
            "[RNA] Searching with: country='%s'  comp='%s'  sbu='%s'  brand='%s'",
            rna_country, filename_comp_code, filename_sbu, brand_lov_id,
        )
        # Pass 1: exact match (country_name + comp_code + sbu + brand_code)
        rna_result = rna.get(rna_country, filename_comp_code, filename_sbu, brand_lov_id)
        log.info("[RNA] Pass 1 result: %s", rna_result)
        # Pass 2: fuzzy (country_name + sbu + brand_code)
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_fuzzy(rna_country, filename_sbu, brand_lov_id)
            log.info("[RNA] Pass 2 result: %s", rna_result)
        # Pass 3: fuzzy with original brand_code fallback
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_fuzzy(rna_country, filename_sbu, brand_code)
            log.info("[RNA] Pass 3 result: %s", rna_result)
        # Pass 4: brand code only — last resort
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_by_brand_only(brand_lov_id) or rna.get_by_brand_only(brand_code)
            log.info("[RNA] Pass 4 (brand-only) result: %s", rna_result)
        brand_type     = rna_result.get("brand_type",     "")
        brand_status   = rna_result.get("brand_status",   "")
        brand_category = rna_result.get("brand_category", "")
        brand_group    = rna_result.get("brand_group",    "")
        log.info(
            "[RNA] Final resolved — brand_type='%s'  brand_status='%s'  brand_category='%s'  brand_group='%s'",
            brand_type, brand_status, brand_category, brand_group,
        )
        if not brand_type and not brand_category:
            # Show all unique country values in RNA to help diagnose mismatch
            rna_countries = sorted({k[0] for k in rna.lookup})
            rna_brands    = sorted({k[3] for k in rna.lookup})
            log.warning("[RNA] ✗ All 4 passes failed — no match found")
            log.warning("[RNA] Countries in RNA: %s", rna_countries)
            log.warning("[RNA] Brands in RNA:    %s", rna_brands)
    else:
        log.warning("[RNA] RNALoader not initialised — brand attributes will be empty")

    # ── Load input file ──────────────────────────────────────────
    loader = SteveMaddenFootwearPOLoader(input_file)
    if loader.df.empty:
        log.warning("[SteveMadden-FootwearPO] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_rows(raw_rows, loader, brand_code, mdd, svm_mapping)

    if not generics:
        log.warning("[SteveMadden-FootwearPO] No valid generics produced")
        return

    # ── Build season classification ID (Clarks pattern) ──────────
    # season is already normalised to e.g. "SP29", "SS27", "FW26"
    sp_code  = season[:2].upper() if len(season) >= 2 else season
    sys_part = season[2:] if len(season) > 2 else ""
    sy       = f"20{sys_part}" if len(sys_part) == 2 else sys_part
    season_id = f"CLH_{brand_code}_{sp_code}{sy}"

    log.info(
        "[SteveMadden-FootwearPO] Season: code=%s  year=%s  season_id=%s",
        sp_code, sy, season_id,
    )

    # ── Generate XML ─────────────────────────────────────────────
    xml_filename = f"{input_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = 0

    # ── Build Classifications block ──────────────────────────────
    cls_el  = build_classifications(brand, brand_code, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))

    log.info("[SteveMadden-FootwearPO] Writing XML → %s", xml_filename)
    with open(xml_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            f'<STEP-ProductInformation '
            f'xmlns="{STIBO_NS}" '
            f'xmlns:xsi="{STIBO_XSI}" '
            f'xsi:schemaLocation="{STIBO_SCHEMA}" '
            f'ExportTime="{export_time}" '
            f'ExportContext="Context1" '
            f'WorkspaceID="Main" '
            f'UseContextLocale="false">\n\n'
        )

        # Write Classifications
        f.write(f"  {cls_str}\n\n")

        f.write("  <Products>\n")

        all_generics = list(generics.values())
        log.info(
            "[SteveMadden-FootwearPO] Processing %d generics",
            len(all_generics),
        )

        for generic in all_generics:
            px = build_product_xml(
                generic, brand, brand_code, comp_code, sbu,
                season, season_id, filename_country_id,
                brand_type=brand_type,
                brand_status=brand_status,
                brand_category=brand_category,
                brand_group=brand_group,
                mdd=mdd,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info(
        "[SteveMadden-FootwearPO] XML written: %s  (%d KB)",
        xml_path.name, file_kb,
    )
    print(
        f"=== STEVE MADDEN FOOTWEAR PO SUMMARY ===\n"
        f"  Input rows : {len(raw_rows)}\n"
        f"  Generics   : {generic_count}\n"
        f"  XML size   : {file_kb} KB",
        flush=True,
    )

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return xml_path


# ======================================================================
# CLI  (local test)
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Steve Madden Footwear PO → Stibo XML")
    parser.add_argument("--brand",        default=BRAND_NAME)
    parser.add_argument("--brand_code",   default=BRAND_CODE)
    parser.add_argument("--comp_code",    default="0888")
    parser.add_argument("--sbu",          default="SP")
    parser.add_argument("--season",       default="SS27")
    parser.add_argument("--country_code", default="")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    run(args)
