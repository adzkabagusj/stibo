"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Asics Apparel & Accessories v1.0║
║   Apparel & Accessories Order Form → Stibo STEP XML             ║
║   (Generic Articles + Colour Variants, with Size)               ║
╚══════════════════════════════════════════════════════════════════╝

Dynamic Attribute Mapping:
    Reads from "New brand mapping file" → "Asics(Inline)" tab
    
    Mapping columns:
    • Col A: Stibo Attribute Name
    • Col B: Stibo Attribute ID
    • Col C: Type (LOV / Text)
    • Col G: Source ("Direct from Principal" / "Formula in system" / other)
    • Col J: Excel Column Name (source column from input file)
    • Col K: Mapping Logic (transformation rules)

Inclusion Rules:
    Include attribute if:
    1. Col G = "Direct from Principal", OR
    2. Col G = "Formula in system" AND Col J has value
    Otherwise: Skip

LOV vs Text:
    • LOV  → use <Value AttributeID="..." ID="..."/>
    • Text → use <Value AttributeID="...">text</Value>
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

INPUT_DIR       = BASE_DIR / "input" / "apparel_accessories"
MDD_DIR         = BASE_DIR / "input" / "mdd"
ATTR_DIR        = BASE_DIR / "input" / "attributes"
XML_OUT_DIR     = BASE_DIR / "output" / "xml"
LOG_DIR         = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME  = "Asics"
BRAND_CODE  = "ASC"          # 3-letter SAP brand code for Asics

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value — returns empty string for None/nan/empty."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT", "#VALUE!", "nan") else s


def _first_sentence(text: str) -> str:
    """Return the first sentence (up to first newline, period, or semicolon)."""
    if not text:
        return ""
    for sep in ("\n", ".", ";"):
        idx = text.find(sep)
        if idx > 0:
            return text[:idx].strip()
    return text.strip()


def _resolve_currency_from_country(country_code: str) -> str:
    """Resolve Retail Price Currency LOV ID from country code."""
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
    Loads LOV tables from the MDD Excel needed for Asics Apparel & Accessories.
    Provides lookup functionality for all LOV types.
    """

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Load all LOV sheets
        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        
        # Specific LOV loaders
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
        self._load_width_lov(wb)
        self._load_sports_category_lov(wb)

        wb.close()
        log.info("[MDD] %d LOV types loaded", len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip().upper()
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
                    self.lovs.setdefault(display_name, {})[str(name).strip().upper()] = (
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
        """Material LOV: display name → LOV ID."""
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
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["BrandTypeLOV"] = lov_map
        self.lovs["AT_BrandType"] = lov_map
        log.info("[MDD] Brand Type LOV loaded — %d entries", len(lov_map))

    def _load_brand_status_lov(self, wb):
        """Brand Status LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Status LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["BrandStatusLOV"] = lov_map
        self.lovs["AT_BrandStatus"] = lov_map
        log.info("[MDD] Brand Status LOV loaded — %d entries", len(lov_map))

    def _load_brand_category_lov(self, wb):
        """Brand Category LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND CATEGORY LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Category LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["BrandCategoryLOV"] = lov_map
        self.lovs["AT_BrandCategory"] = lov_map
        log.info("[MDD] Brand Category LOV loaded — %d entries", len(lov_map))

    def _load_brand_group_lov(self, wb):
        """Brand Group LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            sheet_name = next(
                (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP"),
                None,
            )
        if not sheet_name:
            log.warning("[MDD] Brand Group LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["BrandGroupLOV"] = lov_map
        self.lovs["AT_BrandGroup"] = lov_map
        log.info("[MDD] Brand Group LOV loaded — %d entries", len(lov_map))

    def _load_country_lov(self, wb):
        """Country LOV: Col A = LOV ID (e.g. 'ID'), Col B = display name (e.g. 'Indonesia')."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper()
                  and "ORIGIN" not in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Country LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        id_to_name: dict[str, str] = {}
        name_to_id: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                id_to_name[lov_id.upper()] = display
                name_to_id[display.upper()] = lov_id
        self.lovs["CountryLOV"]         = id_to_name
        self.lovs["CountryLOV_NameToId"] = name_to_id
        log.info("[MDD] Country LOV loaded — %d entries", len(id_to_name))

    def country_id_to_name(self, country_id: str) -> str:
        """Convert country ISO code to full name. e.g. 'ID' -> 'Indonesia'."""
        return self.lovs.get("CountryLOV", {}).get(country_id.strip().upper(), "")

    def _load_width_lov(self, wb):
        """Width LOV: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "WIDTH LOV"),
            next((s for s in wb.sheetnames if "WIDTH" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Width LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["WidthLOV"] = lov_map
        self.lovs["AT_Width"] = lov_map
        log.info("[MDD] Width LOV loaded — %d entries", len(lov_map))

    def _load_sports_category_lov(self, wb):
        """Sports Category LOV: Col A = display value, Col B = LOV ID."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "SPORTS CATEGORY LOV"),
            next((s for s in wb.sheetnames if "SPORTS" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Sports Category LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov_map[display.upper()] = lov_id
        self.lovs["SportsCategoryLOV"] = lov_map
        self.lovs["AT_SportsCategoryEN"] = lov_map
        log.info("[MDD] Sports Category LOV loaded — %d entries", len(lov_map))

    def lookup(self, lov_key: str, display_value: str) -> str:
        """Look up a LOV ID. Returns display_value unchanged if not found."""
        lov = self.lovs.get(lov_key, {})
        if not lov:
            return display_value
        return lov.get(display_value.strip().upper(), display_value)


# ======================================================================
# BRAND MAPPING LOADER (Dynamic Attribute Mapping)
# ======================================================================

class BrandMappingLoader:
    """
    Loads the brand-specific attribute mapping from 'Asics(Inline)' tab.
    
    Structure:
    • Col A: Stibo Attribute Name
    • Col B: Stibo Attribute ID
    • Col C: Type (LOV / Text)
    • Col G: Source ("Direct from Principal" / "Formula in system")
    • Col J: Excel Column Name (source)
    • Col K: Mapping Logic
    
    Inclusion logic:
    - Include if Col G = "Direct from Principal"
    - Include if Col G = "Formula in system" AND Col J has value
    - Otherwise skip
    """
    
    def __init__(self, path: Path, sheet_name: str = "Asics(Inline)"):
        self.path = path
        self.sheet_name = sheet_name
        self.mappings: list[dict] = []
        self._load()
    
    def _load(self):
        log.info("[BrandMapping] Loading from: %s → sheet '%s'", self.path.name, self.sheet_name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        log.info("[BrandMapping] Available sheets in '%s': %s", self.path.name, wb.sheetnames)

        # Fuzzy sheet name resolution: exact → case-insensitive → partial
        resolved_sheet = None
        if self.sheet_name in wb.sheetnames:
            resolved_sheet = self.sheet_name
        else:
            target_norm = self.sheet_name.lower().replace(" ", "")
            for sn in wb.sheetnames:
                if sn.lower().replace(" ", "") == target_norm:
                    resolved_sheet = sn
                    log.info("[BrandMapping] Fuzzy matched sheet '%s' → '%s'", self.sheet_name, sn)
                    break
            if not resolved_sheet:
                for sn in wb.sheetnames:
                    if target_norm in sn.lower().replace(" ", ""):
                        resolved_sheet = sn
                        log.info("[BrandMapping] Partial matched sheet '%s' → '%s'", self.sheet_name, sn)
                        break

        if not resolved_sheet:
            log.warning("[BrandMapping] Sheet '%s' not found. Available: %s",
                        self.sheet_name, wb.sheetnames)
            wb.close()
            return

        ws = wb[resolved_sheet]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.warning("[BrandMapping] Sheet '%s' is empty", resolved_sheet)
            return

        # Find header row (look for "Stibo Attribute Name" or "AttributeID")
        hdr_idx = None
        for i, row in enumerate(rows[:50]):
            if row and any(isinstance(v, str) and 
                          ("STIBO ATTRIBUTE" in v.upper() or "ATTRIBUTEID" in v.upper()) 
                          for v in row if v):
                hdr_idx = i
                break
        
        if hdr_idx is None:
            log.warning("[BrandMapping] Cannot find header row in sheet '%s'", self.sheet_name)
            return
        
        log.info("[BrandMapping] Header found at row %d", hdr_idx)
        
        # Parse mappings
        for row_idx, row in enumerate(rows[hdr_idx + 1:], start=hdr_idx + 2):
            if not row or len(row) < 7:
                continue
            
            attr_name      = _s(row[0])  # Col A
            attr_id        = _s(row[1])  # Col B
            attr_type      = _s(row[2])  # Col C
            source_type    = _s(row[6])  # Col G (0-indexed: col 6)
            excel_col_name = _s(row[9]) if len(row) > 9 else ""  # Col J (0-indexed: col 9)
            mapping_logic  = _s(row[10]) if len(row) > 10 else "" # Col K (0-indexed: col 10)
            
            # Skip if no attribute ID
            if not attr_id:
                continue
            
            # Inclusion logic
            include = False
            if "DIRECT FROM PRINCIPAL" in source_type.upper():
                include = True
            elif "FORMULA IN SYSTEM" in source_type.upper() and excel_col_name:
                include = True
            
            if not include:
                continue
            
            mapping = {
                "attr_name": attr_name,
                "attr_id": attr_id,
                "attr_type": attr_type.upper(),  # LOV or TEXT
                "source_type": source_type,
                "excel_col_name": excel_col_name,
                "mapping_logic": mapping_logic,
                "is_lov": "LOV" in attr_type.upper(),
            }
            
            self.mappings.append(mapping)
            log.debug("[BrandMapping] Row %d: %s → %s (type=%s, source=%s, col=%s)", 
                     row_idx, attr_name, attr_id, attr_type, source_type, excel_col_name)
        
        log.info("[BrandMapping] Loaded %d attribute mappings from '%s'", 
                len(self.mappings), self.sheet_name)
    
    def get_mappings(self) -> list[dict]:
        """Return all loaded attribute mappings."""
        return self.mappings


# ======================================================================
# ASICS MD MAPPINGS LOADER
# ======================================================================

class AsicsMDMappingsLoader:
    """
    Loads Asics MD Mappings from 'Asics MD Mappings' sheet.
    Maps Category 2 values to Sports Category values (apparel/accessories).

    Structure:
    • Col A: Division (filter for 'APP' / 'ACC')
    • Col B: Category 2 value
    • Col D: Sports Category value
    """

    def __init__(self, path: Path, sheet_name: str = "Asics MD Mappings"):
        self.path = path
        self.sheet_name = sheet_name
        self.mappings: dict[str, str] = {}  # Category 2 -> Sports Category
        self._load()
    
    def _load(self):
        log.info("[AsicsMDMappings] Loading from: %s → sheet '%s'", self.path.name, self.sheet_name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        if self.sheet_name not in wb.sheetnames:
            log.warning("[AsicsMDMappings] Sheet '%s' not found. Available: %s", 
                       self.sheet_name, wb.sheetnames)
            wb.close()
            return
        
        ws = wb[self.sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        
        if not rows:
            log.warning("[AsicsMDMappings] Sheet '%s' is empty", self.sheet_name)
            return
        
        # Find header row
        hdr_idx = 0
        for i in range(min(20, len(rows))):
            if rows[i] and any(
                isinstance(v, str) and "DIVISION" in v.upper()
                for v in rows[i] if v
            ):
                hdr_idx = i
                log.info("[AsicsMDMappings] Header found at row %d", hdr_idx)
                break
        
        # Parse mappings (filter for APP / ACC in Col A)
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 4:
                continue

            division  = _s(row[0])  # Col A
            cat2      = _s(row[1])  # Col B (Category 2)
            sports_cat = _s(row[3])  # Col D

            # Only include rows where Division = 'APP' or 'ACC'
            if division.upper() in ("APP", "ACC") and cat2 and sports_cat:
                self.mappings[cat2.upper()] = sports_cat

        log.info("[AsicsMDMappings] Loaded %d APP/ACC mappings from '%s'",
                len(self.mappings), self.sheet_name)

    def get_sports_category(self, category2: str) -> str:
        """Get Sports Category for a given Category 2 value."""
        return self.mappings.get((category2 or "").strip().upper(), "")


# ======================================================================
# RNA LOADER
# ======================================================================

class RNALoader:
    """
    Loads RNA (Source Mapping) data to resolve Brand Type, Brand Category,
    Brand Group and Brand Status.
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
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}

    def get_by_brand_only(self, brand_code: str) -> dict:
        """Match on brand code only."""
        b = (brand_code or "").strip().upper()
        for (_, _, _, kb), val in self.lookup.items():
            if kb == b:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


# ======================================================================
# INPUT FILE LOADER
# ======================================================================

class AsicsApparelAccessoriesLoader:
    """
    Loads Asics Apparel & Accessories data from Excel.
    Automatically detects the correct sheet and header row.
    """

    HEADER_SIGNALS = {"style", "color", "size", "ean", "barcode", "description"}

    def __init__(self, path: Path):
        self.path    = path
        self.df      = pd.DataFrame()
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[AsicsApparelAcc] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[AsicsApparelAcc] Available sheets: %s", wb.sheetnames)

        # Prefer "AW26" sheet; fall back to first sheet
        PREFERRED_SHEETS = ["AW26"]
        target = next(
            (s for s in wb.sheetnames if s.strip() in PREFERRED_SHEETS),
            wb.sheetnames[0],
        )
        log.info("[AsicsApparelAcc] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.warning("[AsicsApparelAcc] Sheet '%s' is empty", target)
            return

        # Locate header row by looking for "Item Code" in any cell
        hdr_idx = 0
        for i in range(min(20, len(rows))):
            if rows[i] and any(
                str(v).strip() in ("Item Code", "ITEM CODE")
                for v in rows[i] if v is not None
            ):
                hdr_idx = i
                log.info("[AsicsApparelAcc] Header detected at row %d", hdr_idx)
                break
        else:
            log.warning("[AsicsApparelAcc] 'Item Code' not found in first 20 rows — defaulting to row 0")
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

        # Filter: keep rows with valid data (at least one non-empty cell)
        df = df.dropna(how='all')
        
        self.df = df.reset_index(drop=True)
        log.info(
            "[AsicsApparelAcc] Loaded %d data rows | %d columns",
            len(self.df), len(header),
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
# DYNAMIC ATTRIBUTE MAPPER
# ======================================================================

def apply_mapping_logic(value: str, logic: str, mdd: MDDLoader = None) -> str:
    """
    Apply transformation logic from Col K (Mapping Logic).
    
    Examples:
    - "Uppercase" → convert to uppercase
    - "First sentence" → extract first sentence
    - "LOV lookup: AT_Gender" → look up in MDD
    - "" → return as-is
    """
    if not logic or not value:
        return value
    
    logic_upper = logic.upper()
    
    # Uppercase transformation
    if "UPPERCASE" in logic_upper or "UPPER" in logic_upper:
        return value.upper()
    
    # First sentence extraction
    if "FIRST SENTENCE" in logic_upper:
        return _first_sentence(value)
    
    # LOV lookup
    if "LOV LOOKUP" in logic_upper and mdd:
        # Extract LOV key from logic (e.g., "LOV lookup: AT_Gender")
        parts = logic.split(":")
        if len(parts) > 1:
            lov_key = parts[1].strip()
            return mdd.lookup(lov_key, value)
    
    # Default: return as-is
    return value


def group_rows_dynamic(
    raw_rows: list[dict],
    loader: AsicsApparelAccessoriesLoader,
    brand_code: str,
    brand_mapping: BrandMappingLoader,
    mdd: MDDLoader | None,
) -> dict[str, dict]:
    """
    Group flat rows into Generic articles using dynamic brand mapping.
    """
    mappings = brand_mapping.get_mappings()
    
    if not mappings:
        log.warning("[Grouper] No brand mappings loaded — proceeding with hardcoded attributes only")
    
    # Identify key columns for grouping
    # (These should be defined in the mapping or we use common defaults)
    col_style = loader.find_col(
        "KEY ITEM", "Item Trading Code",  # Asics-specific
        "Style", "StyleCode", "Style Code", "Article",
        "SKU", "Product Code", "Style #", "Style Number",
        "Model", "Item Code", "Item", "Product", "Reference",
        "Style No", "Style No.", "ArticleCode", "Article Code",
        "Material", "Material Number", "Mat.",
    )
    col_principal_style = loader.find_col(
        "Item Code", "ITEM CODE", "ItemCode",  # for AT_PrincipalStyleCode
    )
    col_working_name = loader.find_col(
        "Working Name",  # Asics-specific for AT_PrincipalStyleDescription and Name
    )
    col_colorway = loader.find_col(
        "Colorway",  # Asics-specific for AT_PrincipalColorName
    )
    col_gender = loader.find_col(
        "Gender Name", "GenderName"  # for AT_PrincipalGenderDescription
    )
    col_aid_srp = loader.find_col(
        "AID SRP",  # Asics-specific for AT_OriginalPrice
    )
    col_sub_silhouette_group = loader.find_col(
        "Sub Silhouette Group",  # Asics-specific for AT_PrincipalMerchandiseHierarchyL1
    )
    col_product_group_hier = loader.find_col(
        "Product Group",  # Asics-specific for AT_PrincipalMerchandiseHierarchyL2
    )
    col_category2 = loader.find_col(
        "Category 2",  # Asics-specific for AT_PrincipalMerchandiseHierarchyL3
    )
    col_width = loader.find_col(
        "Width",  # Asics-specific for AT_Width
    )
    col_color = loader.find_col(
        "Item Color Code", "Colorway", "Color Link",  # Asics-specific
        "Color", "Colour", "ColorCode", "Color Code",
        "Colour Code", "ColourCode", "Color Name", "Colour Name",
        "Colourway",
    )
    col_size  = loader.find_col(
        "Sizes",  # Asics-specific
        "Size", "SizeCode", "Size Code", "Size Name",
        "EU Size", "UK Size", "US Size",
    )
    col_ean   = loader.find_col(
        "Global RID", "AID RID",  # Asics-specific (might be EAN equivalents)
        "EAN", "Barcode", "EAN Code", "EAN13", "GTIN", "UPC",
    )
    col_incoming_month = loader.find_col(
        "Regional RID",  # Asics-specific for AT_IncomingMonth
    )
    col_product_group = loader.find_col(
        "Product Group",  # For determining ParentID (ACCESSORIES vs APPAREL)
    )

    log.info(
        "[Grouper] Key columns → style=%s, principal_style=%s, working_name=%s, colorway=%s, gender=%s, aid_srp=%s, sub_silhouette_group=%s, product_group_hier=%s, category2=%s, width=%s, color=%s, size=%s, ean=%s",
        col_style, col_principal_style, col_working_name, col_colorway, col_gender, col_aid_srp, col_sub_silhouette_group, col_product_group_hier, col_category2, col_width, col_color, col_size, col_ean,
    )
    log.info("[Grouper] Available columns: %s", loader.headers)

    if col_principal_style is None:
        log.warning(
            "[Grouper] Could not find 'Item Color Code' column (for AT_PrincipalStyleCode). "
            "All available columns: %s",
            loader.headers,
        )
        return {}
    
    result: dict[str, dict] = {}
    
    for row in raw_rows:
        # Generate generic key (brand + principal style code from Item Color Code)
        principal_style_code = _s(row.get(col_principal_style, "")) if col_principal_style else ""
        if not principal_style_code:
            continue
        
        generic_key = f"{brand_code}{principal_style_code}"
        
        # First encounter: create generic article
        if generic_key not in result:
            working_name = _s(row.get(col_working_name, "")) if col_working_name else ""
            colorway = _s(row.get(col_colorway, "")) if col_colorway else ""
            gender = _s(row.get(col_gender, "")) if col_gender else ""
            principal_size = _s(row.get(col_size, "")) if col_size else ""
            aid_srp = _s(row.get(col_aid_srp, "")) if col_aid_srp else ""
            sub_silhouette_group = _s(row.get(col_sub_silhouette_group, "")) if col_sub_silhouette_group else ""
            product_group_hier = _s(row.get(col_product_group_hier, "")) if col_product_group_hier else ""
            category2 = _s(row.get(col_category2, "")) if col_category2 else ""
            width = _s(row.get(col_width, "")) if col_width else ""
            # Format incoming_month as DD-MM-YYYY
            # Input can be a datetime object or a string in M/D/YYYY format
            _raw_incoming = row.get(col_incoming_month, "") if col_incoming_month else ""
            if _raw_incoming and hasattr(_raw_incoming, "month"):
                # datetime/date object
                incoming_month = f"{_raw_incoming.day:02d}-{_raw_incoming.month:02d}-{_raw_incoming.year}"
            else:
                _str_incoming = _s(_raw_incoming)
                # Try to parse string in M/D/YYYY format (e.g. "8/1/2026")
                _m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', _str_incoming)
                if _m:
                    incoming_month = f"{int(_m.group(2)):02d}-{int(_m.group(1)):02d}-{_m.group(3)}"
                else:
                    incoming_month = _str_incoming
            product_group = _s(row.get(col_product_group, "")) if col_product_group else ""
            result[generic_key] = {
                "generic_key": generic_key,
                "style_code": _s(row.get(col_style, "")) if col_style else "",
                "principal_style_code": principal_style_code,
                "working_name": working_name,
                "colorway": colorway,
                "gender": gender,
                "principal_size": principal_size,
                "aid_srp": aid_srp,
                "sub_silhouette_group": sub_silhouette_group,
                "product_group_hier": product_group_hier,
                "category2": category2,
                "width": width,
                "incoming_month": incoming_month,
                "product_group": product_group,
                "brand_code": brand_code,
                "attributes": {},  # Will be populated dynamically
                "variants": {},
            }
            
            # Apply all mapped attributes
            for mapping in mappings:
                excel_col = mapping["excel_col_name"]
                attr_id   = mapping["attr_id"]
                is_lov    = mapping["is_lov"]
                logic     = mapping["mapping_logic"]
                
                # Get value from Excel
                raw_value = ""
                if excel_col:
                    # Try exact match first
                    if excel_col in row:
                        raw_value = _s(row[excel_col])
                    else:
                        # Try case-insensitive match
                        found_col = loader.find_col(excel_col)
                        if found_col:
                            raw_value = _s(row[found_col])
                
                # Apply mapping logic
                if raw_value:
                    processed_value = apply_mapping_logic(raw_value, logic, mdd)
                    
                    # If LOV type, look up in MDD
                    if is_lov and mdd:
                        processed_value = mdd.lookup(attr_id, processed_value)
                    
                    result[generic_key]["attributes"][attr_id] = {
                        "value": processed_value,
                        "is_lov": is_lov,
                    }
        
        # Add variant (color + size combination)
        color_code = _s(row.get(col_color, "")) if col_color else "000"
        size_code  = _s(row.get(col_size, ""))  if col_size  else "000"
        ean        = _s(row.get(col_ean, ""))   if col_ean   else ""
        
        variant_key = f"{generic_key}{color_code}{size_code}"
        
        result[generic_key]["variants"][variant_key] = {
            "variant_key": variant_key,
            "color_code": color_code,
            "size_code": size_code,
            "ean": ean,
        }
    
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


# ======================================================================
# CLASSIFICATIONS BUILDER
# ======================================================================

def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_tail   = season_code[2:]
    sea_year   = f"20{sea_tail}" if len(sea_tail) == 2 else sea_tail

    full_season    = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season}"
    batches_parent = f"CLH_{brand}Batches"

    _SEASON_LABELS = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
    }
    sea_name       = _SEASON_LABELS.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
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
# XML BUILDER (Dynamic)
# ======================================================================

def build_product_xml_dynamic(
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
    asics_md_mappings: "AsicsMDMappingsLoader | None" = None,
) -> str:
    """Build XML for one Generic article using dynamic attributes."""
    if not generic.get("principal_style_code"):
        return ""

    # Determine ParentID based on Product Group
    product_group = generic.get("product_group", "").upper()
    if product_group == "ACCESSORIES":
        parent_id = "PPH_E-TempSubCat"  # Accessories
    elif product_group == "APPAREL":
        parent_id = "PPH_A-TempSubCat"  # Apparel
    else:
        # Default to Apparel if not specified
        parent_id = "PPH_A-TempSubCat"
        log.warning(
            "[XML] Product Group '%s' not recognized (expected ACCESSORIES or APPAREL). Defaulting to Apparel.",
            generic.get("product_group", "(empty)"),
        )

    # Generic Article
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_key"]

    # Name (use Working Name column value or fallback to generic key)
    name = generic.get("working_name") or generic["generic_key"]
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name

    # Classification References
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # Generic Values (dynamic attributes)
    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    for attr_id, attr_data in generic["attributes"].items():
        value  = attr_data["value"]
        is_lov = attr_data["is_lov"]
        
        if not value:
            continue
        
        if is_lov:
            _val_lov(gv, attr_id, value)
        else:
            _val_text(gv, attr_id, value)

    # ── AT_PrincipalStyleCode: from Item Color Code column ──
    _val_text(gv, "AT_PrincipalStyleCode", generic["principal_style_code"])

    # ── AT_PrincipalColorCode: last 3 digits from Item Color Code column ──
    principal_color_code = generic["principal_style_code"][-3:] if len(generic["principal_style_code"]) >= 3 else generic["principal_style_code"]
    _val_text(gv, "AT_PrincipalColorCode", principal_color_code)

    # ── AT_PrincipalColorName: from Colorway column ──
    if generic.get("colorway"):
        _val_text(gv, "AT_PrincipalColorName", generic["colorway"])

    # ── AT_PrincipalGenderDescription: from Gender column ──
    if generic.get("gender"):
        _val_text(gv, "AT_PrincipalGenderDescription", generic["gender"])

    # ── AT_OriginalPrice: from AID SRP column ──
    if generic.get("aid_srp"):
        _val_text(gv, "AT_OriginalPrice", generic["aid_srp"])

    # ── AT_CurrentPrice: from AID SRP column ──
    if generic.get("aid_srp"):
        _val_text(gv, "AT_CurrentPrice", generic["aid_srp"])

    # ── AT_PrincipalMerchandiseHierarchyL1: from Sub Silhouette Group column ──
    if generic.get("sub_silhouette_group"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL1", generic["sub_silhouette_group"])

    # ── AT_PrincipalMerchandiseHierarchyL2: from Product Group column ──
    if generic.get("product_group_hier"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL2", generic["product_group_hier"])

    # ── AT_PrincipalMerchandiseHierarchyL3: from Category 2 column ──
    if generic.get("category2"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL3", generic["category2"])

    # ── AT_PrincipalStyleDescription: from Working Name column ──
    if generic.get("working_name"):
        _val_text(gv, "AT_PrincipalStyleDescription", generic["working_name"])

    # ── AT_InboundGenericCode: same as KEY_InboundArticle ──
    _val_text(gv, "AT_InboundGenericCode", generic["generic_key"])

    # ── AT_Brand: look up LOV ID from MDD using brand display name (e.g. 'ASICS' → 'ASC') ──
    _brand_lov = mdd.lovs.get("AT_Brand", {}) if mdd else {}
    _brand_lov_id = _brand_lov.get(brand.upper(), brand_code)
    _val_lov(gv, "AT_Brand", _brand_lov_id)

    # ── AT_Gender: from Gender column with transformation and LOV lookup ──
    if generic.get("gender"):
        gender_raw = generic["gender"]
        # Transform Excel values to Gender LOV display values
        gender_mapping = {
            "MEN": "MALE",
            "WOMEN": "FEMALE",
            "UNISEX": "UNISEX",
        }
        gender_display = gender_mapping.get(gender_raw.upper(), gender_raw.upper())
        
        # Look up LOV ID from Gender LOV (Male → M, Female → F, Unisex → U)
        gender_lov = mdd.lovs.get("AT_Gender", {}) if mdd else {}
        gender_lov_id = gender_lov.get(gender_display, "")
        
        if gender_lov_id:
            _val_lov(gv, "AT_Gender", gender_lov_id)
        else:
            # Fallback: use transformed display value as text
            _val_text(gv, "AT_Gender", gender_display)

    # ── AT_BYGender: from Gender column with transformation and LOV lookup ──
    if generic.get("gender"):
        gender_raw = generic["gender"]
        # Transform Excel values to Gender LOV display values
        gender_mapping = {
            "MEN": "MALE",
            "WOMEN": "FEMALE",
            "UNISEX": "UNISEX",
        }
        gender_display = gender_mapping.get(gender_raw.upper(), gender_raw.upper())
        
        # Look up LOV ID from Gender LOV (Male → M, Female → F, Unisex → U)
        gender_lov = mdd.lovs.get("AT_Gender", {}) if mdd else {}
        gender_lov_id = gender_lov.get(gender_display, "")
        
        if gender_lov_id:
            _val_lov(gv, "AT_BYGender", gender_lov_id)
        else:
            # Fallback: use transformed display value as text
            _val_text(gv, "AT_BYGender", gender_display)

    # ── System-level attributes (brand/season/country/comp/sbu) ──
    log.info(
        "[XML] Brand attrs — brand_type='%s'  brand_status='%s'  brand_category='%s'  brand_group='%s'",
        brand_type, brand_status, brand_category, brand_group,
    )

    # ── Brand Type ───────────────────────────────────────────────
    if brand_type:
        bt_lov_id = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type.upper(), "")
        if bt_lov_id:
            _val_lov(gv, "AT_BrandType", bt_lov_id)
        else:
            _val_text(gv, "AT_BrandType", brand_type)

    # ── Brand Status ─────────────────────────────────────────────
    if brand_status:
        bs_lov_id = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status.upper(), "")
        if bs_lov_id:
            _val_lov(gv, "AT_BrandStatus", bs_lov_id)
        else:
            _val_text(gv, "AT_BrandStatus", brand_status)

    # ── Brand Category ─────────────────────────────────────────── (DISABLED)
    # if brand_category:
    #     bc_lov_id = (mdd.lovs.get("AT_BrandCategory", {}) if mdd else {}).get(brand_category.upper(), "")
    #     if bc_lov_id:
    #         _val_lov(gv, "AT_BrandCategory", bc_lov_id)
    #     else:
    #         _val_text(gv, "AT_BrandCategory", brand_category)

    # ── Brand Group ──────────────────────────────────────────────
    if brand_group:
        bg_lov_id = (mdd.lovs.get("AT_BrandGroup", {}) if mdd else {}).get(brand_group.upper(), "")
        if bg_lov_id:
            _val_lov(gv, "AT_BrandGroup", bg_lov_id)
        else:
            _val_text(gv, "AT_BrandGroup", brand_group)

    # Company + SBU (MultiValue pattern — LOV ID)
    mv_comp = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_comp.set("AttributeID", "AT_CompanyCode")
    v_comp = ET.SubElement(mv_comp, f"{{{STIBO_NS}}}Value")
    v_comp.set("ID", comp_code)

    mv_sbu = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_sbu.set("AttributeID", "AT_SBU")
    v_sbu = ET.SubElement(mv_sbu, f"{{{STIBO_NS}}}Value")
    v_sbu.set("ID", sbu)

    # Season + Season Year
    sea_code  = season[:2].upper() if len(season) >= 2 else season
    _val_lov(gv, "AT_Season", sea_code)

    # Extract year from season (e.g., "SS27" → "2027")
    year_match = re.search(r"20\d{2}", season)
    if not year_match:
        short_match = re.search(r"\d{2}$", season)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    if season_year:
        _val_text(gv, "AT_SeasonYear", season_year)

    # Country (destination country)
    if country_code:
        _val_lov(gv, "AT_Country", country_code.upper())

    # ── Retail Price Currency (from country code) ───────────────
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _val_lov(gv, "AT_RetailPriceCurrency", currency_lov_id)

    # ── FOB Currency (from country code, same as Retail Price Currency) ──
    fob_currency_lov_id = _resolve_currency_from_country(country_code)
    if fob_currency_lov_id:
        _val_lov(gv, "AT_FOBCurrency", fob_currency_lov_id)
    else:
        # Fallback to USD if country code not found
        _val_lov(gv, "AT_FOBCurrency", "USD")

    # ── SAP Product Flag (default A) ─────────────────────────────
    _val_lov(gv, "AT_SAPProductFlag", "A")

    # ── Material Type (default ZINA) ─────────────────────────────
    _val_lov(gv, "AT_MaterialType", "ZINA")

    # ── SAP Article Category (default 01) ────────────────────────
    _val_lov(gv, "AT_SAPArticleCategory", "01")

    # ── SAP Age (default AD) ──────────────────────────────────────
    _val_lov(gv, "AT_SAPAge", "AD")

    # ── Pricing Distribution Channel (default 01) ────────────────
    _val_lov(gv, "AT_PricingDistributionChannel", "01")

    # ── UOM (default EA) ──────────────────────────────────────────
    _val_lov(gv, "AT_UOM", "EA")

    # ── BY Age (default ADULT) ────────────────────────────────────
    _val_lov(gv, "AT_BYAge", "ADULT")

    # ── Country Size (default US) ─────────────────────────────────
    _val_lov(gv, "AT_CountrySize", "US")

    # ── Sports Category EN: from Category 2 via Asics MD Mappings ─────
    _sports_category_en_display = ""   # keep for AT_EComProductNameEN below
    if generic.get("category2") and asics_md_mappings and mdd:
        category2_value = generic["category2"]
        # Step 1: Map Category 2 to Sports Category using Asics MD Mappings
        sports_category = asics_md_mappings.get_sports_category(category2_value)

        if sports_category:
            _sports_category_en_display = sports_category   # save display value
            # Step 2: Look up LOV ID from Sports Category LOV (Col A → Col B)
            sports_cat_lov = mdd.lovs.get("AT_SportsCategoryEN", {})
            sports_cat_lov_id = sports_cat_lov.get(sports_category.upper(), "")

            if sports_cat_lov_id:
                # Step 3: If single digit, pad with leading zero
                if len(sports_cat_lov_id) == 1 and sports_cat_lov_id.isdigit():
                    sports_cat_lov_id = f"0{sports_cat_lov_id}"

                _val_lov(gv, "AT_SportsCategoryEN", sports_cat_lov_id)
                log.info(
                    "[XML] AT_SportsCategoryEN: category2='%s' → sports_cat='%s' → lov_id='%s'",
                    category2_value, sports_category, sports_cat_lov_id,
                )
            else:
                log.warning(
                    "[XML] Sports Category '%s' not found in Sports Category LOV",
                    sports_category,
                )
        else:
            log.warning(
                "[XML] Category 2 '%s' not found in Asics MD Mappings",
                category2_value,
            )

    # ── AT_EComProductNameEN: Brand + Working Name + Gender + Sports Category EN + "-" + Colorway ──
    _ecom_parts = [
        brand,
        generic.get("working_name", ""),
        generic.get("gender", ""),
        _sports_category_en_display,
    ]
    _ecom_prefix = " ".join(p for p in _ecom_parts if p)
    _ecom_colorway = generic.get("colorway", "")
    if _ecom_prefix and _ecom_colorway:
        ecom_product_name = f"{_ecom_prefix} - {_ecom_colorway}"
    elif _ecom_prefix:
        ecom_product_name = _ecom_prefix
    else:
        ecom_product_name = _ecom_colorway
    if ecom_product_name:
        _val_text(gv, "AT_EComProductNameEN", ecom_product_name)
        _val_text(gv, "AT_ShortDescriptionEN", ecom_product_name)
        _val_text(gv, "AT_LongDescriptionEN", ecom_product_name)
        log.info("[XML] AT_EComProductNameEN / AT_ShortDescriptionEN / AT_LongDescriptionEN: '%s'", ecom_product_name)

    # ── EComAgesCategory (default 18+Y) ──────────────────────────
    _val_lov(gv, "AT_EComAgesCategory", "18+Y")

    # ── Article Status (default A) ────────────────────────────────
    _val_lov(gv, "AT_ArticleStatus", "A")

    # ── BY Article Type (default Inline) ──────────────────────────
    _val_lov(gv, "AT_BYArticleType", "Inline")

    # ── AT_IncomingMonth: from AID RID column ─────────────────────
    if generic.get("incoming_month"):
        _val_text(gv, "AT_IncomingMonth", generic["incoming_month"])

    # ── Country Origin (default ID) ───────────────────────────────
    _val_lov(gv, "AT_CountryOrigin", "ID")

    # ── Apparel-specific attributes ──────────────────────────────
    if product_group == "APPAREL":
        # ── Content (default POLYESTER) ───────────────────────────────
        _val_lov(gv, "AT_Content", "POLYESTER")

        # ── Fabric (default JE) ───────────────────────────────────────
        _val_lov(gv, "AT_Fabric", "JE")

        # ── Pattern Print (default PLAIN) ─────────────────────────────
        _val_lov(gv, "AT_PatternPrint", "PLAIN")

    # ── Accessories-specific attributes ───────────────────────────
    elif product_group == "ACCESSORIES":
        # ── Material (default PES) ────────────────────────────────────
        _val_lov(gv, "AT_Material", "PES")

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """
    Main entry point for Asics Apparel & Accessories ETL.
    Called by asics/lambda_function.py after files are downloaded.
    """
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0999")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SS27")
    country_code = getattr(args, "country_code", "")

    log.info(
        "[AsicsApparelAcc] Starting: brand=%s  code=%s  season=%s",
        brand, brand_code, season,
    )

    # Load MDD
    mdd      = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[AsicsApparelAcc] No MDD file found in %s — LOV lookups disabled", MDD_DIR)

    # Load Brand Mapping
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    if not mapping_file:
        raise FileNotFoundError(f"Brand mapping file not found in {ATTR_DIR}")
    
    brand_mapping = BrandMappingLoader(mapping_file, sheet_name="Asics(Inline)")
    
    if not brand_mapping.get_mappings():
        log.warning("[AsicsApparelAcc] No attribute mappings loaded from 'Asics(Inline)' tab — continuing with hardcoded attributes only")

    # Load Asics MD Mappings for Sports Category
    asics_md_mappings = AsicsMDMappingsLoader(mapping_file, sheet_name="Asics MD Mappings")
    log.info("[AsicsApparelAcc] Asics MD Mappings loaded: %d APP/ACC entries", len(asics_md_mappings.mappings))

    # Load RNA for brand attributes
    rna = RNALoader(mapping_file)
    log.info("[AsicsApparelAcc] RNA loaded: %d entries", len(rna.lookup))

    # Load input file
    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        raise FileNotFoundError(f"No Asics apparel & accessories file found in {INPUT_DIR}")

    # ── Extract brand name from filename ──────────────────────────
    # Pattern: 0999-SP-ASICS-Apparel-SP2027-ID-1.xlsx
    #   parts[0]  = comp code   (e.g. '0999')
    #   parts[1]  = SBU         (e.g. 'SP')
    #   parts[2]  = brand name  (e.g. 'ASICS')
    #   parts[-2] = country id  (e.g. 'ID')
    filename_stem = input_file.stem
    brand_from_file = brand  # default fallback
    filename_comp_code = comp_code   # fallback to args
    filename_sbu       = sbu         # fallback to args
    filename_country_id = ""         # will be resolved below

    parts = filename_stem.split("-")
    if len(parts) >= 3:
        brand_token = parts[2].strip().upper()
        if brand_token:
            brand_from_file = brand_token.title()  # e.g. 'Asics'
            log.info(
                "[AsicsApparelAcc] Brand extracted from filename: '%s' → '%s'",
                brand_token, brand_from_file,
            )
    if len(parts) >= 1 and parts[0].strip():
        raw_comp = parts[0].strip().lstrip("0") or "0"  # '0999' → '999'
        filename_comp_code = raw_comp
        log.info("[AsicsApparelAcc] Comp code from filename: '%s' → '%s'", parts[0].strip(), filename_comp_code)
    if len(parts) >= 2 and parts[1].strip():
        filename_sbu = parts[1].strip().upper()   # e.g. 'SP'
        log.info("[AsicsApparelAcc] SBU from filename: '%s'", filename_sbu)
    if len(parts) >= 2:
        # country id is the 2nd-to-last dash token (e.g. '...SP2027-ID-1')
        candidate = parts[-2].strip().upper()
        if re.match(r'^[A-Z]{2}$', candidate):   # 2-letter ISO code
            filename_country_id = candidate
            log.info("[AsicsApparelAcc] Country ID from filename: '%s'", filename_country_id)
    
    # ── Look up Brand LOV ID from MDD ────────────────────────────
    brand_lov_id = brand_code  # fallback to original brand_code if not found
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            brand_key    = brand_from_file.upper()   # e.g. "ASICS"
            looked_up    = brand_lov.get(brand_key)  # e.g. "ASC"
            brand_lov_id = looked_up if looked_up else brand_code
            log.info(
                "[AsicsApparelAcc] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_key, brand_lov_id,
                "(FOUND)" if looked_up else f"(NOT FOUND — fallback to '{brand_code}')",
            )
            if not looked_up:
                log.warning(
                    "[AsicsApparelAcc] '%s' not in Brand LOV. "
                    "Available keys: %s",
                    brand_key, list(brand_lov.keys()),
                )
        else:
            log.warning("[AsicsApparelAcc] Brand LOV not found in MDD — using fallback brand_code='%s'", brand_code)
    else:
        log.warning("[AsicsApparelAcc] No MDD loaded — brand_lov_id defaulting to '%s'", brand_code)
    
    # ── Resolve country name from country ID via MDD ─────────────
    rna_country = ""
    if filename_country_id and mdd:
        rna_country = mdd.country_id_to_name(filename_country_id)
        log.info(
            "[AsicsApparelAcc] Country: ID='%s' → name='%s' (for RNA lookup)",
            filename_country_id, rna_country,
        )
        if not rna_country:
            log.warning("[AsicsApparelAcc] Country ID '%s' not found in Country LOV", filename_country_id)
    if not rna_country:
        rna_country = country_code   # fallback to args

    # ── RNA lookup: brand_type, brand_status, brand_category, brand_group ─────
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

    loader = AsicsApparelAccessoriesLoader(input_file)

    if loader.df.empty:
        log.warning("[AsicsApparelAcc] No data rows found — nothing to process")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_rows_dynamic(raw_rows, loader, brand_lov_id, brand_mapping, mdd)

    if not generics:
        log.warning("[AsicsApparelAcc] No valid generics produced")
        return

    # Build season classification ID
    sp_code  = season[:2].upper() if len(season) >= 2 else season
    sys_part = season[2:] if len(season) > 2 else ""
    sy       = f"20{sys_part}" if len(sys_part) == 2 else sys_part
    season_id = f"CLH_{brand_lov_id}_{sp_code}{sy}"

    # Generate XML
    xml_filename = f"{input_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    # Build Classifications block
    # brand_from_file = display name (e.g. 'Asics') → used in CLH_{brand}Batches, CLH_{brand}Articles
    # brand_lov_id    = MDD LOV ID   (e.g. 'ASC')   → used in CLH_{brand_code}_{season} season ID
    cls_el  = build_classifications(brand_from_file, brand_lov_id, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))

    log.info("[AsicsApparelAcc] Writing XML → %s", xml_filename)
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

        for generic in generics.values():
            px = build_product_xml_dynamic(
                generic, brand_from_file, brand_lov_id, comp_code, sbu,
                season, season_id, country_code,
                brand_type=brand_type,
                brand_status=brand_status,
                brand_category=brand_category,
                brand_group=brand_group,
                mdd=mdd,
                asics_md_mappings=asics_md_mappings,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info(
        "[AsicsApparelAcc] XML written: %s  (%d KB)",
        xml_path.name, file_kb,
    )
    print(
        f"=== ASICS APPAREL & ACCESSORIES SUMMARY ===\n"
        f"  Input rows : {len(raw_rows)}\n"
        f"  Generics   : {generic_count}\n"
        f"  Variants   : {variant_count}\n"
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
    parser = argparse.ArgumentParser(description="Asics Apparel & Accessories → Stibo XML")
    parser.add_argument("--brand",        default=BRAND_NAME)
    parser.add_argument("--brand_code",   default=BRAND_CODE)
    parser.add_argument("--season",       default="SS27")
    parser.add_argument("--country_code", default="")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    run(args)
