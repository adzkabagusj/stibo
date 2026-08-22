"""
╔══════════════════════════════════════════════════════════════════╗
║   CROCS SEA — DYNAMIC LINELIST GENERATOR + STIBO XML v2.0       ║
║  Reads brand mapping file dynamically for attribute mappings     ║
║  and MDD file for LOV lookups                                    ║
╚══════════════════════════════════════════════════════════════════╝

Lambda entry point: run(args, auditor)

Key Changes from v1:
- Dynamically loads attribute mappings from "NEW - Brand mapping files Template.xlsx"
- Reads "Crocs(Inline)" sheet for attribute definitions
- Fetches LOV IDs from MDD file sheets dynamically
- Supports both "Direct from Principal" and "Formula in System" with column mapping
"""

from __future__ import annotations

import logging
import math
import os
import re
from datetime import datetime
from pathlib import Path
from xml.dom import minidom
from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace

import openpyxl
import pandas as pd

log = logging.getLogger(__name__)

# ── Paths (Lambda-aware) ─────────────────────────────────────────
BASE_DIR    = Path(os.environ.get("LAMBDA_TMP_DIR", "/tmp/stibo_workdir_crocs"))
OUT_DIR     = BASE_DIR / "output" / "linelists"
XML_OUT_DIR = BASE_DIR / "output" / "xml"
INPUT_DIR   = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR     = INPUT_DIR / "mdd"
ATTR_DIR    = INPUT_DIR / "attributes"

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
NS           = f"{{{STIBO_NS}}}"

register_namespace("",    STIBO_NS)
register_namespace("xsi", STIBO_XSI)

# ── Brand constants ──────────────────────────────────────────────
CROCS_BRAND_CODE = "CCR"

# ── Country code → MDD COUNTRY name ──────────────────────────────
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
}

# ── Season LOV mapping ────────────────────────────────────────────
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter", "SP": "Spring",
    "SM": "Summer",
}

# ── Gender to Age mapping (for FTW sheet) ────────────────────────
GENDER_TO_AGE = {
    "MEN":           "Adults",
    "WOMEN":         "Adults",
    "UNISEX ADULTS": "Adults",
    "UNISEX ADULT":  "Adults",
    "UNISEX KIDS":   "Children",
}

# ── Gender to LOV mapping (for FTW sheet) ────────────────────────
GENDER_TO_LOV = {
    "MEN":           "Male",
    "WOMEN":         "Female",
    "UNISEX ADULTS": "Unisex",
    "UNISEX ADULT":  "Unisex",
    "UNISEX KIDS":   "Unisex",
}

# ── Attributes to exclude from XML (even if in brand mapping) ─────
EXCLUDED_ATTRIBUTES = {
    "AT_SAPStyleCode",  # Not needed in XML output
    "AT_Generic",
    "AT_Material",      # Use AT_MaterialType instead (hardcoded to ZINA)
    "AT_Gender",        # Handled by system-level logic with custom mapping
    "AT_SAPAge",        # Handled by system-level logic with custom mapping
    "AT_BYGender",      # Handled by system-level logic with custom mapping
    "AT_BYAge",         # Handled by system-level logic with custom mapping
    "AT_PackDetails",   # Handled by system-level logic with strict LOV lookup
    "AT_PrincipalBarcode",  # Not needed in XML output
    "AT_FOB",           # Calculated by formula (SEA MSRP with discount based on Reporting Silhouette)
}


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — DYNAMIC LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """
    Loads Core Attributes + LOVs from the MDD Excel.
    Dynamically reads all LOV sheets for flexible lookup.
    """

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb  = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        # Load Core Attributes
        if "Core Attributes" in wb.sheetnames:
            ws  = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 1:
                hdr  = rows[1]
                col: dict[str, int] = {}
                for i, h in enumerate(hdr):
                    if h:
                        col[str(h).split("\n")[0].strip()] = i
                        col[str(h).strip()] = i

                for row in rows[2:]:
                    aid = row[7] if len(row) > 7 else None
                    if not aid:
                        continue
                    aid = str(aid).strip()

                    def _v(key):
                        idx = col.get(key)
                        if idx is None or idx >= len(row):
                            return None
                        v = row[idx]
                        return str(v).strip() if v else None

                    self.attributes[aid] = {
                        "id":           aid,
                        "name":         _v("PIM Attribute Name"),
                        "source_name":  _v("Source Attribute Name"),
                        "validation":   _v("Validation Base Type"),
                        "lov_name":     _v("Name of LOV"),
                    }

        # Load all LOV sheets
        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        self._load_age_lov(wb)
        self._load_gender_lov(wb)
        self._load_size_lov(wb)
        self._load_color_lov(wb)
        self._load_country_lov(wb)
        self._load_country_origin_lov(wb)
        self._load_brand_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_silhouette_lov(wb)
        self._load_material_upper_lov(wb)
        self._load_material_lov(wb)
        self._load_pack_details_lov(wb)
        
        wb.close()
        log.info("[MDD] %d attributes | %d LOV types loaded", len(self.attributes), len(self.lovs))

    def _load_simple_lovs(self, wb):
        """Load Simple LOVs sheet."""
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = str(val_id).strip() if val_id else str(val_name).strip()

    def _load_age_lov(self, wb):
        """Load Age LOV sheet: 
        - AT_SAPAge: col A = display value, col B = LOV ID
        - AT_BYAge: col G = display value, col F = LOV ID
        """
        sheet_name = next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return

        # Load AT_SAPAge from columns A and B
        sap_age_lov_map = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display = str(row[0]).strip()
            age_lov_id  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            
            if age_display and age_lov_id:
                sap_age_lov_map[age_display.upper()] = age_lov_id

        # Load AT_BYAge from columns G and F
        by_age_lov_map = {}
        for row in rows[1:]:
            if not row or len(row) < 7:
                continue
            age_display = str(row[6]).strip() if row[6] else ""  # Column G (index 6)
            age_lov_id  = str(row[5]).strip() if row[5] else ""  # Column F (index 5)
            
            if age_display and age_lov_id:
                # Store the display value
                by_age_lov_map[age_display.upper()] = age_lov_id
                
                # Also map both "ADULTS" and "ADULT" to the same LOV ID
                if age_display.upper() == "ADULTS":
                    by_age_lov_map["ADULT"] = age_lov_id
                elif age_display.upper() == "ADULT":
                    by_age_lov_map["ADULTS"] = age_lov_id

        self.lovs["AgeLOV"] = sap_age_lov_map  # Default to SAPAge mapping
        self.lovs["AT_SAPAge"] = sap_age_lov_map
        self.lovs["AT_BYAge"] = by_age_lov_map
        log.info("[MDD] Age LOV loaded — SAP: %d entries, BY: %d entries", len(sap_age_lov_map), len(by_age_lov_map))

    def _load_gender_lov(self, wb):
        """Load Gender LOV sheet: col A = display value, col C = LOV ID."""
        sheet_name = next((s for s in wb.sheetnames if "GENDER" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        gender_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            display = str(row[0]).strip()
            lov_id  = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            
            if display and lov_id:
                gender_lov[display.upper()] = lov_id

        self.lovs["GenderLOV"] = gender_lov
        self.lovs["AT_Gender"] = gender_lov
        self.lovs["AT_BYGender"] = gender_lov
        log.info("[MDD] Gender LOV loaded — %d entries", len(gender_lov))

    def _load_size_lov(self, wb):
        """Load Size Code LOV sheet."""
        sheet_name = next((s for s in wb.sheetnames if "SIZE" in s.upper() and "CODE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        size_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            # Usually col A = display, col B = LOV ID
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            
            if display and lov_id:
                size_lov[display.upper()] = lov_id

        self.lovs["SizeCodeLOV"] = size_lov
        self.lovs["AT_Size"] = size_lov
        log.info("[MDD] Size Code LOV loaded — %d entries", len(size_lov))

    def _load_color_lov(self, wb):
        """Load Color Code LOV sheet."""
        sheet_name = next((s for s in wb.sheetnames if "COLOR" in s.upper() and "CODE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        color_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            
            if display and lov_id:
                color_lov[display.upper()] = lov_id

        self.lovs["ColorCodeLOV"] = color_lov
        self.lovs["AT_Color"] = color_lov
        log.info("[MDD] Color Code LOV loaded — %d entries", len(color_lov))

    def _load_country_lov(self, wb):
        """Load Country LOV sheet: Col A = LOV ID, Col B = display value."""
        sheet_name = next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper() and "ORIGIN" not in s.upper()), None)
        if not sheet_name:
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        country_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            
            if display and lov_id:
                country_lov[display.upper()] = lov_id

        self.lovs["CountryLOV"] = country_lov
        self.lovs["AT_Country"] = country_lov
        log.info("[MDD] Country LOV loaded — %d entries", len(country_lov))

    def _load_country_origin_lov(self, wb):
        """Load Country Origin LOV sheet: Col A = LOV ID, Col B = LOV value."""
        sheet_name = next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "ORIGIN" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Country Origin LOV sheet not found")
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        country_origin_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id    = str(row[0]).strip() if row[0] else ""
            lov_value = str(row[1]).strip() if row[1] else ""
            
            if lov_value and lov_id:
                country_origin_lov[lov_value.upper()] = lov_id

        self.lovs["CountryOriginLOV"] = country_origin_lov
        self.lovs["AT_CountryOrigin"] = country_origin_lov
        log.info("[MDD] Country Origin LOV loaded — %d entries", len(country_origin_lov))

    def _load_brand_lov(self, wb):
        """Load Brand LOV sheet."""
        sheet_name = next((s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        brand_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            
            if display and lov_id:
                brand_lov[display.upper()] = lov_id

        self.lovs["BrandLOV"] = brand_lov
        self.lovs["AT_Brand"] = brand_lov
        log.info("[MDD] Brand LOV loaded — %d entries", len(brand_lov))

    def _load_brand_type_lov(self, wb):
        """Load Brand Type LOV sheet: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Brand Type LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        brand_type_lov: dict[str, str] = {}  # display.upper() -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""

            if display and lov_id:
                brand_type_lov[display.upper()] = lov_id

        self.lovs["BrandTypeLOV"] = brand_type_lov
        self.lovs["AT_BrandType"] = brand_type_lov
        log.info("[MDD] Brand Type LOV loaded — %d entries", len(brand_type_lov))

    def _load_silhouette_lov(self, wb):
        """Load Silhouette LOV sheet: Col A = LOV ID (Code), Col B = display value (Silhouette)."""
        sheet_name = next((s for s in wb.sheetnames if "SILHOUETTE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Silhouette LOV sheet not found")
            return
            
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        silhouette_lov: dict[str, str] = {}
        
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            # Col A = LOV ID (Code), Col B = display value (Silhouette)
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            
            if display and lov_id:
                silhouette_lov[display.upper()] = lov_id

        self.lovs["SilhouetteLOV"] = silhouette_lov
        self.lovs["AT_Silhouette"] = silhouette_lov
        log.info("[MDD] Silhouette LOV loaded — %d entries", len(silhouette_lov))

    def _load_material_upper_lov(self, wb):
        """Load Material -Upper LOV sheet: Col A = LOV ID (Code), Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if "MATERIAL" in s.upper() and "UPPER" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Material Upper LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        material_upper_lov: dict[str, str] = {}  # display.upper() -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            # Col A = LOV ID, Col B = display value
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""

            if display and lov_id:
                material_upper_lov[display.upper()] = lov_id

        self.lovs["MaterialUpperLOV"] = material_upper_lov
        self.lovs["AT_MaterialUpper"] = material_upper_lov
        log.info("[MDD] Material Upper LOV loaded — %d entries", len(material_upper_lov))

    def _load_material_lov(self, wb):
        """Load Material LOV sheet: Col A = LOV ID (Code), Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if "MATERIAL" in s.upper() and "LOV" in s.upper() and "UPPER" not in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Material LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        material_lov: dict[str, str] = {}  # display.upper() -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            # Col A = LOV ID, Col B = display value
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""

            if display and lov_id:
                material_lov[display.upper()] = lov_id

        self.lovs["MaterialLOV"] = material_lov
        self.lovs["AT_Material"] = material_lov
        log.info("[MDD] Material LOV loaded — %d entries", len(material_lov))

    def _load_pack_details_lov(self, wb):
        """Load Pack Details LOV sheet: Col A = LOV ID (Code), Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if "PACK" in s.upper() and "DETAILS" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Pack Details LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        pack_details_lov: dict[str, str] = {}  # display.upper() -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            # Col A = LOV ID, Col B = display value
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""

            if display and lov_id:
                pack_details_lov[display.upper()] = lov_id

        self.lovs["PackDetailsLOV"] = pack_details_lov
        self.lovs["AT_PackDetails"] = pack_details_lov
        log.info("[MDD] Pack Details LOV loaded — %d entries", len(pack_details_lov))

    def _load_named_lov_sheets(self, wb):
        """Load any sheet with 'LOV' in name that hasn't been loaded yet."""
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            if any(kw in sn.upper() for kw in ["AGE", "GENDER", "SIZE CODE", "COLOR CODE", "COUNTRY", "BRAND", "SIMPLE", "SILHOUETTE", "MATERIAL"]):
                continue  # Already loaded
                
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
                
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            lov_map = {}
            
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                name = str(row[0]).strip() if row[0] else ""
                code = str(row[1]).strip() if row[1] else ""
                if name:
                    lov_map[name.upper()] = code if code else name

            if lov_map:
                self.lovs[display] = lov_map
                log.info("[MDD] %s loaded — %d entries", display, len(lov_map))

    def get_lov_id(self, attr_id: str, display_value: str) -> str:
        """
        Get LOV ID for a given attribute and display value.
        Searches in multiple LOV maps for flexibility.
        """
        if not display_value:
            return ""
        
        search_val = str(display_value).strip().upper()
        
        # Try attribute-specific LOV first
        if attr_id in self.lovs:
            if search_val in self.lovs[attr_id]:
                return self.lovs[attr_id][search_val]
        
        # Try based on attribute name patterns
        attr_info = self.attributes.get(attr_id, {})
        lov_name = attr_info.get("lov_name")
        
        if lov_name and lov_name in self.lovs:
            if search_val in self.lovs[lov_name]:
                return self.lovs[lov_name][search_val]
        
        # Fallback: search all LOV maps
        for lov_map in self.lovs.values():
            if search_val in lov_map:
                return lov_map[search_val]
        
        # If not found, return the original value
        return display_value


class BrandMappingLoader:
    """
    Loads the "NEW - Brand mapping files Template.xlsx" file.
    Reads the "Crocs(Inline)-NEW" sheet to get dynamic attribute mappings.
    
    FULLY DYNAMIC ATTRIBUTE MAPPING:
    ────────────────────────────────────────────────────────────────
    Column Layout:
    - A: Stibo Attribute Name (display name)
    - B: Stibo Attribute ID (e.g., AT_PrincipalStyleCode)
    - C: Stibo Validation (LOV, Text, Number, etc.)
    - G: Field Mapping Type
    - J: Field Name in the Brand File (Excel column name)
    - K: Mapping Logic (transformations like "last 3 digits")
    
    Inclusion Rules (Col G):
    - "Direct from Principal" → INCLUDE
    - "Formula in system" + has field name (Col J) → INCLUDE
    - "Mapping from Principal" + has field name (Col J) → INCLUDE
    - "Formula in system" without field name → SKIP (system-level only)
    
    All included attributes are processed automatically:
    - LOV types → sent as ID attributes
    - Text types → sent as value attributes
    - Mapping logic (Col K) applied automatically
    """

    def __init__(self, path: Path, sheet_name: str = "Crocs(Inline)-NEW"):
        self.path = path
        self.sheet_name = sheet_name
        self.mappings: list[dict] = []
        self._load()

    def _load(self):
        log.info("[BrandMapping] Loading from: %s, sheet: %s", self.path.name, self.sheet_name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        if self.sheet_name not in wb.sheetnames:
            log.error("[BrandMapping] Sheet '%s' not found", self.sheet_name)
            wb.close()
            return
        
        ws = wb[self.sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        
        # Row 2 (index 1) is the header
        # Data starts from row 3 (index 2)
        for row_idx in range(2, len(rows)):
            row = rows[row_idx]
            if not row or len(row) < 10:
                continue
            
            col_a = row[0]  # Stibo Attribute Name
            col_b = row[1]  # Stibo Attribute ID
            col_c = row[2]  # Stibo Validation
            col_g = row[6]  # Field Mapping Type
            col_j = row[9]  # Field Name in the Brand File
            
            # Skip if no attribute ID
            if not col_b or str(col_b).startswith("#N/A"):
                continue
            
            col_g_str = str(col_g).strip() if col_g else ""
            col_j_str = str(col_j).strip() if col_j else ""
            col_k_str = str(row[10]).strip() if len(row) > 10 and row[10] else ""  # Mapping Logic
            
            col_g_lower = col_g_str.lower()
            
            # Include if:
            # 1. "Direct from Principal" OR
            # 2. "Formula in System" AND has column value in J OR
            # 3. "Mapping from Principal" / "Mapping from the principal data" AND has column value in J
            include = False
            if "direct from principal" in col_g_lower:
                include = True
            elif "formula in system" in col_g_lower and col_j_str:
                include = True
            elif "mapping from" in col_g_lower and "principal" in col_g_lower and col_j_str:
                include = True
            
            if include:
                # For fields like "FW & Acc : Colorway Code", extract the actual field name
                # Handle multi-line cells (e.g., "FW & Acc : Colorway\nCharms: N/A")
                # Create separate mappings for each sheet type
                actual_field = col_j_str
                
                # Check if multi-line (contains newline)
                if '\n' in col_j_str:
                    lines = col_j_str.split('\n')
                    # Process each line to create multiple mappings
                    for line in lines:
                        line = line.strip()
                        if not line or line.upper() == 'N/A':
                            continue
                        
                        # Extract field name after colon
                        if ':' in line:
                            parts = line.split(':', 1)
                            if len(parts) > 1:
                                field_name = parts[1].strip()
                                # Skip if field is N/A
                                if field_name.upper() == 'N/A':
                                    continue
                                
                                mapping = {
                                    'attribute_name': str(col_a).strip() if col_a else "",
                                    'attribute_id': str(col_b).strip(),
                                    'value_type': str(col_c).strip().upper() if col_c else "TEXT",
                                    'mapping_type': col_g_str,
                                    'field_name': field_name,
                                    'mapping_logic': col_k_str,
                                    'sheet_hint': parts[0].strip(),  # e.g., "FW & Acc" or "Charms"
                                }
                                self.mappings.append(mapping)
                else:
                    # Single line - original logic
                    if ':' in col_j_str:
                        # Extract field name after colon (e.g., "FW & Acc : Colorway Code" -> "Colorway Code")
                        parts = col_j_str.split(':', 1)
                        if len(parts) > 1:
                            actual_field = parts[1].strip()
                    
                    mapping = {
                        'attribute_name': str(col_a).strip() if col_a else "",
                        'attribute_id': str(col_b).strip(),
                        'value_type': str(col_c).strip().upper() if col_c else "TEXT",
                        'mapping_type': col_g_str,
                        'field_name': actual_field,
                        'mapping_logic': col_k_str,
                    }
                    self.mappings.append(mapping)
        
        wb.close()
        log.info("[BrandMapping] Loaded %d attribute mappings", len(self.mappings))
        
        # Log first few mappings for verification
        for i, m in enumerate(self.mappings[:5]):
            log.info("  [%d] %s ← %s (Type: %s)", i+1, m['attribute_id'], m['field_name'], m['value_type'])

    def get_mapping_for_field(self, field_name: str) -> dict | None:
        """Get mapping info for a specific brand file field name."""
        for mapping in self.mappings:
            if mapping['field_name'].upper() == field_name.upper():
                return mapping
        return None

    def get_all_mappings(self) -> list[dict]:
        """Get all loaded mappings."""
        return self.mappings


class RNALoader:
    """
    Loads RNA (Source Mapping) data to resolve Brand Type and Brand Category.
    Reads from "Source Mapping related RNA" sheet in the brand mapping file.
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
        # Strip .0 suffix (e.g. "888.0" → "888")
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


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — LINELIST LOADER
# ══════════════════════════════════════════════════════════════════

class CrocsLinelistLoader:
    """
    Loads the Crocs linelist Excel file.
    Reads ALL sheets, auto-detecting the header row per sheet
    (first row containing 'Style Code' or 'Colorway Code').
    """

    # Anchor columns used to detect the header row
    HEADER_ANCHORS = {"style code", "colorway code", "style", "colorway"}

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheets_loaded = []
        self._load()

    def _find_header_row(self, ws) -> int | None:
        """Return 0-based row index of the header row, or None if not found.
        Requires at least 2 anchor columns to be present to avoid false positives.
        """
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            cells = [str(v).strip().lower() for v in row if v and str(v).strip()]
            matches = sum(1 for anchor in self.HEADER_ANCHORS if anchor in cells)
            if matches >= 2:
                return i
        return None

    def _load(self):
        log.info("[Linelist] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        log.info("[Linelist] Found %d sheets: %s", len(wb.sheetnames), wb.sheetnames)

        all_dfs = []

        for sheet_name in wb.sheetnames:
            log.info("[Linelist] Reading sheet: '%s'", sheet_name)
            try:
                ws = wb[sheet_name]
                hdr_idx = self._find_header_row(ws)

                if hdr_idx is None:
                    log.warning("[Linelist] Sheet '%s': no header row found — skipping", sheet_name)
                    continue

                log.info("[Linelist] Sheet '%s': header at row %d (0-indexed)", sheet_name, hdr_idx)

                # Read with correct header row
                df_sheet = pd.read_excel(
                    self.path,
                    sheet_name=sheet_name,
                    header=hdr_idx,
                )

                # Drop rows where ALL non-image columns are empty
                df_sheet.dropna(how="all", inplace=True)

                # Drop columns that are entirely unnamed (Unnamed: x)
                # Coerce column names to string first to handle NaN column names
                df_sheet.columns = [str(c) if not isinstance(c, str) else c for c in df_sheet.columns]
                df_sheet = df_sheet.loc[:, ~df_sheet.columns.str.startswith("Unnamed:")]

                if df_sheet.empty:
                    log.warning("[Linelist] Sheet '%s' has no data rows after header", sheet_name)
                    continue

                df_sheet["_source_sheet"] = sheet_name
                all_dfs.append(df_sheet)
                self.sheets_loaded.append(sheet_name)
                log.info("[Linelist] Sheet '%s': %d rows loaded, columns: %s",
                         sheet_name, len(df_sheet), list(df_sheet.columns[:8]))

            except Exception as e:
                log.error("[Linelist] Failed to read sheet '%s': %s", sheet_name, e)

        wb.close()

        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            log.info("[Linelist] Combined total: %d rows from %d sheets",
                     len(self.df), len(self.sheets_loaded))
        else:
            log.warning("[Linelist] No data loaded from any sheet")


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — XML HELPERS  (mirrors Lotto recap pattern)
# ══════════════════════════════════════════════════════════════════

from xml.etree import ElementTree as ET

# Strip duplicate xmlns injected by ET.tostring
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')
ET.register_namespace("", STIBO_NS)
ET.register_namespace("xsi", STIBO_XSI)


def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element | None:
    """Append <Value AttributeID="..."> to parent."""
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip()  if id_val  else ""
    clean_val = str(value).strip()   if value   else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    elif clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    """Append a <MultiValue> → <Value ID="..."> element."""
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _safe_str(v) -> str:
    """Convert any value to a clean string; skip NaN/None."""
    if v is None:
        return ""
    import math
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _get_field(row_data: dict, field_name: str) -> str:
    """
    Case-insensitive column lookup against row_data.
    Tries exact match first, then falls back to upper-cased comparison.
    """
    # Exact match
    v = row_data.get(field_name)
    if v is not None:
        return _safe_str(v)
    # Case-insensitive fallback
    field_upper = field_name.strip().upper()
    for k, val in row_data.items():
        if str(k).strip().upper() == field_upper:
            return _safe_str(val)
    return ""


def _apply_mapping_logic(value: str, logic: str) -> str:
    """
    Apply mapping logic transformations to a value.
    
    Supported logic patterns:
    - "last 3 digits" / "last 3" -> extract last 3 characters
    - More patterns can be added as needed
    """
    if not value or not logic:
        return value
    
    logic_lower = logic.lower()
    
    # Extract last N digits/characters
    if "last 3" in logic_lower or "last 3 digits" in logic_lower:
        return value[-3:] if len(value) >= 3 else value
    
    # Add more transformation patterns here as needed
    
    return value


def _extract_highest_percentage_material(raw_value: str) -> str:
    """
    Extract the material name with the highest percentage from a composition string.

    Examples:
        "EVA 40% + Poly 600D 60%"  → "Poly 600D"   (60% is highest)
        "Cotton 50% + Polyester 50%" → "Cotton"     (first match when tied)
        "Leather"                   → "Leather"     (no percentage, return as-is)
        "100% EVA"                  → "EVA"

    Returns:
        The material name with highest percentage, or the raw value if no
        percentage pattern found at all.
    """
    if not raw_value:
        return ""

    raw = raw_value.strip()

    # Pattern: find all (material_text, percentage) pairs
    # Handles: "EVA 40%", "Poly 600D 60%", "100% EVA"
    # Split on common delimiters: +, ,, ;, /
    segments = re.split(r"[+,;/]", raw)

    candidates: list[tuple[float, str]] = []  # (percentage, material_name)

    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue

        # Match percentage at end: "EVA 40%" or "Poly 600D 60%"
        m_end = re.search(r"^(.*?)\s*(\d+(?:\.\d+)?)\s*%\s*$", seg)
        if m_end:
            mat = m_end.group(1).strip()
            pct = float(m_end.group(2))
            if mat:
                candidates.append((pct, mat))
            continue

        # Match percentage at start: "40% EVA" or "60% Poly 600D"
        m_start = re.search(r"^\s*(\d+(?:\.\d+)?)\s*%\s*(.+)$", seg)
        if m_start:
            pct = float(m_start.group(1))
            mat = m_start.group(2).strip()
            if mat:
                candidates.append((pct, mat))
            continue

    if not candidates:
        # No percentage found — return raw value as-is
        return raw

    # Return material with highest percentage (first wins on tie)
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _get_material_upper_lov_id(raw_upper: str, material_upper_lov: dict) -> str:
    """
    Resolve AT_MaterialUpper LOV ID from raw 'Upper' column value.

    Logic:
    1. Extract the material with the highest percentage.
    2. Search all LOV display values (Col B) for a match.
    3. If exactly 1 match → return its LOV ID (Col A).
    4. If 0 or >1 matches → return "" (do not send attribute).

    Example:
        raw_upper = "EVA 40% + Poly 600D 60%"
        → highest = "Poly 600D"
        → search LOV col B for "Poly 600D" → not found
        → search for partial matches → only 1 match found → return LOV ID
    """
    if not raw_upper or not material_upper_lov:
        return ""

    highest_material = _extract_highest_percentage_material(raw_upper)
    if not highest_material:
        return ""

    search_key = highest_material.strip().upper()

    # 1. Exact match first
    if search_key in material_upper_lov:
        return material_upper_lov[search_key]

    # 2. Partial match: find all LOV entries where col B contains search_key
    #    or search_key contains col B
    matches: list[str] = []  # collect matching lov_ids
    for display_upper, lov_id in material_upper_lov.items():
        if display_upper in search_key or search_key in display_upper:
            matches.append(lov_id)

    if len(matches) == 1:
        return matches[0]

    # 0 matches or >1 matches → do not send
    if len(matches) > 1:
        log.debug(
            "[MaterialUpper] Multiple LOV matches for '%s' → skipping AT_MaterialUpper",
            highest_material,
        )
    return ""


def _resolve_currency_from_country(country_code: str) -> str:
    """
    Resolve Retail Price Currency LOV ID from country code.
    Based on the country-to-currency mapping.
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


def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    """
    Build the <Classifications> block — mirrors Lotto recap exactly.
      CLH_{brand_code}_{season}          CLS_Season         parent=CLH_{brand}Batches
        CLH_{brand_code}_{season}CA      CLS_ConfirmedArticles
        CLH_{brand_code}_{season}UA      CLS_UnconfirmedArticles
    """
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_tail   = season_code[2:]
    sea_year   = f"20{sea_tail}" if len(sea_tail) == 2 else sea_tail

    full_season   = f"{sea_prefix}{sea_year}"
    season_id     = f"CLH_{brand_code}_{full_season}"
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


def _determine_parent_id(row_data: dict, source_sheet: str) -> str:
    """
    Determine the PPH ParentID based on sheet name / product type.
      FTW sheet  → PPH_F-TempSubCat
      Charms     → PPH_E-TempSubCat   (Equipment)
      ACC        → PPH_E-TempSubCat   (Equipment)
    """
    sheet = (source_sheet or "").lower()
    if "ftw" in sheet or "footwear" in sheet:
        return "PPH_F-TempSubCat"
    if "charm" in sheet:
        return "PPH_E-TempSubCat"
    if "acc" in sheet:
        return "PPH_E-TempSubCat"
    return "PPH_F-TempSubCat"  # default


def build_product_xml(
    row_data:      dict,
    brand:         str,
    brand_code:    str,
    comp_code:     str,
    sbu:           str,
    season_id:     str,
    season:        str,
    country_code:  str,
    brand_type:    str,
    brand_mapping: BrandMappingLoader,
    mdd:           MDDLoader,
) -> str:
    """
    Build one <Product> XML string — mirrors Lotto recap build_product_xml:
      • UserTypeID = PRD_GenericArticle
      • ParentID   = PPH_{div}-TempSubCat
      • KeyValue   KEY_InboundArticle
      • Name
      • ClassificationReference  ×2  (Merchandiser + UnConfirmedForSeason)
      • Values  (dynamic from brand mapping + brand_code/comp/sbu/season)
    """
    # ── Resolve identifying fields ────────────────────────────────
    colorway_code = _safe_str(row_data.get("Colorway Code")) or _safe_str(row_data.get("Style Code", ""))
    style         = _safe_str(row_data.get("Style", ""))
    colorway      = _safe_str(row_data.get("Colorway", ""))
    source_sheet  = _safe_str(row_data.get("_source_sheet", ""))

    if not colorway_code:
        return ""   # skip rows with no identifying code

    # Generic code: brand_code(3) + colorway_code stripped of non-alnum, max 12
    alnum_code  = re.sub(r"[^A-Za-z0-9]", "", colorway_code)[:9].upper()
    generic_code = f"{brand_code[:3].upper()}{alnum_code}"

    parent_id = _determine_parent_id(row_data, source_sheet)

    # ── Element construction ──────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    # FTW → PRD_GenericArticle, ACC & Charms → PRD_SingleArticle
    if "ftw" in source_sheet.lower() or "footwear" in source_sheet.lower():
        g_el.set("UserTypeID", "PRD_GenericArticle")
    else:
        g_el.set("UserTypeID", "PRD_SingleArticle")
    g_el.set("ParentID",   parent_id)

    # KeyValue (like Lotto)
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    # Name
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = style or colorway_code

    # ClassificationReferences (like Lotto)
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # Add generic_code to row_data for use in attributes
    row_data["_generic_code"] = generic_code

    # Values
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_product_values(vals_el, row_data, brand, brand_code, comp_code, sbu, season, country_code, brand_type, brand_mapping, mdd)

    return ET.tostring(g_el, encoding="unicode")


def _add_product_values(
    vals_el:       ET.Element,
    row_data:      dict,
    brand:         str,
    brand_code:    str,
    comp_code:     str,
    sbu:           str,
    season:        str,
    country_code:  str,
    brand_type:    str,
    brand_mapping: BrandMappingLoader,
    mdd:           MDDLoader,
) -> None:
    """
    Write all <Value> elements — fully dynamic from brand mapping file.
    
    DYNAMIC MAPPING LOGIC (from Brand Mapping Excel file):
    ─────────────────────────────────────────────────────────────
    Col A: Stibo Attribute Name (for reference)
    Col B: Stibo Attribute ID (e.g., AT_PrincipalStyleCode)
    Col C: Validation Type (LOV or Text)
           - LOV  → send as ID attribute (<Value AttributeID="..." ID="..."/>)
           - Text → send as value (<Value AttributeID="...">value</Value>)
    Col G: Mapping Type (determines inclusion)
           - "Direct from Principal" → include
           - "Formula in system" with field name in Col J → include
           - "Mapping from Principal" with field name in Col J → include
           - Otherwise → skip (handled separately if needed)
    Col J: Field Name in Excel (e.g., "Colorway Code", "Style")
           - May contain conditional formats like "FW & Acc : Colorway Code"
           - System extracts the actual field name after ":"
    Col K: Mapping Logic (transformations)
           - "last 3 digits" → extract last 3 characters
           - More patterns can be added to _apply_mapping_logic()
    
    PROCESSING ORDER:
    1. Dynamic attributes from brand mapping (with field names)
    2. System-level attributes (Brand, Season, Company, etc.)
    """
    written: set[str] = set()

    def _w(attr_id: str, value: str = "", id_val: str = ""):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val)
            written.add(attr_id)

    def _mw(attr_id: str, id_val: str):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── 1. Dynamic attributes from brand mapping ──────────────────
    source_sheet = _safe_str(row_data.get("_source_sheet", "")).lower()
    
    for mapping in brand_mapping.get_all_mappings():
        attr_id      = mapping["attribute_id"]
        field_name   = mapping["field_name"]
        value_type   = mapping["value_type"]
        mapping_logic = mapping.get("mapping_logic", "")

        # Skip excluded attributes
        if attr_id in EXCLUDED_ATTRIBUTES:
            continue

        if not field_name or attr_id in written:
            continue

        # ── Special handling for AT_PrincipalStyleCode based on sheet type ──
        if attr_id == "AT_PrincipalStyleCode":
            if "charm" in source_sheet:
                # Charms sheet → use "Style Code" column
                field_name = "Style Code"
            else:
                # FW & ACC sheets → use "Colorway Code" column
                field_name = "Colorway Code"

        # ── Special handling for AT_PrincipalMerchandiseHierarchyL2 based on sheet type ──
        # Brand mapping file Col J has a misspelling ("Sillhoute") — override explicitly
        if attr_id == "AT_PrincipalMerchandiseHierarchyL2":
            if "charm" in source_sheet:
                field_name = "Reporting Silhouette"  # Charms column
            else:
                field_name = "Silhouette"             # FTW & ACC correct column name

        # ── Special handling for AT_Silhouette ──
        # Always use "Silhouette" column from input excel (all sheets)
        # If column not present, skip this attribute
        if attr_id == "AT_Silhouette":
            field_name = "Silhouette"

        # ── Special handling for AT_MaterialUpper ──
        # Extract highest-percentage material from "Upper" column, look up in LOV.
        # Send only if exactly 1 LOV match found.
        if attr_id == "AT_MaterialUpper":
            raw_upper = _get_field(row_data, "Upper")
            if not raw_upper:
                continue
            material_upper_lov = mdd.lovs.get("AT_MaterialUpper", {})
            lov_id = _get_material_upper_lov_id(raw_upper, material_upper_lov)
            if lov_id:
                _w("AT_MaterialUpper", id_val=lov_id)
            continue  # handled — skip the generic LOV path below

        # ── Special handling for AT_PrincipalGenderDescription ──
        # Send only for FTW and ACC sheets, map to Gender column
        if attr_id == "AT_PrincipalGenderDescription":
            if "charm" in source_sheet:
                # Skip for Charms sheet
                continue
            # Map to Gender column
            field_name = "Gender"
            gender_value = _get_field(row_data, field_name)
            if gender_value:
                _w("AT_PrincipalGenderDescription", value=gender_value)
            continue  # handled — skip the generic path below

        # ── Special handling for AT_CountryOrigin ──
        # Send only for FTW and ACC sheets, map to COO column
        if attr_id == "AT_CountryOrigin":
            if "charm" in source_sheet:
                # Skip for Charms sheet
                continue
            # Map to COO column
            field_name = "COO"
            coo_value = _get_field(row_data, field_name)
            if coo_value:
                # Look up LOV ID from MDD
                lov_id = mdd.get_lov_id("AT_CountryOrigin", coo_value)
                if lov_id:
                    _w("AT_CountryOrigin", id_val=lov_id)
            continue  # handled — skip the generic path below

        # ── Special handling for AT_Style ──
        # Map to Name column, send as LOV value (text) not LOV ID
        if attr_id == "AT_Style":
            field_name = "Name"
            style_value = _get_field(row_data, field_name)
            if style_value:
                _w("AT_Style", value=style_value)
            continue  # handled — skip the generic path below

        # ── Skip AT_PrincipalMerchandiseHierarchyL3 for Charms sheet ──
        if attr_id == "AT_PrincipalMerchandiseHierarchyL3":
            if "charm" in source_sheet:
                # Skip this attribute for Charms sheet
                continue

        raw_value = _get_field(row_data, field_name)
        if not raw_value:
            continue

        # Apply mapping logic transformation if specified
        transformed_value = _apply_mapping_logic(raw_value, mapping_logic)

        if value_type == "LOV":
            # AT_Silhouette: strict lookup — only send if found in Silhouette LOV
            if attr_id == "AT_Silhouette":
                silhouette_lov = mdd.lovs.get("AT_Silhouette", {})
                lov_id = silhouette_lov.get(transformed_value.strip().upper(), "")
                if lov_id:
                    # Zero-pad single-digit IDs (e.g. "9" → "09")
                    if lov_id.isdigit() and len(lov_id) == 1:
                        lov_id = lov_id.zfill(2)
                    _w(attr_id, id_val=lov_id)
                # else: not found → do not send attribute
            else:
                lov_id = mdd.get_lov_id(attr_id, transformed_value)
                if lov_id:
                    _w(attr_id, id_val=lov_id)
        else:
            _w(attr_id, value=transformed_value)

    # ── 2. Fixed system-level values (not in brand file) ──────────
    # Brand
    _w("AT_Brand", id_val=brand_code)
    _w("AT_BrandGroup", id_val=brand.upper())

    # Company + SBU (using MultiValue)
    _mw("AT_CompanyCode", comp_code)
    _mw("AT_SBU",        sbu)

    # ── Generic Code ─────────────────────────────────────────────
    # Extract generic_code from row_data (should be passed from build_product_xml)
    generic_code = _safe_str(row_data.get("_generic_code", ""))
    if generic_code:
        _w("AT_InboundGenericCode", generic_code)
        # AT_Generic is excluded - managed by STIBO business rules

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = season
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    # Extract year from season (e.g., "FW26" -> "2026")
    year_match = re.search(r"20\d{2}", sea_raw)
    if not year_match:
        short_match = re.search(r"\d{2}$", sea_raw)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    _w("AT_SeasonYear", season_year)

    # ── Country (destination country) ────────────────────────────
    country_val = country_code.strip().upper() if country_code else ""
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)

    # ── Brand Type ───────────────────────────────────────────────
    # Resolved from RNA (Source Mapping) based on country/comp/sbu/brand
    # Look up LOV ID from MDD Brand Type LOV (col A = ID, col B = display)
    if brand_type:
        brand_type_lov = mdd.lovs.get("AT_BrandType", {})
        brand_type_lov_id = brand_type_lov.get(brand_type.strip().upper(), "")
        if brand_type_lov_id:
            _w("AT_BrandType", id_val=brand_type_lov_id)
        else:
            log.warning("[BrandType] No LOV ID found for '%s' — skipping AT_BrandType", brand_type)

    # ── Franchise ────────────────────────────────────────────────
    franchise = _safe_str(row_data.get("Franchise", "")).strip()
    if franchise:
        _w("AT_Franchise", value=franchise)

    # ── Material (AT_Material from Upper column) ─────────────────
    upper_value = _safe_str(row_data.get("Upper", "")).strip()
    if upper_value:
        material_lov = mdd.lovs.get("AT_Material", {})
        material_lov_id = material_lov.get(upper_value.upper(), "")
        if material_lov_id:
            _w("AT_Material", id_val=material_lov_id)
        else:
            log.warning("[Material] No LOV ID found for Upper='%s' — skipping AT_Material", upper_value)

    # ── Product Length Width Height ──────────────────────────────
    # Convert from mm to cm (e.g., "30*15*12.2" → "3*1.5*1.22")
    charm_dimensions = _safe_str(row_data.get("Charm Dimensions", "")).strip()
    if charm_dimensions:
        try:
            # Split by '*' delimiter
            dimensions = charm_dimensions.split('*')
            # Convert each dimension from mm to cm
            dimensions_cm = []
            for dim in dimensions:
                dim_mm = float(dim.strip())
                dim_cm = dim_mm / 10
                # Format to remove unnecessary trailing zeros
                dimensions_cm.append(str(dim_cm).rstrip('0').rstrip('.') if '.' in str(dim_cm) else str(dim_cm))
            # Reconstruct the string
            charm_dimensions_cm = '*'.join(dimensions_cm)
            _w("AT_ProductLengthWidthHeight", value=charm_dimensions_cm)
        except (ValueError, TypeError) as e:
            log.warning("[ProductLengthWidthHeight] Invalid Charm Dimensions value: '%s' — skipping conversion", charm_dimensions)

    # ── Pack Details (AT_PackDetails from Charms Packaging Type) ─
    charms_packaging_type = _safe_str(row_data.get("Charms Packaging Type", "")).strip()
    if charms_packaging_type:
        pack_details_lov = mdd.lovs.get("AT_PackDetails", {})
        pack_details_lov_id = pack_details_lov.get(charms_packaging_type.upper(), "")
        if pack_details_lov_id:
            _w("AT_PackDetails", id_val=pack_details_lov_id)
        # If LOV ID not found, skip attribute (no warning needed)

    # ── Gender (AT_Gender & AT_BYGender) ─────────────────────────
    # Different logic based on sheet type
    source_sheet = _safe_str(row_data.get("_source_sheet", "")).lower()
    
    if "acc" in source_sheet or "charm" in source_sheet:
        # Accessories & Charms → Default to Unisex
        _w("AT_Gender", id_val="U")
        _w("AT_BYGender", id_val="U")
    elif "ftw" in source_sheet or "footwear" in source_sheet:
        # FTW → Extract from Gender column and map to Gender LOV
        gender_value = _safe_str(row_data.get("Gender", "")).strip().upper()
        gender_display = GENDER_TO_LOV.get(gender_value, "")
        
        if gender_display:
            # Lookup LOV IDs from MDD
            gender_id = mdd.get_lov_id("AT_Gender", gender_display)
            by_gender_id = mdd.get_lov_id("AT_BYGender", gender_display)
            
            if gender_id:
                _w("AT_Gender", id_val=gender_id)
            if by_gender_id:
                _w("AT_BYGender", id_val=by_gender_id)

    # ── Age (AT_SAPAge & AT_BYAge) ───────────────────────────────
    # Different logic based on sheet type
    
    if "acc" in source_sheet or "charm" in source_sheet:
        # Accessories & Charms → Default to Adult
        # Use proper LOV IDs from MDD (Adults for SAPAge, ADULT for BYAge)
        adults_lov_id = mdd.get_lov_id("AT_SAPAge", "Adults")
        if adults_lov_id:
            _w("AT_SAPAge", id_val=adults_lov_id)
        _w("AT_BYAge", id_val="ADULT")
    elif "ftw" in source_sheet or "footwear" in source_sheet:
        # FTW → Extract from Gender column and map to Age
        gender_value = _safe_str(row_data.get("Gender", "")).strip().upper()
        age_display = GENDER_TO_AGE.get(gender_value, "")

        # AT_BYAge uses static LOV IDs: ADULT / CHILD — map directly
        _BY_AGE_LOV_ID = {
            "Adults":   "ADULT",
            "Children": "CHILD",
        }

        if age_display:
            # AT_SAPAge — lookup from MDD (e.g. "Adults" → LOV ID)
            sap_age_id = mdd.get_lov_id("AT_SAPAge", age_display)
            if sap_age_id:
                _w("AT_SAPAge", id_val=sap_age_id)

            # AT_BYAge — use direct static mapping to avoid unreliable LOV lookup
            by_age_id = _BY_AGE_LOV_ID.get(age_display, "")
            if by_age_id:
                _w("AT_BYAge", id_val=by_age_id)

    # ── Principal Size (AT_PrincipalSize) ────────────────────────
    # Only for FTW and ACC sheets, not Charms
    if "ftw" in source_sheet or "acc" in source_sheet:
        size_range = _safe_str(row_data.get("Size Range (Colorway)", "")).strip()
        if size_range:
            _w("AT_PrincipalSize", value=size_range)

    # ── Article category (based on sheet type) ──────────────────
    # FTW → Generic (1), ACC → Single (0), Charms → skip
    if "acc" in source_sheet:
        _w("AT_SAPArticleCategory", id_val="0")  # Single
    elif "ftw" in source_sheet or "footwear" in source_sheet:
        _w("AT_SAPArticleCategory", id_val="1")  # Generic
    # Skip for Charms sheet

    # Article type
    _w("AT_BYArticleType", "Inline", id_val="Inline")

    # SAP Product Flag
    _w("AT_SAPProductFlag", "A", id_val="A")

    # UOM
    _w("AT_UOM", "Each", id_val="EA")

    # ── Retail Price Currency (from country code) ───────────────
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _w("AT_RetailPriceCurrency", id_val=currency_lov_id)

    # ── FOB Currency (default USD) ───────────────────────────────
    _w("AT_FOBCurrency", id_val="USD")

    # ── FOB (calculated from SEA MSRP and Reporting Silhouette) ──
    sea_msrp_str = _safe_str(row_data.get("SEA MSRP", "")).strip()
    reporting_silhouette = _safe_str(row_data.get("Reporting Silhouette", "")).strip().upper()
    
    if sea_msrp_str:
        try:
            sea_msrp = float(sea_msrp_str)
            # Apply discount based on Reporting Silhouette
            if reporting_silhouette == "BAG":
                fob_value = sea_msrp * (1 - 0.57)  # 57% discount for Bags
            else:
                fob_value = sea_msrp * (1 - 0.645)  # 64.5% discount for others
            
            # Format to 2 decimal places
            _w("AT_FOB", value=f"{fob_value:.2f}")
        except (ValueError, TypeError):
            log.warning("[FOB] Invalid SEA MSRP value: '%s' — skipping AT_FOB", sea_msrp_str)

    # ── Material Type (default ZINA) ─────────────────────────────
    _w("AT_MaterialType", id_val="ZINA")

    # ── Pricing Distribution Channel (default 01) ────────────────
    _w("AT_PricingDistributionChannel", id_val="01")

    # ── Country Size (FTW only, default US) ──────────────────────
    if "ftw" in source_sheet or "footwear" in source_sheet:
        _w("AT_CountrySize", value="US")

    # ── Fastening (FTW only, default NL) ─────────────────────────
    if "ftw" in source_sheet or "footwear" in source_sheet:
        _w("AT_Fastening", id_val="NL")

    # ── EComProductNameEN (formula: Brand + Style + Gender + Silhouette + "-" + Colorway) ──
    style_val = _safe_str(row_data.get("Style", "")).strip()
    gender_val = _safe_str(row_data.get("Gender", "")).strip()
    colorway_val = _safe_str(row_data.get("Colorway", "")).strip()
    
    # Silhouette: use "Silhouette" for FTW/ACC, "Reporting Silhouette" for Charms
    if "charm" in source_sheet:
        silhouette_val = _safe_str(row_data.get("Reporting Silhouette", "")).strip()
    else:
        silhouette_val = _safe_str(row_data.get("Silhouette", "")).strip()
    
    # Build the product name
    if style_val or gender_val or silhouette_val or colorway_val:
        parts = [p for p in [brand, style_val, gender_val, silhouette_val] if p]
        name_prefix = " ".join(parts)
        
        if colorway_val:
            ecom_product_name = f"{name_prefix} - {colorway_val}"
        else:
            ecom_product_name = name_prefix
        
        if ecom_product_name.strip():
            _w("AT_EComProductNameEN", value=ecom_product_name)



# ══════════════════════════════════════════════════════════════════
# SECTION 4 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point for Crocs linelist processing.
    Mirrors Lotto recap run() structure exactly.
    """
    log.info("[Crocs Linelist] Starting ...")

    # ── Load files ───────────────────────────────────────────────
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if not mdd_file:
        raise FileNotFoundError(f"No MDD file found in {MDD_DIR}")

    brand_mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if not brand_mapping_file:
        raise FileNotFoundError(f"No Brand mapping file found in {ATTR_DIR}")

    linelist_file = next(LINELIST_DIR.glob("*.xlsx"), None)
    if not linelist_file:
        raise FileNotFoundError(f"No linelist file found in {LINELIST_DIR}")

    log.info("═══ SELECTED FILES ════════════════════════════════════════")
    log.info("  MDD           : %s", mdd_file.name)
    log.info("  Brand Mapping : %s", brand_mapping_file.name)
    log.info("  Linelist      : %s", linelist_file.name)
    log.info("═══════════════════════════════════════════════════════")

    mdd           = MDDLoader(mdd_file)
    brand_mapping = BrandMappingLoader(brand_mapping_file)
    rna           = RNALoader(brand_mapping_file)  # Load RNA for brand_type lookup
    linelist      = CrocsLinelistLoader(linelist_file)

    if linelist.df.empty:
        log.warning("[Crocs Linelist] No data rows found")
        return

    # ── Pipeline args ────────────────────────────────────────────
    brand        = getattr(args, "brand",        "Crocs").title()
    brand_code   = getattr(args, "brand_code",   "CCR").upper()
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP").upper()
    season       = getattr(args, "season",       "FW26").upper()
    country_code = getattr(args, "country_code", "ID").upper()

    log.info(
        "Pipeline → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s",
        brand, brand_code, comp_code, sbu, season,
    )

    # ── Resolve Brand Type from RNA ──────────────────────────────
    country_name = _COUNTRY_MAP.get(country_code, country_code)
    
    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        country_name, comp_code, sbu, brand_code,
    )
    
    rna_result = rna.get(country_name, comp_code, sbu, brand_code)
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        rna_result = rna.get(country_name, str(int(comp_code)), sbu, brand_code)
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        log.warning("[RNA] Exact match failed — trying fuzzy (country+sbu+brand)")
        rna_result = rna.get_fuzzy(country_name, sbu, brand_code)
    
    brand_type     = rna_result.get("brand_type",     "")
    brand_category = rna_result.get("brand_category", "")
    
    log.info(
        "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
        country_name, brand_type, brand_category,
    )

    # ── Season ID (same formula as Lotto) ────────────────────────
    sea_prefix = season[:2].upper()
    sea_tail   = season[2:]
    sea_year   = f"20{sea_tail}" if len(sea_tail) == 2 else sea_tail
    season_id  = f"CLH_{brand_code}_{sea_prefix}{sea_year}"

    # ── Log sheet info ────────────────────────────────────────────
    log.info("[Crocs Linelist] Sheets loaded: %s", linelist.sheets_loaded)
    log.info("[Crocs Linelist] Total rows to process: %d", len(linelist.df))

    # ── Build Classifications block ──────────────────────────────
    cls_el  = build_classifications(brand, brand_code, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    del cls_el

    # ── Output file: same stem as input, .xml extension ──────────
    out_name = linelist_file.stem + ".xml"
    out_path = XML_OUT_DIR / out_name
    export_time   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written_count = 0

    log.info("[Crocs Linelist] Streaming XML → %s ...", out_name)

    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        # ── XML declaration + root open ──────────────────────────
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
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

        # ── Classifications block ─────────────────────────────────
        f.write(f"  {cls_str}\n\n")
        del cls_str

        # ── Products ──────────────────────────────────────────────
        f.write("  <Products>\n")

        log.info("[Crocs Linelist] Total products to write: %d", len(linelist.df))

        for _, row in linelist.df.iterrows():
            row_dict = row.to_dict()

            product_xml = build_product_xml(
                row_dict, brand, brand_code, comp_code, sbu,
                season_id, season, country_code, brand_type, brand_mapping, mdd,
            )
            if not product_xml:
                continue

            product_xml = _XMLNS_RE.sub("", product_xml)
            f.write(f"    {product_xml}\n")
            written_count += 1
            del product_xml

        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")

    file_kb = out_path.stat().st_size // 1024
    log.info("✓ XML written → %s  (%dKB)", out_path.name, file_kb)

    print("═══ CROCS LINELIST SUMMARY ═════════════════════════════", flush=True)
    print(f"  Input file      : {linelist_file.name}",           flush=True)
    print(f"  Sheets          : {', '.join(linelist.sheets_loaded)}", flush=True)
    print(f"  Products written: {written_count}",                flush=True)
    print(f"  Output XML      : {out_path.name}",                flush=True)
    print(f"  File size       : {file_kb}KB",                    flush=True)
    print("═══════════════════════════════════════════════════════════", flush=True)

    if auditor:
        try:
            auditor.set_xml_uploads([str(out_path)])
        except AttributeError:
            pass



if __name__ == "__main__":
    import argparse
    import types
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--brand", default="Crocs")
    parser.add_argument("--brand-code", dest="brand_code", default="CRO")
    parser.add_argument("--season", default="FW26")
    parser.add_argument("--country-code", dest="country_code", default="ID")
    
    parsed_args = parser.parse_args()
    args = types.SimpleNamespace(**vars(parsed_args))
    
    run(args)
