"""
╔══════════════════════════════════════════════════════════════════╗
║   PAZZION SEA — DYNAMIC ORDER FORM GENERATOR + STIBO XML v1.0   ║
║  Reads brand mapping file dynamically for attribute mappings     ║
║  and MDD file for LOV lookups                                    ║
╚══════════════════════════════════════════════════════════════════╝

Lambda entry point: run(args, auditor)

Key Features:
- Dynamically loads attribute mappings from "NEW - Brand mapping files Template.xlsx"
- Reads "Pazzion(Inline)" sheet for attribute definitions
- Fetches LOV IDs from MDD file sheets dynamically
- Supports both "Direct from Principal" and "Mapping from Principal" with column mapping
"""

from __future__ import annotations

import logging
import math
import os
import re
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# Configure logging level
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
log = logging.getLogger(__name__)

# ── Paths (Lambda-aware) ─────────────────────────────────────────
BASE_DIR    = Path(os.environ.get("LAMBDA_TMP_DIR", "/tmp/stibo_workdir_pazzion"))
XML_OUT_DIR = BASE_DIR / "output" / "xml"
INPUT_DIR   = BASE_DIR / "input"
ORDERFORM_DIR = INPUT_DIR / "orderform"
MDD_DIR     = INPUT_DIR / "mdd"
ATTR_DIR    = INPUT_DIR / "attributes"

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("",    STIBO_NS)
ET.register_namespace("xsi", STIBO_XSI)

# ── Season LOV mapping ────────────────────────────────────────────
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter", "SP": "Spring",
    "SM": "Summer",
}

# ── Attributes to exclude from XML (even if in brand mapping) ─────
EXCLUDED_ATTRIBUTES = {
    "AT_SAPStyleCode",       # Not needed in XML output
    "AT_Generic",
    "AT_Material",           # Excluded from the *generic* dynamic-mapping loop only —
                              # its brand-mapping field ("Material Description") points at
                              # the label column, not the value. Written separately below
                              # from the captured "Upper" material via MaterialMappingLoader.
    "AT_GenericDescription", # Excluded per business rule
    "AT_Variant",            # Excluded per business rule
    "AT_VariantDescription", # Excluded per business rule
    "AT_HeelHeight",         # Excluded from the *generic* dynamic-mapping loop only —
                              # its brand-mapping field ("Material Description") also
                              # points at the label column, not the value, and would
                              # otherwise win the race and block the correct dedicated
                              # handling below (which uses the captured heel-height value
                              # + HeelHeightMappingLoader + LOV lookup).
    "AT_PrincipalSize",      # Excluded from the *generic* dynamic-mapping loop only —
                              # its brand-mapping field ("Material Description") also
                              # points at the label column, not the value, and would
                              # otherwise win the race and block the correct dedicated
                              # handling below (which uses the captured "_available_sizes"
                              # value from _post_process_colors).
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
        self._load_brand_group_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_heel_height_lov(wb)
        self._load_material_lov(wb)

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
        """Load Country LOV sheet: Col A = ID (e.g. KH, ID, SG), Col B = Country Name.
        Row 1 is the header (ID / Value), data starts from row 2.
        Builds:
          CountryIDToName : { 'KH' -> 'Cambodia', 'ID' -> 'Indonesia', ... }
          CountryLOV      : { 'CAMBODIA' -> 'KH', 'INDONESIA' -> 'ID', ... }
        """
        # Prefer exact 'Country LOV' sheet; fall back to any sheet with COUNTRY+LOV (not ORIGIN)
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            next(
                (s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper() and "ORIGIN" not in s.upper()),
                None,
            ),
        )
        if not sheet_name:
            log.warning("[MDD] Country LOV sheet not found")
            return

        log.info("[MDD] Country LOV — using sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        country_lov: dict[str, str] = {}        # Name (upper) -> ID
        country_id_to_name: dict[str, str] = {}  # ID (upper)   -> Name

        for row in rows[1:]:  # skip header row
            if not row or len(row) < 2:
                continue
            country_id   = str(row[0]).strip() if row[0] else ""
            country_name = str(row[1]).strip() if row[1] else ""

            # Skip header-like rows that may have leaked through
            if country_id.upper() in ("ID", "CODE") and country_name.upper() in ("VALUE", "NAME", "COUNTRY"):
                continue

            if country_id and country_name:
                country_id_to_name[country_id.upper()] = country_name
                country_lov[country_name.upper()] = country_id

        self.lovs["CountryLOV"]      = country_lov
        self.lovs["AT_Country"]      = country_lov
        self.lovs["CountryIDToName"] = country_id_to_name
        log.info("[MDD] Country LOV loaded — %d entries: %s",
                 len(country_id_to_name), list(country_id_to_name.items()))

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
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND TYPE LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Brand Type LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}  # display_value (upper) -> lov_id

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

    def _load_brand_group_lov(self, wb):
        """Load Brand Group LOV sheet: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            # Also try sheet named exactly "Brand Group" (without LOV)
            sheet_name = next(
                (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP"),
                None,
            )
        if not sheet_name:
            log.warning("[MDD] Brand Group LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}  # display_value (upper) -> lov_id

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

    def _load_brand_status_lov(self, wb):
        """Load Brand Status LOV sheet: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            # Also try sheet named exactly "Brand Status" (without LOV)
            sheet_name = next(
                (s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS"),
                None,
            )
        if not sheet_name:
            log.warning("[MDD] Brand Status LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}  # display_value (upper) -> lov_id

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

    def _load_heel_height_lov(self, wb):
        """Load Heel Height LOV sheet: Col A = LOV ID, Col B = display value."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "HEEL HEIGHT LOV"),
            next((s for s in wb.sheetnames if "HEEL" in s.upper() and "HEIGHT" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Heel Height LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}  # display_value (upper) -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id

        self.lovs["HeelHeightLOV"] = lov_map
        self.lovs["AT_HeelHeight"] = lov_map
        log.info("[MDD] Heel Height LOV loaded — %d entries", len(lov_map))

    def _load_material_lov(self, wb):
        """Load Material LOV sheet: Col A = LOV ID (Code), Col B = display value (Material name)."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "MATERIAL LOV"),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Material LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}  # display_value (upper) -> lov_id

        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id

        self.lovs["MaterialLOV"] = lov_map
        self.lovs["AT_Material"] = lov_map
        log.info("[MDD] Material LOV loaded — %d entries", len(lov_map))

    def _load_named_lov_sheets(self, wb):
        """Load any sheet with 'LOV' in name that hasn't been loaded yet."""
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            if sn.strip().upper() == "MATERIAL LOV":
                continue  # Already loaded (dedicated ID/Name-ordered loader above)
            if any(kw in sn.upper() for kw in ["AGE", "GENDER", "SIZE CODE", "COLOR CODE", "COUNTRY", "BRAND", "SIMPLE"]):
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

    def get_country_name_by_id(self, country_id: str) -> str:
        """Get country name from country ID (e.g., 'KH' -> 'Cambodia').
        Uses the reverse mapping loaded from Country LOV sheet.
        """
        if not country_id:
            return ""
        
        search_id = country_id.strip().upper()
        
        # Try reverse mapping first
        if "CountryIDToName" in self.lovs:
            if search_id in self.lovs["CountryIDToName"]:
                result = self.lovs["CountryIDToName"][search_id]
                log.info("[MDD] Country ID '%s' resolved to '%s'", country_id, result)
                return result
            else:
                all_entries = list(self.lovs["CountryIDToName"].items())
                log.warning("[MDD] Country ID '%s' NOT found in CountryIDToName map (%d entries): %s",
                            country_id, len(all_entries), all_entries)
        else:
            log.warning("[MDD] CountryIDToName map NOT loaded — MDD Country LOV sheet missing?")
        
        # If not found, return the original ID
        return country_id


class HeelHeightMappingLoader:
    """
    Loads the Heel Height Mapping from Brand Mapping file.
    Maps raw heel height values (e.g., '0.5CM', '1CM') to descriptive values (e.g., 'Flat', 'Low').
    
    Expected sheet structure:
    - Sheet name: 'Pazzion MD Mapping-Heel Height' or similar
    - Column B: Heel Height in CM (e.g., '0.5CM', '1CM', '1.5CM')
    - Column D: Pazzion value (e.g., 'Flat', 'Flat', 'Flat')
    """
    
    def __init__(self, mapping_file_path: Path):
        self.path = mapping_file_path
        self.mapping: dict[str, str] = {}  # raw_value (upper) -> pazzion_value
        self._load()
    
    def _load(self):
        log.info("[HeelHeightMapping] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        # Find the Pazzion heel height mapping sheet
        sheet_name = next(
            (s for s in wb.sheetnames if "PAZZION" in s.upper() and "HEEL" in s.upper() and "MAPPING" in s.upper()),
            None,
        )
        
        if not sheet_name:
            log.warning("[HeelHeightMapping] No Pazzion Heel Height mapping sheet found")
            wb.close()
            return
        
        log.info("[HeelHeightMapping] Found sheet: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        
        # Skip header row (row 1 typically has column names)
        for row in rows[1:]:
            if not row or len(row) < 4:
                continue
            
            raw_cm = str(row[1]).strip() if row[1] else ""  # Column B: Heel Height in CM
            pazzion_value = str(row[3]).strip() if row[3] else ""  # Column D: Pazzion
            
            if raw_cm and pazzion_value:
                self.mapping[raw_cm.upper()] = pazzion_value
        
        wb.close()
        log.info("[HeelHeightMapping] Loaded %d mappings", len(self.mapping))
    
    def get_pazzion_value(self, raw_value: str) -> str:
        """
        Get Pazzion heel height value from raw input.
        
        Args:
            raw_value: Raw heel height from Excel (e.g., '-', '0.5CM', '1CM')
        
        Returns:
            Pazzion value (e.g., 'Flat', 'Low', 'Medium', 'High') or empty string.
            Returns empty string if raw_value is "-" (attribute should not be sent).
        """
        if not raw_value:
            return ""
        
        search_val = raw_value.strip().upper()
        
        # If value is "-", return empty (don't send attribute)
        if search_val == "-":
            return ""
        
        return self.mapping.get(search_val, "")


class MaterialMappingLoader:
    """
    Loads the Material Mapping from the Brand Mapping file.
    Maps the raw "(Upper) Material" description text from the order form
    (e.g. 'Cowhide Leather', 'Lambskin') to the standardized Pazzion/MDD
    material category (e.g. 'Leather', 'Fabrics') used for AT_Material.

    Expected sheet: 'Pazzion MD Mapping-Material'
    - Footwear:     Col A = raw material, Col B = mapped Pazzion value
    - Non-Footwear: Col D = raw material, Col E = mapped Pazzion value
    """

    # A couple of source values don't exactly match the MDD Material LOV
    # display names (e.g. the sheet uses singular "Fabric"); normalize here.
    _DISPLAY_ALIASES = {"FABRIC": "FABRICS"}

    def __init__(self, mapping_file_path: Path):
        self.path = mapping_file_path
        self.footwear_map: dict[str, str] = {}     # raw_value (upper) -> Pazzion value
        self.nonfootwear_map: dict[str, str] = {}   # raw_value (upper) -> Pazzion value
        self._load()

    def _load(self):
        log.info("[MaterialMapping] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames if "PAZZION" in s.upper() and "MATERIAL" in s.upper() and "MAPPING" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MaterialMapping] No Pazzion Material mapping sheet found")
            wb.close()
            return

        log.info("[MaterialMapping] Found sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))

        # Skip the two header rows (division labels + column labels)
        for row in rows[2:]:
            if not row:
                continue
            fw_raw  = str(row[0]).strip() if len(row) > 0 and row[0] else ""
            fw_val  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            nfw_raw = str(row[3]).strip() if len(row) > 3 and row[3] else ""
            nfw_val = str(row[4]).strip() if len(row) > 4 and row[4] else ""

            if fw_raw and fw_val:
                self.footwear_map[fw_raw.upper()] = fw_val
            if nfw_raw and nfw_val:
                self.nonfootwear_map[nfw_raw.upper()] = nfw_val

        wb.close()
        log.info("[MaterialMapping] Loaded %d footwear / %d non-footwear mappings",
                 len(self.footwear_map), len(self.nonfootwear_map))

    def get_pazzion_value(self, raw_value: str, is_footwear: bool = True) -> str:
        """
        Get Pazzion material value from raw "Upper" material text.
        Returns "" if raw_value is empty, "-", or not found in the mapping.
        """
        if not raw_value:
            return ""

        search_val = raw_value.strip().upper()
        if search_val == "-":
            return ""

        table  = self.footwear_map if is_footwear else self.nonfootwear_map
        mapped = table.get(search_val, "")
        if not mapped:
            return ""
        return self._DISPLAY_ALIASES.get(mapped.upper(), mapped)


class BrandMappingLoader:
    """
    Loads the "NEW - Brand mapping files Template.xlsx" file.
    Reads the "Pazzion(Inline)" sheet to get dynamic attribute mappings.
    
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
    - "Mapping from Principal" → INCLUDE
    - "Formula in system" + has field name (Col J) → INCLUDE
    - "Formula in system" without field name → SKIP (system-level only)
    
    All included attributes are processed automatically:
    - LOV types → sent as ID attributes
    - Text types → sent as value attributes
    - Mapping logic (Col K) applied automatically
    """

    def __init__(self, path: Path, sheet_name: str = "Pazzion(Inline)"):
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
            
            # Skip if no attribute ID or if it's #N/A
            if not col_b or str(col_b).startswith("#N/A"):
                continue
            
            col_g_str = str(col_g).strip() if col_g else ""
            col_j_str = str(col_j).strip() if col_j else ""
            col_k_str = str(row[10]).strip() if len(row) > 10 and row[10] else ""  # Mapping Logic

            col_g_lower = col_g_str.lower()

            # Include if (case-insensitive):
            # 1. "Direct from Principal" OR
            # 2. "Formula in system" AND has column value in J OR
            # 3. "Mapping from Principal" AND has column value in J
            include = False
            if "direct from principal" in col_g_lower:
                include = True
            elif "formula in system" in col_g_lower and col_j_str:
                include = True
            elif "mapping from" in col_g_lower and "principal" in col_g_lower and col_j_str:
                include = True
            
            if include:
                # For fields like "FW & Acc : Article", extract the actual field name
                actual_field = col_j_str
                if ':' in col_j_str:
                    # Extract field name after colon
                    parts = col_j_str.split(':')
                    if len(parts) > 1:
                        actual_field = parts[-1].strip()
                
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
        log.info("[RNA] Columns found → country=%s  comp=%s  sbu=%s  brandcode=%s  btype=%s  bcat=%s",
                 c_country, c_comp, c_sbu, c_bcode, c_btype, c_bcat)
        # Log first 3 keys for verification
        for k, v in list(self.lookup.items())[:3]:
            log.info("[RNA] Sample stored key: %s → %s", k, v)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        """Return {brand_type, brand_category} for the given keys, or empty strings."""
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
        # pass 1: country + sbu + brand
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        # pass 2: country + brand only (first match)
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — ORDER FORM LOADER
# ══════════════════════════════════════════════════════════════════

class PazzionOrderFormLoader:
    """
    Loads the Pazzion order form Excel file.
    Reads ALL sheets, auto-detecting the header row per sheet
    (first row containing key columns like 'Article No.', 'Colours', 'Unit Price').
    """

    # Anchor columns used to detect the header row
    HEADER_ANCHORS = {"article no", "article no.", "colours", "unit price", "retail price", "material description"}

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
        log.info("[OrderForm] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        log.info("[OrderForm] Found %d sheets: %s", len(wb.sheetnames), wb.sheetnames)

        all_dfs = []

        for sheet_name in wb.sheetnames:
            log.info("[OrderForm] Reading sheet: '%s'", sheet_name)
            try:
                ws = wb[sheet_name]
                hdr_idx = self._find_header_row(ws)

                if hdr_idx is None:
                    log.warning("[OrderForm] Sheet '%s': no header row found — skipping", sheet_name)
                    continue

                log.info("[OrderForm] Sheet '%s': header at row %d (0-indexed)", sheet_name, hdr_idx)

                # Read with correct header row
                df_sheet = pd.read_excel(
                    self.path,
                    sheet_name=sheet_name,
                    header=hdr_idx,
                )

                # Drop rows where ALL non-image columns are empty
                df_sheet.dropna(how="all", inplace=True)
                df_sheet.columns = [str(c) if not isinstance(c, str) else c for c in df_sheet.columns]

                if df_sheet.empty:
                    log.warning("[OrderForm] Sheet '%s' has no data rows after header", sheet_name)
                    continue

                df_sheet["_source_sheet"] = sheet_name
                # Post-process: collect colors from sub-rows into parent article row.
                # NOTE: must run BEFORE dropping unnamed columns — the MATERIAL
                # DESCRIPTION *value* column (Heel Height / Available Sizes / Upper
                # material) has no header text, so it shows up as "Unnamed: N" and
                # would otherwise be stripped before this method ever gets to read it.
                df_sheet = self._post_process_colors(df_sheet)

                # Now safe to drop columns that are entirely unnamed (Unnamed: x)
                df_sheet = df_sheet.loc[:, ~df_sheet.columns.astype(str).str.startswith("Unnamed:")]
                all_dfs.append(df_sheet)
                self.sheets_loaded.append(sheet_name)
                log.info("[OrderForm] Sheet '%s': %d rows loaded, columns: %s",
                         sheet_name, len(df_sheet), list(df_sheet.columns[:8]))

            except Exception as e:
                log.error("[OrderForm] Failed to read sheet '%s': %s", sheet_name, e)

        wb.close()

        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            log.info("[OrderForm] Combined total: %d rows from %d sheets",
                     len(self.df), len(self.sheets_loaded))
        else:
            log.warning("[OrderForm] No data loaded from any sheet")

    def _post_process_colors(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        The Pazzion order form has colors in sub-rows below each article:
          Row N:   ARTICLE NO.=3306-3, COLOURS=(empty)
          Row N+1: ARTICLE NO.=(empty), COLOURS=BLACK
          Row N+2: ARTICLE NO.=(empty), COLOURS=WHITE

        Also extracts material description fields:
          Row N+3: ARTICLE NO.=(empty), MATERIAL DESCRIPTION Col1="Heel Height", Col2="-"
          Row N+4: ARTICLE NO.=(empty), MATERIAL DESCRIPTION Col1="Available Sizes", Col2="35-41"

        This method collects all colors from sub-rows, stores them
        comma-separated in the parent article row's COLOURS column,
        extracts Heel Height and Available Sizes values,
        then drops the sub-rows.
        """
        # Find article and colour column names (case-insensitive)
        article_col = next(
            (c for c in df.columns if str(c).strip().upper() in ("ARTICLE NO.", "ARTICLE NO", "ARTICLE")),
            None,
        )
        colour_col = next(
            (c for c in df.columns if str(c).strip().upper() == "COLOURS"),
            None,
        )
        
        # Find material description columns (Col 7 = label, Col 8 = value)
        # Column names might be "MATERIAL DESCRIPTION" and unnamed column after it
        mat_desc_label_col = None
        mat_desc_value_col = None
        
        for idx, col in enumerate(df.columns):
            col_str = str(col).strip().upper()
            if "MATERIAL DESCRIPTION" in col_str:
                mat_desc_label_col = col
                # The value column is the next column (skip if it's a size column or starts with number)
                if idx + 1 < len(df.columns):
                    next_col = str(df.columns[idx + 1]).strip()
                    # Skip if next column looks like a size column (numeric or starts with digit)
                    if not next_col.isdigit() and not (next_col and next_col[0].isdigit()):
                        mat_desc_value_col = df.columns[idx + 1]
                break
        
        log.info("[OrderForm] Material description columns: label='%s', value='%s'",
                 mat_desc_label_col, mat_desc_value_col)

        if not article_col or not colour_col:
            log.warning("[OrderForm] _post_process_colors: article_col=%s, colour_col=%s — skipping",
                        article_col, colour_col)
            return df

        def _sv(v):
            """Safe string conversion."""
            if v is None:
                return ""
            import math as _math
            if isinstance(v, float) and _math.isnan(v):
                return ""
            s = str(v).strip()
            return "" if s.lower() in ("nan", "none", "nat") else s

        rows_out   = []       # list of dicts for article rows
        colors_buf = []       # colors collected for current article
        heel_height_val = ""  # heel height for current article
        available_sizes_val = ""  # available sizes for current article
        material_upper_val = ""   # "Upper" material composition for current article
        division_val = ""     # "Division" (e.g. Footwear/Bag) for current article
        category_val = ""     # "Category" (e.g. COVERED FLATS) for current article

        for _, row in df.iterrows():
            article_val = _sv(row.get(article_col, ""))
            colour_val  = _sv(row.get(colour_col, ""))

            # Extract material description label and value
            mat_label = _sv(row.get(mat_desc_label_col, "")) if mat_desc_label_col else ""
            mat_value = _sv(row.get(mat_desc_value_col, "")) if mat_desc_value_col else ""

            if article_val:
                # Flush previous article with its collected data
                if rows_out:
                    rows_out[-1][colour_col] = ",".join(colors_buf) if colors_buf else ""
                    rows_out[-1]["_heel_height"] = heel_height_val
                    rows_out[-1]["_available_sizes"] = available_sizes_val
                    rows_out[-1]["_material_upper"] = material_upper_val
                    rows_out[-1]["_division"] = division_val
                    rows_out[-1]["_category"] = category_val

                row_dict = row.to_dict()
                rows_out.append(row_dict)
                # Seed colors: include the colour on the article row itself if present
                colors_buf = [colour_val] if colour_val else []
                # Reset material description fields
                heel_height_val = ""
                available_sizes_val = ""
                material_upper_val = ""
                division_val = ""
                category_val = ""
            else:
                # Sub-row — capture colour values
                if colour_val:
                    colors_buf.append(colour_val)

            # Capture material description fields — the first one ("Upper") lands
            # on the article's own row, the rest (Heel Height, Available Sizes, ...)
            # land on sub-rows below it, so this must run for BOTH row kinds.
            if mat_label and mat_value:
                mat_label_lower = mat_label.lower().strip()
                if "heel height" in mat_label_lower:
                    heel_height_val = mat_value
                    log.info("[OrderForm] Captured heel height: label='%s', value='%s'", mat_label, mat_value)
                elif "available size" in mat_label_lower:
                    available_sizes_val = mat_value
                    log.info("[OrderForm] Captured available sizes: label='%s', value='%s'", mat_label, mat_value)
                elif mat_label_lower in ("upper", "material"):
                    # Footwear rows label this "Upper"; Bag rows label the
                    # exact same slot "Material" instead (confirmed 1:1
                    # against every "Material"-labeled row's Division ==
                    # "Bag") — both need capturing into the same field.
                    material_upper_val = mat_value
                    log.info("[OrderForm] Captured %s material: label='%s', value='%s'",
                             "bag" if mat_label_lower == "material" else "upper", mat_label, mat_value)
                elif mat_label_lower == "division":
                    division_val = mat_value
                    log.info("[OrderForm] Captured division: label='%s', value='%s'", mat_label, mat_value)
                elif mat_label_lower == "category":
                    category_val = mat_value
                    log.info("[OrderForm] Captured category: label='%s', value='%s'", mat_label, mat_value)

        # Flush the last article
        if rows_out:
            rows_out[-1][colour_col] = ",".join(colors_buf) if colors_buf else ""
            rows_out[-1]["_heel_height"] = heel_height_val
            rows_out[-1]["_available_sizes"] = available_sizes_val
            rows_out[-1]["_material_upper"] = material_upper_val
            rows_out[-1]["_division"] = division_val
            rows_out[-1]["_category"] = category_val

        result = pd.DataFrame(rows_out)
        print(f"[POST PROCESS] {len(df)} rows → {len(result)} article rows after color grouping", flush=True)
        return result


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

# Strip duplicate xmlns injected by ET.tostring
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


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
    Tries exact match first, then falls back to a normalized (whitespace-
    and case-insensitive) comparison.

    Different sheets in the same order form sometimes spell a header
    slightly differently (e.g. "UNIT PRICE\n(SGD)" vs "Unit Price \n(SGD)").
    After the sheets are concatenated, EVERY row carries both columns —
    NaN in whichever one didn't belong to its own source sheet — so the
    first normalized-match key is not necessarily the right one. Keep
    scanning and return the first NON-EMPTY match instead of stopping at
    the first key encountered.
    """
    # Exact match
    v = row_data.get(field_name)
    if v is not None:
        s = _safe_str(v)
        if s:
            return s

    # Normalize field name (remove newlines, extra spaces, case-insensitive)
    def normalize(s):
        return " ".join(str(s).replace("\n", " ").split()).upper()

    field_normalized = normalize(field_name)

    for k, val in row_data.items():
        if normalize(k) == field_normalized:
            s = _safe_str(val)
            if s:
                return s
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


def _derive_season(season_code: str) -> tuple[str, str]:
    """'SS26' → ('SS', '2026'). Handles bare prefix or already 4-digit year."""
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


def _season_id_full(brand_code: str, season_code: str) -> str:
    """Build classification ID like 'CLH_PZZ_SS2026'."""
    pfx, yr = _derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"


def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    """
    Build the <Classifications> block.
      CLH_{brand_code}_{season}          CLS_Season         parent=CLH_{brand}Batches
        CLH_{brand_code}_{season}CA      CLS_ConfirmedArticles
        CLH_{brand_code}_{season}UA      CLS_UnconfirmedArticles
    """
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    brand_token  = brand.title().replace(" ", "")
    pfx, yr      = _derive_season(season_code)
    season_id    = _season_id_full(brand_code, season_code)
    batches_par  = f"CLH_{brand_token}Batches"

    _SEASON_LABELS = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
    }
    sea_name = _SEASON_LABELS.get(pfx, pfx)
    display  = f"{brand} {pfx} {yr}".strip()
    short    = f"{pfx} {yr}".strip()

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


def _determine_parent_id(row_data: dict, source_sheet: str) -> str:
    """
    Determine the PPH ParentID based on the row's captured Division value
    (from the "Material Description" sub-row — "Footwear" or "Bag" in the
    current order form; captured in _post_process_colors as _division),
    falling back to sheet-name heuristics only when a row has no Division
    captured at all.

    The order form's own sheet names (e.g. "SS'26") never contain
    "footwear"/"accessories", so relying on sheet name alone silently
    defaulted EVERY row — including all "Bag" rows — to Footwear. That
    broke both AT_Material (Bags need the brand mapping sheet's "ACC (non
    FW)" reference column, not the Footwear one — lookups against the
    wrong table just come back empty) and AT_CountrySize (which must use
    a different default ID for Bags/Accessories — see below).

    PPH division letter convention (shared across brands — see e.g.
    dr_marten/retail_price_master_main.py's DIVISION_PARENT_MAP):
        Accessories - E   Footwear - F   Apparel - A
        Equipment   - Q   Toys     - T
    "A" is Apparel, NOT Accessories — Bags must map to PPH_E-TempSubCat.
    """
    division = _safe_str(row_data.get("_division", "")).strip().lower()
    if division:
        if "footwear" in division:
            return "PPH_F-TempSubCat"
        # "Bag" (and anything else non-footwear) -> Accessories bucket,
        # matching the brand mapping sheet's "ACC (non FW)" material split.
        return "PPH_E-TempSubCat"

    sheet = (source_sheet or "").lower()
    if "footwear" in sheet or "fw" in sheet:
        return "PPH_F-TempSubCat"
    if "acc" in sheet or "accessories" in sheet:
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
    brand_category: str,
    brand_group:   str,
    brand_status:  str,
    brand_mapping: BrandMappingLoader,
    mdd:           MDDLoader,
    heel_ht_mapping: HeelHeightMappingLoader,
    material_mapping: MaterialMappingLoader,
    color:         str = "",
) -> str:
    """
    Build one <Product> XML string:
      • UserTypeID = PRD_GenericArticle
      • ParentID   = PPH_{div}-TempSubCat
      • KeyValue   KEY_InboundArticle
      • Name
      • ClassificationReference  ×2  (Merchandiser + UnConfirmedForSeason)
      • Values  (dynamic from brand mapping + brand_code/comp/sbu/season)
    """
    # ── Resolve identifying fields ────────────────────────────────
    # Try "Article No." first (current format), fallback to "Article" (legacy)
    article_code = _safe_str(row_data.get("Article No.") or row_data.get("ARTICLE NO.") or row_data.get("Article"))
    source_sheet = _safe_str(row_data.get("_source_sheet", ""))

    if not article_code:
        return ""   # skip rows with no identifying code

    # Generic code: brand_code(3) + article_code (uppercased) + color (uppercased)
    article_upper = article_code.upper()
    color_upper   = color.upper().strip() if color else ""
    generic_code  = f"{brand_code[:3].upper()}{article_upper}{color_upper}"

    parent_id = _determine_parent_id(row_data, source_sheet)

    # ── Principal Style Description (per brand mapping: "Article No. +
    #    Colours + MD Material Mapping") + <Name>, which mirrors it ──────
    # e.g. "3306-3 BLACK Leather". Resolve material the same way the
    # AT_Material block does (footwear vs "ACC (non FW)" reference table).
    material_upper_raw = _safe_str(row_data.get("_material_upper", ""))
    pazzion_material_desc = ""
    if material_upper_raw and material_upper_raw != "-":
        is_footwear_flag = parent_id == "PPH_F-TempSubCat"
        pazzion_material_desc = material_mapping.get_pazzion_value(
            material_upper_raw, is_footwear=is_footwear_flag,
        )
    style_description = " ".join(
        p for p in (article_code, color.strip() if color else "", pazzion_material_desc) if p
    ).strip()
    row_data["_style_description"] = style_description

    # ── Element construction ──────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    # KeyValue
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    # Name — mirrors AT_PrincipalStyleDescription (see below)
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = style_description

    # ClassificationReferences
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # Add generic_code and color to row_data for use in attributes
    row_data["_generic_code"] = generic_code
    row_data["_color"] = color

    # Values
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_product_values(vals_el, row_data, brand, brand_code, comp_code, sbu, season, country_code, brand_type, brand_group, brand_status, brand_mapping, mdd, heel_ht_mapping, material_mapping)

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
    brand_group:   str,
    brand_status:  str,
    brand_mapping: BrandMappingLoader,
    mdd:           MDDLoader,
    heel_ht_mapping: HeelHeightMappingLoader,
    material_mapping: MaterialMappingLoader,
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
           - "Mapping from Principal" → include
           - "Formula in system" with field name in Col J → include
           - Otherwise → skip (handled separately if needed)
    Col J: Field Name in Excel (e.g., "Article", "Colours", "Size Chart")
    Col K: Mapping Logic (transformations)
    
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

        # For AT_PrincipalColorName, always use the specific color
        # tied to this product's KEY_InboundArticle (stored in _color)
        if attr_id == "AT_PrincipalColorName":
            raw_value = _safe_str(row_data.get("_color", ""))
        # AT_FOB is handled in section 2 below — skip here
        elif attr_id == "AT_FOB":
            continue
        else:
            raw_value = _get_field(row_data, field_name)
        if not raw_value:
            continue

        # Apply mapping logic transformation if specified
        transformed_value = _apply_mapping_logic(raw_value, mapping_logic)

        if value_type == "LOV":
            lov_id = mdd.get_lov_id(attr_id, transformed_value)
            if lov_id:
                _w(attr_id, id_val=lov_id)
        else:
            _w(attr_id, value=transformed_value)

    # ── 2. Fixed system-level values (not in brand file) ──────────
    # Brand
    _w("AT_Brand", id_val=brand_code)

    # Company + SBU (using MultiValue)
    _mw("AT_CompanyCode", comp_code)
    _mw("AT_SBU",        sbu)

    # ── SAP Age (default to AD = Adult) ──────────────────────────
    _w("AT_SAPAge", id_val="AD")

    # ── Gender (default to F = Female) ───────────────────────────
    _w("AT_Gender", id_val="F")

    # ── BY Gender (default to F = Female) ────────────────────────
    _w("AT_BYGender", id_val="F")

    # ── BY Age (default to ADULT) ─────────────────────────────────
    _w("AT_BYAge", id_val="ADULT")

    # ── Country Origin (default to CN = China) ───────────────────
    _w("AT_CountryOrigin", id_val="CN")

    # ── Pricing Distribution Channel (default to 01) ─────────────
    _w("AT_PricingDistributionChannel", id_val="01")

    # ── Generic Code ─────────────────────────────────────────────
    generic_code = _safe_str(row_data.get("_generic_code", ""))
    if generic_code:
        _w("AT_InboundGenericCode", value=generic_code)

    # ── Season + Season Year ─────────────────────────────────────
    # Parse season code (e.g., "SS26" → prefix="SS", year="2026")
    sea_raw = (season or "").strip().upper()
    sea_prefix = sea_raw[:2] if len(sea_raw) >= 2 else sea_raw
    
    # Get season LOV value and ID
    sea_label = LOV_SEASON.get(sea_prefix, sea_prefix)
    _w("AT_Season", sea_label, id_val=sea_prefix)
    
    # Extract year from season code (SS26 → 2026)
    season_year = ""
    if len(sea_raw) >= 4:
        year_part = sea_raw[2:]
        if len(year_part) == 2 and year_part.isdigit():
            season_year = f"20{year_part}"
        elif len(year_part) == 4 and year_part.isdigit():
            season_year = year_part
        else:
            # Try regex to extract year
            year_match = re.search(r"20\d{2}", sea_raw)
            if not year_match:
                short_match = re.search(r"\d{2}$", sea_raw)
                season_year = f"20{short_match.group()}" if short_match else ""
            else:
                season_year = year_match.group()
    
    if season_year:
        _w("AT_SeasonYear", value=season_year)

    # ── Country ──────────────────────────────────────────────────
    if country_code:
        _w("AT_Country", id_val=country_code)

    # ── Brand Type ───────────────────────────────────────────────
    # brand_type comes from RNA BRANDTYPE_DETAIL column
    # Look up its LOV ID from the Brand Type LOV sheet
    if brand_type:
        bt_lov_id = mdd.lovs.get("AT_BrandType", {}).get(brand_type.upper(), "")
        if bt_lov_id:
            _w("AT_BrandType", id_val=bt_lov_id)
        else:
            _w("AT_BrandType", value=brand_type)

    # ── Brand Group ──────────────────────────────────────────────
    # brand_group comes from RNA BRANDGROUP column
    # Look up its LOV ID from the Brand Group LOV sheet
    if brand_group:
        bg_lov_id = mdd.lovs.get("AT_BrandGroup", {}).get(brand_group.upper(), "")
        if bg_lov_id:
            _w("AT_BrandGroup", id_val=bg_lov_id)
        else:
            _w("AT_BrandGroup", value=brand_group)

    # ── Brand Status ─────────────────────────────────────────────
    # brand_status comes from RNA BRAND STATUS column
    # Look up its LOV ID from the Brand Status LOV sheet
    if brand_status:
        bs_lov_id = mdd.lovs.get("AT_BrandStatus", {}).get(brand_status.upper(), "")
        if bs_lov_id:
            _w("AT_BrandStatus", id_val=bs_lov_id)
        else:
            _w("AT_BrandStatus", value=brand_status)

    # ── Heel Height ──────────────────────────────────────────────
    # Extract heel height from row_data (collected in _post_process_colors)
    # Process: raw value -> mapping file (CM to Pazzion) -> LOV lookup -> LOV ID
    # Note: If raw value is "-", skip this attribute entirely
    heel_height_raw = _safe_str(row_data.get("_heel_height", ""))
    log.info("[Heel Height] Raw value from _heel_height: '%s'", heel_height_raw)
    
    if heel_height_raw and heel_height_raw != "-":
        # Step 1: Map raw value (e.g., "2.5cm", "0.5CM") to Pazzion value (e.g., "Flat", "Low")
        pazzion_heel_value = heel_ht_mapping.get_pazzion_value(heel_height_raw)
        log.info("[Heel Height] Mapped to Pazzion value: '%s'", pazzion_heel_value)
        
        if pazzion_heel_value:
            # Step 2: Look up LOV ID from Heel Height LOV sheet
            heel_ht_lov_id = mdd.lovs.get("AT_HeelHeight", {}).get(pazzion_heel_value.upper(), "")
            log.info("[Heel Height] LOV ID lookup: '%s' -> '%s'", pazzion_heel_value.upper(), heel_ht_lov_id)
            
            if heel_ht_lov_id:
                _w("AT_HeelHeight", id_val=heel_ht_lov_id)
            else:
                # Fallback: send as text value if LOV not found
                log.warning("[Heel Height] LOV ID not found for '%s', sending as text", pazzion_heel_value)
                _w("AT_HeelHeight", value=pazzion_heel_value)

    # ── Principal Size (from "Available Sizes" material description) ──
    # Text attribute — sent as-is per brand mapping ("Populated the value
    # from principal data as is"). Captured in _post_process_colors.
    available_sizes_raw = _safe_str(row_data.get("_available_sizes", ""))
    if available_sizes_raw and available_sizes_raw != "-":
        _w("AT_PrincipalSize", value=available_sizes_raw)

    # ── Principal Merchandise Hierarchy L1 (from "Division") ──────
    # Text attribute — sent as-is per brand mapping (field: text, "Rows:
    # Division" within the Material Description block). Captured in
    # _post_process_colors. Only present in the updated order form
    # template (MAA added the Division/Category rows); older files
    # without them simply won't populate this attribute.
    division_raw = _safe_str(row_data.get("_division", ""))
    if division_raw and division_raw != "-":
        _w("AT_PrincipalMerchandiseHierarchyL1", value=division_raw)

    # ── Principal Merchandise Hierarchy L2 (from "Category") ──────
    # Text attribute — sent as-is, same pattern as L1/Division above
    # (field: text, "Rows: Category" within the Material Description
    # block). Captured in _post_process_colors as _category.
    category_raw = _safe_str(row_data.get("_category", ""))
    if category_raw and category_raw != "-":
        _w("AT_PrincipalMerchandiseHierarchyL2", value=category_raw)

    # ── Principal Style Description ("Article No. + Colours + MD
    #    Material Mapping" per brand mapping) — composed once in
    #    build_product_xml (needs the resolved material + parent_id
    #    before Name/Values are split apart) and stashed onto row_data.
    style_description = _safe_str(row_data.get("_style_description", ""))
    if style_description:
        _w("AT_PrincipalStyleDescription", value=style_description)

    # ── Material (from "Upper" material description) ─────────────
    # Raw upper material text -> Pazzion material category (via MD mapping
    # file, footwear/non-footwear specific) -> Material LOV ID.
    material_upper_raw = _safe_str(row_data.get("_material_upper", ""))
    if material_upper_raw and material_upper_raw != "-":
        is_footwear = _determine_parent_id(row_data, source_sheet) == "PPH_F-TempSubCat"
        pazzion_material = material_mapping.get_pazzion_value(material_upper_raw, is_footwear=is_footwear)
        if pazzion_material:
            material_lov_id = mdd.lovs.get("AT_Material", {}).get(pazzion_material.upper(), "")
            if material_lov_id:
                _w("AT_Material", id_val=material_lov_id)
            else:
                log.warning("[Material] LOV ID not found for '%s' (raw='%s'), sending as text",
                            pazzion_material, material_upper_raw)
                _w("AT_Material", value=pazzion_material)
        else:
            log.warning("[Material] No mapping found for raw upper material '%s' (footwear=%s)",
                        material_upper_raw, is_footwear)

    # ── Country Size ──────────────────────────────────────────────
    # Footwear -> 'EU'. Accessories/Bags -> 'NS' (No Size), per the
    # brand mapping sheet's "Acc: No Size" cell text.
    if _determine_parent_id(row_data, source_sheet) == "PPH_F-TempSubCat":
        _w("AT_CountrySize", id_val="EU")
    else:
        _w("AT_CountrySize", id_val="NS")

    # ── Retail Price Currency (from country code) ───────────────
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _w("AT_RetailPriceCurrency", id_val=currency_lov_id)

    # ── FOB Currency (default SGD) ───────────────────────────────
    _w("AT_FOBCurrency", id_val="SGD")

    # ── FOB Price (from "Unit Price (SGD)" column) ────────────────
    fob_value = _get_field(row_data, "Unit Price (SGD)")
    if fob_value:
        _w("AT_FOB", value=fob_value)

    # ── SAP Product Flag (default A) ─────────────────────────────
    _w("AT_SAPProductFlag", id_val="A")

    # ── SAP Article Category (default 1 = Generic) ───────────────
    _w("AT_SAPArticleCategory", id_val="1")

    # ── BCI (default COMMERCIAL) ─────────────────────────────────
    # Confirmed default ID with MAA — no LOV sheet for it in the MDD.
    _w("AT_BCI", id_val="COMMERCIAL")

    # ── Material Type (default ZINA) ──────────────────────────────
    # Confirmed default ID with MAA — no LOV sheet for it in the MDD.
    _w("AT_MaterialType", id_val="ZINA")

    # ── Heel Type — intentionally NOT sent (per MAA: no automated source;
    #    mapping doc lists it as Manual Input with no field in the order form).

    # ── Principal Merchandise Hierarchy L1 — intentionally NOT sent yet
    #    (per MAA: logic still pending discussion, new template received
    #    11/08/2026).

    # ── Pack Details (default S) ─────────────────────────────────
    _w("AT_PackDetails", id_val="S")

    # ── ECom Ages Category (default 18+Y) ────────────────────────
    _w("AT_EComAgesCategory", id_val="18+Y")

    # ── Article Status (default A) ───────────────────────────────
    _w("AT_ArticleStatus", id_val="A")

    # ── Unit of Measure (default EA = Each) ──────────────────────
    _w("AT_UOM", id_val="EA")
    # ── BY Article Type (default Inline) ──────────────────────────
    by_article_type_lov_id = mdd.get_lov_id("AT_BYArticleType", "Inline")
    _w("AT_BYArticleType", id_val=by_article_type_lov_id)



# ══════════════════════════════════════════════════════════════════
# SECTION 4 — MAIN RUN FUNCTION
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point for Pazzion order form processing.
    """
    log.info("[Pazzion Order Form] Starting ...")

    # ── Load files ───────────────────────────────────────────────
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if not mdd_file:
        raise FileNotFoundError(f"No MDD file found in {MDD_DIR}")

    brand_mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if not brand_mapping_file:
        raise FileNotFoundError(f"No Brand mapping file found in {ATTR_DIR}")

    orderform_file = next(ORDERFORM_DIR.glob("*.xlsx"), None)
    if not orderform_file:
        raise FileNotFoundError(f"No order form file found in {ORDERFORM_DIR}")

    log.info("═══ SELECTED FILES ════════════════════════════════════════")
    log.info("  MDD           : %s", mdd_file.name)
    log.info("  Brand Mapping : %s", brand_mapping_file.name)
    log.info("  Order Form    : %s", orderform_file.name)
    log.info("═══════════════════════════════════════════════════════════")

    mdd             = MDDLoader(mdd_file)
    brand_mapping   = BrandMappingLoader(brand_mapping_file)
    rna             = RNALoader(brand_mapping_file)
    heel_ht_mapping = HeelHeightMappingLoader(brand_mapping_file)
    material_mapping = MaterialMappingLoader(brand_mapping_file)
    orderform       = PazzionOrderFormLoader(orderform_file)

    if orderform.df.empty:
        log.warning("[Pazzion Order Form] No data rows found")
        return

    # ── Pipeline args ────────────────────────────────────────────
    brand        = getattr(args, "brand",        "Pazzion").title()
    brand_code   = getattr(args, "brand_code",   "PZZ").upper()
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP").upper()
    season       = getattr(args, "season",       "SS26").upper()

    # ── Parse country_code from input filename ───────────────────
    # Filename convention: ...-{COUNTRY_CODE}-{N}.xlsx
    # e.g. 0888-FF-PAZZION PZZ-Order Form MAA-SP2028-ID-1.xlsx → 'ID'
    _fname_stem = orderform_file.stem  # without .xlsx
    _fname_parts = _fname_stem.replace(" ", "-").split("-")
    _country_from_file = ""
    # Walk parts from the end; look for a 2-letter alphabetic segment
    for _part in reversed(_fname_parts):
        _p = _part.strip().upper()
        if len(_p) == 2 and _p.isalpha():
            _country_from_file = _p
            break

    if _country_from_file:
        log.info("[Pazzion] Country code parsed from filename: '%s' (file: %s)",
                 _country_from_file, orderform_file.name)
        country_code = _country_from_file
    else:
        country_code = getattr(args, "country_code", "SG").upper()
        log.info("[Pazzion] Country code from args (not found in filename): '%s'", country_code)

    log.info(
        "Pipeline → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s",
        brand, brand_code, comp_code, sbu, season,
    )

    # ── Resolve Brand Type from RNA ──────────────────────────────
    # Get country name from MDD Country LOV (ID -> Name mapping)
    country_name = mdd.get_country_name_by_id(country_code)
    
    log.info(
        "[RNA] Country code '%s' resolved to '%s' from MDD Country LOV",
        country_code, country_name,
    )
    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        country_name, comp_code, sbu, brand_code,
    )
    
    # ── Pass 1: exact match with original comp_code ──────────────
    log.info("[RNA] Pass 1 — exact match key: ('%s', '%s', '%s', '%s')",
             country_name, comp_code, sbu, brand_code)
    rna_result = rna.get(country_name, comp_code, sbu, brand_code)
    log.info("[RNA] Pass 1 result: brand_type='%s'  brand_category='%s'",
             rna_result.get("brand_type"), rna_result.get("brand_category"))

    # ── Pass 2: strip leading zeros from comp_code ────────────────
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        try:
            comp_code_stripped = str(int(comp_code))
        except (ValueError, TypeError):
            comp_code_stripped = comp_code
        log.info("[RNA] Pass 2 — stripped comp_code key: ('%s', '%s', '%s', '%s')",
                 country_name, comp_code_stripped, sbu, brand_code)
        rna_result = rna.get(country_name, comp_code_stripped, sbu, brand_code)
        log.info("[RNA] Pass 2 result: brand_type='%s'  brand_category='%s'",
                 rna_result.get("brand_type"), rna_result.get("brand_category"))

    # ── Pass 3: fuzzy match (country + sbu + brand, any comp_code) ─
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        log.warning("[RNA] Pass 3 — exact match failed, trying fuzzy (country+sbu+brand)")
        log.info("[RNA] Pass 3 — fuzzy key: country='%s'  sbu='%s'  brand='%s'",
                 country_name, sbu, brand_code)
        rna_result = rna.get_fuzzy(country_name, sbu, brand_code)
        log.info("[RNA] Pass 3 result: brand_type='%s'  brand_category='%s'",
                 rna_result.get("brand_type"), rna_result.get("brand_category"))

    # ── Log RNA lookup table stats for debugging ──────────────────
    log.info("[RNA] Total entries in lookup table: %d", len(rna.lookup))
    # Show a few sample keys to help debug
    sample_keys = list(rna.lookup.keys())[:5]
    for k in sample_keys:
        log.info("[RNA] Sample key: %s → %s", k, rna.lookup[k])

    brand_type     = rna_result.get("brand_type",     "")
    brand_category = rna_result.get("brand_category", "")
    brand_group    = rna_result.get("brand_group",    "")
    brand_status   = rna_result.get("brand_status",   "")

    if brand_type or brand_category:
        log.info("[RNA] ✓ Resolved — brand_type='%s'  brand_category='%s'  brand_group='%s'  brand_status='%s'",
                 brand_type, brand_category, brand_group, brand_status)
    else:
        log.warning("[RNA] ✗ All 3 passes failed — brand_type and brand_category will be EMPTY")
        log.warning("[RNA] Searched for: country='%s'  comp='%s'  sbu='%s'  brand='%s'",
                    country_name, comp_code, sbu, brand_code)

    # ── Season ID ─────────────────────────────────────────────────
    season_id = _season_id_full(brand_code, season)

    # ── Filter out category header rows (rows with no Article No.) ──
    log.info("[Pazzion Order Form] Sheets loaded: %s", orderform.sheets_loaded)
    log.info("[Pazzion Order Form] Total rows before filtering: %d", len(orderform.df))
    
    # Filter rows that have an article code
    filtered_df = orderform.df[
        orderform.df.apply(
            lambda row: bool(_safe_str(row.get("Article No.") or row.get("ARTICLE NO.") or row.get("Article"))),
            axis=1
        )
    ].copy()
    
    log.info("[Pazzion Order Form] Total rows to process after filtering: %d", len(filtered_df))

    # ── Build Classifications block ──────────────────────────────
    cls_el  = build_classifications(brand, brand_code, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    del cls_el

    # ── Output file: same stem as input, .xml extension ──────────
    out_name = orderform_file.stem + ".xml"
    out_path = XML_OUT_DIR / out_name
    export_time   = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    written_count = 0

    log.info("[Pazzion Order Form] Streaming XML → %s ...", out_name)

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
            f' WorkspaceID="Main"'
            f' UseContextLocale="false">\n'
        )

        # ── Classifications block ─────────────────────────────────
        f.write(f"{cls_str}\n")
        del cls_str

        # ── Products ──────────────────────────────────────────────
        f.write("  <Products>\n")

        for _, row in filtered_df.iterrows():
            row_dict = row.to_dict()
            source_sheet = _safe_str(row_dict.get("_source_sheet", ""))

            # Extract colors — post-processing has already collected sub-row colors
            # into the COLOURS column as comma-separated string
            colors_raw = _safe_str(row_dict.get("Colours") or row_dict.get("COLOURS"))
            article_code = _safe_str(row_dict.get("Article No.") or row_dict.get("ARTICLE NO.") or row_dict.get("Article"))

            if colors_raw:
                colors = [c.strip() for c in colors_raw.replace('\n', ',').split(',') if c.strip()]
            else:
                colors = [""]

            # Generate one product per color
            for color in colors:
                product_xml = build_product_xml(
                    row_dict, brand, brand_code, comp_code, sbu,
                    season_id, season, country_code, brand_type, brand_category,
                    brand_group, brand_status, brand_mapping, mdd, heel_ht_mapping,
                    material_mapping, color,
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

    print("═══ PAZZION ORDER FORM SUMMARY ════════════════════════════", flush=True)
    print(f"  Input file      : {orderform_file.name}",           flush=True)
    print(f"  Sheets          : {', '.join(orderform.sheets_loaded)}", flush=True)
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
    parser.add_argument("--brand", default="Pazzion")
    parser.add_argument("--brand-code", dest="brand_code", default="PAZ")
    parser.add_argument("--season", default="SS26")
    parser.add_argument("--country-code", dest="country_code", default="SG")
    
    parsed_args = parser.parse_args()
    args = types.SimpleNamespace(**vars(parsed_args))
    
    run(args)
