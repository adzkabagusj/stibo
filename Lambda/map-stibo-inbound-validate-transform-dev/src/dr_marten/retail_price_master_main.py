"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — DR Marten Inline  v2.0          ║
║   Retail Price Master → Stibo STEP XML (Article Planning)       ║
║                                                                  ║
║   FULLY DYNAMIC MAPPING ENGINE                                   ║
║   All mappings are read from the Brand Mapping Template Excel    ║
║   No hardcoded field-to-attribute mappings                       ║
╚══════════════════════════════════════════════════════════════════╝

Source file format : "SEA_SS27 B2B Suggested Retail Price Master - 20260526_OUTPUT FINAL.xlsx"
Mapping file       : "NEW - Brand mapping files Template.xlsx" → Sheet: "DR.Martens(Inline)"

Dynamic Mapping Features:
  - Reads all attribute mappings from Excel
  - Supports direct field copy (source field → attribute)
  - Supports default values from mapping logic
  - Supports RNA-based lookups (BrandType, BrandCategory, etc.)
  - Supports formula-based derivations (Generic code, etc.)
  - No Python code changes needed to add/modify attribute mappings
"""

from __future__ import annotations

import logging
import math
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, Callable, Any
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
# TEST LIMIT — Set to a number to limit rows, set to None to process all
# ══════════════════════════════════════════════════════════════════════════════
TEST_ROW_LIMIT = None   # Process all rows (production)
# TEST_ROW_LIMIT = 5      # Process only first 5 rows (for testing)

# ── TEST ROW RANGE — pick a specific 1-based, inclusive range of rows ─────────
# Row numbers are 1-based and refer to data rows (row 1 = first data row).
# Just comment / uncomment the line you want:
TEST_ROW_RANGE = None   # Disabled — using TEST_ROW_LIMIT = 5

# TEST_ROW_RANGE = (496, 501)   # Process only rows 496–501 (1-based, inclusive)

# ══════════════════════════════════════════════════════════════════════════════
# PATHS — driven by LAMBDA_TMP_DIR
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR     = BASE_DIR / "input"
LINELIST_DIR  = INPUT_DIR / "linelist"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"
MAPPING_DIR   = INPUT_DIR / "mapping"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, MAPPING_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("dr_marten.retail_price_master_main")
if not log.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
# Ensure the named logger also has INFO level set (for Lambda compatibility)
log.setLevel(logging.INFO)


# ══════════════════════════════════════════════════════════════════════════════
# STIBO / STEP XML CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')

# Brand constants
BRAND_CODE = "DRM"
BRAND_NAME = "DR MARTEN"
BRAND_GROUP = "DR. MARTENS"

# Country code → Country name (used for RNA lookup)
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
    "KH": "CAMBODIA",
    "HK": "HONG KONG",
    "GB": "UNITED KINGDOM",
    "US": "UNITED STATES",
}


# ══════════════════════════════════════════════════════════════════════════════
# LOV TABLES (Fallback defaults — MDD takes precedence)
# ══════════════════════════════════════════════════════════════════════════════
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_SEASON = {"SS": "Spring-Summer", "FW": "Fall-Winter", "AL": "All Season", "AW": "Autumn-Winter"}

# ── ParentID / SAP Product Division mapping ───────────────────────────────────
# ParentID = PPH_{division_letter}-TempSubCat
#
# Step 1 — Source column 'Product General' from the input Excel is resolved via
#          the 'DR_MARTENS MD Mappings' tab:
#               FW  → "F - FOOTWEAR"
#               NFW → "E - ACCESSORIES"
# Step 2 — The resulting Division value is mapped to a single letter (per the
#          Products ParentID rule):
#               Accessories - E
#               Footwear    - F
#               Apparel     - A
#               Equipment   - Q
#               Toys        - T
#               TEST DIV    - X   (also the fallback for any unknown Division)
PRODUCT_GENERAL_TO_DIVISION: dict[str, str] = {
    "FW":  "F",   # Footwear
    "NFW": "E",   # Non-Footwear → Accessories
}

DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E",
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "EQUIPMENT":   "Q",
    "TOYS":        "T",
    "TEST DIV":    "X",
}

# NOTE: Gender is resolved dynamically from the MDD "Gender LOV" tab
#       (see MDDLoader._load_gender_lov / MDDLoader.resolve_gender_lov_id).
#       No hardcoded gender mapping is used.

# Country → Currency mapping (fallback)
COUNTRY_CURRENCY: dict[str, str] = {
    "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD", "MY": "MYR",
    "VN": "VND", "KH": "USD", "US": "USD", "GB": "GBP", "HK": "HKD",
}


# ══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "") -> Optional[ET.Element]:
    """Create a Value element under parent."""
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


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    """Create a MultiValue element with a single Value ID for LOV-based multi-valued attributes."""
    if not id_val or str(id_val).strip() in ("", "None", "nan"):
        return
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val).strip())


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
    """'SS27' → ('SS', '2027')."""
    s = (season_code or "").strip().upper()
    if len(s) < 2:
        return s, ""
    prefix = s[:2]
    tail = s[2:]
    if len(tail) == 2 and tail.isdigit():
        return prefix, f"20{tail}"
    if len(tail) == 4 and tail.isdigit():
        return prefix, tail
    return prefix, tail


def season_id_full(brand_code: str, season_code: str) -> str:
    pfx, yr = derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"


def _clean_str(v) -> str:
    """Clean and convert value to string. Handles whole-number floats by removing .0 suffix."""
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        # If float is a whole number (like 44727002.0), convert to int first to remove .0
        if isinstance(v, float) and not math.isnan(v) and v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if s in ("None", "nan", "NaN", "N/A", "#N/A"):
        return ""
    return s


def _extract_digits(code: str, max_len: int = 8) -> str:
    """Extract only digits from a code string, max length."""
    digits = re.sub(r"[^\d]", "", str(code or ""))
    return digits[:max_len]


# ══════════════════════════════════════════════════════════════════════════════
# DYNAMIC MAPPING RULE
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class MappingRule:
    """Represents a single attribute mapping rule from the Excel template."""
    stibo_attr_name: str
    stibo_attr_id: str
    validation_type: str
    cluster: str
    grouping: str
    description: str
    field_mapping_type: str
    source_field: str
    mapping_logic: str
    
    # Parsed mapping instructions
    is_direct_copy: bool = False
    default_value: str = ""
    is_rna_lookup: bool = False
    is_formula: bool = False
    formula_type: str = ""
    is_lov_lookup: bool = False
    
    def __post_init__(self):
        self._parse_mapping_logic()
    
    def _parse_mapping_logic(self):
        """Parse the mapping logic text to determine transformation type."""
        logic_lower = (self.mapping_logic or "").lower()
        src_lower = (self.source_field or "").lower()
        
        # Has a source field = direct copy (with possible transformation)
        if self.source_field and self.source_field not in ("None", "N/A", ""):
            self.is_direct_copy = True
        
        # Check for default values in logic
        default_patterns = [
            r"default\s*[:\-]?\s*['\"]?([^'\"(\n]+)['\"]?",
            r"default\s*[:\-]?\s*(.+?)(?:\s*\(|$|\n)",
        ]
        for pattern in default_patterns:
            match = re.search(pattern, logic_lower, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                # Clean up common suffixes
                val = re.sub(r"\s*\(.*$", "", val).strip()
                val = val.strip("'\"")
                if val and val not in ("", "none"):
                    self.default_value = val
                    break
        
        # Check for RNA-based lookups
        if "compcode" in logic_lower or "based on compcode" in logic_lower:
            self.is_rna_lookup = True
        
        # Check for LOV validation
        if self.validation_type.lower() == "lov":
            self.is_lov_lookup = True
        
        # Check for formula-based derivations
        if "formula" in logic_lower or "brand code" in logic_lower:
            self.is_formula = True
            if "generic" in self.stibo_attr_id.lower():
                self.formula_type = "generic_code"
            elif "variant" in self.stibo_attr_id.lower():
                self.formula_type = "variant_code"
            elif "description" in self.stibo_attr_id.lower():
                self.formula_type = "description"


# ══════════════════════════════════════════════════════════════════════════════
# DYNAMIC MAPPING LOADER
# ══════════════════════════════════════════════════════════════════════════════
class DynamicMappingLoader:
    """
    Loads mapping rules from the Brand Mapping Template Excel file.
    All attribute mappings are read dynamically from the DR.Martens(Inline) sheet.
    """
    
    MAPPING_SHEET_NAME = "DR.Martens(Inline)"
    MD_MAPPINGS_SHEET_NAME = "DR.Martens MD Mappings"
    
    # Column indices in the mapping template (0-based)
    COL_STIBO_ATTR_NAME = 0
    COL_STIBO_ATTR_ID = 1
    COL_VALIDATION = 2
    COL_CLUSTER = 3
    COL_GROUPING = 4
    COL_DESCRIPTION = 5
    COL_FIELD_MAPPING_TYPE = 6
    COL_BRAND_FILE_NAME = 7
    COL_BRAND_FILE_SHEET = 8
    COL_SOURCE_FIELD = 9
    COL_MAPPING_LOGIC = 10
    
    def __init__(self, mapping_file_path: Path):
        self.path = mapping_file_path
        self.rules: list[MappingRule] = []
        self.rules_by_id: dict[str, MappingRule] = {}
        self.md_mappings_by_gender: dict[str, str] = {}   # Gender → BY Gender (Col G)
        self.md_mappings_sap_gender: dict[str, str] = {}  # Gender → SAP Gender (Col F)
        self._load()
    
    def _load(self):
        log.info("[DynamicMapping] Loading from: %s", self.path.name)
        
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.error("[DynamicMapping] Failed to open workbook: %s", e)
            return
        
        if self.MAPPING_SHEET_NAME not in wb.sheetnames:
            log.error("[DynamicMapping] Sheet '%s' not found. Available: %s",
                      self.MAPPING_SHEET_NAME, wb.sheetnames)
            wb.close()
            return
        
        log.info("[DynamicMapping] Using sheet: '%s'", self.MAPPING_SHEET_NAME)
        ws = wb[self.MAPPING_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        
        # Skip header row (index 0), data starts at index 1
        for i, row in enumerate(rows[1:], start=1):
            if not row or not row[self.COL_STIBO_ATTR_ID]:
                continue
            
            attr_id = self._clean(row[self.COL_STIBO_ATTR_ID])
            if not attr_id or attr_id.startswith("#"):
                continue
            
            source_field = self._clean(row[self.COL_SOURCE_FIELD])
            # Clean up source field - handle "N/A", "None", etc.
            if source_field.lower() in ("n/a", "none", ""):
                source_field = ""
            # Handle multi-line source fields (e.g., "FW : Gender\nAcc : Default 'Unisex'")
            if "\n" in source_field:
                # Take the first non-conditional part
                parts = source_field.split("\n")
                for part in parts:
                    if ":" in part:
                        # Format like "FW : Gender" → extract "Gender"
                        subparts = part.split(":")
                        if len(subparts) >= 2:
                            source_field = subparts[1].strip()
                            break
                    elif part.strip() and part.strip().lower() not in ("acc", "fw"):
                        source_field = part.strip()
                        break
            
            rule = MappingRule(
                stibo_attr_name=self._clean(row[self.COL_STIBO_ATTR_NAME]),
                stibo_attr_id=attr_id,
                validation_type=self._clean(row[self.COL_VALIDATION]) or "text",
                cluster=self._clean(row[self.COL_CLUSTER]),
                grouping=self._clean(row[self.COL_GROUPING]),
                description=self._clean(row[self.COL_DESCRIPTION]),
                field_mapping_type=self._clean(row[self.COL_FIELD_MAPPING_TYPE]),
                source_field=source_field,
                mapping_logic=self._clean(row[self.COL_MAPPING_LOGIC]),
            )
            
            self.rules.append(rule)
            # Store by ID (first occurrence wins for duplicates like AT_Createdon)
            if attr_id not in self.rules_by_id:
                self.rules_by_id[attr_id] = rule
        
        self._load_md_mappings_gender(wb)
        wb.close()
        log.info("[DynamicMapping] Loaded %d mapping rules", len(self.rules))
        
        # Log rules with direct source field mappings
        direct_mappings = [r for r in self.rules if r.is_direct_copy]
        log.info("[DynamicMapping] %d rules with direct source field mappings", len(direct_mappings))
    
    def _load_md_mappings_gender(self, wb):
        """
        Load the 'DR.Martens MD Mappings' sheet for SAP Gender and BY Gender lookups.
        Column E (index 4) = Gender source value (WOMENS, UNISEX, GIRLS, KIDS)
        Column F (index 5) = SAP Gender LOV display value (Female, Unisex, etc.)
        Column G (index 6) = BY Gender LOV display value (Women, Unisex, Female, etc.)
        """
        if self.MD_MAPPINGS_SHEET_NAME not in wb.sheetnames:
            log.warning("[DynamicMapping] '%s' sheet not found, Gender mapping unavailable",
                        self.MD_MAPPINGS_SHEET_NAME)
            return
        
        ws = wb[self.MD_MAPPINGS_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        
        # Find the header row and data rows for Gender mapping
        for i, row in enumerate(rows):
            if not row or len(row) < 7:
                continue
            
            # Look for the mapping table (header with "Gender", "SAP Gender", "BY Gender")
            if (str(row[4] or "").strip().upper() == "GENDER" and 
                str(row[6] or "").strip().upper() == "BY GENDER"):
                # Found header row, load data from next rows
                for data_row in rows[i+1:]:
                    if not data_row or len(data_row) < 7:
                        continue
                    gender = self._clean(data_row[4])      # Column E — source Gender value
                    sap_gender = self._clean(data_row[5])  # Column F — SAP Gender display
                    by_gender = self._clean(data_row[6])   # Column G — BY Gender display
                    if gender:
                        key = gender.upper()
                        if sap_gender:
                            self.md_mappings_sap_gender[key] = sap_gender
                        if by_gender:
                            self.md_mappings_by_gender[key] = by_gender
                break
        
        log.info("[DynamicMapping] Loaded %d Gender → SAP Gender mappings from MD Mappings sheet",
                 len(self.md_mappings_sap_gender))
        log.info("[DynamicMapping] Loaded %d Gender → BY Gender mappings from MD Mappings sheet",
                 len(self.md_mappings_by_gender))
    
    @staticmethod
    def _clean(v) -> str:
        if v is None:
            return ""
        s = str(v).strip()
        if s in ("None", "nan", "\xa0", " "):
            return ""
        return s
    
    def resolve_sap_gender(self, gender: str) -> str:
        """Resolve Gender source value to SAP Gender display (Col F) via MD Mappings sheet."""
        if not gender:
            return ""
        return self.md_mappings_sap_gender.get(gender.upper(), "")

    def resolve_by_gender(self, gender: str) -> str:
        """Resolve Gender value to BY Gender via MD Mappings sheet."""
        if not gender:
            return ""
        return self.md_mappings_by_gender.get(gender.upper(), "")
    
    def get_rule(self, attr_id: str) -> MappingRule | None:
        return self.rules_by_id.get(attr_id)
    
    def get_rules_with_source(self) -> list[MappingRule]:
        """Return only rules that have a direct source field mapping."""
        return [r for r in self.rules if r.is_direct_copy and r.source_field]
    
    def get_rules_with_default(self) -> list[MappingRule]:
        """Return rules that have default values."""
        return [r for r in self.rules if r.default_value]
    
    def get_rules_with_rna(self) -> list[MappingRule]:
        """Return rules that require RNA lookup."""
        return [r for r in self.rules if r.is_rna_lookup]
    
    def get_all_source_fields(self) -> set[str]:
        """Return all unique source field names that need to be read from input."""
        return {r.source_field for r in self.rules if r.source_field}


# ══════════════════════════════════════════════════════════════════════════════
# RNA LOADER — Brand Type / Brand Category Lookup
# ══════════════════════════════════════════════════════════════════════════════
class RNALoader:
    """Loads Source Mapping Related RNA tab from the Attributes List workbook."""

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
        self.path = path
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
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found")
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
            log.warning("[RNA] Cannot find header row")
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
        c_comp = _find("COMPANY CODE", "COMP CODE", "COMPCODE")
        c_sbu = _find("SBU")
        c_bcode = _find("BRANDCODE", "BRAND CODE", "REPORTING BRAND CODE MAPPED")
        c_btype = _find("BRANDTYPE_DETAIL", "BRAND TYPE", "BRANDTYPE")
        c_bcat = _find("BRANDCATEGORY", "BRAND CATEGORY")
        c_bgroup = _find("BRANDGROUP", "BRAND GROUP")
        c_bstatus = _find("BRAND STATUS", "BRANDSTATUS")

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
                "BRANDGROUP": _cell(row, c_bgroup),
                "BRANDTYPE_DETAIL": _cell(row, c_btype),
                "BRANDCATEGORY": _cell(row, c_bcat),
                "BRAND STATUS": _cell(row, c_bstatus),
            }
            count += 1

        wb.close()
        log.info("[RNA] %d entries loaded", count)

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {})

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
        return {}


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER — Core Attributes + LOV Lookups
# ══════════════════════════════════════════════════════════════════════════════
class MDDLoader:
    """Loads Core Attributes + LOVs from the MDD Excel."""

    def __init__(self, path: Path):
        self.path = path
        self.attributes: dict[str, dict] = {}
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Load Core Attributes
        if "Core Attributes" in wb.sheetnames:
            ws = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 2:
                hdr = rows[1]
                col: dict[str, int] = {}
                for i, h in enumerate(hdr):
                    if h:
                        col[str(h).split("\n")[0].strip()] = i

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
                        "id": aid,
                        "name": _v("PIM Attribute Name"),
                        "validation": _v("Validation Base Type"),
                        "lov_name": _v("Name of LOV"),
                    }

        self._load_lov_sheets(wb)
        self._load_retail_price_currency(wb)
        self._load_gender_lov(wb)
        self._load_age_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_brand_group_lov(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper() and sn not in ("Brand", "Gender", "Age", "UOM", "Season"):
                continue

            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue

            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            sub = self.lovs.setdefault(display, {})

            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                v0 = str(row[0]).strip() if row[0] is not None else ""
                v1 = str(row[1]).strip() if row[1] is not None else ""

                if v0:
                    sub[v0] = v1 if v1 else v0
                    sub[v0.upper()] = v1 if v1 else v0
                if v1:
                    sub[v1] = v0 if v0 else v1
                    sub[v1.upper()] = v0 if v0 else v1

    def _load_retail_price_currency(self, wb):
        country_id_to_name: dict[str, str] = {}
        country_sheet = next(
            (s for s in wb.sheetnames
             if "COUNTRY" in s.upper() and "LOV" in s.upper()
             and "SIZE" not in s.upper() and "ORIGIN" not in s.upper()),
            None,
        )
        if country_sheet:
            for row in wb[country_sheet].iter_rows(min_row=2, values_only=True):
                if row and row[0] and len(row) > 1 and row[1]:
                    code = str(row[0]).strip().upper()
                    name = str(row[1]).strip()
                    if code and name:
                        country_id_to_name[code] = name

        retail_lov: dict[str, str] = {}
        currency_sheet = next(
            (s for s in wb.sheetnames if "RETAIL PRICE CURRENCY" in s.upper()),
            None,
        )
        if currency_sheet:
            for row in wb[currency_sheet].iter_rows(min_row=2, values_only=True):
                if row and row[0] and len(row) > 1 and row[1]:
                    name = str(row[0]).strip().upper()
                    curr = str(row[1]).strip()
                    if name and curr:
                        retail_lov[name] = curr

        self.lovs["CountryIdToName"] = country_id_to_name
        self.lovs["RetailPriceCurrency"] = retail_lov

    def _load_gender_lov(self, wb):
        """
        Load the 'Gender LOV' tab dynamically (no hardcoded mapping).

        Sheet layout:
            Col 0 = 'Values of LOV'          (e.g. "Male")
            Col 1 = 'Value ID of LOV'        (e.g. "M")  ← AT_Gender LOV ID
            Col 2 = 'SAP Gender Code'        (e.g. "M")

        Builds a case-insensitive lookup where the display value, the LOV ID
        and the SAP gender code all resolve to the AT_Gender LOV ID.
        """
        gender_lov: dict[str, str] = {}
        sheet = next(
            (s for s in wb.sheetnames
             if "GENDER" in s.upper() and "LOV" in s.upper()
             and "SIZE" not in s.upper()),
            None,
        )
        if sheet:
            for row in wb[sheet].iter_rows(values_only=True):
                if not row:
                    continue
                display = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                lov_id  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                sap     = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                # Skip header / non-data rows
                if not lov_id or lov_id.upper().startswith("VALUE ID"):
                    continue
                for key in (display, lov_id, sap):
                    if key:
                        gender_lov[key.upper()] = lov_id
        self.lovs["GenderLOV"] = gender_lov
        log.info("[MDD] Gender LOV loaded: %d entries", len(gender_lov))

    def resolve_gender_lov_id(self, value: str) -> str:
        """Resolve any source gender value to the AT_Gender LOV ID via the MDD Gender LOV tab."""
        if not value:
            return ""
        return self.lovs.get("GenderLOV", {}).get(str(value).strip().upper(), "")

    def _load_age_lov(self, wb):
        """
        Load the 'Age LOV' tab dynamically (no hardcoded mapping).

        Sheet layout:
            Col 0 = 'Age'                 (e.g. "Adults")
            Col 1 = 'Article Description' (e.g. "AD")  ← valid Stibo LOV_Age ID
            Col 2 = 'SAP Age Code'        (e.g. "A")

        Builds a case-insensitive lookup where the Age name, the Article
        Description code and the SAP Age Code all resolve to the Article
        Description code (the value Stibo's LOV_Age expects).
        """
        age_lov: dict[str, str] = {}
        sheet = next(
            (s for s in wb.sheetnames
             if "AGE" in s.upper() and "LOV" in s.upper()
             and "E-COM" not in s.upper() and "IMAGE" not in s.upper()),
            None,
        )
        if sheet:
            for row in wb[sheet].iter_rows(values_only=True):
                if not row:
                    continue
                age_name = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                art_desc = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                sap_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                # Skip header / non-data rows
                if not art_desc or art_desc.upper().startswith("ARTICLE"):
                    continue
                for key in (age_name, art_desc, sap_code):
                    if key:
                        age_lov[key.upper()] = art_desc
        self.lovs["AgeLOV"] = age_lov
        log.info("[MDD] Age LOV loaded: %d entries", len(age_lov))

    def resolve_age_lov_id(self, value: str) -> str:
        """Resolve any source age value to the LOV_Age ID (Article Description) via the MDD Age LOV tab."""
        if not value:
            return ""
        return self.lovs.get("AgeLOV", {}).get(str(value).strip().upper(), "")

    def _load_brand_status_lov(self, wb):
        """
        Load the 'Brand Status LOV' tab.

        Sheet layout (as shown in screenshot):
            Col A = Code          (e.g. "A", "D")   ← LOV ID to send to Stibo
            Col B = Brand Status  (e.g. "ACTIVE", "DISCONTINUED")  ← value from RNA

        Builds a lookup: Brand Status label (upper) → Code (Col A)
        so that resolve_brand_status_lov_id("ACTIVE") → "A"
        """
        brand_status_lov: dict[str, str] = {}
        sheet = next(
            (s for s in wb.sheetnames if "BRAND STATUS" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if sheet:
            for row in wb[sheet].iter_rows(values_only=True):
                if not row:
                    continue
                code   = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                label  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                # Skip header row
                if not code or code.upper() == "CODE":
                    continue
                if label:
                    brand_status_lov[label.upper()] = code
                # also allow lookup by code itself
                brand_status_lov[code.upper()] = code
        self.lovs["BrandStatusLOV"] = brand_status_lov
        log.info("[MDD] Brand Status LOV loaded: %d entries", len(brand_status_lov))

    def resolve_brand_status_lov_id(self, value: str) -> str:
        """
        Resolve a Brand Status label (e.g. 'ACTIVE') or code ('A')
        to the LOV Code (Col A) via the MDD 'Brand Status LOV' tab.
        """
        if not value:
            return ""
        return self.lovs.get("BrandStatusLOV", {}).get(str(value).strip().upper(), "")

    def _load_brand_type_lov(self, wb):
        """
        Load the 'Brand Type LOV' tab.

        Sheet layout:
            Col A = Code        (e.g. "MAA", "NONMAA")  ← LOV ID to send to Stibo
            Col B = Brand Type  (e.g. "MAA BRAND")      ← value from RNA BRANDTYPE_DETAIL

        Builds a lookup: Brand Type label (upper) → Code (Col A)
        so that resolve_brand_type_lov_id("MAA BRAND") → "MAA".
        """
        brand_type_lov: dict[str, str] = {}
        sheet = next(
            (s for s in wb.sheetnames if "BRAND TYPE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if sheet:
            for row in wb[sheet].iter_rows(values_only=True):
                if not row:
                    continue
                code  = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                label = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                # Skip header row
                if not code or code.upper() == "CODE":
                    continue
                if label:
                    brand_type_lov[label.upper()] = code
                # also allow lookup by code itself
                brand_type_lov[code.upper()] = code
        self.lovs["BrandTypeLOV"] = brand_type_lov
        log.info("[MDD] Brand Type LOV loaded: %d entries", len(brand_type_lov))

    def resolve_brand_type_lov_id(self, value: str) -> str:
        """
        Resolve a Brand Type label (e.g. 'MAA BRAND') or code ('MAA')
        to the LOV Code (Col A) via the MDD 'Brand Type LOV' tab.
        """
        if not value:
            return ""
        return self.lovs.get("BrandTypeLOV", {}).get(str(value).strip().upper(), "")

    def _load_brand_group_lov(self, wb):
        """
        Load the 'Brand Group LOV' tab.

        Sheet layout:
            Col A = Code         (e.g. "DR. MARTENS", "LOTTO")  ← LOV ID to send to Stibo
            Col B = Brand Group  (e.g. "DR. MARTENS", "LOTTO")  ← value from RNA BRANDGROUP

        Builds a lookup: Brand Group label (upper) → Code (Col A)
        so that resolve_brand_group_lov_id("DR. MARTENS") → "DR. MARTENS".
        """
        brand_group_lov: dict[str, str] = {}
        sheet = next(
            (s for s in wb.sheetnames if "BRAND GROUP" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if sheet:
            for row in wb[sheet].iter_rows(values_only=True):
                if not row:
                    continue
                code  = str(row[0]).strip() if len(row) > 0 and row[0] else ""
                label = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                # Skip header row
                if not code or code.upper() == "CODE":
                    continue
                if label:
                    brand_group_lov[label.upper()] = code
                # also allow lookup by code itself
                brand_group_lov[code.upper()] = code
        self.lovs["BrandGroupLOV"] = brand_group_lov
        log.info("[MDD] Brand Group LOV loaded: %d entries", len(brand_group_lov))

    def resolve_brand_group_lov_id(self, value: str) -> str:
        """
        Resolve a Brand Group label (e.g. 'DR. MARTENS') or code
        to the LOV Code (Col A) via the MDD 'Brand Group LOV' tab.
        """
        if not value:
            return ""
        return self.lovs.get("BrandGroupLOV", {}).get(str(value).strip().upper(), "")

    def lookup_lov(self, lov_name: str, value: str) -> str:
        if not value:
            return ""
        lov = self.lovs.get(lov_name, {})
        return lov.get(value, lov.get(value.upper(), value))

    def resolve_currency_from_country_code(self, country_code: str) -> str:
        if not country_code:
            return "USD"
        code = country_code.strip().upper()
        name = self.lovs.get("CountryIdToName", {}).get(code, "")
        if not name:
            return COUNTRY_CURRENCY.get(code, "USD")
        curr = self.lovs.get("RetailPriceCurrency", {}).get(name.upper(), "")
        return curr if curr else COUNTRY_CURRENCY.get(code, "USD")


# ══════════════════════════════════════════════════════════════════════════════
# SOURCE DATA LOADER — Retail Price Master
# ══════════════════════════════════════════════════════════════════════════════
class RetailPriceMasterLoader:
    """Loads the DR Marten Retail Price Master source data."""

    HEADER_ROW_INDEX = 6  # 0-based (row 7 in Excel)

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.headers: list[str] = []
        self.column_map: dict[str, str] = {}  # normalized_name → actual_column_name
        self._load()

    def _load(self):
        log.info("[Source] Loading: %s", self.path.name)

        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = wb.sheetnames[0] if wb.sheetnames else None
        if not sheet_name:
            log.error("[Source] No sheets found")
            wb.close()
            return

        log.info("[Source] Using sheet: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        if len(rows) <= self.HEADER_ROW_INDEX + 1:
            log.error("[Source] Not enough rows")
            wb.close()
            return

        # Get headers and handle duplicates
        raw_headers = rows[self.HEADER_ROW_INDEX]
        seen = {}
        unique_headers = []
        for i, h in enumerate(raw_headers):
            col_name = _clean_str(h) if h else f"col_{i}"
            if col_name in seen:
                seen[col_name] += 1
                col_name = f"{col_name}_{seen[col_name]}"
            else:
                seen[col_name] = 0
            unique_headers.append(col_name)
            # Build column map for flexible lookups
            normalized = col_name.lower().strip()
            self.column_map[normalized] = col_name

        self.headers = unique_headers
        log.info("[Source] Headers: %s", self.headers[:15])
        
        # DEBUG: Check for Product Code columns
        product_code_cols = [h for h in self.headers if "Product Code" in h]
        if product_code_cols:
            log.info("[DEBUG] ═══════════════════════════════════════════════════════")
            log.info("[DEBUG] Found %d 'Product Code' column(s): %s", 
                     len(product_code_cols), product_code_cols)
            log.info("[DEBUG] ═══════════════════════════════════════════════════════")

        # Create DataFrame
        data_rows = rows[self.HEADER_ROW_INDEX + 1:]
        self.df = pd.DataFrame(data_rows, columns=self.headers)
        
        # DEBUG: Show raw data types and values from Product Code columns
        for col in product_code_cols:
            if col in self.df.columns:
                log.info("[DEBUG] Column '%s' - First 5 raw values from Excel:", col)
                for idx in range(min(5, len(self.df))):
                    val = self.df.loc[idx, col]
                    log.info("[DEBUG]   Row %d: value=%s | type=%s | repr=%r", 
                             idx, val, type(val).__name__, val)

        # Filter empty rows
        key_col = "Product Code"
        if key_col in self.df.columns:
            col_data = self.df[key_col]
            if isinstance(col_data, pd.DataFrame):
                col_data = col_data.iloc[:, 0]
            key_series = col_data.astype(str).str.strip()
            self.df = self.df[col_data.notna() & (~key_series.isin(["", "None", "nan"]))]

        self.df = self.df.reset_index(drop=True)
        wb.close()
        log.info("[Source] %d rows loaded", len(self.df))

    def get_rows(self) -> list[dict]:
        return self.df.to_dict(orient="records")

    def find_column(self, field_name: str) -> str | None:
        """Find the actual column name for a field, handling variations."""
        if not field_name:
            return None
        
        # Direct match (exact)
        if field_name in self.headers:
            return field_name
        
        # Normalized exact match
        normalized = field_name.lower().strip()
        if normalized in self.column_map:
            return self.column_map[normalized]
        
        # Case-insensitive exact match (check all headers)
        for header in self.headers:
            if header.lower().strip() == normalized:
                return header
        
        # Partial match ONLY if the lengths are similar (avoid "Category" matching "Product Category")
        # Only allow partial match if searching for a longer string that contains a shorter header
        # NOT the other way around (which causes the bug)
        for header in self.headers:
            header_norm = header.lower().strip()
            # Only match if the field_name contains the header AND they're reasonably similar length
            # This prevents "Product Category" from matching "Category"
            if header_norm in normalized and len(header_norm) >= len(normalized) * 0.7:
                return header
        
        # Handle common typos (Product Categoty vs Product Category)
        for header in self.headers:
            header_norm = header.lower().strip()
            # Levenshtein-style: if only 1-2 chars different
            if abs(len(header_norm) - len(normalized)) <= 2:
                # Check character-by-character similarity
                matches = sum(1 for a, b in zip(header_norm, normalized) if a == b)
                if matches >= len(normalized) - 2 and matches >= len(header_norm) - 2:
                    return header
        
        return None


# ══════════════════════════════════════════════════════════════════════════════
# DYNAMIC TRANSFORMATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
class DynamicTransformationEngine:
    """
    Applies mapping rules dynamically to transform source data.
    ALL transformations are driven by the mapping Excel — nothing hardcoded.
    """
    
    def __init__(
        self,
        mapping: DynamicMappingLoader,
        mdd: MDDLoader | None,
        rna: RNALoader | None,
        source_loader: RetailPriceMasterLoader,
        context: dict,
    ):
        self.mapping = mapping
        self.mdd = mdd
        self.rna = rna
        self.source = source_loader
        self.context = context  # brand_code, country_code, comp_code, sbu, season, etc.
        
        # Cache RNA data
        self._rna_cache: dict = {}
        if rna:
            country_code = context.get("country_code", "")
            country_name = _COUNTRY_MAP.get(country_code.upper(), country_code.upper()) if country_code else ""
            comp_code    = context.get("comp_code", "")
            sbu          = context.get("sbu", "")
            brand        = context.get("brand_code", "")

            # Try 1: exact match with comp_code as-is (e.g. "0888")
            result = rna.get(country_name, comp_code, sbu, brand)
            # Try 2: comp_code stripped of leading zeros (e.g. "888")
            if not any(result.values()) and comp_code.strip().isdigit():
                result = rna.get(country_name, str(int(comp_code)), sbu, brand)
            # Try 3: fuzzy — country + sbu + brand (ignores comp_code)
            if not any(result.values()):
                result = rna.get_fuzzy(country_name, sbu, brand)
            # Try 4: fuzzy — country + brand only
            if not any(result.values()):
                result = rna.get_fuzzy(country_name, "", brand)

            self._rna_cache = result
    
    def transform_row(self, row: dict) -> dict:
        """
        Transform a single source row into mapped article attributes.
        All mappings are driven by the rules loaded from Excel.
        """
        attrs = {}
        
        # Process each mapping rule
        for rule in self.mapping.rules:
            value = self._apply_rule(rule, row, attrs)
            if value:
                attrs[rule.stibo_attr_id] = value
        
        # Generate derived attributes (Generic, Variant, etc.)
        self._apply_formulas(attrs, row)
        
        # Add context attributes (these are always set)
        attrs["AT_Country"] = self.context.get("country_code", "")
        attrs["AT_CompanyCode"] = self.context.get("comp_code", "")
        attrs["AT_SBU"] = self.context.get("sbu", "")
        attrs["AT_Brand"] = self.context.get("brand_code", BRAND_CODE)
        
        return attrs
    
    def _apply_rule(self, rule: MappingRule, row: dict, attrs: dict) -> str:
        """Apply a single mapping rule and return the resulting value."""
        
        # 1. RNA-based lookup (BrandGroup, BrandType, BrandCategory, BrandStatus)
        if rule.is_rna_lookup and rule.source_field:
            rna_key = rule.source_field.upper().replace(" ", "_")
            # Map source field names to RNA keys
            rna_key_map = {
                "BRANDGROUP": "BRANDGROUP",
                "BRANDTYPE_DETAIL": "BRANDTYPE_DETAIL",
                "BRANDCATEGORY": "BRANDCATEGORY",
                "BRAND_STATUS": "BRAND STATUS",
                "BRAND STATUS": "BRAND STATUS",
            }
            actual_key = rna_key_map.get(rna_key, rna_key)
            return self._rna_cache.get(actual_key, "")
        
        # 2. Direct field copy from source
        if rule.is_direct_copy and rule.source_field:
            actual_col = self.source.find_column(rule.source_field)
            if actual_col:
                raw_value = row.get(actual_col, "")
                
                # DEBUG: Log Product Code extraction
                if rule.stibo_attr_id == "AT_PrincipalStyleCode":
                    log.info("[DEBUG] ═══════════════════════════════════════════════════════")
                    log.info("[DEBUG] AT_PrincipalStyleCode Mapping:")
                    log.info("[DEBUG]   Source field from mapping: '%s'", rule.source_field)
                    log.info("[DEBUG]   Actual column found: '%s'", actual_col)
                    log.info("[DEBUG]   Raw value from Excel: %s (type=%s)", raw_value, type(raw_value).__name__)
                    log.info("[DEBUG] ═══════════════════════════════════════════════════════")
                
                return self._transform_value(rule, raw_value)
        
        # 3. Default value from mapping logic
        if rule.default_value:
            return self._normalize_default(rule)
        
        # 4. Skip if no mapping available
        return ""
    
    def _transform_value(self, rule: MappingRule, raw_value) -> str:
        """Transform a raw value based on the rule's validation type and mapping logic."""
        value = _clean_str(raw_value)
        if not value:
            return rule.default_value if rule.default_value else ""
        
        attr_id = rule.stibo_attr_id.upper()
        
        # Force specific LOV values regardless of source data
        if attr_id == "AT_PRICINGDISTRIBUTIONCHANNEL":
            return "01"
        
        # Gender transformation
        if "GENDER" in attr_id:
            mapped = self.mdd.resolve_gender_lov_id(value)
            # AT_Gender handled separately in _apply_formulas (conditional on Category column)
            if attr_id == "AT_GENDER":
                return ""
            if attr_id == "AT_SAPGENDER":
                return mapped
            # AT_BYGender handled separately in _apply_formulas (conditional on Category column)
            elif attr_id == "AT_BYGENDER":
                return ""
            return value
        
        # Age transformation
        if "AGE" in attr_id:
            if attr_id == "AT_BYAGE":
                # Default to Adult for DR Marten
                return "ADULT"
            # AT_SAPAge handled separately in _apply_formulas (conditional on Category column)
            if attr_id == "AT_SAPAGE":
                return ""
            return value
        
        # Season transformation
        if attr_id == "AT_SEASON":
            pfx, _ = derive_season(value)
            return pfx
        if attr_id == "AT_SEASONYEAR":
            _, yr = derive_season(value)
            return yr
        
        # Numeric fields (FOB, prices)
        if "PRICE" in attr_id or "FOB" in attr_id or "COST" in attr_id:
            try:
                return str(float(value))
            except (ValueError, TypeError):
                return ""
        
        # Style code truncation
        if attr_id == "AT_SAPSTYLECODE":
            return value[:9] if value else ""
        
        return value
    
    def _normalize_default(self, rule: MappingRule) -> str:
        """Normalize default values from mapping logic."""
        default = rule.default_value.lower()
        attr_id = rule.stibo_attr_id.upper()
        
        # Specific attribute defaults (highest priority)
        if attr_id == "AT_PRICINGDISTRIBUTIONCHANNEL":
            return "01"
        
        if attr_id == "AT_SAPPRODUCTFLAG":
            return "A"
        
        if attr_id == "AT_PACKDETAILS":
            return "S"
        
        if attr_id == "AT_ARTICLESTATUS":
            return "A"
        
        # Common default normalizations
        if "adult" in default:
            if "BYAGE" in attr_id:
                return "ADULT"
            return "AD"
        
        if "unisex" in default:
            if "BYGENDER" in attr_id:
                return "Unisex"
            return "U"
        
        if "ea" in default:
            return "EA"
        
        if "usd" in default:
            return "USD"
        
        if "generic" in default:
            return "1"
        
        if ("intercompany" in default or "zina" in default) and "MATERIALTYPE" in attr_id:
            return "ZINA"
        
        if "inline" in default:
            return "Inline"
        
        if "retailer" in default:
            return "01"
        
        # Return default as-is with first letter capitalized
        return rule.default_value.strip()
    
    def _apply_formulas(self, attrs: dict, row: dict):
        """Apply formula-based derivations for Generic, Variant, Description, etc."""
        
        # Get base values
        product_code = attrs.get("AT_PrincipalStyleCode", "")
        product_name = attrs.get("AT_PrincipalStyleDescription", "")
        
        # Generate Generic code: BRAND_CODE + Product Code (digits only, max 8)
        if product_code:
            product_digits = _extract_digits(product_code, 8)
            generic_code = f"{BRAND_CODE}{product_digits}"
            attrs["AT_Generic"] = generic_code
            attrs["AT_GenericDescription"] = f"{BRAND_CODE} {product_name}"[:80]
        
        # Retail Price Currency (from MDD based on country)
        if self.mdd and not attrs.get("AT_RetailPriceCurrency"):
            country_code = self.context.get("country_code", "")
            currency = self.mdd.resolve_currency_from_country_code(country_code)
            attrs["AT_RetailPriceCurrency"] = currency if currency else "USD"
        
        # Season — context (parsed from filename) always takes precedence
        ctx_season = self.context.get("season", "")
        if ctx_season:
            pfx, yr = derive_season(ctx_season)
            attrs["AT_Season"] = pfx
            attrs["AT_SeasonYear"] = yr
        else:
            # Fallback: extract from source data "Launch Season" column
            season_field = self.source.find_column("Launch Season")
            if season_field:
                season_value = _clean_str(row.get(season_field, ""))
                if season_value:
                    pfx, yr = derive_season(season_value)
                    if not attrs.get("AT_Season"):
                        attrs["AT_Season"] = pfx
                    if not attrs.get("AT_SeasonYear"):
                        attrs["AT_SeasonYear"] = yr
        
        # Copy CurrentPrice from OriginalPrice if not set
        if attrs.get("AT_OriginalPrice") and not attrs.get("AT_CurrentPrice"):
            attrs["AT_CurrentPrice"] = attrs["AT_OriginalPrice"]

        # SAP Product Division / ParentID source — from the "Product General" column.
        #   'DR_MARTENS MD Mappings' tab: FW → "F - FOOTWEAR", NFW → "E - ACCESSORIES".
        #   The resolved single-letter division is stored internally and consumed by
        #   XMLGenerator._determine_parent_id (stripped before XML output).
        pg_field = self.source.find_column("Product General")
        div_letter = ""
        pg_raw = ""
        if pg_field:
            pg_raw = _clean_str(row.get(pg_field, "")).upper()
            if pg_raw:
                # 1. Direct Product General code (FW / NFW)
                div_letter = PRODUCT_GENERAL_TO_DIVISION.get(pg_raw, "")
                # 2. Fallback — Product General already holds a Division name
                if not div_letter:
                    for name, letter in DIVISION_PARENT_MAP.items():
                        if name in pg_raw:
                            div_letter = letter
                            break
        attrs["_sap_product_division"] = div_letter  # internal — stripped before output

        # Derive category from Product General (FW → footwear, NFW → accessories)
        # This drives AT_SAPAge, AT_Gender, AT_CountrySize and the internal _category flag.
        if div_letter == "F" or pg_raw == "FW":
            category_value = "footwear"
        elif div_letter == "E" or pg_raw == "NFW":
            category_value = "accessories"
        else:
            category_value = pg_raw.lower()   # pass through whatever is there
        attrs["_category"] = category_value   # internal — used by XMLGenerator, stripped before output

        # AT_SAPAge — conditional on Product General derived category
        #   - footwear OR accessories → AT_SAPAge = Adult (LOV ID via MDD Age LOV → "AD")
        #   - anything else           → do NOT emit AT_SAPAge
        if category_value in ("footwear", "accessories"):
            adult_id = self.mdd.resolve_age_lov_id("Adults") if self.mdd else ""
            attrs["AT_SAPAge"] = adult_id or "AD"
        else:
            attrs.pop("AT_SAPAge", None)

        # AT_Gender — conditional on Product General derived category
        #   - accessories → hardcoded 'U' (Unisex)
        #   - footwear    → Gender column → MD Mappings Col F (SAP Gender) → MDD Gender LOV ID
        #                   e.g. "WOMENS" → Col F: "Female" → MDD LOV → "F"
        #   - anything else → do NOT emit AT_Gender
        if category_value == "accessories":
            attrs["AT_Gender"] = "U"
        elif category_value == "footwear":
            gender_field = self.source.find_column("Gender")
            gender_raw = _clean_str(row.get(gender_field, "")) if gender_field else ""
            if gender_raw:
                # Step 1: Lookup Col F (SAP Gender display) from MD Mappings
                sap_gender_display = self.mapping.resolve_sap_gender(gender_raw)
                log.info("[AT_Gender] '%s' → MD Mappings Col F: '%s'", gender_raw, sap_gender_display)
                if sap_gender_display:
                    # Step 2: Resolve SAP Gender display to LOV ID via MDD Gender LOV
                    gender_code = self.mdd.resolve_gender_lov_id(sap_gender_display) if self.mdd else ""
                    log.info("[AT_Gender] MDD Gender LOV: '%s' → '%s'", sap_gender_display, gender_code)
                    if gender_code:
                        attrs["AT_Gender"] = gender_code
                    else:
                        attrs.pop("AT_Gender", None)
                else:
                    attrs.pop("AT_Gender", None)
            else:
                attrs.pop("AT_Gender", None)
        else:
            attrs.pop("AT_Gender", None)
        
        # AT_BYGender — conditional on Product General derived category
        #   - accessories → hardcoded 'U' (Unisex)
        #   - footwear    → Gender column → MD Mappings lookup → MDD Gender LOV resolution
        #   - anything else → do NOT emit AT_BYGender
        log.info("[DEBUG BY_GENDER] ═══════════════════════════════════════════════════════")
        log.info("[DEBUG BY_GENDER] Starting AT_BYGender resolution")
        log.info("[DEBUG BY_GENDER] Category value: '%s'", category_value)
        log.info("[DEBUG BY_GENDER] Product Code: '%s'", product_code)
        
        if category_value == "accessories":
            log.info("[DEBUG BY_GENDER] Category is 'accessories' → Setting AT_BYGender = 'U'")
            attrs["AT_BYGender"] = "U"
        elif category_value == "footwear":
            log.info("[DEBUG BY_GENDER] Category is 'footwear' → Starting Gender lookup")
            gender_field = self.source.find_column("Gender")
            log.info("[DEBUG BY_GENDER] Gender column found: '%s'", gender_field)
            
            gender_raw = _clean_str(row.get(gender_field, "")) if gender_field else ""
            log.info("[DEBUG BY_GENDER] Raw Gender value from source: '%s'", gender_raw)
            
            if gender_raw:
                # Step 1: Lookup in MD Mappings (Gender → BY Gender)
                by_gender_value = self.mapping.resolve_by_gender(gender_raw)
                log.info("[DEBUG BY_GENDER] Step 1 - MD Mappings lookup: '%s' → '%s'", 
                         gender_raw, by_gender_value)
                
                if by_gender_value:
                    # Step 2: Resolve BY Gender value to LOV ID via MDD Gender LOV
                    by_gender_code = self.mdd.resolve_gender_lov_id(by_gender_value) if self.mdd else ""
                    log.info("[DEBUG BY_GENDER] Step 2 - MDD Gender LOV resolution: '%s' → '%s'", 
                             by_gender_value, by_gender_code)
                    
                    if by_gender_code:
                        log.info("[DEBUG BY_GENDER] ✓ SUCCESS - AT_BYGender set to: '%s'", by_gender_code)
                        attrs["AT_BYGender"] = by_gender_code
                    else:
                        log.warning("[DEBUG BY_GENDER] ✗ FAILED - MDD Gender LOV resolution returned empty")
                        log.warning("[DEBUG BY_GENDER] MDD available: %s", self.mdd is not None)
                        if self.mdd:
                            log.warning("[DEBUG BY_GENDER] GenderLOV entries: %s", 
                                       list(self.mdd.lovs.get("GenderLOV", {}).keys())[:10])
                        attrs.pop("AT_BYGender", None)
                else:
                    log.warning("[DEBUG BY_GENDER] ✗ FAILED - MD Mappings lookup returned empty")
                    log.warning("[DEBUG BY_GENDER] Available MD Gender mappings: %s", 
                               dict(list(self.mapping.md_mappings_by_gender.items())[:10]))
                    attrs.pop("AT_BYGender", None)
            else:
                log.warning("[DEBUG BY_GENDER] ✗ FAILED - Gender field is empty or not found")
                attrs.pop("AT_BYGender", None)
        else:
            log.info("[DEBUG BY_GENDER] Category is '%s' (not footwear/accessories) → Not emitting AT_BYGender", 
                     category_value)
            attrs.pop("AT_BYGender", None)
        
        log.info("[DEBUG BY_GENDER] Final AT_BYGender in attrs: %s", attrs.get("AT_BYGender", "NOT SET"))
        log.info("[DEBUG BY_GENDER] ═══════════════════════════════════════════════════════")

        # Set SAP Article Category default (Generic)
        if not attrs.get("AT_SAPArticleCategory"):
            attrs["AT_SAPArticleCategory"] = "1"
        
        # Set UOM default
        if not attrs.get("AT_UOM"):
            attrs["AT_UOM"] = "EA"
        
        # Set FOB Currency default
        if not attrs.get("AT_FOBCurrency"):
            attrs["AT_FOBCurrency"] = "USD"
        
        # Set Material Type default
        if not attrs.get("AT_MaterialType"):
            attrs["AT_MaterialType"] = "ZINA"
        
        # Set SAP Product Flag default
        if not attrs.get("AT_SAPProductFlag"):
            attrs["AT_SAPProductFlag"] = "A"
        
        # Set BY Article Type default
        if not attrs.get("AT_BYArticleType"):
            attrs["AT_BYArticleType"] = "Inline"

        # AT_BrandType — RNA 'BRANDTYPE_DETAIL' label → resolved to LOV Code via MDD Brand Type LOV
        #   e.g. RNA gives 'MAA BRAND' → MDD Brand Type LOV → Col A = 'MAA'
        brand_type_raw = self._rna_cache.get("BRANDTYPE_DETAIL", "")
        if brand_type_raw and self.mdd:
            brand_type_id = self.mdd.resolve_brand_type_lov_id(brand_type_raw)
            if brand_type_id:
                attrs["AT_BrandType"] = brand_type_id

        # AT_BrandStatus — RNA 'BRAND STATUS' label → resolved to LOV Code via MDD Brand Status LOV
        #   e.g. RNA gives 'ACTIVE' → MDD Brand Status LOV → Col A = 'A'
        brand_status_raw = self._rna_cache.get("BRAND STATUS", "")
        if brand_status_raw and self.mdd:
            brand_status_id = self.mdd.resolve_brand_status_lov_id(brand_status_raw)
            if brand_status_id:
                attrs["AT_BrandStatus"] = brand_status_id

        # AT_BrandGroup — RNA 'BRANDGROUP' label → resolved to LOV Code via MDD Brand Group LOV
        #   e.g. RNA gives 'DR. MARTENS' → MDD Brand Group LOV → Col A = 'DR. MARTENS'
        brand_group_raw = self._rna_cache.get("BRANDGROUP", "")
        if brand_group_raw and self.mdd:
            brand_group_id = self.mdd.resolve_brand_group_lov_id(brand_group_raw)
            if brand_group_id:
                attrs["AT_BrandGroup"] = brand_group_id


# ══════════════════════════════════════════════════════════════════════════════
# XML GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
class XMLGenerator:
    """Generates Stibo STEP XML from mapped articles (Nike 360 pattern)."""

    def __init__(self, articles: list[dict], context: dict):
        self.articles = articles
        self.context = context
        self.season_code = context.get("season", "SS27")

    def build_classifications(self) -> ET.Element:
        """Build Classifications with Season, Confirmed, and Unconfirmed sub-classifications (Lotto pattern)."""
        cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

        # Parse season code
        sea_prefix = self.season_code[:2].upper() if len(self.season_code) >= 2 else self.season_code
        sea_year_short = self.season_code[2:] if len(self.season_code) > 2 else ""
        sea_year = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

        # Build season IDs
        full_season_code = f"{sea_prefix}{sea_year}"
        season_id = f"CLH_{BRAND_CODE}_{full_season_code}"
        batches_parent = "CLH_DrMartensBatches"  # Parent must be "CLH_{Brand}Batches" format

        # Season display names
        season_label_map = {
            "SS": "Spring Summer", "FW": "Fall Winter",
            "AW": "Autumn Winter", "HO": "Holiday",
            "SP": "Spring", "SM": "Summer",
        }
        sea_name = season_label_map.get(sea_prefix, sea_prefix)
        season_display = f"{BRAND_NAME} {sea_name} {sea_year}".strip()
        season_short = f"{sea_prefix} {sea_year}".strip()

        # Main season classification
        season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
        season_cls.set("ID", season_id)
        season_cls.set("UserTypeID", "CLS_Season")
        season_cls.set("ParentID", batches_parent)
        ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

        # Confirmed Articles sub-classification
        confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        confirmed.set("ID", f"{season_id}CA")
        confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
        ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

        # Unconfirmed Articles sub-classification
        unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
        unconfirmed.set("ID", f"{season_id}UA")
        unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
        ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

        return cls_root

    def _determine_parent_id(self, art: dict) -> str:
        """
        Determine ParentID = PPH_{division_letter}-TempSubCat.

        The division letter is resolved in DynamicMapper._apply_formulas from the
        'Product General' input column via the 'DR_MARTENS MD Mappings' tab:
            FW  → F (Footwear)   NFW → E (Accessories)
        and mapped per the Products ParentID rule (E/F/A/Q/T). Any unknown /
        'TEST DIV' value falls back to X.
        """
        div_letter = (art.get("_sap_product_division", "") or "").strip().upper()
        if not div_letter:
            div_letter = "X"  # TEST DIV / unknown fallback
        return f"PPH_{div_letter}-TempSubCat"

    def _analyze_style_colors(self) -> dict[str, set[str]]:
        """
        Analyze all articles to find which style codes have multiple colors.
        Returns a dict mapping style_code -> set of color_codes.
        """
        style_colors: dict[str, set[str]] = {}
        for art in self.articles:
            style_code = art.get("AT_PrincipalStyleCode", "")
            color_code = art.get("AT_PrincipalColorCode", "") or ""
            if style_code:
                if style_code not in style_colors:
                    style_colors[style_code] = set()
                if color_code:
                    style_colors[style_code].add(color_code)
        return style_colors

    def _build_inbound_key(self, art: dict, style_colors: dict[str, set[str]]) -> str:
        """
        Build KEY_InboundArticle (= AT_InboundGenericCode).
        
        Logic:
          - If one Principal Style Code has ONE color:
              AT_InboundGenericCode = Brand Code + Principal Style Code
          
          - If one Principal Style Code has MULTIPLE colors:
              AT_InboundGenericCode = Brand Code + Principal Style Code + Principal Color Code
        """
        style_code = art.get("AT_PrincipalStyleCode", "")
        color_code = art.get("AT_PrincipalColorCode", "") or ""
        
        if not style_code:
            return ""
        
        # Check if this style has multiple colors
        colors_for_style = style_colors.get(style_code, set())
        has_multiple_colors = len(colors_for_style) > 1
        
        if has_multiple_colors and color_code:
            # Multiple colors for this style → include color code
            return f"{BRAND_CODE}{style_code}{color_code}"
        else:
            # Single color (or no color) → just brand + style
            return f"{BRAND_CODE}{style_code}"

    def build_products(self) -> ET.Element:
        """Build Products following Nike 360 pattern."""
        products_root = ET.Element(f"{{{STIBO_NS}}}Products")
        season_id = season_id_full(BRAND_CODE, self.season_code)
        
        # Analyze which styles have multiple colors
        style_colors = self._analyze_style_colors()

        for art in self.articles:
            # Build the inbound key using the correct logic
            key_article = self._build_inbound_key(art, style_colors)
            if not key_article or key_article == BRAND_CODE:
                continue

            # Determine parent based on product type
            parent_id = self._determine_parent_id(art)

            # Product element (no ID attribute, like Nike)
            product = ET.SubElement(products_root, f"{{{STIBO_NS}}}Product")
            product.set("UserTypeID", "PRD_GenericArticle")
            product.set("ParentID", parent_id)

            # KeyValue with KEY_InboundArticle
            kv = ET.SubElement(product, f"{{{STIBO_NS}}}KeyValue")
            kv.set("KeyID", "KEY_InboundArticle")
            kv.text = key_article

            # Name element (from AT_PrincipalStyleDescription)
            name_text = art.get("AT_PrincipalStyleDescription", "")
            if name_text:
                ET.SubElement(product, f"{{{STIBO_NS}}}Name").text = name_text

            # Classification Reference - Merchandiser (Lotto pattern)
            cr_merch = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
            cr_merch.set("ClassificationID", "CLH_DrMartensArticles")
            cr_merch.set("Type", "CPL_Merchandiser")

            # Classification Reference - UnConfirmed for Season (Lotto pattern)
            cr_unconf = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
            cr_unconf.set("ClassificationID", f"{season_id}UA")
            cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

            # Values element with all attributes
            values = ET.SubElement(product, f"{{{STIBO_NS}}}Values")
            
            # Add AT_InboundGenericCode (same as KEY_InboundArticle) — written explicitly here
            _val(values, "AT_InboundGenericCode", key_article)
            
            # Add multi-valued attributes (AT_CompanyCode, AT_SBU use MultiValue wrapper)
            _multival(values, "AT_CompanyCode", art.get("AT_CompanyCode", ""))
            _multival(values, "AT_SBU", art.get("AT_SBU", ""))
            
            # Add LOV-based attributes with ID (AT_Country, AT_Brand)
            country_code = art.get("AT_Country", "")
            if country_code:
                _val(values, "AT_Country", "", id_val=country_code)
            
            brand_code = art.get("AT_Brand", BRAND_CODE)
            if brand_code:
                _val(values, "AT_Brand", "", id_val=brand_code)
            
            # AT_BrandGroup — LOV attribute, send as ID (from RNA BRANDGROUP → MDD Brand Group LOV)
            brand_group = art.get("AT_BrandGroup", "")
            if brand_group:
                _val(values, "AT_BrandGroup", "", id_val=brand_group)

            # AT_BrandType — LOV attribute, send as ID (RNA BRANDTYPE_DETAIL is the LOV code)
            brand_type = art.get("AT_BrandType", "")
            if brand_type:
                _val(values, "AT_BrandType", "", id_val=brand_type)

            # AT_BrandStatus — LOV attribute, send as ID (resolved from RNA Brand Status via MDD LOV)
            #   e.g. RNA 'ACTIVE' → MDD Brand Status LOV → ID='A'
            brand_status = art.get("AT_BrandStatus", "")
            if brand_status:
                _val(values, "AT_BrandStatus", "", id_val=brand_status)

            # AT_PricingDistributionChannel — LOV attribute, must use ID not text
            dist_channel = art.get("AT_PricingDistributionChannel", "01")
            if dist_channel:
                _val(values, "AT_PricingDistributionChannel", "", id_val=dist_channel)
            
            # AT_FOBCurrency — always USD for DR Marten
            _val(values, "AT_FOBCurrency", "", id_val="USD")

            # AT_RetailPriceCurrency — derived from country code
            if country_code:
                currency_id = COUNTRY_CURRENCY.get(country_code.strip().upper(), "USD")
                _val(values, "AT_RetailPriceCurrency", "", id_val=currency_id)
            
            # AT_SAPProductFlag — LOV attribute, send as ID not text
            sap_flag = art.get("AT_SAPProductFlag", "A")
            if sap_flag:
                _val(values, "AT_SAPProductFlag", "", id_val=sap_flag)
            
            # AT_MaterialType — LOV attribute, send as ID not text (e.g., "ZINA")
            material_type = art.get("AT_MaterialType", "ZINA")
            if material_type:
                _val(values, "AT_MaterialType", "", id_val=material_type)
            
            # AT_SAPArticleCategory — LOV attribute, send as ID not text (e.g., "1")
            sap_article_cat = art.get("AT_SAPArticleCategory", "1")
            if sap_article_cat:
                _val(values, "AT_SAPArticleCategory", "", id_val=sap_article_cat)
            
            # AT_UOM — LOV attribute, send as ID not text (e.g., "EA")
            uom = art.get("AT_UOM", "EA")
            if uom:
                _val(values, "AT_UOM", "", id_val=uom)
            
            # AT_BYAge — LOV attribute, send as ID not text (e.g., "ADULT")
            by_age = art.get("AT_BYAge", "ADULT")
            if by_age:
                _val(values, "AT_BYAge", "", id_val=by_age)
            
            # AT_PackDetails — LOV attribute, send as ID not text
            # Not applicable for Footwear — skip entirely for FW articles
            category_val = str(art.get("_category", "")).lower()
            pack_details = art.get("AT_PackDetails", "S")
            if pack_details and "footwear" not in category_val:
                _val(values, "AT_PackDetails", "", id_val=pack_details)
            
            # AT_CountrySize — LOV attribute, UK size scale, footwear only (not accessories)
            country_size = art.get("AT_CountrySize", "")
            if country_size and "footwear" in category_val:
                _val(values, "AT_CountrySize", "", id_val=str(country_size).strip().upper())
            
            # AT_ArticleStatus — LOV attribute, send as ID not text
            article_status = art.get("AT_ArticleStatus", "A")
            if article_status:
                _val(values, "AT_ArticleStatus", "", id_val=article_status)
            
            # AT_SAPAge — LOV attribute, send as ID only when set
            # (set conditionally in _apply_formulas based on Category == "Footwear/Accessories")
            sap_age = art.get("AT_SAPAge", "")
            if sap_age:
                _val(values, "AT_SAPAge", "", id_val=sap_age)
            
            # AT_Gender — LOV attribute, send as ID not text
            gender = art.get("AT_Gender", "")
            if gender:
                _val(values, "AT_Gender", "", id_val=gender)
            
            # AT_SAPGender — LOV attribute, send as ID not text
            sap_gender = art.get("AT_SAPGender", "")
            if sap_gender:
                _val(values, "AT_SAPGender", "", id_val=sap_gender)
            
            # AT_BYGender — LOV attribute, send as ID not text
            by_gender = art.get("AT_BYGender", "")
            if by_gender:
                _val(values, "AT_BYGender", "", id_val=by_gender)
            
            # AT_Season — LOV attribute, send as ID only
            sea_raw = art.get("AT_Season", "")
            if sea_raw:
                sea_code = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw.upper()
                _val(values, "AT_Season", "", id_val=sea_code)
            
            # AT_SeasonYear — text value (e.g., "2027")
            season_year = art.get("AT_SeasonYear", "")
            if season_year:
                _val(values, "AT_SeasonYear", str(season_year))
            
            # Add all other mapped attributes
            # Skip these attributes from output XML:
            #   - AT_InboundGenericCode: already written above (avoid duplicate)
            #   - AT_SAPStyleCode: excluded per request
            #   - AT_Generic: excluded per request
            #   - AT_Createdon, AT_Createdby: excluded per request
            #   - AT_CompanyCode, AT_SBU, AT_Country, AT_Brand: already written above as LOV/MultiValue
            #   - AT_PricingDistributionChannel: already written above with id_val
            skip_attrs = {
                "AT_InboundGenericCode", "AT_SAPStyleCode", "AT_Generic", 
                "AT_Createdon", "AT_Createdby",
                "AT_CompanyCode", "AT_SBU", "AT_Country", "AT_Brand", "AT_BrandGroup",
                "AT_BrandType", "AT_BrandStatus", "AT_BrandCategory",
                "_category",
                "_sap_product_division",
                "AT_PricingDistributionChannel",
                "AT_FOBCurrency", "AT_RetailPriceCurrency",
                "AT_SAPProductFlag",
                "AT_MaterialType", "AT_SAPArticleCategory",
                "AT_UOM", "AT_BYAge", "AT_PackDetails", "AT_CountrySize",
                "AT_ArticleStatus",
                "AT_SAPAge",
                "AT_Season", "AT_SeasonYear",
                "AT_Gender", "AT_SAPGender", "AT_BYGender",
                "AT_GenericDescription",
            }
            
            for attr_id, value in art.items():
                if value and str(value).strip() and attr_id.startswith("AT_"):
                    if attr_id in skip_attrs:
                        continue
                    _val(values, attr_id, str(value))

        return products_root

    def generate(self, output_path: Path):
        export_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

        with open(output_path, "w", encoding="utf-8") as f:
            open_step_xml(f, export_time)

            cls_elem = self.build_classifications()
            xml_str = ET.tostring(cls_elem, encoding="unicode")
            f.write(strip_xmlns(xml_str))
            f.write("\n")

            prod_elem = self.build_products()
            xml_str = ET.tostring(prod_elem, encoding="unicode")
            f.write(strip_xmlns(xml_str))
            f.write("\n")

            close_step_xml(f)

        log.info("[XML] Generated: %s (%d articles)", output_path.name, len(self.articles))
        return output_path


# ══════════════════════════════════════════════════════════════════════════════
# FILENAME METADATA PARSER
# ══════════════════════════════════════════════════════════════════════════════
def _parse_meta_from_filename(path: Path, mdd: "MDDLoader | None" = None) -> dict:
    """
    Parse comp_code, sbu, brand_code, season, country_code from the linelist filename.
    
    Expected format (like Nike 360):
        <comp_code>-<sbu>-<brand>-<file_type>-<season>-<country>-<seq>.xlsx
        Example: 0193-CH-DRMARTEN-Retail Price Master-SS27-ID-1.xlsx
    
    Alternative format (DR Marten specific):
        SEA_SS27 B2B Suggested Retail Price Master - 20260526_OUTPUT FINAL.xlsx
        - Season extracted from SEA_SS27 or similar patterns
        - Country defaults to ID (Indonesia) for SEA files
    
    Returns dict with: brand_code, comp_code, sbu, season, country_code
    Returns empty dict if parsing fails completely.
    """
    stem = path.stem
    log.info("[Filename] Parsing metadata from: '%s'", stem)
    
    result = {
        "brand_code": BRAND_CODE,
        "comp_code": "",
        "sbu": "",
        "season": "",
        "country_code": "",
    }
    
    # Try standard format first: <comp_code>-<sbu>-<brand>-...
    parts = re.split(r"\s*-\s*", stem)
    
    if len(parts) >= 5:
        # Standard format: 0193-CH-DRMARTEN-Retail Price Master-SS27-ID-1
        comp_code = parts[0].strip()
        sbu = parts[1].strip()
        
        # Check if first part looks like a comp_code (4 digits)
        if re.match(r"^\d{4}$", comp_code):
            result["comp_code"] = comp_code
            result["sbu"] = sbu
            
            # Find season token (e.g., SP27, SS26, FW2027)
            for p in parts:
                if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
                    result["season"] = p.strip().upper()
                    break
            
            # Find country code (2-3 letter uppercase after season)
            season_found = False
            for p in parts:
                p_clean = p.strip().upper()
                if season_found and re.match(r"^[A-Z]{2,3}$", p_clean) and not p_clean.isdigit():
                    result["country_code"] = p_clean
                    break
                if re.match(r"^[A-Z]{2}\d{2,4}$", p_clean, re.IGNORECASE):
                    season_found = True
            
            log.info("[Filename] Standard format parsed → comp=%s sbu=%s season=%s country=%s",
                     result["comp_code"], result["sbu"], result["season"], result["country_code"])
            return result
    
    # Try DR Marten specific format: SEA_SS27 B2B Suggested...
    # Extract season from patterns like SEA_SS27, SS27, FW26, etc.
    season_match = re.search(r"[_\s]?([SF][WSP]\d{2,4})", stem, re.IGNORECASE)
    if season_match:
        result["season"] = season_match.group(1).upper()
    
    # Check for SEA prefix (South East Asia) — season only; no hardcoded comp/sbu/country
    # (country_code, comp_code, sbu must come from the filename tokens or args)

    # Try to find comp_code pattern (4 digits)
    comp_match = re.search(r"\b(\d{4})\b", stem)
    if comp_match and not result["comp_code"]:
        # Only use if it looks like a comp code (not a date like 20260526)
        potential_comp = comp_match.group(1)
        if not re.search(r"202\d", potential_comp):  # Not a year
            result["comp_code"] = potential_comp
    
    log.info("[Filename] DR Marten format parsed → comp=%s sbu=%s season=%s country=%s",
             result["comp_code"], result["sbu"], result["season"], result["country_code"])
    
    return result


# ══════════════════════════════════════════════════════════════════════════════
# MAIN RUN FUNCTION
# ══════════════════════════════════════════════════════════════════════════════
def run(args: dict, auditor=None) -> dict:
    """
    Main entry point for the DR Marten Retail Price Master ETL.
    FULLY DYNAMIC — all mappings read from Excel.
    """
    log.info("=" * 60)
    log.info("  DR Marten Retail Price Master ETL (DYNAMIC v2.0)")
    log.info("=" * 60)

    # Extract paths
    linelist_path = Path(args.get("linelist_path", ""))
    mdd_path = Path(args.get("mdd_path", ""))
    attributes_path = Path(args.get("attributes_path", ""))
    mapping_path = Path(args.get("mapping_path", ""))

    # ── Parse metadata from filename ─────────────────────────────────────────
    # Extract comp_code, sbu, season, country_code from input filename
    file_meta = {}
    if linelist_path and linelist_path.name:
        file_meta = _parse_meta_from_filename(linelist_path)
    
    # Context: filename values take precedence, then args (no hardcoded fallbacks)
    context = {
        "brand_code": BRAND_CODE,
        "country_code": file_meta.get("country_code") or args.get("country_code") or "",
        "comp_code": file_meta.get("comp_code") or args.get("comp_code") or "",
        "sbu": file_meta.get("sbu") or args.get("sbu") or "",
        "season": file_meta.get("season") or "",
    }
    
    log.info("[Context] brand=%s  comp=%s  sbu=%s  season=%s  country=%s",
             context["brand_code"], context["comp_code"], context["sbu"],
             context["season"], context["country_code"])

    # ── Load mapping rules (REQUIRED) ────────────────────────────────────────
    if not mapping_path.exists():
        log.error("[Mapping] File not found: %s", mapping_path)
        return {"status": "error", "message": f"Mapping file not found: {mapping_path}"}
    
    mapping = DynamicMappingLoader(mapping_path)
    if not mapping.rules:
        log.error("[Mapping] No rules loaded from mapping file")
        return {"status": "error", "message": "No mapping rules found"}
    
    log.info("[Mapping] Source fields used: %s", mapping.get_all_source_fields())

    # ── Load supporting files ────────────────────────────────────────────────
    mdd: MDDLoader | None = None
    rna: RNALoader | None = None

    if mdd_path.exists():
        mdd = MDDLoader(mdd_path)

    if attributes_path.exists():
        rna = RNALoader(attributes_path)

    # ── Load source data ─────────────────────────────────────────────────────
    if not linelist_path.exists():
        log.error("[Source] File not found: %s", linelist_path)
        return {"status": "error", "message": f"Source file not found: {linelist_path}"}

    source_loader = RetailPriceMasterLoader(linelist_path)
    rows = source_loader.get_rows()

    if not rows:
        log.error("[Source] No data rows found")
        return {"status": "error", "message": "No data rows found"}

    # Apply test row limit if defined and > 0
    if TEST_ROW_LIMIT and TEST_ROW_LIMIT > 0:
        rows = rows[:TEST_ROW_LIMIT]
        log.info("[TEST MODE] Limited to %d rows", TEST_ROW_LIMIT)

    # Apply test row range if defined (1-based, inclusive) — e.g. (1, 5) or (9, 10)
    if TEST_ROW_RANGE:
        start, end = TEST_ROW_RANGE
        rows = rows[start - 1:end]
        log.info("[TEST MODE] Limited to rows %d–%d (%d rows)", start, end, len(rows))

    # Fallback: Extract season from data if not found in filename
    if not context["season"] or context["season"] == "SS27":
        season_col = source_loader.find_column("Launch Season")
        if season_col and rows:
            season_from_data = _clean_str(rows[0].get(season_col, ""))
            if season_from_data:
                context["season"] = season_from_data
                log.info("[Context] Season from data: %s", context["season"])

    # ── Transform using dynamic engine ───────────────────────────────────────
    engine = DynamicTransformationEngine(mapping, mdd, rna, source_loader, context)
    mapped_articles: list[dict] = []

    for row in rows:
        try:
            article = engine.transform_row(row)
            if article.get("AT_Generic"):
                mapped_articles.append(article)
        except Exception as e:
            log.warning("[Transform] Error: %s", e)
            continue

    log.info("[Transform] %d articles mapped", len(mapped_articles))

    # ── Generate XML ─────────────────────────────────────────────────────────
    # Use input file stem as output filename (like Nike 360)
    output_filename = linelist_path.stem + ".xml"
    output_path = XML_OUT_DIR / output_filename

    generator = XMLGenerator(mapped_articles, context)
    generator.generate(output_path)

    if auditor:
        auditor.set_etl_counts(
            mapped=len(mapped_articles),
            written=len(mapped_articles),
            skipped=0,
            variants=0,
            xml_file=output_filename,
        )
        auditor.set_etl_status("success")

    return {
        "status": "success",
        "output_files": [str(output_path)],
        "article_count": len(mapped_articles),
    }


# ══════════════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="DR Marten Retail Price Master ETL (Dynamic)")
    parser.add_argument("--linelist", required=True, help="Path to source Excel file")
    parser.add_argument("--mdd", help="Path to MDD file")
    parser.add_argument("--attributes", help="Path to attributes file")
    parser.add_argument("--mapping", required=True, help="Path to brand mapping template")
    parser.add_argument("--country", default="ID", help="Country code")
    parser.add_argument("--comp", default="0888", help="Company code")
    parser.add_argument("--sbu", default="SP", help="SBU code")

    args_parsed = parser.parse_args()

    result = run({
        "linelist_path": args_parsed.linelist,
        "mdd_path": args_parsed.mdd or "",
        "attributes_path": args_parsed.attributes or "",
        "mapping_path": args_parsed.mapping,
        "country_code": args_parsed.country,
        "comp_code": args_parsed.comp,
        "sbu": args_parsed.sbu,
    })

    print(f"Result: {result}")

