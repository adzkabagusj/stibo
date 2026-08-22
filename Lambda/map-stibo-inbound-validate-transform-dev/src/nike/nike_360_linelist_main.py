"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Nike 360 Inline v1.0           ║
║   Dynamic Mapping Engine: Reads mapping logic from Excel       ║
╚══════════════════════════════════════════════════════════════════╝

This module implements a DYNAMIC mapping engine that:
  1. Reads mapping rules from the Brand Mapping Template Excel file
  2. Applies transformations based on the "Mapping Logic" column
  3. Generates Stibo STEP XML without hardcoded field mappings

Input files:
  - Brand Mapping Template (NIK 360 tab) → defines all attribute mappings
  - Nike 360 Order Form Excel → source data
  - MDD (Master Data Dictionary) → LOV lookups
  - Attributes List → RNA (Brand Type/Category) lookups

Key Features:
  - No hardcoded field mappings — all read from mapping Excel
  - Supports various mapping logic patterns:
    • Direct field copy ("as is")
    • Default values ("Default : X")
    • Formula-based derivations
    • LOV lookups
    • RNA-based lookups (BrandType, BrandCategory, etc.)
"""

from __future__ import annotations

import logging
import os
import re
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Callable
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# ── Directory setup ──────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent.parent)))

INPUT_DIR      = BASE_DIR / "input"
LINELIST_DIR   = INPUT_DIR / "linelist"
MDD_DIR        = INPUT_DIR / "mdd"
ATTR_DIR       = INPUT_DIR / "attributes"
MAPPING_DIR    = INPUT_DIR / "mapping"

OUTPUT_DIR     = BASE_DIR / "output"
XML_OUT_DIR    = OUTPUT_DIR / "xml"
LOG_DIR        = OUTPUT_DIR / "logs"

for _d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, MAPPING_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"nike360_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════
# LOV TABLES (Fallback defaults — MDD takes precedence)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {"AD": "Adults", "CH": "Children", "IN": "Infant", "AA": "All Ages", "JR": "Junior"}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter", "AL": "All Season",
    "HO": "Holiday", "AW": "Autumn-Winter", "SP": "Spring", "SM": "Summer",
}

# Note: Country mapping is now loaded dynamically from MDD's "Country LOV" sheet
# No hardcoded country mappings — all from MDD


# ══════════════════════════════════════════════════════════════════
# STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace("", STIBO_NS)


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — DYNAMIC MAPPING LOADER
# ══════════════════════════════════════════════════════════════════

class MappingRule:
    """Represents a single attribute mapping rule from the Excel template."""
    
    def __init__(
        self,
        stibo_attr_name: str,
        stibo_attr_id: str,
        validation_type: str,
        cluster: str,
        grouping: str,
        description: str,
        source_type: str,
        source_field: str,
        source_sheet: str,
        mapping_logic: str,
    ):
        self.stibo_attr_name = stibo_attr_name
        self.stibo_attr_id = stibo_attr_id
        self.validation_type = validation_type.lower() if validation_type else "text"
        self.cluster = cluster
        self.grouping = grouping
        self.description = description
        self.source_type = source_type  # "Mapping from principal", "Manual Input", "Formula in System"
        self.source_field = source_field  # Column name in source Excel
        self.source_sheet = source_sheet
        self.mapping_logic = mapping_logic
        
        # Parsed mapping instructions
        self.is_direct_copy = False
        self.default_value: str | None = None
        self.is_formula = False
        self.is_lov_lookup = False
        self.is_rna_lookup = False
        self.is_derived_from_image = False
        self.formula_type: str | None = None
        
        self._parse_mapping_logic()
    
    def _parse_mapping_logic(self):
        """Parse the mapping logic text to determine transformation type."""
        logic = (self.mapping_logic or "").strip().lower()
        
        # Check for direct copy
        if "as is" in logic or "populated the value from principal" in logic:
            self.is_direct_copy = True
        
        # Check for default values
        default_match = re.search(r"default\s*[:\-]?\s*(.+?)(?:\s*\(|$)", logic, re.IGNORECASE)
        if default_match:
            self.default_value = default_match.group(1).strip()
            # Clean up common patterns
            self.default_value = re.sub(r"\s*\(.*$", "", self.default_value).strip()
        
        # Check for LOV lookup
        if self.validation_type == "lov" or "mapping refers to lov" in logic:
            self.is_lov_lookup = True
        
        # Check for RNA-based lookups
        if "mapping based on compcode" in logic or "compcode and sbu" in logic:
            self.is_rna_lookup = True
        
        # Check for formula-based derivations
        if "formula" in logic or "digits brand code" in logic:
            self.is_formula = True
            if "3 digits brand code + up to 8 digits" in logic:
                self.formula_type = "generic_code"
            elif "variant" in self.stibo_attr_id.lower():
                self.formula_type = "variant_code"
        
        # Check for image-based derivation
        if "defined from image" in logic or "images by ai" in logic:
            self.is_derived_from_image = True
    
    def __repr__(self):
        return f"MappingRule({self.stibo_attr_id}, src={self.source_field}, logic={self.mapping_logic[:30]}...)"


class DynamicMappingLoader:
    """
    Loads mapping rules from the Brand Mapping Template Excel file.
    All attribute mappings are read dynamically — nothing is hardcoded.
    """
    
    MAPPING_SHEET_KEYWORDS = ["NIK 360", "NIKE 360", "NIK360"]
    
    # Column indices in the mapping template (0-based)
    COL_STIBO_ATTR_NAME = 0
    COL_STIBO_ATTR_ID = 1
    COL_VALIDATION = 2
    COL_CLUSTER = 3
    COL_GROUPING = 4
    COL_DESCRIPTION = 5
    COL_SOURCE_TYPE = 6
    COL_SOURCE_FIELD_1 = 7  # Sometimes used
    COL_SOURCE_SHEET = 8
    COL_SOURCE_FIELD = 9   # Main source field column
    COL_MAPPING_LOGIC = 10
    
    def __init__(self, mapping_file_path: Path):
        self.path = mapping_file_path
        self.rules: list[MappingRule] = []
        self.rules_by_id: dict[str, MappingRule] = {}
        self.source_sheet_name: str = ""
        self._load()
    
    def _load(self):
        log.info("[DynamicMapping] Loading from: %s", self.path.name)
        
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.error("[DynamicMapping] Failed to open workbook: %s", e)
            return
        
        # Find the NIK 360 sheet
        sheet_name = None
        for sn in wb.sheetnames:
            if any(kw in sn.upper() for kw in self.MAPPING_SHEET_KEYWORDS):
                sheet_name = sn
                break
        
        if not sheet_name:
            log.error("[DynamicMapping] No NIK 360 sheet found. Available: %s", wb.sheetnames)
            wb.close()
            return
        
        log.info("[DynamicMapping] Using sheet: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        
        # Skip header row(s) — start from row 2 (index 2)
        for i, row in enumerate(rows[2:], start=2):
            if not row or not row[self.COL_STIBO_ATTR_ID]:
                continue
            
            attr_id = self._clean(row[self.COL_STIBO_ATTR_ID])
            if not attr_id or attr_id.startswith("#"):  # Skip #N/A entries
                continue
            
            rule = MappingRule(
                stibo_attr_name=self._clean(row[self.COL_STIBO_ATTR_NAME]),
                stibo_attr_id=attr_id,
                validation_type=self._clean(row[self.COL_VALIDATION]),
                cluster=self._clean(row[self.COL_CLUSTER]),
                grouping=self._clean(row[self.COL_GROUPING]),
                description=self._clean(row[self.COL_DESCRIPTION]),
                source_type=self._clean(row[self.COL_SOURCE_TYPE]),
                source_field=self._clean(row[self.COL_SOURCE_FIELD]),
                source_sheet=self._clean(row[self.COL_SOURCE_SHEET]),
                mapping_logic=self._clean(row[self.COL_MAPPING_LOGIC]),
            )
            
            # Track source sheet name
            if rule.source_sheet and rule.source_sheet not in (" ", "\xa0"):
                self.source_sheet_name = rule.source_sheet
            
            self.rules.append(rule)
            self.rules_by_id[attr_id] = rule
        
        wb.close()
        log.info("[DynamicMapping] Loaded %d mapping rules", len(self.rules))
        log.info("[DynamicMapping] Source sheet: %s", self.source_sheet_name)
    
    @staticmethod
    def _clean(v) -> str:
        if v is None:
            return ""
        s = str(v).strip()
        if s in ("None", "nan", "\xa0", " "):
            return ""
        return s
    
    def get_source_fields(self) -> set[str]:
        """Return all source field names that need to be read from input file."""
        return {r.source_field for r in self.rules if r.source_field}
    
    def get_rules_with_source(self) -> list[MappingRule]:
        """Return only rules that have a source field mapping."""
        return [r for r in self.rules if r.source_field]
    
    def get_rule(self, attr_id: str) -> MappingRule | None:
        return self.rules_by_id.get(attr_id)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MDD LOADER
# ══════════════════════════════════════════════════════════════════

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
        
        # Load all LOV sheets
        self._load_lov_sheets(wb)
        self._load_retail_price_currency(wb)
        
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
                if v1:
                    sub[v1] = v0 if v0 else v1
    
    def _load_retail_price_currency(self, wb):
        """
        Load Country LOV and Retail Price Currency LOV dynamically from MDD.
        No hardcoded country mappings — all from MDD sheets.
        
        Logic:
        1. Country LOV: country_code → country_name (e.g., "ID" → "INDONESIA")
        2. Retail Price Currency LOV: country_name → currency_code (e.g., "INDONESIA" → "IDR")
        """
        # Step 1: Load Country LOV (country_code → country_name)
        # Note: Must exclude "Country Size LOV" and "Country Origin LOV" — only want "Country LOV"
        country_id_to_name: dict[str, str] = {}
        country_sheet = next(
            (s for s in wb.sheetnames 
             if "COUNTRY" in s.upper() and "LOV" in s.upper()
             and "SIZE" not in s.upper() and "ORIGIN" not in s.upper()),
            None,
        )
        if country_sheet:
            log.info("[MDD] Using Country LOV sheet: '%s'", country_sheet)
            for row in wb[country_sheet].iter_rows(min_row=2, values_only=True):
                if row and row[0] and len(row) > 1 and row[1]:
                    country_code = str(row[0]).strip().upper()
                    country_name = str(row[1]).strip()
                    if country_code and country_name:
                        country_id_to_name[country_code] = country_name
            log.info("[MDD] Country LOV loaded: %d entries", len(country_id_to_name))
        
        # Step 2: Load Retail Price Currency LOV (country_name → currency_code)
        retail_price_currency_lov: dict[str, str] = {}
        currency_sheet = next(
            (s for s in wb.sheetnames if "RETAIL PRICE CURRENCY" in s.upper()),
            None,
        )
        if currency_sheet:
            for row in wb[currency_sheet].iter_rows(min_row=2, values_only=True):
                if row and row[0] and len(row) > 1 and row[1]:
                    country_name = str(row[0]).strip().upper()
                    currency_code = str(row[1]).strip()
                    if country_name and currency_code:
                        retail_price_currency_lov[country_name] = currency_code
            log.info("[MDD] Retail Price Currency LOV loaded: %d entries", len(retail_price_currency_lov))
        
        # Store both LOVs for dynamic lookup
        self.lovs["CountryIdToName"] = country_id_to_name
        self.lovs["RetailPriceCurrency"] = retail_price_currency_lov
    
    def lookup_lov(self, lov_name: str, value: str) -> str:
        """Look up a value in a LOV and return the ID."""
        if not value:
            return ""
        lov = self.lovs.get(lov_name, {})
        return lov.get(value, lov.get(value.upper(), value))
    
    def resolve_currency_from_country_code(self, country_code: str) -> str:
        """
        Resolve currency code from country code using dynamic MDD lookups.
        
        Logic:
        1. country_code (e.g., "ID") → country_name (e.g., "INDONESIA") via CountryIdToName
        2. country_name (e.g., "INDONESIA") → currency_code (e.g., "IDR") via RetailPriceCurrency
        
        Returns: currency_code (e.g., "IDR") or empty string if not found.
        """
        if not country_code:
            return ""
        
        country_code_upper = country_code.strip().upper()
        
        # Step 1: country_code → country_name
        country_id_to_name = self.lovs.get("CountryIdToName", {})
        country_name = country_id_to_name.get(country_code_upper, "")
        if not country_name:
            log.warning("[MDD] Country code '%s' not found in Country LOV", country_code_upper)
            return ""
        
        # Step 2: country_name → currency_code
        retail_currency_lov = self.lovs.get("RetailPriceCurrency", {})
        country_name_upper = country_name.strip().upper()
        
        # Try exact match first
        if country_name_upper in retail_currency_lov:
            return retail_currency_lov[country_name_upper]
        
        # Try partial match (e.g., "INDONESIAN RUPIAH" contains root of "INDONESIA")
        # Also handles "PHILIPPINE PESO" matching "PHILIPPINES" (remove trailing 'S' or 'N')
        country_root = country_name_upper.rstrip("S").rstrip("N")  # PHILIPPINES -> PHILIPPINE, INDONESIA -> INDONESIA
        for lov_entry, currency in retail_currency_lov.items():
            lov_entry_upper = lov_entry.upper()
            # Check if country name or its root appears in the LOV entry
            if (country_name_upper in lov_entry_upper or 
                country_root in lov_entry_upper or
                lov_entry_upper.split()[0].rstrip("N") in country_name_upper):  # INDONESIAN -> INDONESIA
                return currency
        
        log.warning("[MDD] Country name '%s' not found in Retail Price Currency LOV", country_name)
        return ""


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — RNA LOADER (Brand Type/Category)
# ══════════════════════════════════════════════════════════════════

class RNALoader:
    """Loads Brand Type / Brand Category from RNA sheet."""
    
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
            log.warning("[RNA] No RNA sheet found in %s", self.path.name)
            wb.close()
            return
        
        log.info("[RNA] Using sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        
        # Find header row
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
        c_comp = _find("COMPCODE", "COMPANY CODE", "COMP CODE")
        c_sbu = _find("SBU")
        c_bcode = _find("BRANDCODE", "BRAND CODE")
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
                "brand_type": _cell(row, c_btype),
                "brand_category": _cell(row, c_bcat),
                "brand_group": _cell(row, c_bgroup),
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
        return self.lookup.get(key, {
            "brand_type": "", "brand_category": "", "brand_group": ""
        })
    
    def get_fuzzy(self, country_name: str, sbu: str, brand_code: str) -> dict:
        """Fuzzy lookup ignoring company code."""
        c = self._norm_country(country_name)
        b = (brand_code or "").strip().upper()
        s = (sbu or "").strip().upper()
        
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        
        return {"brand_type": "", "brand_category": "", "brand_group": ""}


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — SOURCE DATA LOADER
# ══════════════════════════════════════════════════════════════════

class Nike360SourceLoader:
    """Loads the Nike 360 Order Form source data."""
    
    PREFERRED_SHEETS = ["ORDER", "Sheet1", "Data"]
    
    def __init__(self, path: Path, target_sheet: str = ""):
        self.path = path
        self.target_sheet = target_sheet
        self.df = pd.DataFrame()
        self.headers: list[str] = []
        self._load()
    
    def _load(self):
        log.info("[Source] Loading: %s", self.path.name)
        
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        # Find target sheet
        target = None
        if self.target_sheet and self.target_sheet in wb.sheetnames:
            target = self.target_sheet
        else:
            target = next(
                (s for s in self.PREFERRED_SHEETS if s in wb.sheetnames),
                wb.sheetnames[0] if wb.sheetnames else None
            )
        
        if not target:
            log.error("[Source] No valid sheet found")
            wb.close()
            return
        
        log.info("[Source] Using sheet: '%s'", target)
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        
        # Find header row (first row with data)
        hdr_idx = 0
        for i, row in enumerate(rows[:5]):
            if row and any(v for v in row):
                hdr_idx = i
                break
        
        # Build header list
        self.headers = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        
        # Create DataFrame
        self.df = pd.DataFrame(rows[hdr_idx + 1:], columns=self.headers)
        
        # Filter out empty rows
        key_col = self.headers[0] if self.headers else None
        if key_col:
            self.df = self.df[
                self.df[key_col].notna() &
                (~self.df[key_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]
        
        wb.close()
        self.df = self.df.reset_index(drop=True)
        log.info("[Source] %d rows loaded, columns: %s", len(self.df), self.headers)


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — DYNAMIC TRANSFORMATION ENGINE
# ══════════════════════════════════════════════════════════════════

class TransformationEngine:
    """
    Applies mapping rules dynamically to transform source data.
    All transformations are driven by the mapping Excel — nothing hardcoded.
    """
    
    def __init__(
        self,
        mapping: DynamicMappingLoader,
        mdd: MDDLoader | None,
        rna: RNALoader | None,
        context: dict,
    ):
        self.mapping = mapping
        self.mdd = mdd
        self.rna = rna
        self.context = context  # brand_code, comp_code, sbu, country_code, season, etc.
    
    def transform_row(self, source_row: dict) -> dict:
        """
        Transform a single source row into a mapped article dict.
        Returns a dict with Stibo attribute IDs as keys.
        """
        result: dict[str, str] = {}
        
        for rule in self.mapping.rules:
            value = self._apply_rule(rule, source_row, result)
            if value is not None:
                result[rule.stibo_attr_id] = value
        
        # Add computed/derived attributes
        self._add_computed_attributes(result, source_row)
        
        return result
    
    def _apply_rule(self, rule: MappingRule, source_row: dict, partial_result: dict) -> str | None:
        """Apply a single mapping rule and return the transformed value."""
        
        # 1. RNA-based lookups (Brand Type, Brand Category, etc.)
        if rule.is_rna_lookup and self.rna:
            return self._apply_rna_lookup(rule, source_row)

        # 1b. Special case: AT_PrincipalGenderDescription
        #     Default = "Unisex", EXCEPT when PRODUCT LINE is "Headwear" → "Female"
        if rule.stibo_attr_id == "AT_PrincipalGenderDescription":
            product_line = self._get_source_value(source_row, "PRODUCT LINE")
            if "headwear" in product_line.lower():
                return "Female"
            return "Unisex"

        # 1b2. Special case: AT_Gender (LOV ID)
        #      Default = "U" (Unisex), EXCEPT when PRODUCT LINE is "Headwear" → "F" (Female)
        if rule.stibo_attr_id == "AT_Gender":
            product_line = self._get_source_value(source_row, "PRODUCT LINE")
            if "headwear" in product_line.lower():
                return "F"
            return "U"

        # 1b3. Special case: AT_BYGender (LOV ID)
        #      Default = "U" (Unisex), EXCEPT when PRODUCT LINE is "Headwear" → "F" (Female)
        if rule.stibo_attr_id == "AT_BYGender":
            product_line = self._get_source_value(source_row, "PRODUCT LINE")
            if "headwear" in product_line.lower():
                return "F"
            return "U"

        # 1c. Special case: AT_PrincipalAgeDescription
        #     Default = "Adults" (as text value, not ID)
        if rule.stibo_attr_id == "AT_PrincipalAgeDescription":
            return "Adults"

        # 1d. Special case: AT_SAPAge
        #     Default = "AD" (as ID for Adults)
        if rule.stibo_attr_id == "AT_SAPAge":
            return "AD"

        # 1e. Special case: AT_BYAge
        #     Default = "Adult" (as ID from BY Age LOV)
        if rule.stibo_attr_id == "AT_BYAge":
            return "ADULT"

        # 1f. Special case: AT_BYArticleType
        #     Default = "Inline" (as ID)
        if rule.stibo_attr_id == "AT_BYArticleType":
            return "Inline"

        # 2. Get source value (if applicable)
        source_value = ""
        if rule.source_field:
            source_value = self._get_source_value(source_row, rule.source_field)
        
        # 3. Apply default if no source value
        if not source_value and rule.default_value:
            return self._resolve_default(rule.default_value, source_row)
        
        # 4. Direct copy
        if rule.is_direct_copy and source_value:
            return self._clean_value(source_value)
        
        # 5. Formula-based transformations
        if rule.is_formula:
            return self._apply_formula(rule, source_row, partial_result)
        
        # 6. LOV lookup
        if rule.is_lov_lookup and source_value and self.mdd:
            lov_name = self._get_lov_name(rule.stibo_attr_id)
            return self.mdd.lookup_lov(lov_name, source_value)
        
        # 7. Context-based values (from portal/system)
        if "portal" in (rule.mapping_logic or "").lower():
            return self._get_context_value(rule.stibo_attr_id)
        
        # 8. Return source value if available
        if source_value:
            return self._clean_value(source_value)
        
        return None
    
    def _get_source_value(self, source_row: dict, field_name: str) -> str:
        """Get value from source row, handling various column name formats."""
        # Try exact match
        if field_name in source_row:
            return str(source_row[field_name]).strip() if source_row[field_name] else ""
        
        # Try case-insensitive match
        field_upper = field_name.upper()
        for key, val in source_row.items():
            if str(key).strip().upper() == field_upper:
                return str(val).strip() if val else ""
        
        # Try partial match (for fields like "ITEM DESCRIPTION " with trailing space)
        for key, val in source_row.items():
            if field_upper in str(key).strip().upper():
                return str(val).strip() if val else ""
        
        return ""
    
    def _clean_value(self, value: str) -> str:
        """Clean and normalize a value."""
        if not value:
            return ""
        s = str(value).strip()
        if s in ("None", "nan", "NaT", "\xa0"):
            return ""
        # Remove leading apostrophe (common in Excel for text numbers)
        if s.startswith("'"):
            s = s[1:]
        return s
    
    def _resolve_default(self, default: str, source_row: dict) -> str:
        """Resolve default value, which may have conditional logic."""
        default_lower = default.lower()
        
        # Handle special defaults
        if "unisex" in default_lower:
            return "U"
        if "adults" in default_lower:
            return "AD"
        if "ea" in default_lower:
            return "EA"
        if "sg" in default_lower:
            return "SG"
        if "local production" in default_lower or "5" in default_lower:
            return "5"
        if "direct" in default_lower or "zhaw" in default_lower:
            return "ZHAW"
        if "generic" in default_lower or "1" in default_lower:
            return "1"
        if "inline" in default_lower:
            return "Inline"
        
        return default
    
    def _apply_rna_lookup(self, rule: MappingRule, source_row: dict) -> str | None:
        """Apply RNA-based lookup for Brand Type, Brand Category, etc."""
        if not self.rna:
            return None
        
        # Get country_name from MDD's Country LOV (dynamic, not hardcoded)
        country_code = self.context.get("country_code", "")
        country_name = ""
        if self.mdd and country_code:
            country_id_to_name = self.mdd.lovs.get("CountryIdToName", {})
            country_name = country_id_to_name.get(country_code.strip().upper(), "")
        
        rna_result = self.rna.get(
            country_name,
            self.context.get("comp_code", ""),
            self.context.get("sbu", ""),
            self.context.get("brand_code", ""),
        )
        
        # If no exact match, try fuzzy
        if not any(rna_result.values()):
            rna_result = self.rna.get_fuzzy(
                country_name,
                self.context.get("sbu", ""),
                self.context.get("brand_code", ""),
            )
        
        # Map attribute ID to RNA field
        attr_id_lower = rule.stibo_attr_id.lower()
        if "brandtype" in attr_id_lower:
            return rna_result.get("brand_type", "")
        elif "brandcategory" in attr_id_lower:
            return rna_result.get("brand_category", "")
        elif "brandgroup" in attr_id_lower:
            return rna_result.get("brand_group", "")
        
        return None
    
    def _apply_formula(self, rule: MappingRule, source_row: dict, partial_result: dict) -> str | None:
        """Apply formula-based transformation."""
        brand_code = self.context.get("brand_code", "")
        
        # Get Principal Style Code and Color Code
        style_code = self._get_source_value(source_row, "STYLE CODE") or self._get_source_value(source_row, "SKU NO.")
        color_code = self._get_source_value(source_row, "COLOR CODE") or self._get_source_value(source_row, "PRINCIPAL COLOR CODE")
        
        if rule.formula_type == "generic_code" or "generic" in rule.stibo_attr_id.lower():
            # Inbound Generic Code = Brand Code + Principal Style Code
            # If multiple colors for same style: Brand Code + Principal Style Code + Principal Color Code
            if color_code:
                return f"{brand_code}{style_code}{color_code}"
            return f"{brand_code}{style_code}"
        
        if "variant" in rule.stibo_attr_id.lower():
            # Brand Code + Style + Color Code + Size
            size = self._get_source_value(source_row, "SIZE")
            size_clean = self._normalize_size_code(size)
            return f"{brand_code}{style_code}{color_code}{size_clean}"
        
        if "sapstylecode" in rule.stibo_attr_id.lower():
            # Principal Style Code
            return style_code
        
        return None
    
    def _normalize_size_code(self, size: str) -> str:
        """Convert size to 3-character code."""
        if not size:
            return "000"
        
        s = size.strip().upper()
        
        # Size code mapping
        size_map = {
            "OSFM": "OSS", "OS": "OSS", "O/S": "OSS", "ONE SIZE": "OSS",
            "XXS": "XXS", "XS": "XSS", "S": "SML", "M": "MED", "L": "LGE",
            "XL": "XLL", "XXL": "XXL", "2XL": "2XL", "3XL": "3XL",
        }
        
        if s in size_map:
            return size_map[s]
        
        # Numeric sizes (e.g., "03", "07")
        if s.isdigit():
            return s.zfill(3)[-3:]
        
        # Keep first 3 alphanumeric chars
        alnum = re.sub(r"[^A-Z0-9]", "", s)
        return alnum[:3].ljust(3, "0") if alnum else "000"
    
    def _get_lov_name(self, attr_id: str) -> str:
        """Derive LOV name from attribute ID."""
        # Remove AT_ prefix and common suffixes
        name = attr_id.replace("AT_", "")
        return name
    
    def _get_context_value(self, attr_id: str) -> str | None:
        """Get value from context (portal/system input)."""
        attr_lower = attr_id.lower()
        
        if "country" in attr_lower and "origin" not in attr_lower:
            return self.context.get("country_code", "")
        if "companycode" in attr_lower:
            return self.context.get("comp_code", "")
        if "sbu" in attr_lower:
            return self.context.get("sbu", "")
        if "brand" in attr_lower and "type" not in attr_lower and "category" not in attr_lower:
            return self.context.get("brand_code", "")
        if "season" in attr_lower and "year" not in attr_lower:
            return self.context.get("season", "")[:2] if self.context.get("season") else ""
        if "seasonyear" in attr_lower:
            season = self.context.get("season", "")
            if len(season) >= 4:
                return season[2:]
            return ""
        
        return None
    
    def _add_computed_attributes(self, result: dict, source_row: dict):
        """Add computed/derived attributes not directly from mapping."""
        brand_code = self.context.get("brand_code", "")
        brand_name = self.context.get("brand", "")
        
        # Ensure critical attributes are present
        if "AT_Brand" not in result or not result["AT_Brand"]:
            result["AT_Brand"] = brand_code
        
        if "AT_BrandGroup" not in result or not result["AT_BrandGroup"]:
            result["AT_BrandGroup"] = brand_name.upper()
        
        if "AT_BYIndicator" not in result:
            result["AT_BYIndicator"] = "Y"
        
        if "AT_SAPIndicator" not in result:
            result["AT_SAPIndicator"] = "N"

        if "AT_BYArticleType" not in result:
            result["AT_BYArticleType"] = "Inline"
        
        # Store source row reference for variant building
        result["_source_row"] = source_row


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — XML BUILDER
# ══════════════════════════════════════════════════════════════════

def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element | None:
    """Create a Value element."""
    clean_id = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip() if value else ""
    
    if clean_id in ("", "None", "nan"):
        clean_id = ""
    if clean_val in ("", "None", "nan"):
        clean_val = ""
    
    if not clean_id and not clean_val:
        return None
    
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    
    if clean_id:
        el.set("ID", clean_id)
    if clean_val:
        el.text = clean_val
    
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    """Create a MultiValue element."""
    if not id_val:
        return
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def build_classifications(
    brand: str, brand_code: str, season_code: str
) -> ET.Element:
    """Build the Classifications block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    
    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand.capitalize()}Batches"
    
    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "SP": "Spring", "SM": "Summer", "AL": "All Season",
    }
    sea_name = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short = f"{sea_prefix} {sea_year}".strip()
    
    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", batches_parent)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display
    
    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"
    
    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"
    
    return cls_root


def build_product_xml(
    article: dict,
    brand: str,
    brand_code: str,
    comp_code: str,
    sbu: str,
    season_id: str,
    mapping: DynamicMappingLoader,
    mdd: MDDLoader | None = None,
    country_code: str = "",
) -> str:
    """Build XML for a single product using dynamic mapping rules."""
    
    # Determine parent ID based on product type
    product_line = article.get("_source_row", {}).get("PRODUCT LINE", "")
    if "FOOTWEAR" in str(product_line).upper():
        div_letter = "F"
    elif "APPAREL" in str(product_line).upper():
        div_letter = "A"
    else:
        div_letter = "E"  # Equipment/Accessories
    
    parent_id = f"PPH_{div_letter}-TempSubCat"
    
    # Get key article value (AT_InboundGenericCode)
    # Formula: Brand Code + Principal Style Code
    # If multiple colors exist for same style: Brand Code + Principal Style Code + Principal Color Code
    key_article = article.get("AT_InboundGenericCode", "")
    if not key_article:
        style_code = article.get("AT_PrincipalStyleCode", "")
        color_code = article.get("AT_PrincipalColorCode", "")
        key_article = f"{brand_code}{style_code}{color_code}" if color_code else f"{brand_code}{style_code}"
    
    # Build product element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", parent_id)
    
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article
    
    # Product name
    name = article.get("AT_PrincipalStyleDescription", "") or article.get("AT_GenericDescription", "")
    if name:
        ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name
    
    # Classification references
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.capitalize()}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")
    
    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")
    
    # Values element
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    
    # Write attributes based on mapping rules
    written: set[str] = set()

    # Attributes explicitly excluded from XML output
    EXCLUDED_ATTRS: set[str] = {
        "AT_SAPStyleCode",          # Not to be sent in XML (length validation error in Stibo)
        "AT_SportsCategoryEN",      # TBC with MD — not to be sent in XML
        "AT_Generic",               # Handled by Stibo internally
        "AT_GenericDescription",    # Handled by Stibo internally
        "AT_Variant",               # Handled by Stibo internally
        "AT_VariantDescription",    # Handled by Stibo internally
        "AT_Size",                  # Not to be sent in XML
        "AT_PrincipalSize",         # Not to be sent in XML
        "AT_PrincipalBarcode",      # Not to be sent in XML
        "AT_SAPIndicator",          # Not to be sent in XML
        "AT_BYIndicator",           # Not to be sent in XML
        "AT_Createdon",             # Not to be sent in XML as it handled by stibo
        "AT_Createdby",             # Not to be sent in XML as it handled by stibo
    }
    
    # Multi-valued attributes first
    multi_attrs = {"AT_SBU", "AT_CompanyCode"}
    for attr_id in multi_attrs:
        if attr_id in article and article[attr_id]:
            _multival(vals_el, attr_id, article[attr_id])
            written.add(attr_id)

    # AT_InboundGenericCode — same logic as KEY_InboundArticle
    # Formula: Brand Code + Principal Style Code
    # If multiple colors for same style: Brand Code + Principal Style Code + Principal Color Code
    _val(vals_el, "AT_InboundGenericCode", value=key_article)
    written.add("AT_InboundGenericCode")
    
    # AT_ArticleStatus — always send with ID='A' (no text value)
    _val(vals_el, "AT_ArticleStatus", id_val="A")
    written.add("AT_ArticleStatus")
    
    # AT_FOBCurrency and AT_RetailPriceCurrency — derived from country_code using hardcoded mapping
    # Mapping: Country LOV ID → Retail Price Currency LOV ID
    COUNTRY_TO_CURRENCY = {
        "ID": "IDR",  # Indonesia → Indonesian Rupiah
        "PH": "PHP",  # Philippines → Philippine Peso
        "TH": "THB",  # Thailand → Thailand Baht
        "SG": "SGD",  # Singapore → Singapore Dollar
        "MY": "MYR",  # Malaysia → Malaysian Ringgit
        "VN": "VND",  # Vietnam → Vietnamese Dong
        "KH": "USD",  # Cambodia → United States Dollar
    }
    if country_code:
        currency_id = COUNTRY_TO_CURRENCY.get(country_code.strip().upper(), "")
        if currency_id:
            _val(vals_el, "AT_FOBCurrency", id_val=currency_id)
            written.add("AT_FOBCurrency")
            _val(vals_el, "AT_RetailPriceCurrency", id_val=currency_id)
            written.add("AT_RetailPriceCurrency")
    
    # Single-valued attributes
    for rule in mapping.rules:
        attr_id = rule.stibo_attr_id
        if attr_id in written or attr_id in EXCLUDED_ATTRS or attr_id.startswith("#"):
            continue
        
        value = article.get(attr_id, "")
        if not value:
            continue
        
        # Special case: AT_BrandCategory and AT_BrandType should pass value, not ID
        if attr_id in ("AT_BrandCategory", "AT_BrandType"):
            _val(vals_el, attr_id, value=value)
        # Determine if we should use ID or text value
        elif rule.validation_type == "lov":
            _val(vals_el, attr_id, id_val=value)
        else:
            _val(vals_el, attr_id, value=value)
        
        written.add(attr_id)
    
    # Add any remaining attributes not in mapping but in article
    for attr_id, value in article.items():
        if attr_id.startswith("_") or attr_id in written or attr_id in EXCLUDED_ATTRS:
            continue
        if not value or not attr_id.startswith("AT_"):
            continue
        _val(vals_el, attr_id, value=str(value))
    
    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def _parse_meta_from_filename(path: Path, mdd: "MDDLoader | None" = None) -> dict:
    """
    Parse comp_code, sbu, brand, brand_code, season, country_code from the
    linelist filename — same logic as lambda_function.py.
    Expected: <comp_code> - <sbu> - <brand> - <file_type> - <season> [- <country>] [- <seq>]
    Returns empty dict if parsing fails (caller should abort).
    """
    stem  = path.stem
    parts = re.split(r"\s*-\s*", stem)

    if len(parts) < 5:
        log.error("[Filename] '%s' has only %d parts — cannot parse metadata", stem, len(parts))
        return {}

    comp_code = parts[0].strip()
    sbu       = parts[1].strip()

    # ── Extract brand name + brand code from parts[2] ────
    # Supports both:
    #   "NIKE"         → brand="Nike",  brand_code="NIK"  (default)
    #   "NIKE NIK"     → brand="Nike",  brand_code="NIK"  (embedded)
    brand      = "Nike"   # default
    brand_code = "NIK"    # default
    raw_brand_token = parts[2].strip()
    brand_token_match = re.match(
        r"^([A-Za-z][A-Za-z\s]*?)\s+([A-Z]{3})$",
        raw_brand_token,
        re.IGNORECASE,
    )
    if brand_token_match:
        # e.g. "NIKE NIK" → brand="Nike", brand_code="NIK" (from filename)
        brand      = brand_token_match.group(1).strip().title()
        brand_code = brand_token_match.group(2).upper()
        log.info(
            "  Brand token '%s' → brand=%r  brand_code=%r (embedded in filename)",
            raw_brand_token, brand, brand_code,
        )
    else:
        # No embedded code — use defaults: brand="Nike", brand_code="NIK"
        brand = raw_brand_token.title()
        log.info(
            "  Brand token '%s' → brand=%r  brand_code=%r (default fallback)",
            raw_brand_token, brand, brand_code,
        )

    # Locate season token (e.g. SP27, SS26, FW2027)
    season     = ""
    season_idx = None
    for i, p in enumerate(parts):
        if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
            season     = p.strip().upper()
            season_idx = i
            break

    if not season or season_idx is None:
        log.error("[Filename] No season token found in '%s'", stem)
        return {}

    # country_code = first 2-3 letter uppercase-only token after the season
    trailing     = parts[season_idx + 1:]
    country_code = ""
    for p in trailing:
        p = p.strip()
        if re.match(r"^[A-Z]{2,3}$", p, re.IGNORECASE) and not p.isdigit():
            country_code = p.upper()
            break

    log.info(
        "[Filename] Parsed → comp_code=%s  sbu=%s  brand=%s  brand_code=%s  season=%s  country=%s",
        comp_code, sbu, brand, brand_code, season, country_code,
    )
    return {
        "brand":        brand,
        "brand_code":   brand_code,
        "comp_code":    comp_code,
        "sbu":          sbu,
        "season":       season,
        "country_code": country_code,
    }


def run(args=None, auditor=None):
    """Main entry point for Nike 360 Linelist processing."""
    
    def first(d: Path, pattern: str = "*.xlsx") -> Path | None:
        files = list(d.glob(pattern))
        return files[0] if files else None
    
    def find_mapping_file() -> Path | None:
        """Search for mapping file in multiple locations."""
        # 1. Check MAPPING_DIR first (Lambda downloads here)
        f = first(MAPPING_DIR)
        if f:
            log.info("[Mapping] Found in MAPPING_DIR: %s", f.name)
            return f
        
        # 2. Check BASE_DIR with various patterns
        for pattern in ["*mapping*Template*.xlsx", "*Brand*mapping*.xlsx", "*mapping*.xlsx"]:
            files = list(BASE_DIR.glob(pattern))
            if files:
                log.info("[Mapping] Found in BASE_DIR: %s", files[0].name)
                return files[0]
        
        # 3. Check INPUT_DIR root
        f = first(INPUT_DIR, "*mapping*.xlsx")
        if f:
            log.info("[Mapping] Found in INPUT_DIR: %s", f.name)
            return f
        
        # 4. Check all input subdirectories
        for subdir in INPUT_DIR.iterdir():
            if subdir.is_dir():
                for pattern in ["*mapping*.xlsx", "*Template*.xlsx"]:
                    files = list(subdir.glob(pattern))
                    if files:
                        log.info("[Mapping] Found in %s: %s", subdir.name, files[0].name)
                        return files[0]
        
        return None
    
    # Load required files
    mdd_f = first(MDD_DIR)
    attr_f = first(ATTR_DIR)
    mapping_f = find_mapping_file()
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))
    
    # Validate required files
    for label, val in [("MDD", mdd_f), ("Linelist", ll_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            return None
    
    # Load MDD
    mdd = MDDLoader(mdd_f) if mdd_f else None
    
    # Load RNA from brand mapping file (downloaded to ATTR_DIR — lotto pattern)
    rna = None
    if attr_f and attr_f.exists():
        rna = RNALoader(attr_f)
        log.info("[RNA] Loaded from: %s", attr_f.name)
    else:
        log.warning("[RNA] No brand mapping file found — AT_BrandType/AT_BrandCategory will be empty.")
    
    # Load mapping rules
    if not mapping_f:
        log.error("No mapping template file found — aborting.")
        log.error("  Searched locations:")
        log.error("    MAPPING_DIR: %s  (exists: %s)", MAPPING_DIR, MAPPING_DIR.exists())
        log.error("    BASE_DIR:    %s  (exists: %s)", BASE_DIR, BASE_DIR.exists())
        log.error("    INPUT_DIR:   %s  (exists: %s)", INPUT_DIR, INPUT_DIR.exists())
        if INPUT_DIR.exists():
            log.error("  INPUT_DIR contents:")
            for item in INPUT_DIR.iterdir():
                if item.is_dir():
                    log.error("    %s/ → %s", item.name, [f.name for f in item.glob("*")])
                else:
                    log.error("    %s", item.name)
        return None
    
    mapping = DynamicMappingLoader(mapping_f)
    if not mapping.rules:
        log.error("No mapping rules loaded — aborting.")
        return None

    # Parse all metadata from the linelist filename (no fallbacks)
    meta = _parse_meta_from_filename(ll_files[0], mdd)
    if not meta:
        log.error("Cannot parse metadata from filename — aborting.")
        return None

    brand        = meta["brand"]
    brand_code   = meta["brand_code"]
    comp_code    = meta["comp_code"]
    sbu          = meta["sbu"]
    season       = meta["season"]
    country_code = meta["country_code"]

    log.info(
        "Filename metadata → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  country=%s",
        brand, brand_code, comp_code, sbu, season, country_code,
    )

    # Build context for transformation
    context = {
        "brand":        brand,
        "brand_code":   brand_code,
        "comp_code":    comp_code,
        "sbu":          sbu,
        "season":       season,
        "country_code": country_code,
    }
    
    # Create transformation engine
    engine = TransformationEngine(mapping, mdd, rna, context)
    
    # Process each linelist file
    all_articles: list[dict] = []
    
    for ll_path in ll_files:
        log.info("─── Processing: %s ───", ll_path.name)
        
        # Load source data
        source = Nike360SourceLoader(ll_path, target_sheet=mapping.source_sheet_name)
        if source.df.empty:
            log.warning("[Source] Empty dataframe — skipping.")
            continue
        
        # source.df = source.df.head(5)  # DEBUG for testing: Uncomment to process only first 5 rows        
        log.info("Processing %d rows...", len(source.df))
        
        # Transform each row
        for _, row in source.df.iterrows():
            row_dict = row.to_dict()
            article = engine.transform_row(row_dict)
            all_articles.append(article)
    
    if not all_articles:
        log.error("No articles processed — aborting.")
        return None
    
    log.info("Total articles: %d", len(all_articles))
    
    # Build season ID
    sea_prefix = season[:2].upper() if len(season) >= 2 else season
    sea_year_short = season[2:] if len(season) > 2 else ""
    sea_year = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id = f"CLH_{brand_code}_{sea_prefix}{sea_year}"
    
    # Build output XML
    out_name = ll_files[0].stem + ".xml" if ll_files else "nike360_output.xml"
    out_path = XML_OUT_DIR / out_name
    
    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    cls_el = build_classifications(brand, brand_code, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    
    written_count = 0
    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
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
        f.write(f"  {cls_str}\n\n")
        
        f.write("  <Products>\n")
        for art in all_articles:
            product_xml = build_product_xml(
                art, brand, brand_code,
                comp_code, sbu, season_id, mapping,
                mdd=mdd, country_code=country_code,
            )
            product_xml = _XMLNS_RE.sub("", product_xml)
            f.write(f"    {product_xml}\n")
            written_count += 1
        
        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")
    
    file_kb = out_path.stat().st_size // 1024
    log.info("✓ XML written → %s  (%dKB, %d products)", out_path, file_kb, written_count)
    
    print("═══ NIKE 360 SUMMARY ═══════════════════════════════", flush=True)
    print(f"  Input rows     : {len(all_articles)}", flush=True)
    print(f"  Products XML   : {written_count}", flush=True)
    print(f"  Output file    : {out_path}", flush=True)
    print(f"  File size      : {file_kb}KB", flush=True)
    print("════════════════════════════════════════════════════", flush=True)
    
    if auditor:
        auditor.set_xml_uploads([str(out_path)])
    
    return {
        "output": str(out_path),
        "total": len(all_articles),
        "success": written_count,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    import argparse

    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — Nike 360 Inline (Dynamic Mapping)"
    )
    p.add_argument("--input", "-i", help="Input linelist file path")
    p.add_argument("--mapping", "-m", help="Mapping template file path")

    args = p.parse_args()

    # If input file provided, copy to input dir
    if hasattr(args, "input") and args.input:
        import shutil
        src = Path(args.input)
        if src.exists():
            shutil.copy(src, LINELIST_DIR / src.name)

    if hasattr(args, "mapping") and args.mapping:
        import shutil
        src = Path(args.mapping)
        if src.exists():
            shutil.copy(src, MAPPING_DIR / src.name)

    result = run()
    if result:
        log.info("Done ✓")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
