"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Birkenstock Inline  v1.2        ║
║   OFS Mitra → Stibo STEP XML (Article Planning)                 ║
║                                                                  ║
║   FULLY DYNAMIC MAPPING ENGINE                                   ║
║   All mappings are read from the Brand Mapping Template Excel    ║
║   No hardcoded field-to-attribute mappings                       ║
╚══════════════════════════════════════════════════════════════════╝

Source file format : "OFS_FC_Mitra_ss27.xlsx"
Mapping file       : "NEW - Brand mapping files Template.xlsx" → Sheet: "Birken(Inline)"

Change log v1.2:
  - AT_BrandGroup / AT_BrandType: added from RNA lookup (Nike parity)
  - AT_SAPStyleCode: max 9 chars of principal style code
  - AT_Generic: BCK + max 9 alphanumeric chars of style code
  - AT_GenericDescription: REMOVED
  - AT_Color: written as LOV ID from "Color" column
  - AT_BYGender: written as LOV ID from Gender column (M/F/U)
  - AT_PrincipalGenderCode / AT_PrincipalGenderDescription: from "Gender" column
  - AT_BYAge: written as LOV ID
  - AT_CountryOrigin: written as LOV ID (default DE for Birkenstock/Germany)
  - AT_FOBCurrency: default EUR
  - AT_SAPProductFlag: default A
  - AT_HeelHeight: LOV lookup (F/H/L/M)
  - AT_HeelType: LOV lookup (BL/EF/EW/FL/HW/KT/PH/PL/WD)
  - AT_Occasion: LOV lookup (CASUAL/FORMAL/LIFESTYLE/OCCASION/OUTDOOR)
  - REMOVED: AT_MainEANIndicator, AT_SAPAge, AT_Gender, AT_BCI,
             AT_Silhouette, AT_Material, AT_PatternPrint,
             AT_PricingDistributionChannel,
             AT_PrincipalMerchandiseHierarchyL1/L2/L3,
             AT_ShortDescriptionEN, AT_LongDescriptionEN,
             AT_GenericDescription
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
from typing import Optional, Callable, Any
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
# TEST LIMIT
# ══════════════════════════════════════════════════════════════════════════════
# TEST_ROW_RANGE = (40, 41)
# TEST_ROW_LIMIT = None

# ══════════════════════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"
MAPPING_DIR  = INPUT_DIR / "mapping"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, MAPPING_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("birkenstock.ofs_mitra")
if not log.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


# ══════════════════════════════════════════════════════════════════════════════
# STIBO / STEP XML CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(:\w+)?="[^"]*"')

BRAND_CODE = "BCK"
BRAND_NAME = "BIRKENSTOCK"


# ══════════════════════════════════════════════════════════════════════════════
# LOV TABLES
# ══════════════════════════════════════════════════════════════════════════════
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "AW": "Autumn-Winter",
}

GENDER_MAP: dict[str, str] = {
    "MENS": "M", "MEN": "M", "MALE": "M", "M": "M",
    "WOMENS": "F", "WOMEN": "F", "FEMALE": "F", "LADIES": "F", "F": "F",
    "UNISEX": "U", "KIDS": "U", "YOUTH": "U", "INFANT": "U",
    "JUNIORS": "U", "U": "U",
}

COUNTRY_CURRENCY: dict[str, str] = {
    "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD", "MY": "MYR",
    "VN": "VND", "KH": "USD", "US": "USD", "GB": "GBP", "HK": "HKD",
}
# ── Country → Country Size system LOV code (footwear only) ────────────────
COUNTRY_SIZE_SYSTEM: dict[str, str] = {
    "ID": "EU",   # Indonesia → EUR sizing, LOV code "UE"
    # Add other countries here as their mapping is confirmed
    # "MY": "...", "TH": "...", "SG": "...", "PH": "...", "VN": "...",
}

# ── Heel Type LOV ─────────────────────────────────────────────────────────────
# Code → display label
LOV_HEEL_TYPE: dict[str, str] = {
    "BL": "Block",
    "EF": "Espadrille Flat",
    "EW": "Espadrille Wedge",
    "FL": "Flat",
    "HW": "Hidden Wedge",
    "KT": "Kitten",
    "PH": "Pencil Heel",
    "PL": "Platform",
    "WD": "Wedge",
}
# Reverse: display label → code (for lookup by label from source)
LOV_HEEL_TYPE_REV: dict[str, str] = {v.upper(): k for k, v in LOV_HEEL_TYPE.items()}

# ── Heel Height LOV ───────────────────────────────────────────────────────────
LOV_HEEL_HEIGHT: dict[str, str] = {
    "F": "Flat",
    "H": "High",
    "L": "Low",
    "M": "Medium",
}
LOV_HEEL_HEIGHT_REV: dict[str, str] = {v.upper(): k for k, v in LOV_HEEL_HEIGHT.items()}

# ── Occasion LOV ──────────────────────────────────────────────────────────────
LOV_OCCASION: dict[str, str] = {
    "CASUAL":        "Casual",
    "ESPADRILLEFLAT":"Espadrille Flat",
    "FORMAL":        "Formal",
    "LIFESTYLE":     "Lifestyle",
    "OCCASION":      "Occasion",
    "OUTDOOR":       "Outdoor",
}
LOV_OCCASION_REV: dict[str, str] = {v.upper(): k for k, v in LOV_OCCASION.items()}
BY_AGE_FROM_PPG: dict[str, str] = {
    "ACTIVE ADULTS EVA":     "Adults",
    "ACTIVE ADULTS PU":      "Adults",
    "ACTIVE KIDS EVA":       "Kids",
    "ACTIVE KIDS PU":        "Kids",
    "CLASSIC ADULTS LEA":    "Adults",
    "CLASSIC ADULTS TEX":    "Adults",
    "CLASSIC KIDS LEA":      "Kids",
    "CLASSIC KIDS TEX":      "Kids",
    "HOMESHOES ADULTS":      "Adults",
    "PAPILLIO":              "Adults",
    "PROFESSIONAL FOOTWEAR": "Adults",
    "SHOES ADULTS":          "Adults",
    "LEGWEAR ADULTS":        "Adults",
    "ORTHOPEDICS":           "Adults",
    "SHOE CARE":             "Adults",
}

# ── Distributor → country code ─────────────────────────────────────────────
MITRA_COUNTRY_MAP: dict[str, str] = {
    "MITRA": "ID",
    "SGP":   "SG",
    "MYS":   "MY",
    "THA":   "TH",
    "PHL":   "PH",
    "VNM":   "VN",
}
# ── Country code → RNA sheet's full country name ──────────────────────────
RNA_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
}

SBU_MAP: dict[str, str] = {
    "FC": "SP",
    "SP": "SP",
    "AP": "AP",
    "FW": "FW",
}
# ── Division → ParentID letter (PPH_{letter}-TempSubCat) ─────────────────────
DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E",
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "EQUIPMENT":   "Q",
    "HARDWARE":    "Q",
    "TOYS":        "T",
}

def _get_division_code(division: str) -> str:
    """Return the ParentID letter for a division; 'X' if unrecognised."""
    div_upper = (division or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in div_upper:
            return code
    return "X"

# ══════════════════════════════════════════════════════════════════════════════
# LOV RESOLVER HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _resolve_heel_type(raw: str) -> str:
    """Given raw source value, return the LOV code (e.g. 'BL', 'FL')."""
    if not raw:
        return ""
    upper = raw.strip().upper()
    # Already a valid code?
    if upper in LOV_HEEL_TYPE:
        return upper
    # Try reverse lookup by label
    return LOV_HEEL_TYPE_REV.get(upper, raw.strip())

def _strip_gender_code_prefix(raw: str) -> str:
    """'03 - Women' → 'WOMEN'; leaves already-clean values untouched."""
    if not raw:
        return raw
    return re.sub(r"^\d+\s*-\s*", "", raw.strip()).upper()

def _resolve_heel_height(raw: str) -> str:
    """Given raw source value, return the LOV code (e.g. 'F', 'H', 'L', 'M')."""
    if not raw:
        return ""
    upper = raw.strip().upper()
    if upper in LOV_HEEL_HEIGHT:
        return upper
    return LOV_HEEL_HEIGHT_REV.get(upper, raw.strip())


def _resolve_occasion(raw: str) -> str:
    """Given raw source value, return the LOV code (e.g. 'CASUAL', 'FORMAL')."""
    if not raw:
        return ""
    upper = raw.strip().upper().replace(" ", "")
    if upper in LOV_OCCASION:
        return upper
    # Try reverse lookup
    return LOV_OCCASION_REV.get(raw.strip().upper(), raw.strip().upper())


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "") -> Optional[ET.Element]:
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
    if not id_val or str(id_val).strip() in ("", "None", "nan"):
        return
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val).strip())


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


def _clean_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v).strip()
    if s in ("None", "nan", "NaN", "N/A", "#N/A"):
        return ""
    return s


def _style_clean(code: str, max_len: int = 9) -> str:
    """Strip non-alphanumeric chars, uppercase, max length."""
    return re.sub(r"[^A-Z0-9]", "", (code or "").upper())[:max_len]


# ══════════════════════════════════════════════════════════════════════════════
# MAPPING RULE
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class MappingRule:
    stibo_attr_name:    str
    stibo_attr_id:      str
    validation_type:    str
    cluster:            str
    grouping:           str
    description:        str
    field_mapping_type: str
    source_field:       str
    mapping_logic:      str

    is_direct_copy: bool = False
    default_value:  str  = ""
    is_rna_lookup:  bool = False
    is_formula:     bool = False
    formula_type:   str  = ""
    is_lov_lookup:  bool = False

    def __post_init__(self):
        self._parse_mapping_logic()

    def _parse_mapping_logic(self):
        logic_lower = (self.mapping_logic or "").lower()

        if self.source_field and self.source_field not in ("None", "N/A", ""):
            self.is_direct_copy = True

        default_patterns = [
            r"default\s*[:\-]?\s*['\"]?([^'\"(\n]+)['\"]?",
            r"default\s*[:\-]?\s*(.+?)(?:\s*\(|$|\n)",
        ]
        # Match against the ORIGINAL-case mapping_logic (not logic_lower),
        # so the captured default value keeps its real casing.
        for pattern in default_patterns:
            match = re.search(pattern, self.mapping_logic or "", re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                val = re.sub(r"\s*\(.*$", "", val).strip().strip("'\"")
                if val and val.lower() not in ("", "none"):
                    self.default_value = val
                    break

        if "compcode" in logic_lower or "based on compcode" in logic_lower:
            self.is_rna_lookup = True

        if self.validation_type.lower() == "lov":
            self.is_lov_lookup = True

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
    MAPPING_SHEET_NAME = "Birken(Inline)"

    COL_STIBO_ATTR_NAME    = 0
    COL_STIBO_ATTR_ID      = 1
    COL_VALIDATION         = 2
    COL_CLUSTER            = 3
    COL_GROUPING           = 4
    COL_DESCRIPTION        = 5
    COL_FIELD_MAPPING_TYPE = 6
    COL_BRAND_FILE_NAME    = 7
    COL_BRAND_FILE_SHEET   = 8
    COL_SOURCE_FIELD       = 9
    COL_MAPPING_LOGIC      = 10

    def __init__(self, mapping_file_path: Path):
        self.path = mapping_file_path
        self.rules: list[MappingRule] = []
        self.rules_by_id: dict[str, MappingRule] = {}
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

        ws   = wb[self.MAPPING_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))

        for row in rows[1:]:
            if not row or not row[self.COL_STIBO_ATTR_ID]:
                continue
            attr_id = self._clean(row[self.COL_STIBO_ATTR_ID])
            if not attr_id or attr_id.startswith("#"):
                continue

            source_field = self._clean(row[self.COL_SOURCE_FIELD])
            if source_field.lower() in ("n/a", "none", ""):
                source_field = ""

            if "\n" in source_field:
                parts = source_field.split("\n")
                for part in parts:
                    if ":" in part:
                        subparts = part.split(":")
                        if len(subparts) >= 2:
                            source_field = subparts[1].strip()
                            break
                    elif part.strip() and part.strip().lower() not in ("acc", "fw", "ofs"):
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
            if attr_id not in self.rules_by_id:
                self.rules_by_id[attr_id] = rule

        # ── Load Upper Material → Material LOV from "Birken-MD Mappings" ──────
        self.material_lov = self._load_material_lov(wb)
        self.silhouette_lov = self._load_silhouette_lov(wb)
        self.sap_age_lov = self._load_sap_age_lov(wb)
        self.sap_product_division_lov = self._load_sap_product_division_lov(wb)
        self.ecom_short_desc_lov, self.ecom_long_desc_lov = self._load_ecom_description_lov(wb)
        self.gender_std_lov = self._load_gender_std_lov(wb)

        wb.close()
        log.info("[DynamicMapping] Loaded %d rules (%d direct), %d material LOV entries",
                 len(self.rules), len(self.get_rules_with_source()), len(self.material_lov))
        
    
    def _load_gender_std_lov(self, wb) -> dict[str, str]:
        """Raw source Gender (e.g. 'MENS'/'WOMENS'/'KIDS') → standardized label
        (e.g. 'Male'/'Female'/'Unisex'), from the Birken-MD Mappings sheet."""
        sheet_name = next(
            (s for s in wb.sheetnames
             if "BIRKEN" in s.upper() and ("MD MAPPING" in s.upper() or "MD-MAPPING" in s.upper())),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] Gender LOV: 'Birken-MD Mappings' sheet not found")
            return {}
    
        ws, rows = wb[sheet_name], list(wb[sheet_name].iter_rows(values_only=True))
        header_idx = col_gender = col_std = None
    
        RAW_HEADER = "GENDER"
        STD_HEADER = "SAP GENDER"
    
        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                cell_upper = str(cell).strip().upper() if cell else ""
                if cell_upper == RAW_HEADER:
                    col_gender, header_idx = j, i
                if cell_upper == STD_HEADER and header_idx == i:
                    col_std = j
            if col_gender is not None and col_std is not None:
                break
            col_gender = None
    
        if header_idx is None or col_gender is None or col_std is None:
            log.warning("[DynamicMapping] Could not locate Gender std-mapping columns in '%s'", sheet_name)
            return {}
    
        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_gender >= len(row) or col_std >= len(row):
                continue
            raw_val, std_val = row[col_gender], row[col_std]
            if raw_val and std_val:
                lov[str(raw_val).strip().upper()] = str(std_val).strip()
    
        log.info("[DynamicMapping] Gender std LOV loaded from '%s': %d entries", sheet_name, len(lov))
        return lov
    
    def _load_ecom_description_lov(self, wb) -> tuple[dict[str, str], dict[str, str]]:
        sheet_name = next(
            (s for s in wb.sheetnames if "ECOM MAPPING" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] 'Birken Ecom Mappings' sheet not found")
            return {}, {}
    
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
    
        def _norm(cell) -> str:
            """Collapse newlines/multi-spaces so wrapped headers match cleanly."""
            if not cell:
                return ""
            return re.sub(r"\s+", " ", str(cell).strip()).upper()
    
        header_idx = None
        col_model  = None
        col_short  = None
        col_long   = None
    
        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                cell_norm = _norm(cell)
                if cell_norm == "MODEL":
                    col_model  = j
                    header_idx = i
                if cell_norm == "ECOM SHORT DESCRIPTION EN" and header_idx == i:
                    col_short = j
                if cell_norm == "ECOM LONG DESCRIPTION EN" and header_idx == i:
                    col_long = j
            if col_model is not None and (col_short is not None or col_long is not None):
                break
            col_model = None  # reset if only half-matched
    
        if header_idx is None or col_model is None:
            log.warning(
                "[DynamicMapping] Could not locate 'Model' header column in sheet '%s'",
                sheet_name,
            )
            return {}, {}
    
        short_lov: dict[str, str] = {}
        long_lov:  dict[str, str] = {}
    
        for row in rows[header_idx + 1:]:
            if not row or col_model >= len(row):
                continue
            model_val = row[col_model]
            if not model_val:
                continue
            key = re.sub(r"\s+", " ", str(model_val).strip()).upper()
            if not key:
                continue
    
            if col_short is not None and col_short < len(row) and row[col_short]:
                val = str(row[col_short]).strip()
                if val:
                    short_lov[key] = val
    
            if col_long is not None and col_long < len(row) and row[col_long]:
                val = str(row[col_long]).strip()
                if val:
                    long_lov[key] = val
    
        log.info(
            "[DynamicMapping] Ecom Description LOV loaded from '%s': %d short / %d long",
            sheet_name, len(short_lov), len(long_lov),
        )
        if not short_lov and not long_lov:
            log.warning(
                "[DynamicMapping] Ecom Description LOV is EMPTY — check header text "
                "in sheet '%s' (row %d): %s", sheet_name, header_idx + 1, rows[header_idx],
            )
        return short_lov, long_lov

    def _load_silhouette_lov(self, wb) -> dict[str, str]:
        """
        Load Model → Silhouette mapping from 'Birken-MD Mappings' sheet.
        Scans for header row containing 'Model' / 'Silhouette' column labels.
        """
        sheet_name = next(
            (s for s in wb.sheetnames
             if "BIRKEN" in s.upper()
             and ("MD MAPPING" in s.upper() or "MD-MAPPING" in s.upper())),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] Silhouette LOV: 'Birken-MD Mappings' sheet not found")
            return {}

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        header_idx    = None
        col_model     = None
        col_silhouette = None

        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                cell_upper = str(cell).strip().upper() if cell else ""
                if cell_upper == "MODEL":
                    col_model  = j
                    header_idx = i
                if cell_upper == "SILHOUETTE" and header_idx == i:
                    col_silhouette = j
            if col_model is not None and col_silhouette is not None:
                break
            col_model = None  # reset if only half-matched

        if header_idx is None or col_model is None or col_silhouette is None:
            log.warning(
                "[DynamicMapping] Could not locate 'Model'/'Silhouette' "
                "header columns in sheet '%s'", sheet_name,
            )
            return {}

        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_model >= len(row) or col_silhouette >= len(row):
                continue
            model_val     = row[col_model]
            silhouette_val = row[col_silhouette]
            if not model_val or not silhouette_val:
                continue
            key = str(model_val).strip().upper()
            val = str(silhouette_val).strip()
            if key and val:
                lov[key] = val

        log.info("[DynamicMapping] Silhouette LOV loaded from '%s': %d entries",
                 sheet_name, len(lov))
        return lov
    
    def _load_sap_age_lov(self, wb) -> dict[str, str]:
        """
        Load Product Planning Group → SAP Age mapping from 'Birken-MD Mappings' sheet.
        Scans for header row containing 'Product Planning Group' / 'SAP Age' column labels.
        """
        sheet_name = next(
            (s for s in wb.sheetnames
             if "BIRKEN" in s.upper()
             and ("MD MAPPING" in s.upper() or "MD-MAPPING" in s.upper())),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] SAP Age LOV: 'Birken-MD Mappings' sheet not found")
            return {}
    
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
    
        header_idx  = None
        col_ppg     = None
        col_sap_age = None
    
        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                cell_upper = str(cell).strip().upper() if cell else ""
                if cell_upper == "PRODUCT PLANNING GROUP":
                    col_ppg    = j
                    header_idx = i
                if cell_upper == "SAP AGE" and header_idx == i:
                    col_sap_age = j
            if col_ppg is not None and col_sap_age is not None:
                break
            col_ppg = None  # reset if only half-matched
    
        if header_idx is None or col_ppg is None or col_sap_age is None:
            log.warning(
                "[DynamicMapping] Could not locate 'Product Planning Group'/'SAP Age' "
                "header columns in sheet '%s'", sheet_name,
            )
            return {}
    
        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_ppg >= len(row) or col_sap_age >= len(row):
                continue
            ppg_val     = row[col_ppg]
            sap_age_val = row[col_sap_age]
            if not ppg_val or not sap_age_val:
                continue
            key = str(ppg_val).strip().upper()
            val = str(sap_age_val).strip()
            if key and val:
                lov[key] = val
    
        log.info("[DynamicMapping] SAP Age LOV loaded from '%s': %d entries",
                 sheet_name, len(lov))
        return lov
    
    
    def _load_sap_product_division_lov(self, wb) -> dict[str, str]:
        """
        Load Product Planning Group → SAP Product Division mapping from
        'Birken-MD Mappings' sheet. Used to derive AT ParentID (PPH_{letter}) for XML output.
        """
        sheet_name = next(
            (s for s in wb.sheetnames
             if "BIRKEN" in s.upper()
             and ("MD MAPPING" in s.upper() or "MD-MAPPING" in s.upper())),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] SAP Product Division LOV: 'Birken-MD Mappings' sheet not found")
            return {}
    
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
    
        header_idx = None
        col_ppg    = None
        col_spd    = None
    
        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                cell_upper = str(cell).strip().upper() if cell else ""
                if cell_upper == "PRODUCT PLANNING GROUP":
                    col_ppg    = j
                    header_idx = i
                if cell_upper == "SAP PRODUCT DIVISION" and header_idx == i:
                    col_spd = j
            if col_ppg is not None and col_spd is not None:
                break
            col_ppg = None  # reset if only half-matched
    
        if header_idx is None or col_ppg is None or col_spd is None:
            log.warning(
                "[DynamicMapping] Could not locate 'Product Planning Group'/'SAP Product Division' "
                "header columns in sheet '%s'", sheet_name,
            )
            return {}
    
        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_ppg >= len(row) or col_spd >= len(row):
                continue
            ppg_val = row[col_ppg]
            spd_val = row[col_spd]
            if not ppg_val or not spd_val:
                continue
            key = str(ppg_val).strip().upper()
            val = str(spd_val).strip().upper()
            if key and val:
                lov[key] = val
    
        log.info("[DynamicMapping] SAP Product Division LOV loaded from '%s': %d entries",
                 sheet_name, len(lov))
        return lov

    def _load_material_lov(self, wb) -> dict[str, str]:
        """
        Dynamically load the Upper Material → Material mapping from the
        'Birken-MD Mappings' sheet. Scans for the header row containing
        'Upper Material' / 'Material' column labels, wherever they sit.
        """
        sheet_name = next(
            (s for s in wb.sheetnames
             if "BIRKEN" in s.upper() and ("MD MAPPING" in s.upper() or "MD-MAPPING" in s.upper())),
            None,
        )
        if not sheet_name:
            log.warning("[DynamicMapping] 'Birken-MD Mappings' sheet not found. Available: %s",
                        wb.sheetnames)
            return {}

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Find the header row + column indices for "Upper Material" and "Material"
        header_idx  = None
        col_upper   = None
        col_material = None

        for i, row in enumerate(rows):
            if not row:
                continue
            for j, cell in enumerate(row):
                if cell and str(cell).strip().upper() == "UPPER MATERIAL":
                    col_upper = j
                    header_idx = i
                if cell and str(cell).strip().upper() == "MATERIAL" and header_idx == i:
                    col_material = j
            if col_upper is not None and col_material is not None:
                break
            col_upper = None  # reset if only half-matched on this row

        if header_idx is None or col_upper is None or col_material is None:
            log.warning(
                "[DynamicMapping] Could not locate 'Upper Material'/'Material' "
                "header columns in sheet '%s'", sheet_name,
            )
            return {}

        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_upper >= len(row) or col_material >= len(row):
                continue
            upper_val    = row[col_upper]
            material_val = row[col_material]
            if not upper_val or not material_val:
                continue
            key = str(upper_val).strip().upper()
            val = str(material_val).strip()
            if key and val:
                lov[key] = val

        log.info("[DynamicMapping] Material LOV loaded from '%s': %d entries",
                 sheet_name, len(lov))
        return lov

    @staticmethod
    def _clean(v) -> str:
        if v is None:
            return ""
        s = str(v).strip()
        return "" if s in ("None", "nan", "\xa0", " ") else s

    def get_rule(self, attr_id: str) -> MappingRule | None:
        return self.rules_by_id.get(attr_id)

    def get_rules_with_source(self) -> list[MappingRule]:
        return [r for r in self.rules if r.is_direct_copy and r.source_field]

    def get_rules_with_default(self) -> list[MappingRule]:
        return [r for r in self.rules if r.default_value]

    def get_rules_with_rna(self) -> list[MappingRule]:
        return [r for r in self.rules if r.is_rna_lookup]

    def get_all_source_fields(self) -> set[str]:
        return {r.source_field for r in self.rules if r.source_field}


# ══════════════════════════════════════════════════════════════════════════════
# RNA LOADER
# ══════════════════════════════════════════════════════════════════════════════
class RNALoader:
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

    def _load(self):
        log.info("[RNA] Loading from: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames
             if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] Sheet not found")
            wb.close()
            return

        rows    = list(wb[sheet_name].iter_rows(values_only=True))
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
        c_comp    = _find("COMPANY CODE", "COMP CODE", "COMPCODE")
        c_sbu     = _find("SBU")
        c_bcode   = _find("BRANDCODE", "BRAND CODE", "REPORTING BRAND CODE MAPPED")
        c_btype   = _find("BRANDTYPE_DETAIL", "BRAND TYPE", "BRANDTYPE")
        c_bcat    = _find("BRANDCATEGORY", "BRAND CATEGORY")
        c_bgroup  = _find("BRANDGROUP", "BRAND GROUP")
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
                "BRANDGROUP":       _cell(row, c_bgroup),
                "BRANDTYPE_DETAIL": _cell(row, c_btype),
                "BRANDCATEGORY":    _cell(row, c_bcat),
                "BRAND STATUS":     _cell(row, c_bstatus),
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
# MDD LOADER
# ══════════════════════════════════════════════════════════════════════════════
class MDDLoader:

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        if "Core Attributes" in wb.sheetnames:
            ws   = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 2:
                hdr: dict[str, int] = {}
                for i, h in enumerate(rows[1]):
                    if h:
                        hdr[str(h).split("\n")[0].strip()] = i

                for row in rows[2:]:
                    aid = row[7] if len(row) > 7 else None
                    if not aid:
                        continue
                    aid = str(aid).strip()

                    def _v(key):
                        idx = hdr.get(key)
                        if idx is None or idx >= len(row):
                            return None
                        v = row[idx]
                        return str(v).strip() if v else None

                    self.attributes[aid] = {
                        "id":         aid,
                        "name":       _v("PIM Attribute Name"),
                        "validation": _v("Validation Base Type"),
                        "lov_name":   _v("Name of LOV"),
                    }
        self._load_lov_sheets(wb)
        self._load_retail_price_currency(wb)
        self.age_lov_by_label        = self._load_named_two_col_lov(wb, "Age LOV",        "ID",   "BY Age")
        self.age_desc_lov            = self._load_age_article_desc_lov(wb)
        self.material_lov_by_label   = self._load_named_two_col_lov(wb, "Material LOV",   "Code", "Material")
        self.silhouette_lov_by_label = self._load_named_two_col_lov(wb, "Silhouette LOV", "Code", "Silhouette")
        self.gender_lov_by_label = self._load_named_two_col_lov(
            wb, "Gender LOV",
            "SAP Gender Code",   # code col (C) — this IS AT_Gender now
            "Values of LOV",     # label col (A)
        )
        self.country_origin_lov_by_label = self._load_named_two_col_lov(
            wb, "Country Origin LOV",
            "Code",        # code col (A) — e.g. DE
            "Country Origin Name",   # label col (B) — e.g. Germany, 
        )
        # in MDDLoader._load(), alongside the other _load_named_two_col_lov calls:
        self.brand_type_lov_by_label = self._load_named_two_col_lov(
            wb, "Brand Type LOV",
            "Code",         # code col (A) -> MAA
            "Brand Type",   # label col (B) -> MAA BRAND
        )
        self.brand_status_lov_by_label = self._load_named_two_col_lov(
            wb, "Brand Status LOV",
            "Code",           # code col (A) -> A / D
            "Brand Status",   # label col (B) -> ACTIVE / DISCONTINUED
        )
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))
        
    def _load_named_two_col_lov(self, wb, sheet_keyword: str, code_header: str, label_header: str) -> dict[str, str]:
        sheet_name = next(
            (s for s in wb.sheetnames if sheet_keyword.upper() in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] '%s' sheet not found", sheet_keyword)
            return {}
    
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
    
        def _norm(cell) -> str:
            if not cell:
                return ""
            return re.sub(r"\s+", " ", str(cell).strip()).upper()
    
        code_target  = _norm(code_header)
        label_target = _norm(label_header)
    
        header_idx = None
        col_code = None
        col_label = None
    
        for i, row in enumerate(rows):
            if not row:
                continue
            row_col_code = None
            row_col_label = None
            for j, cell in enumerate(row):
                cell_norm = _norm(cell)
                if cell_norm == code_target and row_col_code is None:
                    row_col_code = j
                if cell_norm == label_target and row_col_label is None:
                    row_col_label = j
            if row_col_code is not None and row_col_label is not None:
                header_idx = i
                col_code = row_col_code
                col_label = row_col_label
                break
    
        if header_idx is None or col_code is None or col_label is None:
            log.warning(
                "[MDD] Could not locate '%s'/'%s' header columns in sheet '%s'",
                code_header, label_header, sheet_name,
            )
            return {}
    
        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_code >= len(row) or col_label >= len(row):
                continue
            code_val  = row[col_code]
            label_val = row[col_label]
            if not code_val or not label_val:
                continue
            key = str(label_val).strip().upper()
            val = str(code_val).strip()
            if key and val:
                lov[key] = val
    
        log.info("[MDD] '%s' LOV loaded from '%s': %d entries", sheet_keyword, sheet_name, len(lov))
        return lov
    
    def _load_age_article_desc_lov(self, wb) -> dict[str, str]:
        """
        Load Age -> Article Description mapping from the 'Age LOV' sheet's
        first table (columns: Age | Article Description | SAP Age Code).
        Returns dict: UPPERCASE Age label -> Article Description code.
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "AGE LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] 'Age LOV' sheet not found for Age→ArticleDescription table")
            return {}
    
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
    
        header_idx = None
        col_age = None
        col_artdesc = None
    
        for i, row in enumerate(rows):
            if not row:
                continue
            row_norm = {j: (str(c).strip().upper() if c else "") for j, c in enumerate(row)}
            if "AGE" in row_norm.values() and "ARTICLE DESCRIPTION" in row_norm.values():
                for j, v in row_norm.items():
                    if v == "AGE" and col_age is None:
                        col_age = j
                    if v == "ARTICLE DESCRIPTION" and col_artdesc is None:
                        col_artdesc = j
                header_idx = i
                break
    
        if header_idx is None or col_age is None or col_artdesc is None:
            log.warning(
                "[MDD] Could not locate 'Age'/'Article Description' header columns in sheet '%s'",
                sheet_name,
            )
            return {}
    
        lov: dict[str, str] = {}
        for row in rows[header_idx + 1:]:
            if not row or col_age >= len(row) or col_artdesc >= len(row):
                continue
            age_val  = row[col_age]
            desc_val = row[col_artdesc]
            if not age_val or not desc_val:
                continue
            key = str(age_val).strip().upper()
            val = str(desc_val).strip()
            if key and val:
                lov[key] = val
    
        log.info("[MDD] Age→ArticleDescription LOV loaded from '%s': %d entries", sheet_name, len(lov))
        return lov
    
    def resolve_age_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.age_lov_by_label.get(display_value.strip().upper(), "")
        
        
    def resolve_sap_age_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.age_desc_lov.get(display_value.strip().upper(), "")

    def resolve_material_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.material_lov_by_label.get(display_value.strip().upper(), "")

    def resolve_silhouette_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.silhouette_lov_by_label.get(display_value.strip().upper(), "")
        
        
    # new method, next to resolve_gender_id / resolve_material_id:
    def resolve_brand_type_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.brand_type_lov_by_label.get(display_value.strip().upper(), "")

    def resolve_brand_status_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.brand_status_lov_by_label.get(display_value.strip().upper(), "")

    def resolve_country_origin_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.country_origin_lov_by_label.get(display_value.strip().upper(), "")
        
    def resolve_gender_id(self, display_value: str) -> str:
        if not display_value:
            return ""
        return self.gender_lov_by_label.get(display_value.strip().upper(), "")

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

        self.lovs["CountryIdToName"]     = country_id_to_name
        self.lovs["RetailPriceCurrency"] = retail_lov

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

    def lookup_lov_id(self, lov_name: str, display_value: str) -> str:
        """
        Given a display value, return its LOV ID code.
    
        Returns "" (not the raw display value) when no match is found in the
        MDD LOV sheet, so callers can reliably tell "not in LOV" apart from
        "found" and decide to omit the attribute entirely rather than send
        an invalid/unmapped value into the XML.
        """
        if not display_value:
            return ""
        lov = self.lovs.get(lov_name, {})
        if not lov:
            return ""
        val = lov.get(display_value)
        if val is None:
            val = lov.get(display_value.strip().upper())
        if val is None:
            return ""
        val = str(val).strip()
        return val if val not in ("", "None", "nan") else ""


# ══════════════════════════════════════════════════════════════════════════════
# SOURCE LOADER
# ══════════════════════════════════════════════════════════════════════════════
class OFSMitraLoader:
    SOURCE_SHEET_NAME = "Detailed Planning"

    HEADER_DETECTION_KEYS = {
        "article #", "article number", "article no", "article code",
        "style code", "style number", "style no",
        "product code", "sku", "item number", "item no",
        "description", "brand", "gender", "color", "colour",
        "size range", "upper material", "model",
    }
    MAX_SCAN_ROWS = 20

    def __init__(self, path: Path):
        self.path              = path
        self.df                = pd.DataFrame()
        self.headers:          list[str]      = []
        self.column_map:       dict[str, str] = {}
        self.header_row_index: int            = 0
        self._load()

    def _detect_header_row(self, rows: list[tuple]) -> int:
        for i, row in enumerate(rows[: self.MAX_SCAN_ROWS]):
            if not row:
                continue
            cells = {
                str(c).strip().lower()
                for c in row
                if c is not None and str(c).strip()
            }
            matches = cells & self.HEADER_DETECTION_KEYS
            if len(matches) >= 2:
                log.info(
                    "[Source] Header row auto-detected at row %d (matched: %s)",
                    i + 1, sorted(matches),
                )
                return i
        log.warning(
            "[Source] Header not found in first %d rows — defaulting to row 0",
            self.MAX_SCAN_ROWS,
        )
        return 0

    def _load(self):
        log.info("[Source] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        if self.SOURCE_SHEET_NAME in wb.sheetnames:
            sheet_name = self.SOURCE_SHEET_NAME
        else:
            sheet_name = next(
                (s for s in wb.sheetnames
                 if "detailed" in s.lower() or "planning" in s.lower()),
                None,
            )
            if not sheet_name:
                log.error("[Source] Sheet '%s' not found. Available: %s",
                          self.SOURCE_SHEET_NAME, wb.sheetnames)
                wb.close()
                return

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            log.error("[Source] Sheet is empty")
            wb.close()
            return

        self.header_row_index = self._detect_header_row(rows)

        if len(rows) <= self.header_row_index + 1:
            log.error("[Source] Not enough rows after header")
            wb.close()
            return

        raw_headers    = rows[self.header_row_index]
        seen           = {}
        unique_headers = []

        for i, h in enumerate(raw_headers):
            col_name = _clean_str(h) if h else f"col_{i}"
            if col_name in seen:
                seen[col_name] += 1
                col_name = f"{col_name}_{seen[col_name]}"
            else:
                seen[col_name] = 0
            unique_headers.append(col_name)
            self.column_map[col_name.lower().strip()] = col_name

        self.headers = unique_headers
        log.info("[Source] Headers at row %d (%d cols): %s",
                 self.header_row_index + 1, len(self.headers), self.headers[:15])

        data_rows = rows[self.header_row_index + 1:]
        self.df   = pd.DataFrame(data_rows, columns=self.headers)

        key_candidates = [
            "Article #", "Article Number", "Article No", "Style Code",
            "Product Code", "SKU", unique_headers[0],
        ]
        key_col = next((c for c in key_candidates if c in self.df.columns), None)
        if key_col:
            col_data   = self.df[key_col]
            key_series = col_data.astype(str).str.strip()
            self.df    = self.df[
                col_data.notna() & (~key_series.isin(["", "None", "nan"]))
            ]
            log.info("[Source] Filtered on key column '%s'", key_col)
        else:
            log.warning("[Source] No key column found. Available: %s", self.headers[:10])

        self.df = self.df.reset_index(drop=True)
        wb.close()
        log.info("[Source] %d data rows loaded", len(self.df))

    def get_rows(self) -> list[dict]:
        return self.df.to_dict(orient="records")

    def find_column(self, field_name: str) -> str | None:
        if not field_name:
            return None

        # 1. Direct
        if field_name in self.headers:
            return field_name

        # 2. Normalised exact
        normalized = field_name.lower().strip()
        if normalized in self.column_map:
            return self.column_map[normalized]

        # 3. Case-insensitive scan
        for header in self.headers:
            if header.lower().strip() == normalized:
                return header

        # 4. Partial containment
        for header in self.headers:
            header_norm = header.lower().strip()
            if (header_norm in normalized
                    and len(header_norm) >= len(normalized) * 0.7):
                return header

        # 5. Fuzzy (≤ 2 char diff)
        for header in self.headers:
            header_norm = header.lower().strip()
            if abs(len(header_norm) - len(normalized)) <= 2:
                matches = sum(
                    1 for a, b in zip(header_norm, normalized) if a == b
                )
                if (matches >= len(normalized) - 2
                        and matches >= len(header_norm) - 2):
                    return header

        return None


# ══════════════════════════════════════════════════════════════════════════════
# TRANSFORMATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
class DynamicTransformationEngine:

    def __init__(
        self,
        mapping:       DynamicMappingLoader,
        mdd:           MDDLoader | None,
        rna:           RNALoader | None,
        source_loader: OFSMitraLoader,
        context:       dict,
    ):
        self.mapping = mapping
        self.mdd     = mdd
        self.rna     = rna
        self.source  = source_loader
        self.context = context

        self._rna_cache: dict = {}
        if rna:
            country_code = context.get("country_code", "ID")
            country_name = RNA_COUNTRY_MAP.get(country_code.upper(), country_code)
            self._rna_cache = rna.get_fuzzy(
                country_name,
                context.get("sbu",        "SP"),
                context.get("brand_code", BRAND_CODE),
            )
            log.info(
                "[Engine] RNA lookup country_code=%s → country_name=%s | cache=%s",
                country_code, country_name, self._rna_cache,
            )
    def transform_row(self, row: dict) -> dict:
        attrs: dict = {}

        for rule in self.mapping.rules:
            value = self._apply_rule(rule, row, attrs)
            if value:
                attrs[rule.stibo_attr_id] = value

        # ── Inject RNA attributes explicitly (Brand parity with Nike) ─────────
        if self._rna_cache:
            attrs.setdefault("AT_BrandGroup",    self._rna_cache.get("BRANDGROUP",       ""))
            attrs.setdefault("AT_BrandType",     self._rna_cache.get("BRANDTYPE_DETAIL", ""))
            attrs.setdefault("AT_BrandCategory", self._rna_cache.get("BRANDCATEGORY",    ""))
            attrs.setdefault("AT_BrandStatus",   self._rna_cache.get("BRAND STATUS",     ""))
            
        if self.mdd:
            bt_raw = attrs.get("AT_BrandType", "")
            if bt_raw:
                bt_id = self.mdd.resolve_brand_type_id(bt_raw)
                if bt_id:
                    attrs["AT_BrandType"] = bt_id
                else:
                    log.warning(
                        "[BrandTypeLOV] AT_BrandType value '%s' not found in MDD 'Brand Type LOV' sheet — "
                        "AT_BrandType will NOT be sent for this row", bt_raw,
                    )
                    attrs.pop("AT_BrandType", None)

            bs_raw = attrs.get("AT_BrandStatus", "")
            if bs_raw:
                bs_id = self.mdd.resolve_brand_status_id(bs_raw)
                if bs_id:
                    attrs["AT_BrandStatus"] = bs_id
                else:
                    log.warning(
                        "[BrandStatusLOV] AT_BrandStatus value '%s' not found in MDD 'Brand Status LOV' sheet — "
                        "AT_BrandStatus will NOT be sent for this row", bs_raw,
                    )
                    attrs.pop("AT_BrandStatus", None)

        self._apply_formulas(attrs, row)

        # Context (always overwrite)
        attrs["AT_Country"]     = self.context.get("country_code", "ID")
        attrs["AT_CompanyCode"] = self.context.get("comp_code",    "0888")
        attrs["AT_SBU"]         = self.context.get("sbu",          "SP")
        attrs["AT_Brand"]       = self.context.get("brand_code",   BRAND_CODE)

        # AT_Season / AT_SeasonYear: always derived from the filename-parsed
        # season (context["season"]), never from row/source data. Overrides
        # anything the row-based fallback in _apply_formulas may have set.
        season_from_filename = self.context.get("season", "")
        if season_from_filename:
            pfx, yr = derive_season(season_from_filename)
            if pfx:
                attrs["AT_Season"] = pfx
            if yr:
                attrs["AT_SeasonYear"] = yr

        return attrs

    def _apply_rule(self, rule: MappingRule, row: dict, attrs: dict) -> str:
        # 1. RNA lookup
        if rule.is_rna_lookup and rule.source_field:
            rna_key_map = {
                "BRANDGROUP":       "BRANDGROUP",
                "BRANDTYPE_DETAIL": "BRANDTYPE_DETAIL",
                "BRANDCATEGORY":    "BRANDCATEGORY",
                "BRAND_STATUS":     "BRAND STATUS",
                "BRAND STATUS":     "BRAND STATUS",
            }
            rna_key = rule.source_field.upper().replace(" ", "_")
            return self._rna_cache.get(rna_key_map.get(rna_key, rna_key), "")

        # 2. Direct field copy
        if rule.is_direct_copy and rule.source_field:
            actual_col = self.source.find_column(rule.source_field)
            if actual_col:
                raw_value = row.get(actual_col, "")
                return self._transform_value(rule, raw_value)

        # 3. Default value
        if rule.default_value:
            return self._normalize_default(rule)

        return ""

    def _transform_value(self, rule: MappingRule, raw_value) -> str:
        value = _clean_str(raw_value)
        if not value:
            return rule.default_value if rule.default_value else ""

        attr_id = rule.stibo_attr_id.upper()

        # Gender → store as LOV code (M / F / U)
        if "GENDER" in attr_id:
            return GENDER_MAP.get(_strip_gender_code_prefix(value), "U")

        # Heel Type / Heel Height / Occasion → resolve to LOV code here,
        # at capture time, instead of relying on a second lookup pass later
        if attr_id == "AT_HEELTYPE":
            return _resolve_heel_type(value)
        if attr_id == "AT_HEELHEIGHT":
            return _resolve_heel_height(value)
        if attr_id == "AT_OCCASION":
            return _resolve_occasion(value)

        # BY Age → normalize any raw source text to ADULT / CHILD
        if attr_id == "AT_BYAGE":
            return BY_AGE_FROM_PPG.get(value.upper(), "Adults")

        # Season
        if attr_id == "AT_SEASON":
            pfx, _ = derive_season(value)
            return pfx
        if attr_id == "AT_SEASONYEAR":
            _, yr = derive_season(value)
            return yr

        # Prices → numeric string
        if "PRICE" in attr_id or "FOB" in attr_id or "COST" in attr_id:
            try:
                return str(float(value))
            except (ValueError, TypeError):
                return ""

        # SAP style code: max 9 chars
        if attr_id == "AT_SAPSTYLECODE":
            return value[:9]

        return value

    def _normalize_default(self, rule: MappingRule) -> str:
        default = rule.default_value.lower()

        if "adult" in default:
            return "ADULT"
        if "unisex" in default:
            return "U"
        if "ea" in default:
            return "EA"
        if "usd" in default:
            return "USD"
        if "eur" in default:
            return "EUR"
        if "generic" in default:
            return "1"
        if "intercompany" in default or "zina" in default:
            return "ZINA"
        if "inline" in default:
            return "Inline"
        if "retailer" in default:
            return "Retailer"
        if "commercial" in default:     
            return "COMMERCIAL"

        return rule.default_value.strip()

    def _apply_formulas(self, attrs: dict, row: dict):

        # ── AT_PrincipalStyleDescription: join ALL description columns ─────────
        DESC_CANDIDATES = [
            "Description", "Product Description", "Style Description",
            "Article Description", "Description 1", "Description 2",
            "Description 3", "Short Description", "Long Description",
            "Product Name", "Style Name", "Article Name",
        ]
        # ── Product Division: determine footwear vs non-footwear ──────────────
        # Footwear divisions: Active, Classic, Homeshoes, Orthopedics, Papillio,
        #                      Professional, Shoes
        # Non-footwear:        Accessories, Cosmetics, Shoe Care
        NON_FOOTWEAR_DIVISIONS = {
            "ACCESSORIES", "COSMETICS", "SHOE CARE",
        }
        division_col = self.source.find_column("Product Division")
        if division_col:
            raw_division = _clean_str(row.get(division_col, "")).upper()
            attrs["_is_footwear"] = raw_division not in NON_FOOTWEAR_DIVISIONS
        else:
            attrs["_is_footwear"] = True  # default to footwear if column missing
            
        desc_parts = []
        for candidate in DESC_CANDIDATES:
            col = self.source.find_column(candidate)
            if col:
                val = _clean_str(row.get(col, ""))
                if val and val not in desc_parts:
                    desc_parts.append(val)
        if desc_parts:
            attrs["AT_PrincipalStyleDescription"] = " ".join(desc_parts)

        # ── AT_Color: pick from "Color" column → resolve LOV ID ───────────────
        color_col = self.source.find_column("Color")
        if color_col:
            raw_color = _clean_str(row.get(color_col, ""))
            if raw_color:
                if self.mdd:
                    lov_id = (
                        self.mdd.lookup_lov_id("Color Code", raw_color)
                        or self.mdd.lookup_lov_id("Colour Code", raw_color)
                    )
                    if lov_id and lov_id.strip().upper() in ("ACTIVE", "INACTIVE", "Y", "N", "YES", "NO"):
                        log.warning(...)
                        lov_id = ""
                    if lov_id:
                        attrs["AT_Color"] = lov_id
                        log.debug(...)
                    else:
                        log.warning(
                            "[Color] '%s' not found in any Color Code LOV sheet in the MDD file — "
                            "AT_Color will NOT be sent for this row", raw_color,
                        )
                        attrs.pop("AT_Color", None)     
                else:
                    log.warning(...)
                    attrs.pop("AT_Color", None)       

        # ── AT_PrincipalGenderCode + AT_PrincipalGenderDescription ────────────
        gender_col = self.source.find_column("Gender")
        if gender_col:
            raw_gender = _clean_str(row.get(gender_col, ""))
            if raw_gender:
                gender_code  = GENDER_MAP.get(_strip_gender_code_prefix(raw_gender), "U")
                gender_label = LOV_GENDER.get(gender_code, raw_gender)
                attrs["AT_PrincipalGenderCode"]        = gender_code   # M / F / U
                attrs["AT_PrincipalGenderDescription"] = gender_label  # Male / Female / Unisex
                log.debug("[Gender] raw='%s' → code='%s' label='%s'",
                          raw_gender, gender_code, gender_label)
        # ── AT_Gender: Gender → Birken-MD Mapping (SAP Gender) → Gender LOV ──────
        if gender_col and raw_gender:
            std_gender = self.mapping.gender_std_lov.get(raw_gender.upper(), raw_gender)
            if self.mdd:
                gender_id = self.mdd.resolve_gender_id(std_gender)
                if gender_id:
                    attrs["AT_Gender"] = gender_id
                else:
                    log.warning(
                        "[Gender] '%s' (raw '%s') not found in MDD 'Gender LOV' sheet — "
                        "AT_Gender will NOT be sent for this row", std_gender, raw_gender,
                    )
                    attrs.pop("AT_Gender", None)
            
            

        # ── AT_SAPAge: resolve from "Product Planning Group" via dynamic MD Mappings ──
        ppg_col = self.source.find_column("Product Planning Group")
        if ppg_col:
            raw_ppg = _clean_str(row.get(ppg_col, "")).upper()
            if raw_ppg:
                sap_age_val = self.mapping.sap_age_lov.get(raw_ppg, "")
                if sap_age_val:
                    attrs["AT_SAPAge"] = sap_age_val
                else:
                    log.warning(
                        "[SAPAge] Unmapped 'Product Planning Group' value: '%s' — "
                        "not found in Birken-MD Mappings sheet", raw_ppg,
                    )
                # ── SAP Product Division: used to derive ParentID (PPH_{letter}) ──
                sap_division = self.mapping.sap_product_division_lov.get(raw_ppg, "")
                if sap_division:
                    attrs["_sap_product_division"] = sap_division
                else:
                    log.warning(
                        "[SAPProductDivision] Unmapped 'Product Planning Group' value: '%s' — "
                        "not found in Birken-MD Mappings sheet (SAP Product Division table)", raw_ppg,
                    )
        
        # ── Age LOV validation: resolve AT_BYAge / AT_SAPAge against MDD "Age LOV" sheet ──
        if self.mdd:
            by_age_raw = attrs.get("AT_BYAge", "")
            if by_age_raw:
                by_age_id = self.mdd.resolve_age_id(by_age_raw)
                if by_age_id:
                    attrs["AT_BYAge"] = by_age_id
                else:
                    log.warning(
                        "[AgeLOV] AT_BYAge value '%s' not found in MDD 'Age LOV' sheet — "
                        "AT_BYAge will NOT be sent for this row", by_age_raw,
                    )
                    attrs.pop("AT_BYAge", None)
        
            sap_age_raw = attrs.get("AT_SAPAge", "")
            if sap_age_raw:
                sap_age_id = self.mdd.resolve_sap_age_id(sap_age_raw)
                if sap_age_id:
                    attrs["AT_SAPAge"] = sap_age_id
                else:
                    log.warning(
                        "[AgeLOV] AT_SAPAge value '%s' not found in MDD 'Age LOV' sheet — "
                        "AT_SAPAge will NOT be sent for this row", sap_age_raw,
                    )
                    attrs.pop("AT_SAPAge", None)
        

        # ── AT_HeelType: resolve to LOV code ──────────────────────────────────
        heel_type_col = self.source.find_column("Heel Type")
        if heel_type_col:
            raw_ht = _clean_str(row.get(heel_type_col, ""))
            if raw_ht:
                attrs["AT_HeelType"] = _resolve_heel_type(raw_ht)
                log.debug("[HeelType] raw='%s' → code='%s'", raw_ht, attrs["AT_HeelType"])
            # else: Accessories/Apparel — none of these four are sent
            
        # ── AT_Material: resolve from "Upper Material" via dynamic MD Mappings ──
        upper_material_col = self.source.find_column("Upper Material")
        if upper_material_col:
            raw_material = _clean_str(row.get(upper_material_col, "")).upper()
            if raw_material:
                material_code = self.mapping.material_lov.get(raw_material, "")
                if material_code:
                    attrs["AT_Material"] = material_code
                else:
                    log.warning(
                        "[Material] Unmapped 'Upper Material' value: '%s' — "
                        "not found in Birken-MD Mappings sheet", raw_material,
                    )
        
        # ── Material LOV validation: resolve AT_Material against MDD "Material LOV" sheet ──
        if self.mdd:
            material_raw = attrs.get("AT_Material", "")
            if material_raw:
                material_id = self.mdd.resolve_material_id(material_raw)
                if material_id:
                    attrs["AT_Material"] = material_id
                else:
                    log.warning(
                        "[MaterialLOV] AT_Material value '%s' not found in MDD 'Material LOV' sheet — "
                        "AT_Material will NOT be sent for this row", material_raw,
                    )
                    attrs.pop("AT_Material", None)
                    
                    
        co_raw = _clean_str(attrs.get("AT_CountryOrigin", ""))
        if not co_raw:
            attrs["AT_CountryOrigin"] = "DE"   # Birkenstock default
        elif co_raw.upper() == "DE":
            attrs["AT_CountryOrigin"] = "DE"   # already a code
        elif self.mdd:
            co_id = self.mdd.resolve_country_origin_id(co_raw)
            attrs["AT_CountryOrigin"] = co_id if co_id else "DE"
        # ── AT_SAPStyleCode / AT_Generic ──────────────────────────────────────
        product_code = attrs.get("AT_PrincipalStyleCode", "")

        if product_code:
            # AT_SAPStyleCode: max 9 chars
            attrs["AT_SAPStyleCode"] = product_code[:9]

            # AT_Generic: <brand_code> + max 9 alphanumeric chars
            brand_code = self.context.get("brand_code", BRAND_CODE)
            style_9 = _style_clean(product_code, 9)
            attrs["AT_Generic"] = f"{brand_code}{style_9}"

        # ── Retail Price Currency from country ────────────────────────────────
        if self.mdd and not attrs.get("AT_RetailPriceCurrency"):
            country_code = self.context.get("country_code", "ID")
            currency     = self.mdd.resolve_currency_from_country_code(country_code)
            attrs["AT_RetailPriceCurrency"] = currency or "USD"
            
            
        # ── AT_FOB: Distributor Wholesale Price [EUR] deducted by Discount % ──
        dist_price_col = (
            self.source.find_column("Distributor Wholesale Price [EUR]")
            or self.source.find_column("Distributor Wholesale Price")
        )
        discount_col = (
            self.source.find_column("Discount %")
            or self.source.find_column("Discount")
        )

        if dist_price_col:
            raw_price = _clean_str(row.get(dist_price_col, ""))
            try:
                price_val = float(raw_price) if raw_price else None
            except (ValueError, TypeError):
                price_val = None

            if price_val is not None:
                discount_frac = 0.0
                if discount_col:
                    raw_discount = _clean_str(row.get(discount_col, ""))
                    if raw_discount:
                        cleaned = raw_discount.replace("%", "").strip()
                        try:
                            discount_num = float(cleaned)
                            # "10" or "10%" → 0.10 ; already-fractional "0.1" stays as-is
                            discount_frac = discount_num / 100 if discount_num > 1 else discount_num
                        except (ValueError, TypeError):
                            discount_frac = 0.0

                fob_value = price_val * (1 - discount_frac)
                attrs["AT_FOB"] = f"{fob_value:.2f}"
                log.debug(
                    "[FOB] price=%.2f discount=%.4f → FOB=%.2f",
                    price_val, discount_frac, fob_value,
                )
            else:
                log.warning(
                    "[FOB] Could not parse 'Distributor Wholesale Price [EUR]' value: '%s'",
                    raw_price,
                )
        else:
            log.warning("[FOB] Source column 'Distributor Wholesale Price [EUR]' not found")

        # ── Copy CurrentPrice from OriginalPrice ──────────────────────────────
        if attrs.get("AT_OriginalPrice") and not attrs.get("AT_CurrentPrice"):
            attrs["AT_CurrentPrice"] = attrs["AT_OriginalPrice"]
            

        # ── Mandatory defaults ────────────────────────────────────────────────
        attrs.setdefault("AT_SAPArticleCategory", "1")
        attrs.setdefault("AT_UOM",                "EA")
        attrs.setdefault("AT_FOBCurrency",        "EUR")
        attrs.setdefault("AT_MaterialType",       "ZINA")
        attrs.setdefault("AT_SAPProductFlag",     "A")
        # attrs.setdefault("AT_BYArticleType",      "Inline")
        attrs.setdefault("AT_BYAge",              "ADULT")
        attrs.setdefault("AT_PatternPrint",       "PLAIN") 
        attrs.setdefault("AT_BCI",                "COMMERCIAL")
        attrs["AT_ArticleStatus"] = "A"
        # ── AT_PatternPrint: force proper LOV ID (always uppercase) ───────────
        pp_raw = _clean_str(attrs.get("AT_PatternPrint", "")).upper()
        attrs["AT_PatternPrint"] = pp_raw if pp_raw else "PLAIN"
        
        # ── AT_PricingDistributionChannel: force LOV ID "01" (always) ─────────
        attrs["AT_PricingDistributionChannel"] = "01"
        
        
         # ── AT_PrincipalHierarchy L1–L5 ──────────────────────────────────────
        _HIERARCHY_MAP = [
            ("AT_PrincipalMerchandiseHierarchyL1", "Product Division"),
            ("AT_PrincipalMerchandiseHierarchyL2", "Product Planning Group"),
            ("AT_PrincipalMerchandiseHierarchyL3", "Product Story"),
            ("AT_PrincipalMerchandiseHierarchyL4", "ASPA Tier Segmentation"),
            ("AT_PrincipalMerchandiseHierarchyL5", "Model"),
        ]
        for attr_id, source_col in _HIERARCHY_MAP:
            col = self.source.find_column(source_col)
            if col:
                val = _clean_str(row.get(col, ""))
                if val:
                    attrs[attr_id] = val
                    log.debug("[Hierarchy] %s = '%s' (from '%s')", attr_id, val, col)
        

        # ── AT_Silhouette: Model column → Birken-MD Mappings lookup ──────────
        model_col = self.source.find_column("Model")
        if model_col:
            raw_model = re.sub(r"\s+", " ", _clean_str(row.get(model_col, "")).strip()).upper()
            if raw_model:
                silhouette_val = self.mapping.silhouette_lov.get(raw_model, "")
                if silhouette_val:
                    attrs["AT_Silhouette"] = silhouette_val
                    log.debug("[Silhouette] Model='%s' → '%s'", raw_model, silhouette_val)
                else:
                    log.warning(
                        "[Silhouette] Unmapped Model value: '%s' — "
                        "not found in Birken-MD Mappings sheet", raw_model,
                    )
                    attrs.pop("AT_Silhouette", None)

        # ── Silhouette LOV validation: resolve AT_Silhouette against MDD "Silhouette LOV" sheet ──
        if self.mdd:
            silhouette_raw = attrs.get("AT_Silhouette", "")
            if silhouette_raw:
                silhouette_id = self.mdd.resolve_silhouette_id(silhouette_raw)
                if silhouette_id:
                    # Zero-pad single-digit codes (e.g. '9' → '09')
                    sid = str(silhouette_id).strip()
                    if len(sid) == 1 and sid.isdigit():
                        sid = f"0{sid}"
                    attrs["AT_Silhouette"] = sid
                else:
                    log.warning(
                        "[SilhouetteLOV] AT_Silhouette value '%s' not found in MDD 'Silhouette LOV' sheet — "
                        "AT_Silhouette will NOT be sent for this row", silhouette_raw,
                    )
                    attrs.pop("AT_Silhouette", None)
        # ── AT_ShortDescriptionEN / AT_LongDescriptionEN: Model → Birken Ecom Mappings ──
        if model_col:
            if raw_model:
                short_desc = self.mapping.ecom_short_desc_lov.get(raw_model, "")
                if short_desc:
                    attrs["AT_ShortDescriptionEN"] = short_desc
                else:
                    log.warning(
                        "[EcomShortDesc] Unmapped Model value: '%s' — "
                        "not found in Birken Ecom Mappings sheet", raw_model,
                    )
                    attrs.pop("AT_ShortDescriptionEN", None)   # ← clears the stale raw copy
        
                long_desc = self.mapping.ecom_long_desc_lov.get(raw_model, "")
                if long_desc:
                    attrs["AT_LongDescriptionEN"] = long_desc
                else:
                    log.warning(
                        "[EcomLongDesc] Unmapped Model value: '%s' — "
                        "not found in Birken Ecom Mappings sheet", raw_model,
                    )
                    attrs.pop("AT_LongDescriptionEN", None)   # ← clears the stale raw copy

# ══════════════════════════════════════════════════════════════════════════════
# XML GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

# Attributes that must be REMOVED entirely from the output XML
_REMOVED_ATTRS: frozenset[str] = frozenset({
    "AT_MainEANIndicator",
    "AT_GenericDescription",  
    "AT_BYArticleType",                  
})

# Attributes written explicitly in build_products() — skip in generic loop
_EXPLICIT_ATTRS: frozenset[str] = frozenset({
    "AT_InboundGenericCode",
    "AT_Generic",
    "AT_SAPStyleCode",
    "AT_Createdon", "AT_Createdby",
    "AT_CompanyCode",
    "AT_SBU",
    "AT_Country",
    "AT_Brand",
    "AT_BrandGroup",
    "AT_BrandType",
    "AT_BrandCategory",
    "AT_BrandStatus",
    "AT_CountrySize",
    "AT_PackDetails",
    "AT_Silhouette",
    # ── Principal Hierarchy ───────────────────────────────────────────────────
    "AT_PrincipalMerchandiseHierarchyL1",
    "AT_PrincipalMerchandiseHierarchyL2",
    "AT_PrincipalMerchandiseHierarchyL3",
    "AT_PrincipalMerchandiseHierarchyL4",
    "AT_PrincipalMerchandiseHierarchyL5",
    "AT_FOBCurrency",
    "AT_RetailPriceCurrency",
    "AT_SAPProductFlag",
    "AT_MaterialType",
    "AT_SAPArticleCategory",
    "AT_UOM",
    "AT_BYAge",
    "AT_BYGender",
    "AT_Color",
    "AT_PrincipalGenderCode",
    "AT_PrincipalGenderDescription",
    "AT_Gender",
    "AT_CountryOrigin",
    "AT_Material",
    "AT_HeelHeight",
    "AT_HeelType",
    "AT_Occasion",
    "AT_ArticleStatus",
    "AT_CountrySize",
    "AT_PackDetails",
    "AT_PatternPrint",
    "AT_SAPAge", 
    "AT_PricingDistributionChannel",
    "AT_Season",
    "AT_BCI", 
})


class XMLGenerator:
    def __init__(self, articles: list[dict], context: dict):
        self.articles    = articles
        self.context     = context
        self.season_code = context.get("season", "SS27")
        self.brand_code  = context.get("brand_code", BRAND_CODE)

    # ── Classifications ───────────────────────────────────────────────────────
    def build_classifications(self) -> ET.Element:
        cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

        pfx, yr        = derive_season(self.season_code)
        season_id      = season_id_full(self.brand_code, self.season_code)
        batches_par    = "CLH_BirkenstockBatches"
        sea_name       = LOV_SEASON.get(pfx, pfx)
        season_display = f"Birkenstock {sea_name} {yr}".strip()
        season_short   = f"{pfx} {yr}".strip()

        season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
        season_cls.set("ID",         season_id)
        season_cls.set("UserTypeID", "CLS_Season")
        season_cls.set("ParentID",   batches_par)
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

    # ── Parent hierarchy ──────────────────────────────────────────────────────
    def _determine_parent_id(self, art: dict) -> str:
        # Primary: SAP Product Division, resolved via Product Planning Group
        # against the Birken-MD Mappings sheet.
        sap_division = (art.get("_sap_product_division", "") or "").strip().upper()
        if sap_division:
            if sap_division == "J":
                sap_division = "E"
            return f"PPH_{sap_division}-TempSubCat"

        # Fallback: derive from raw Product Division text / footwear flag,
        # in case Product Planning Group didn't resolve for this row.
        division = art.get("AT_PrincipalMerchandiseHierarchyL1", "")
        letter = _get_division_code(division)
        if letter == "X":
            letter = "F" if art.get("_is_footwear", True) else "E"

        return f"PPH_{letter}-TempSubCat"

    # ── Inbound key ───────────────────────────────────────────────────────────
    def _analyze_style_colors(self) -> dict[str, set[str]]:
        style_colors: dict[str, set[str]] = {}
        for art in self.articles:
            sc = art.get("AT_PrincipalStyleCode", "")
            cc = art.get("AT_PrincipalColorCode", "") or ""
            if sc:
                style_colors.setdefault(sc, set())
                if cc:
                    style_colors[sc].add(cc)
        return style_colors

    def _build_inbound_key(self, art: dict, style_colors: dict[str, set[str]]) -> str:
        style_code = art.get("AT_PrincipalStyleCode", "")
        color_code = art.get("AT_PrincipalColorCode", "") or ""
        if not style_code:
            return ""
        has_multi = len(style_colors.get(style_code, set())) > 1
        if has_multi and color_code:
            return f"{self.brand_code}{style_code}{color_code}"
        return f"{self.brand_code}{style_code}"

    # ── Products ──────────────────────────────────────────────────────────────
    def build_products(self) -> ET.Element:
        products_root = ET.Element(f"{{{STIBO_NS}}}Products")
        season_id     = season_id_full(self.brand_code, self.season_code)
        style_colors  = self._analyze_style_colors()

        for art in self.articles:
            key_article = self._build_inbound_key(art, style_colors)
            if not key_article or key_article == self.brand_code:
                continue

            parent_id    = self._determine_parent_id(art)
            country_code = art.get("AT_Country", "")

            product = ET.SubElement(products_root, f"{{{STIBO_NS}}}Product")
            product.set("UserTypeID", "PRD_GenericArticle")
            product.set("ParentID",   parent_id)

            kv = ET.SubElement(product, f"{{{STIBO_NS}}}KeyValue")
            kv.set("KeyID", "KEY_InboundArticle")
            kv.text = key_article

            name_text = art.get("AT_PrincipalStyleDescription", "")
            if name_text:
                ET.SubElement(product, f"{{{STIBO_NS}}}Name").text = name_text

            cr_merch = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
            cr_merch.set("ClassificationID", "CLH_BirkenstockArticles")
            cr_merch.set("Type",             "CPL_Merchandiser")

            cr_unconf = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
            cr_unconf.set("ClassificationID", f"{season_id}UA")
            cr_unconf.set("Type",             "CPL_UnConfirmedForSeason")

            values = ET.SubElement(product, f"{{{STIBO_NS}}}Values")

            # ── System / org ──────────────────────────────────────────────────
            _val(values, "AT_InboundGenericCode", key_article)
            _multival(values, "AT_CompanyCode", art.get("AT_CompanyCode", ""))
            _multival(values, "AT_SBU",         art.get("AT_SBU",         ""))

            if country_code:
                _val(values, "AT_Country", "", id_val=country_code)

            brand_val = art.get("AT_Brand", self.brand_code)
            if brand_val:
                _val(values, "AT_Brand", "", id_val=brand_val)
                
            # ── Season: LOV ID (from filename, e.g. 'SS', 'FW', 'SM') ─────────
            season_lov_id = art.get("AT_Season", "")
            if season_lov_id:
                _val(values, "AT_Season", "", id_val=season_lov_id)

            # ── Brand Group / Type / Category (RNA — Nike parity) ─────────────
            brand_group = art.get("AT_BrandGroup", "")
            if brand_group:
                _val(values, "AT_BrandGroup", "", id_val=brand_group)

            brand_type = art.get("AT_BrandType", "")
            if brand_type:
                _val(values, "AT_BrandType", "", id_val=brand_type)

            brand_category = art.get("AT_BrandCategory", "")
            if brand_category:
                _val(values, "AT_BrandCategory", "", id_val=brand_category)
                
            brand_status = art.get("AT_BrandStatus", "")
            if brand_status:
                _val(values, "AT_BrandStatus", "", id_val=brand_status)

            # ── Currencies ────────────────────────────────────────────────────
            _val(values, "AT_FOBCurrency", "", id_val="EUR")

            if country_code:
                rpc_id = COUNTRY_CURRENCY.get(country_code.strip().upper(), "USD")
                _val(values, "AT_RetailPriceCurrency", "", id_val=rpc_id)

            # ── SAP / article flags ───────────────────────────────────────────
            _val(values, "AT_SAPProductFlag",     "", id_val="A")
            _val(values, "AT_MaterialType",       "", id_val=art.get("AT_MaterialType",       "ZINA"))
            _val(values, "AT_SAPArticleCategory", "", id_val=art.get("AT_SAPArticleCategory", "1"))
            _val(values, "AT_UOM",                "", id_val=art.get("AT_UOM",                "EA"))
            
            # ── BCI: LOV ID (default COMMERCIAL) ───────────────────────────────
            bci_id = art.get("AT_BCI", "COMMERCIAL")
            if bci_id:
                _val(values, "AT_BCI", "", id_val=bci_id)

            # ── BY Age ────────────────────────────────────────────────────────
            by_age = art.get("AT_BYAge", "ADULT")
            if by_age:
                _val(values, "AT_BYAge", "", id_val=by_age)
                
            # ── SAP Age: LOV ID (from Product Planning Group lookup) ───────────
            sap_age_id = art.get("AT_SAPAge", "")
            if sap_age_id:
                _val(values, "AT_SAPAge", "", id_val=sap_age_id)

            # ── BY Gender: LOV ID (M / F / U) ─────────────────────────────────
            by_gender_code  = art.get("AT_BYGender", "U")
            by_gender_label = LOV_GENDER.get(by_gender_code, by_gender_code)
            if by_gender_code:
                _val(values, "AT_BYGender", by_gender_label, id_val=by_gender_code)

            # ── Principal Gender Code + Description ───────────────────────────
            principal_gender = art.get("AT_PrincipalGenderCode", "U")
            if principal_gender:
                _val(values, "AT_PrincipalGenderCode", "", id_val=principal_gender)
            
            gender_id = art.get("AT_Gender", "")
            if gender_id:
                _val(values, "AT_Gender", "", id_val=gender_id)
                

            principal_gender_desc = art.get("AT_PrincipalGenderDescription", "")
            if principal_gender_desc:
                _val(values, "AT_PrincipalGenderDescription", principal_gender_desc)

            # ── Color: LOV ID ─────────────────────────────────────────────────
            color_id = art.get("AT_Color", "")  # ← removed AT_PrincipalColorCode fallback
            if color_id:
                _val(values, "AT_Color", "", id_val=color_id)

            # ── Country of Origin ─────────────────────────────────────────────
            # Birkenstock = Germany → LOV ID "DE" — confirm with Stibo/MDD team
            country_origin = art.get("AT_CountryOrigin", "DE")
            if country_origin:
                _val(values, "AT_CountryOrigin", "", id_val=country_origin)
                
            # ── Material: LOV ID from Upper Material (dynamic MD Mappings) ────
            material_id = art.get("AT_Material", "")
            if material_id:
                _val(values, "AT_Material", "", id_val=material_id)
                
                
            # ── Pattern/Print: LOV ID (default PLAIN for footwear + accessories) ──
            pattern_print_id = art.get("AT_PatternPrint", "")
            if pattern_print_id:
                _val(values, "AT_PatternPrint", "", id_val=pattern_print_id)
                
                
            # ── Pricing Distribution Channel: LOV ID (default 01) ─────────────────
            pricing_channel_id = art.get("AT_PricingDistributionChannel", "")
            if pricing_channel_id:
                _val(values, "AT_PricingDistributionChannel", "", id_val=pricing_channel_id)
                
            # ── Silhouette: text value from Model → MD Mappings lookup ────────
            silhouette_val = art.get("AT_Silhouette", "")
            if silhouette_val:
                _val(values, "AT_Silhouette", "", id_val=silhouette_val)

            # ── Footwear-only: Heel Height / Heel Type / Occasion / CountrySize ──
            is_footwear = art.get("_is_footwear", True)

            if is_footwear:
                heel_height_code = _resolve_heel_height(art.get("AT_HeelHeight", ""))
                if heel_height_code:
                    _val(values, "AT_HeelHeight", "", id_val=heel_height_code)

                heel_type_code = _resolve_heel_type(art.get("AT_HeelType", ""))
                if heel_type_code:
                    _val(values, "AT_HeelType", "", id_val=heel_type_code)

                occasion_code = _resolve_occasion(art.get("AT_Occasion", ""))
                if occasion_code:
                    _val(values, "AT_Occasion", "", id_val=occasion_code)

                country_size = COUNTRY_SIZE_SYSTEM.get(country_code.strip().upper(), "")
                if country_size:
                    _val(values, "AT_CountrySize", "", id_val=country_size)
            # else: Accessories/Cosmetics/Legwear/Shoe Care — none of these four sent

            # ── Article Status (not gated) ─────────────────────────────────────
            article_status = art.get("AT_ArticleStatus", "")
            if article_status:
                _val(values, "AT_ArticleStatus", "", id_val=article_status)
                
                
                
            # ── Principal Hierarchy L1–L5 + Silhouette ────────────────────────
            for h_attr in (
                "AT_PrincipalMerchandiseHierarchyL1",
                "AT_PrincipalMerchandiseHierarchyL2",
                "AT_PrincipalMerchandiseHierarchyL3",
                "AT_PrincipalMerchandiseHierarchyL4",
                "AT_PrincipalMerchandiseHierarchyL5",
            ):
                h_val = art.get(h_attr, "")
                if h_val and str(h_val).strip() not in ("", "None", "nan"):
                    _val(values, h_attr, str(h_val).strip())
                    
            # ── All remaining text AT_ attributes ─────────────────────────────
            _skip = _EXPLICIT_ATTRS | _REMOVED_ATTRS
            for attr_id, value in art.items():
                if not attr_id.startswith("AT_"):
                    continue
                if attr_id in _skip:
                    continue
                v = str(value).strip() if value else ""
                if v and v not in ("None", "nan"):
                    _val(values, attr_id, v)

        return products_root

    # ── Generate ──────────────────────────────────────────────────────────────
    def generate(self, output_path: Path):
        export_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

        with open(output_path, "w", encoding="utf-8") as f:
            open_step_xml(f, export_time)

            cls_elem = self.build_classifications()
            f.write(strip_xmlns(ET.tostring(cls_elem, encoding="unicode")))
            f.write("\n")

            prod_elem = self.build_products()
            f.write(strip_xmlns(ET.tostring(prod_elem, encoding="unicode")))
            f.write("\n")

            close_step_xml(f)

        log.info("[XML] Generated: %s (%d articles)", output_path.name, len(self.articles))
        return output_path


# ══════════════════════════════════════════════════════════════════════════════
# FILENAME METADATA PARSER
# ══════════════════════════════════════════════════════════════════════════════
def _parse_meta_from_filename(path: Path) -> dict:
    """
    MAA naming convention (dash-separated, parsed from the RIGHT so any
    embedded dashes/spaces in the file_type segment don't break parsing):

    {comp_code}-{sbu}-{brand}-{file_type...}-{multi_mono}-{season}-{country}-{seq}.xlsx

    Example:
    0888-SP-BIRKENSTOCK-OFS-FC-PT Mitra-SS27 MAA Sport Inline-Multi-SM28-MY-1.xlsx
    """
    stem = path.stem
    log.info("[Filename] Parsing: '%s'", stem)

    result = {
        "brand_code":   BRAND_CODE,
        "comp_code":    "0888",
        "sbu":          "SP",
        "season":       "",
        "country_code": "ID",
        "multi_mono":   "MONO",
    }

    parts = stem.split("-")

    if len(parts) < 5:
        # Legacy underscore-based test filenames (e.g. OFS_FC_Mitra_ss27)
        parts_us = stem.split("_")
        if len(parts_us) >= 2:
            result["sbu"] = SBU_MAP.get(parts_us[1].strip().upper(), "SP")
        if len(parts_us) >= 3:
            distributor = parts_us[2].strip().upper()
            country = MITRA_COUNTRY_MAP.get(distributor)
            if country:
                result["country_code"] = country
            elif re.match(r"^[A-Z]{2,3}$", distributor):
                result["country_code"] = distributor
        if len(parts_us) >= 4:
            season_token = parts_us[3].strip().upper()
            if re.match(r"^[A-Z]{2}\d{2,4}$", season_token):
                result["season"] = season_token
        log.info(
            "[Filename] (legacy underscore format) → comp=%s sbu=%s season=%s country=%s",
            result["comp_code"], result["sbu"], result["season"], result["country_code"],
        )
        return result

    # ── Parse from the right: ...-{multi_mono}-{season}-{country}-{seq} ──────
    country_tok    = parts[-2].strip().upper()
    season_tok     = parts[-3].strip().upper()
    multi_mono_tok = parts[-4].strip()

    result["comp_code"] = parts[0].strip()
    result["sbu"]       = SBU_MAP.get(parts[1].strip().upper(), parts[1].strip().upper())

    # parts[2] is the brand segment, e.g. "BIRKENSTOCK (BCX)" or "BIRKENSTOCK BCX"
    brand_seg = parts[2].strip().upper() if len(parts) > 2 else ""

    brand_match = re.search(r"\(([A-Z0-9]{2,6})\)", brand_seg)
    if brand_match:
        result["brand_code"] = brand_match.group(1)
    else:
        brand_match = re.search(r"BIRKENSTOCK\s+([A-Z0-9]{2,6})$", brand_seg)
        result["brand_code"] = brand_match.group(1) if brand_match else BRAND_CODE

    if re.match(r"^[A-Z]{2,3}$", country_tok):
        result["country_code"] = country_tok
    else:
        mapped = MITRA_COUNTRY_MAP.get(country_tok)
        if mapped:
            result["country_code"] = mapped

    if re.match(r"^[A-Z]{2}\d{2,4}$", season_tok):
        result["season"] = season_tok

    result["multi_mono"] = multi_mono_tok

    log.info(
        "[Filename] → comp=%s  sbu=%s  season=%s  country=%s  multi_mono=%s",
        result["comp_code"], result["sbu"], result["season"],
        result["country_code"], result["multi_mono"],
    )
    return result


# ══════════════════════════════════════════════════════════════════════════════
# OUTPUT FILENAME BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def _build_output_filename(context: dict, seq: int = 1) -> str:
    """
    MAA convention:
    {comp_code}-{sbu}-{brand}-{file_type}-{multi_mono}-{season_full}-{country}-{seq}.xml

    Example: 0888-SP-BCK-Birkenstock Inline-MONO-SS2027-ID-1.xml
    """
    comp_code  = context.get("comp_code",    "0888")
    sbu        = context.get("sbu",          "SP")
    brand      = context.get("brand_code",   BRAND_CODE)
    file_type  = context.get("file_type",    "Birkenstock Inline")
    multi_mono = context.get("multi_mono",   "MONO")
    country    = context.get("country_code", "ID").upper()
    season_raw = context.get("season",       "SS27").upper()

    pfx, yr     = derive_season(season_raw)
    season_full = f"{pfx}{yr}"

    return (
        f"{comp_code}-{sbu}-{brand}-{file_type}"
        f"-{multi_mono}-{season_full}-{country}-{seq}.xml"
    )


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def run(args: dict, auditor=None) -> dict:
    log.info("=" * 60)
    log.info("  Birkenstock OFS Mitra ETL (v1.2)")
    log.info("=" * 60)

    linelist_path   = Path(args.get("linelist_path",   ""))
    mdd_path        = Path(args.get("mdd_path",        ""))
    attributes_path = Path(args.get("attributes_path", ""))
    mapping_path    = Path(args.get("mapping_path",    ""))

    # ── Context ───────────────────────────────────────────────────────────────
    file_meta = {}
    if linelist_path and linelist_path.name:
        file_meta = _parse_meta_from_filename(linelist_path)

    context = {
        "brand_code":   file_meta.get("brand_code") or args.get("brand_code") or BRAND_CODE,
        "country_code": file_meta.get("country_code") or args.get("country_code") or "ID",
        "comp_code":    file_meta.get("comp_code")    or args.get("comp_code")    or "0888",
        "sbu":          file_meta.get("sbu")          or args.get("sbu")          or "SP",
        "season":       file_meta.get("season")       or args.get("season")       or "SS27",
        "file_type":    args.get("file_type",    "Birkenstock Inline"),
        "multi_mono":   file_meta.get("multi_mono") or args.get("multi_mono", "MONO"),
    }

    log.info(
        "[Context] brand=%s  comp=%s  sbu=%s  season=%s  country=%s",
        context["brand_code"], context["comp_code"], context["sbu"],
        context["season"],     context["country_code"],
    )

    # ── Mapping (required) ────────────────────────────────────────────────────
    if not mapping_path.exists():
        log.error("[Mapping] Not found: %s", mapping_path)
        return {"status": "error", "message": f"Mapping file not found: {mapping_path}"}

    mapping = DynamicMappingLoader(mapping_path)
    if not mapping.rules:
        log.error("[Mapping] No rules loaded")
        return {"status": "error", "message": "No mapping rules found"}

    log.info("[Mapping] %d direct field mappings:", len(mapping.get_rules_with_source()))
    for r in mapping.get_rules_with_source()[:15]:
        log.info("  %-40s → '%s'", r.stibo_attr_id, r.source_field)

    # ── Supporting files ──────────────────────────────────────────────────────
    mdd: MDDLoader | None = None
    rna: RNALoader | None = None

    if mdd_path.exists():
        mdd = MDDLoader(mdd_path)
    else:
        log.warning("[MDD] Not found: %s", mdd_path)

    if attributes_path.exists():
        rna = RNALoader(attributes_path)
    else:
        log.warning("[RNA] Not found: %s", attributes_path)

    # ── Source data ───────────────────────────────────────────────────────────
    if not linelist_path.exists():
        log.error("[Source] Not found: %s", linelist_path)
        return {"status": "error", "message": f"Source file not found: {linelist_path}"}

    source_loader = OFSMitraLoader(linelist_path)
    rows          = source_loader.get_rows()

    if not rows:
        log.error("[Source] No data rows")
        return {"status": "error", "message": "No data rows found"}

    # ── Debug: column vs mapping resolution ──────────────────────────────────
    log.info("[DEBUG] Source columns (%d):", len(source_loader.headers))
    for col in source_loader.headers:
        sample = _clean_str(rows[0].get(col, "")) if rows else ""
        log.info("  %-45s → '%s'", col, sample[:50])

    log.info("[DEBUG] Mapping resolution:")
    for rule in mapping.get_rules_with_source():
        actual = source_loader.find_column(rule.source_field)
        status = "✓ FOUND" if actual else "✗ MISSING"
        log.info("  %-40s source=%-30s %s → actual='%s'",
                 rule.stibo_attr_id, f"'{rule.source_field}'", status, actual or "—")

    # ── Test row limit ────────────────────────────────────────────────────────
    # if TEST_ROW_RANGE:
    #     start, end = TEST_ROW_RANGE
    #     # Convert 1-indexed inclusive range to 0-indexed slice
    #     rows = rows[start - 1 : end]
    #     log.info("[TEST MODE] Limited to data rows %d–%d (%d rows)",
    #               start, end, len(rows))

    # ── Season fallback from data ─────────────────────────────────────────────
    if not context["season"]:
        season_col = (
            source_loader.find_column("Season")
            or source_loader.find_column("Launch Season")
        )
        if season_col and rows:
            season_from_data = _clean_str(rows[0].get(season_col, ""))
            if season_from_data:
                context["season"] = season_from_data
                log.info("[Context] Season from data: %s", context["season"])

    # ── Transform ─────────────────────────────────────────────────────────────
    engine           = DynamicTransformationEngine(mapping, mdd, rna, source_loader, context)
    mapped_articles: list[dict] = []
    skipped          = 0

    for i, row in enumerate(rows):
        try:
            article = engine.transform_row(row)

            if article.get("AT_Generic"):
                mapped_articles.append(article)
            else:
                skipped += 1
                if skipped <= 3:
                    log.warning(
                        "[Filter] Row %d skipped — AT_Generic empty. "
                        "AT_PrincipalStyleCode='%s'  keys=%s",
                        i + 1,
                        article.get("AT_PrincipalStyleCode", "EMPTY"),
                        [k for k, v in article.items() if v][:8],
                    )
        except Exception as e:
            log.warning("[Transform] Row %d error: %s", i + 1, e)
            skipped += 1
            continue

    log.info("[Transform] %d mapped | %d skipped", len(mapped_articles), skipped)

    if not mapped_articles:
        log.error("[Transform] No articles generated — check source field mapping")
        return {
            "status":        "error",
            "message":       "No articles generated",
            "article_count": 0,
        }

    # ── Generate XML ──────────────────────────────────────────────────────────
    output_filename = _build_output_filename(context, seq=args.get("seq", 1))
    output_path     = XML_OUT_DIR / output_filename

    generator = XMLGenerator(mapped_articles, context)
    generator.generate(output_path)

    log.info("[Output] %s", output_filename)

    # ── Audit ─────────────────────────────────────────────────────────────────
    if auditor:
        auditor.set_etl_counts(
            mapped=len(mapped_articles),
            written=len(mapped_articles),
            skipped=skipped,
            variants=0,
            xml_file=output_filename,
        )
        auditor.set_etl_status("success")

    return {
        "status":        "success",
        "output_files":  [str(output_path)],
        "article_count": len(mapped_articles),
    }


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Birkenstock OFS Mitra ETL v1.2")
    parser.add_argument("--linelist",   required=True)
    parser.add_argument("--mdd")
    parser.add_argument("--attributes")
    parser.add_argument("--mapping",    required=True)
    parser.add_argument("--country",    default="ID")
    parser.add_argument("--comp",       default="0888")
    parser.add_argument("--sbu",        default="SP")

    cli = parser.parse_args()

    result = run({
        "linelist_path":   cli.linelist,
        "mdd_path":        cli.mdd        or "",
        "attributes_path": cli.attributes or "",
        "mapping_path":    cli.mapping,
        "country_code":    cli.country,
        "comp_code":       cli.comp,
        "sbu":             cli.sbu,
    })

    print(f"Result: {result}")