"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Clarks Order Form Footwear v1.0     ║
║   Order Form Footwear (Footwear) → Stibo STEP XML                ║
║   (Generic Articles + Colour Variants, No Size)                  ║
╚══════════════════════════════════════════════════════════════════╝

Input file pattern:
    RAW Order Form Clarks Footwear SS27.xlsx

Sheet:
    "SEA Order form"  — header row at row 2 (0-indexed row 1)

Column layout (Clarks Order Form Footwear — brand mapping sheet "Clarks(Inline)"):

    Actual column roles (verified from file):
    SlotNumber          → Groups colour variants / KEY_InboundArticle base
                          (6-digit style number, e.g. 141502)
    ProductNumber       → AT_PrincipalStyleCode / AT_PrincipalBarcode / EAN barcode
                          (8-digit EAN, e.g. 26189642; unique per colour row)
    ColourID            → SAP Colour Slot ID (e.g. 312007; last 3 digits = color code)
    StyleName           → AT_PrincipalStyleDescription
    MediumColour        → AT_PrincipalColorName  (per variant)
    GenderName          → AT_PrincipalGenderDescription (1st sentence)
                          AT_PrincipalMerchandiseHierarchyL3
    Gender              → AT_Gender    (LOV — SAP label → MDD Gender LOV)
                          AT_BYGender  (LOV — BY label → MDD Gender LOV)
    Country of Origins  → AT_CountryOrigin  (LOV)
    BusinessUnitDesc    → AT_PrincipalMerchandiseHierarchyL1
    program             → AT_PrincipalMerchandiseHierarchyL4
                          AT_Style  (LOV — MD will provide mapping)
                          AT_Collection1  (LOV — raw program value, no MD Mapping lookup)
    UpperMaterial       → AT_Material  (LOV)
    BrandName           → AT_Brand  (LOV)
    DivisionName/BU     → context info

Article structure:
  One Generic per unique SlotNumber (style)
  One Variant per row (each ColourID / MediumColour = one colour variant)
  No size → size code defaults to "000"

KEY formulas:
    KEY_InboundArticle  =  BrandCode + PrincipalStyleCode (ProductNumber)
    KEY_InboundVariant  =  KEY_InboundArticle + SAP_ColorCode(3) + "000"

SAP Color Code derivation:
  • MDD "Color Code LOV" lookup: MediumColour → LOV ID
  • Fallback: last 3 digits of ColourID (e.g. 312007 → "007")

Care Instruction EN (AT_CareInstructionEN):
  Derived from `Prod Type` column value per brand mapping logic:
  • Sandals / Shoes / Canvas / Trainers / Boots → footwear care wording
  • All others                                  → not sent (attribute omitted)
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

INPUT_DIR       = BASE_DIR / "input" / "order_form_footwear"
MDD_DIR         = BASE_DIR / "input" / "mdd"
ATTR_DIR        = BASE_DIR / "input" / "attributes"
XML_OUT_DIR     = BASE_DIR / "output" / "xml"
LOG_DIR         = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME  = "Clarks"
BRAND_CODE  = "CKS"          # 3-letter SAP brand code for Clarks

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')

# ======================================================================
# GENDER SAP MAPPING  (GenderName → SAP Gender LOV)
# ======================================================================
_GENDER_TO_SAP: dict[str, str] = {    # Footwear
    "WOMENS": "Female",
    "MENS":   "Male",
    "BOYS":   "Male",
    "GIRLS":  "Female"
}

# ======================================================================
# BY GENDER MAPPING  (Gender → BY Gender LOV display value)
# ======================================================================
# Keyed on the FULL Gender value (upper-cased), not just the 1st word,
# so phrases like "Mens Accessories" / "Womens Bag" map explicitly.
_GENDER_TO_BYGENDER: dict[str, str] = {
    "WOMENS":             "Women",
    "MENS":               "Men",
    "BOYS":               "Boys",
    "GIRLS":              "Girls",
    "MENS ACCESSORIES":   "Men",
    "WOMENS ACCESSORIES": "Women",
    "WOMENS BAG":         "Women",
}

# ======================================================================
# CARE INSTRUCTION EN LOGIC  (from program value)
# ======================================================================
# Footwear product values (exact match)
_FOOTWEAR_PROGRAM_VALUES = {
    "sandals",
    "shoes",
    "canvas",
    "trainers",
    "boots"
}

_CARE_FOOTWEAR = (
    "Clean regularly with a soft, dry cloth to remove dust and dirt. "
    "For deeper cleaning, use suitable cleaner for the material. "
    "Avoid prolonged exposure to water, direct sunlight, and excessive heat. "
    "If shoes become wet, let them air dry naturally at room temperature. Do not use heaters or dryers. "
    "Apply a leather conditioner periodically to maintain softness and prevent cracking (leather shoes) "
    "Store in a cool, dry place and use shoe trees or stuff with paper to help maintain shape. "
    "Rotate footwear between wears to allow the material to breathe and recover."
)


def _care_instruction_en(prod_type_val: str) -> str:
    """Return Care Instruction EN text based on Prod Type column value.

    Rules:
      - Sandals / Shoes / Canvas / Trainers / Boots → footwear care instructions
      - Anything else → not sent (empty string)
    """
    if not prod_type_val:
        return ""

    p = str(prod_type_val).strip().lower()

    # Check for footwear products
    if p in _FOOTWEAR_PROGRAM_VALUES:
        return _CARE_FOOTWEAR

    return ""


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value — returns empty string for None/nan/empty."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT", "#VALUE!") else s


def _format_date(v) -> str:
    """Format date value to DD/MM/YYYY string, preserving original format if possible."""
    if v is None:
        return ""
    # If it's already a string, return as-is
    if isinstance(v, str):
        s = v.strip()
        return "" if s in ("None", "nan", "NaT", "#VALUE!") else s
    # If it's a datetime object, format as DD/MM/YYYY
    try:
        from datetime import datetime
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y")
    except:
        pass
    # Fallback to string conversion
    return _s(v)


def _sap_color_code(medium_colour: str, colour_id, color_lov: dict | None) -> str:
    """
    Derive a 3-character SAP colour code.

    Priority:
      1. MDD Color Code LOV lookup:  MediumColour → LOV ID
      2. Fallback: last 3 digits of ColourID (e.g. 26189642 → '642')
      3. Fallback: first 3 uppercase chars of MediumColour
    """
    # 1. MDD lookup
    if color_lov:
        key = _s(medium_colour).upper()
        if key in color_lov:
            return color_lov[key][:3]

    # 2. Last 3 digits of ColourID (e.g. 312007 → '007')
    cid = _s(colour_id)
    digits = re.sub(r"[^0-9]", "", cid)
    if len(digits) >= 3:
        return digits[-3:]

    # 3. First 3 chars of colour name
    alpha = re.sub(r"[^A-Za-z0-9]", "", _s(medium_colour)).upper()
    if alpha:
        return alpha[:3].ljust(3, "0")

    return "000"


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


# ======================================================================
# MDD LOADER
# ======================================================================

class MDDLoader:
    """
    Loads LOV tables from the MDD Excel needed for Clarks Footwear.
    Loads: Color Code LOV, Country Origin LOV, Material LOV, Gender LOV.
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
        self._load_brand_group_lov(wb)
        self._load_country_lov(wb)
        self._load_heel_height_lov(wb)
        self._load_collection1_lov(wb)

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
        """Brand LOV: col A = Value ID of LOV, col B = Values of LOV (display name).
        Lookup key: display name (col B) → LOV ID (col A).
        e.g. row: A='CKS'  B='CLARKS' → stored as {'CLARKS': 'CKS'}
        """
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
            lov_id  = str(row[0]).strip() if row[0] else ""   # Col A = Value ID of LOV
            display = str(row[1]).strip() if row[1] else ""   # Col B = Values of LOV
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandLOV"] = lov
        self.lovs["AT_Brand"] = lov
        log.info("[MDD] Brand LOV: %d entries loaded", len(lov))
        log.info("[MDD] Brand LOV sample (first 10): %s",
                 dict(list(lov.items())[:10]))

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

    def _load_brand_group_lov(self, wb):
        """Brand Group LOV: Col A = LOV ID, Col B = display value."""
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

    def _load_country_lov(self, wb):
        """Country LOV: Col A = LOV ID (e.g. 'ID'), Col B = display name (e.g. 'Indonesia').
        Stores both id->name and name->id for bidirectional lookup.
        """
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper()
                  and "ORIGIN" not in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Country LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        id_to_name: dict[str, str] = {}   # 'ID' -> 'Indonesia'
        name_to_id: dict[str, str] = {}   # 'INDONESIA' -> 'ID'
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                id_to_name[lov_id.upper()] = display
                name_to_id[display.upper()] = lov_id
        self.lovs["CountryLOV"]         = id_to_name   # id -> name
        self.lovs["CountryLOV_NameToId"] = name_to_id  # name -> id
        log.info("[MDD] Country LOV loaded — %d entries", len(id_to_name))

    def country_id_to_name(self, country_id: str) -> str:
        """Convert country ISO code to full name. e.g. 'ID' -> 'Indonesia'."""
        return self.lovs.get("CountryLOV", {}).get(country_id.strip().upper(), "")

    def _load_heel_height_lov(self, wb):
        """Heel Height LOV: Col A = LOV ID (e.g. 'F'), Col B = display name (e.g. 'Flat').
        Lookup key: display name (col B) → LOV ID (col A).
        """
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
            lov_id  = str(row[0]).strip() if row[0] else ""   # Col A = LOV ID
            display = str(row[1]).strip() if row[1] else ""   # Col B = display name
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["HeelHeightLOV"] = lov_map
        self.lovs["AT_HeelHeight"] = lov_map
        log.info("[MDD] Heel Height LOV loaded — %d entries", len(lov_map))

    def _load_collection1_lov(self, wb):
        """Collection1 LOV: from 'Clarks MD Mapping' sheet.
        Col H (index 7) = Program value (lookup key)
        Col I (index 8) = Collection1 LOV ID
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "CLARKS" in s.upper() and "MD" in s.upper() and "MAPPING" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Clarks MD Mapping sheet not found — Collection1 LOV skipped")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        # Find header row
        hdr_idx = 0
        for i, row in enumerate(rows[:20]):
            if row and any(isinstance(v, str) and "PROGRAM" in str(v).upper() for v in row):
                hdr_idx = i
                break
        lov_map: dict[str, str] = {}  # program value (upper) -> Collection1 LOV ID
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 9:
                continue
            program_val     = str(row[7]).strip() if row[7] else ""  # Col H (index 7)
            collection1_lov = str(row[8]).strip() if row[8] else ""  # Col I (index 8)
            if program_val and collection1_lov:
                lov_map[program_val.upper()] = collection1_lov
        self.lovs["Collection1LOV"] = lov_map
        self.lovs["AT_Collection1"] = lov_map
        log.info("[MDD] Collection1 LOV loaded from '%s' — %d entries", sheet_name, len(lov_map))

    def lookup(self, lov_key: str, display_value: str) -> str:
        """Look up a LOV ID. Returns display_value unchanged if not found."""
        lov = self.lovs.get(lov_key, {})
        if not lov:
            return display_value
        return lov.get(display_value.strip().upper(), display_value)


# ======================================================================
# RNA LOADER
# ======================================================================

class RNALoader:
    """
    Loads RNA (Source Mapping) data to resolve Brand Type, Brand Category,
    Brand Group and Brand Status.
    Reads from "Source Mapping related RNA" sheet in the brand mapping file.
    Also loads "Clarks MD Mapping" sheet for heel height value-to-LOV mapping.
    """
    RNA_SHEET_KEYWORDS = ["SOURCE MAPPING", "RNA", "BRAND MAPPING"]
    CLARKS_MD_MAPPING_KEYWORDS = ["CLARKS MD MAPPING", "MD MAPPING"]

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
        self.heel_height_mapping: dict[str, str] = {}  # numeric value -> LOV name (e.g. '5' -> 'Flat')
        self.style_mapping: dict[str, str] = {}  # program value -> Style LOV value (e.g. 'Umbrella' -> 'Umbrella')
        self._load()
        self._load_heel_height_mapping()
        self._load_style_mapping()

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
        c_bgroup  = _find("BRANDGROUP", "BRAND GROUP", "AT_BRANDGROUP")
        c_bstatus = _find("BRAND STATUS", "BRANDSTATUS", "AT_BRANDSTATUS")

        missing = [nm for nm, idx in [
            ("Country", c_country), ("CompCode", c_comp), ("SBU", c_sbu),
            ("BrandCode", c_bcode), ("BrandType", c_btype),
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
        return self.lookup.get(key, {"brand_type": "", "brand_group": "", "brand_status": ""})

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
        return {"brand_type": "", "brand_group": "", "brand_status": ""}

    def get_by_brand_only(self, brand_code: str) -> dict:
        """Pass 4: match on brand code only — ignores country, comp_code, sbu."""
        b = (brand_code or "").strip().upper()
        for (_, _, _, kb), val in self.lookup.items():
            if kb == b:
                return val
        return {"brand_type": "", "brand_group": "", "brand_status": ""}

    def _load_heel_height_mapping(self):
        """Load Heel Height mapping from 'Clarks MD Mapping' sheet.
        Col A = numeric heel height value (e.g. 5, 10, 15...)
        Col B = LOV name (e.g. 'Flat', 'Low', 'Medium', 'High')
        """
        log.info("[RNA] Loading Heel Height mapping from: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.warning("[RNA] Failed to open workbook for heel height mapping: %s", e)
            return

        log.info("[RNA] Available sheets: %s", wb.sheetnames)
        
        # Try multiple search strategies to find the sheet
        sheet_name = None
        
        # Strategy 1: Exact match (case-insensitive, normalized whitespace)
        for s in wb.sheetnames:
            normalized = re.sub(r'\s+', ' ', s.strip()).upper()
            if normalized == "CLARKS MD MAPPING":
                sheet_name = s
                log.info("[RNA] Found exact match sheet: '%s'", s)
                break
        
        # Strategy 2: Contains "MD" and "MAPPING"
        if not sheet_name:
            for s in wb.sheetnames:
                s_upper = s.upper()
                if "MD" in s_upper and "MAPPING" in s_upper:
                    sheet_name = s
                    log.info("[RNA] Found MD+MAPPING match: '%s'", s)
                    break
        
        # Strategy 3: Contains "CLARKS" and "MAPPING"
        if not sheet_name:
            for s in wb.sheetnames:
                s_upper = s.upper()
                if "CLARKS" in s_upper and "MAPPING" in s_upper:
                    sheet_name = s
                    log.info("[RNA] Found CLARKS+MAPPING match: '%s'", s)
                    break
        
        if not sheet_name:
            log.warning("[RNA] 'Clarks MD Mapping' sheet not found in %s. Available sheets: %s", 
                       self.path.name, wb.sheetnames)
            wb.close()
            return

        log.info("[RNA] Using sheet for heel height: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))

        # Find the header row containing "Heel Height"
        hdr_idx = None
        for i, r in enumerate(rows):
            if r and any(isinstance(v, str) and "HEEL" in v.upper() and "HEIGHT" in v.upper() for v in r if v):
                hdr_idx = i
                break

        if hdr_idx is None:
            log.warning("[RNA] Cannot find 'Heel Height' header in sheet '%s'", sheet_name)
            wb.close()
            return

        # Read value mappings starting from the row after header
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 2:
                continue
            val_raw = row[0]
            lov_name = row[1]
            if val_raw is not None and lov_name:
                # Normalize the numeric value (handle floats like 5.0 -> '5')
                val_str = str(val_raw).strip()
                if val_str.endswith(".0"):
                    val_str = val_str[:-2]
                lov_name_str = str(lov_name).strip()
                if val_str and lov_name_str:
                    self.heel_height_mapping[val_str] = lov_name_str

        wb.close()
        log.info("[RNA] Heel Height mapping loaded — %d entries", len(self.heel_height_mapping))

    def get_heel_height_lov_name(self, heel_height_value: str) -> str:
        """Get LOV name (Flat/Low/Medium/High) for a numeric heel height value."""
        val = str(heel_height_value).strip()
        if val.endswith(".0"):
            val = val[:-2]
        return self.heel_height_mapping.get(val, "")

    def _load_style_mapping(self):
        """Load Style mapping from 'Clarks MD Mapping' sheet.
        Col H = program value (e.g. 'Umbrella', 'Syn Cosmetic Bag', etc.)
        Col I = Style LOV value (e.g. 'Umbrella', 'Pouch', 'Backpack', etc.)
        """
        log.info("[RNA] Loading Style mapping from: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as e:
            log.warning("[RNA] Failed to open workbook for style mapping: %s", e)
            return

        log.info("[RNA] Available sheets: %s", wb.sheetnames)
        
        # Try multiple search strategies to find the sheet
        sheet_name = None
        
        # Strategy 1: Exact match (case-insensitive, normalized whitespace)
        for s in wb.sheetnames:
            normalized = re.sub(r'\s+', ' ', s.strip()).upper()
            if normalized == "CLARKS MD MAPPING":
                sheet_name = s
                log.info("[RNA] Found exact match sheet: '%s'", s)
                break
        
        # Strategy 2: Contains "MD" and "MAPPING"
        if not sheet_name:
            for s in wb.sheetnames:
                s_upper = s.upper()
                if "MD" in s_upper and "MAPPING" in s_upper:
                    sheet_name = s
                    log.info("[RNA] Found MD+MAPPING match: '%s'", s)
                    break
        
        # Strategy 3: Contains "CLARKS" and "MAPPING"
        if not sheet_name:
            for s in wb.sheetnames:
                s_upper = s.upper()
                if "CLARKS" in s_upper and "MAPPING" in s_upper:
                    sheet_name = s
                    log.info("[RNA] Found CLARKS+MAPPING match: '%s'", s)
                    break
        
        if not sheet_name:
            log.warning("[RNA] 'Clarks MD Mapping' sheet not found in %s. Available sheets: %s", 
                       self.path.name, wb.sheetnames)
            wb.close()
            return

        log.info("[RNA] Using sheet for style mapping: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))

        # Find the header row containing "program" and "Style"
        hdr_idx = None
        for i, r in enumerate(rows):
            if r and len(r) >= 9:  # Need at least columns up to I (index 8)
                # Check if this row has "program" in col H (index 7) and "Style" in col I (index 8)
                col_h = str(r[7]).strip().lower() if r[7] else ""
                col_i = str(r[8]).strip().lower() if r[8] else ""
                if "program" in col_h and "style" in col_i:
                    hdr_idx = i
                    break

        if hdr_idx is None:
            log.warning("[RNA] Cannot find 'program' and 'Style' headers in sheet '%s'", sheet_name)
            wb.close()
            return

        # Read value mappings starting from the row after header
        # Col H = index 7 (program), Col I = index 8 (Style)
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 9:
                continue
            program_val = row[7]  # Col H
            style_val = row[8]    # Col I
            if program_val is not None and style_val:
                program_str = str(program_val).strip()
                style_str = str(style_val).strip()
                if program_str and style_str:
                    # Store with uppercase key for case-insensitive lookup
                    self.style_mapping[program_str.upper()] = style_str

        wb.close()
        log.info("[RNA] Style mapping loaded — %d entries", len(self.style_mapping))

    def get_style_lov_value(self, program_value: str) -> str:
        """Get Style LOV value for a program value."""
        val = str(program_value).strip().upper()
        return self.style_mapping.get(val, "")


# ======================================================================
# INPUT FILE LOADER
# ======================================================================

class ClarksOrderFormFWLoader:
    """
    Loads Clarks Order Form Footwear data from the 'Order Form' sheet.

    Header row: row 2 (index 1 in 0-based)
    Data rows : row 3 onwards

    Key columns (confirmed from file inspection):
        SlotNumber      (col F / index 5)  — principal style code (groups variants)
        ProductNumber   (col H / index 7)  — EAN barcode (unique per colour)
        ColourID        (col G / index 6)  — SAP colour slot ID (last 3 = color code)
        StyleName       (col P / index 15) — style description
        MediumColour    (col Q / index 16) — colour name (per variant)
        GenderName      (col L / index 11) — gender description
        Country of Origins (col T / index 19) — country of origin
        BusinessUnitDesc(col U / index 20) — merchandise hierarchy L1
        program         (col O / index 14) — L4 / style / collection1
        UpperMaterial   (col X / index 23) — material
        BrandName       (col I / index 8)  — brand name
        BU              (col K / index 10) — business unit
        SeasonCode      (col C / index 2)  — season
    """

    # Signal words that identify the header row
    HEADER_SIGNALS = {"slotnumber", "stylename", "mediumcolour", "sourcecountry", "productnumber"}

    def __init__(self, path: Path):
        self.path           = path
        self.df             = pd.DataFrame()
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Clarks-OrderFormFW] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[Clarks-OrderFormFW] Available sheets: %s", wb.sheetnames)

        # Only use "SEA Order Form" sheet for footwear (case-insensitive, strip whitespace)
        target = None
        for sn in wb.sheetnames:
            if sn.strip().lower() == "sea order form":
                target = sn
                break
        if target is None:
            wb.close()
            raise ValueError(f"[Clarks-OrderFormFW] Required sheet 'SEA Order form' not found. Available sheets: {wb.sheetnames}")
        log.info("[Clarks-OrderFormFW] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.warning("[Clarks-OrderFormFW] Sheet '%s' is empty", target)
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

        log.info("[Clarks-OrderFormFW] Header at row %d (0-indexed)", hdr_idx)
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

        # ── Filter: keep rows with a valid Product Number ──────────
        prod_col = self._find_col(header, ["Product Number", "ProductNumber", "productnumber"])
        if prod_col:
            df = df[
                df[prod_col].notna()
                & (~df[prod_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        self.df = df.reset_index(drop=True)
        log.info(
            "[Clarks-OrderFormFW] Loaded %d data rows | headers: %s",
            len(self.df), header[:25],
        )

    @staticmethod
    def _find_col(headers: list[str], candidates: list[str]) -> str | None:
        """Case-insensitive column name search, with whitespace-normalised fallback."""
        # Pass 1: exact match
        for c in candidates:
            if c in headers:
                return c
        # Pass 2: case-insensitive exact
        for c in candidates:
            for h in headers:
                if h.lower() == c.lower():
                    return h
        # Pass 3: normalise all whitespace (newlines, multiple spaces → single space)
        def _norm(s: str) -> str:
            return re.sub(r'\s+', ' ', s).strip().lower()
        for c in candidates:
            nc = _norm(c)
            for h in headers:
                if _norm(h) == nc:
                    return h
        return None

    def find_col(self, *candidates: str) -> str | None:
        return self._find_col(self.headers, list(candidates))


# ======================================================================
# GROUPER — builds Generic + Variant dict from flat rows
# ======================================================================

def group_rows(
    raw_rows: list[dict],
    loader: ClarksOrderFormFWLoader,
    brand_code: str,
    mdd: MDDLoader | None,
    rna: RNALoader | None = None,
) -> dict[str, dict]:
    """
    Group flat rows into Generic articles keyed by KEY_InboundArticle.

    KEY_InboundArticle  =  brand_code + PrincipalStyleCode (ProductNumber/EAN)
    KEY_InboundVariant  =  KEY_InboundArticle + sap_color(3) + "000"
    """
    # SlotID = principal style code (groups colour variants into one generic)
    col_slot_number  = loader.find_col("SlotID", "SlotNumber", "Slot Number", "Slot ID")
    # Product Number = EAN barcode (one per colour row)
    col_ean          = loader.find_col("Product Number", "ProductNumber")
    col_style_name   = loader.find_col("Stylename", "StyleName", "Style Name", "Style_Name")
    col_colour       = loader.find_col("Colour Name", "ColourName")
    col_colour_id    = loader.find_col("Colour ID", "ColourID", "ColourId")
    col_gender       = loader.find_col("Gender")
    col_country      = loader.find_col("Country of Origins", "Country Of Origins", "Country of Origin", "CountryOfOrigin", "SourceCountry", "Source Country")
    col_bu_desc      = loader.find_col("BusinessUnitDesc", "BusinessUnit Desc")
    col_program      = loader.find_col("program", "Program")
    col_material     = loader.find_col("UpperMaterial", "Upper Material")
    col_brand_name   = loader.find_col("BrandName", "Brand Name", "Brand_Name")
    col_bu           = loader.find_col("BU")
    col_season       = loader.find_col("SeasonCode", "Season Code")
    col_product_type = loader.find_col("ProductTypeDesc", "Product Type Desc")
    col_size_range   = loader.find_col("Size Range", "SizeRange", "Size_Range")
    col_unit_price   = loader.find_col("ID UNIT PRICE (USD)", "ID UNIT PRICE USD", "ID_UNIT_PRICE_USD", "ID UNIT PRICE")
    col_collection   = loader.find_col("Collection")
    col_tube_gender  = loader.find_col("Tube /\nGenderName", "Tube /  GenderName", "Tube / GenderName", "Tube/GenderName", "Tube/  GenderName", "Tube", "TubeGenderName")
    col_prod_type    = loader.find_col("Prod Type", "ProdType", "Prod_Type")
    col_heel_height  = loader.find_col("Heel Height", "HeelHeight", "Heel_Height")
    col_launch_month = loader.find_col("MYSG Launch Month", "MYSGLaunchMonth", "MYSG_Launch_Month", "Launch Month")
    col_technology   = loader.find_col("Technology")

    log.info(
        "[Grouper] Columns → slot=%s, ean=%s, style=%s, colour=%s, colourID=%s, "
        "gender=%s, country=%s, bu_desc=%s, program=%s, material=%s",
        col_slot_number, col_ean, col_style_name, col_colour, col_colour_id,
        col_gender, col_country, col_bu_desc, col_program, col_material,
    )

    color_lov = mdd.lovs.get("ColorCodeLOV") if mdd else None

    result: dict[str, dict] = {}

    for row in raw_rows:
        # ── EAN barcode (ProductNumber column) ───────────────────
        ean = _s(row.get(col_ean, "")) if col_ean else ""
        if not ean:
            continue

        product_number = ean  # Product Number for AT_PrincipalStyleCode
        slot_number = _s(row.get(col_slot_number, "")) if col_slot_number else ""

        # ── Generic key: BrandCode + PrincipalStyleCode ──────────
        generic_key = f"{brand_code}{product_number}"

        # ── Derive SAP color code (3 chars) ──────────────────────
        medium_colour = _s(row.get(col_colour, ""))   if col_colour   else ""
        colour_id     = row.get(col_colour_id, "")     if col_colour_id else ""
        sap_color     = _sap_color_code(medium_colour, colour_id, color_lov)

        # ── Variant key: generic + color(3) + "000" (no size) ─────
        variant_key   = f"{generic_key}{sap_color}000"

        # ── Generic-level attributes (captured on first encounter) ─
        if generic_key not in result:
            style_name   = _s(row.get(col_style_name, ""))   if col_style_name   else ""
            gender_raw   = _s(row.get(col_gender, ""))        if col_gender       else ""
            bu_desc      = _s(row.get(col_bu_desc, ""))       if col_bu_desc      else ""
            program      = _s(row.get(col_program, ""))       if col_program      else ""
            material_raw = _s(row.get(col_material, ""))      if col_material     else ""
            brand_name   = _s(row.get(col_brand_name, ""))    if col_brand_name   else ""
            bu           = _s(row.get(col_bu, ""))            if col_bu           else ""
            season_code  = _s(row.get(col_season, ""))        if col_season       else ""
            product_type = _s(row.get(col_product_type, "")) if col_product_type else ""
            size_range   = _s(row.get(col_size_range, ""))    if col_size_range   else ""
            unit_price   = _s(row.get(col_unit_price, ""))    if col_unit_price   else ""
            collection   = _s(row.get(col_collection, ""))    if col_collection   else ""
            tube_gender  = _s(row.get(col_tube_gender, ""))   if col_tube_gender  else ""
            prod_type    = _s(row.get(col_prod_type, ""))     if col_prod_type    else ""
            heel_height_raw = _s(row.get(col_heel_height, "")) if col_heel_height else ""
            # Format launch_month as DD-MON-YYYY (e.g. "04-jan-2027")
            # Input can be a datetime object or a string in M/D/YYYY format
            _raw_launch = row.get(col_launch_month, "") if col_launch_month else ""
            if _raw_launch and hasattr(_raw_launch, "month"):
                # datetime/date object
                launch_month = _raw_launch.strftime("%d-%b-%Y").lower()
            else:
                _str_launch = _s(_raw_launch)
                # Try to parse string in M/D/YYYY format (e.g. "1/4/2027")
                _m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', _str_launch)
                if _m:
                    _dt = datetime(int(_m.group(3)), int(_m.group(1)), int(_m.group(2)))
                    launch_month = _dt.strftime("%d-%b-%Y").lower()
                else:
                    launch_month = _str_launch
            technology   = _s(row.get(col_technology, ""))       if col_technology   else ""

            # Gender: 1st sentence of GenderName
            gender_desc = _first_sentence(gender_raw)

            # Gender SAP mapping
            # No fallback: if the raw value isn't mapped (or the mapped label
            # isn't found in the MDD LOV), leave it blank so AT_Gender is
            # simply not sent rather than defaulting to a guessed value.
            gender_key  = gender_raw.strip().upper()
            sap_gender  = _GENDER_TO_SAP.get(gender_key, "")
            if sap_gender and mdd:
                gender_lov = mdd.lovs.get("GenderLOV", {})
                sap_gender = gender_lov.get(sap_gender.upper(), "")

            # BY Gender mapping (two-step: dictionary -> MDD LOV), same
            # pattern as AT_Gender above but keyed on the FULL Gender value
            # ("Mens Accessories" -> "Men") rather than a single word.
            # No fallback: unmapped values leave AT_BYGender blank.
            bygender_key = gender_raw.strip().upper()
            by_gender    = _GENDER_TO_BYGENDER.get(bygender_key, "")
            if by_gender and mdd:
                gender_lov = mdd.lovs.get("GenderLOV", {})
                by_gender  = gender_lov.get(by_gender.upper(), "")

            # Country of Origin LOV lookup
            country_raw = ""
            # (country is variant-level in this file but use from first row for generic)
            country_raw_val = _s(row.get(col_country, "")) if col_country else ""
            if mdd:
                co_lov   = mdd.lovs.get("CountryOriginLOV", {})
                country_raw = co_lov.get(country_raw_val.upper(), country_raw_val)
            else:
                country_raw = country_raw_val

            # Material LOV lookup
            if mdd:
                mat_lov  = mdd.lovs.get("MaterialLOV", {})
                material_id = mat_lov.get(material_raw.upper(), material_raw)
            else:
                material_id = material_raw

            # Heel Height LOV lookup (2-step: value -> LOV name -> LOV ID)
            heel_height_lov_id = ""
            if heel_height_raw:
                # Step 1: Get LOV name from brand mapping (Clarks MD Mapping sheet)
                heel_height_lov_name = ""
                if rna:
                    heel_height_lov_name = rna.get_heel_height_lov_name(heel_height_raw)
                # Step 2: Get LOV ID from MDD (Heel Height LOV sheet)
                if heel_height_lov_name and mdd:
                    hh_lov = mdd.lovs.get("HeelHeightLOV", {})
                    heel_height_lov_id = hh_lov.get(heel_height_lov_name.upper(), "")
                log.info(
                    "[Grouper] Heel height: raw='%s' -> lov_name='%s' -> lov_id='%s'",
                    heel_height_raw, heel_height_lov_name, heel_height_lov_id,
                )

            # Care instruction from Prod Type
            care_en = _care_instruction_en(prod_type)

            # Style LOV value from program (via RNA mapping)
            style_value = ""
            if program and rna:
                style_value = rna.get_style_lov_value(program)
                log.debug(
                    "[Grouper] Style: program='%s' -> style_value='%s'",
                    program, style_value,
                )

            # Brand LOV
            brand_lov_id = brand_name
            if mdd:
                b_lov    = mdd.lovs.get("BrandLOV", {})
                brand_lov_id = b_lov.get(brand_name.upper(), brand_name)

            # Collection1: raw "program" value, sent as the LOV value
            # (no "Clarks MD Mapping" lookup)
            collection1_lov_id = program

            result[generic_key] = {
                # Key identifiers
                "generic_key":      generic_key,
                "product_number":   slot_number,   # SlotNumber is the style/product number
                "brand_code":       brand_code,

                # Mapped attributes — Generic level
                "AT_InboundGenericCode":           generic_key,
                "AT_PrincipalStyleCode":            product_number,
                "AT_PrincipalStyleDescription":     style_name,
                "AT_PrincipalColorName":            medium_colour,
                "AT_PrincipalGenderDescription":    gender_desc,
                "AT_PrincipalGenderCode":           "",   # N/A per mapping
                "AT_PrincipalAgeDescription":       "",   # N/A per mapping
                "AT_PrincipalAgeCode":              "",   # N/A per mapping
                "AT_PrincipalSize":                 size_range,
                "AT_PrincipalMerchandiseHierarchyL1": bu,
                "AT_PrincipalMerchandiseHierarchyL2": collection,
                "AT_PrincipalMerchandiseHierarchyL3": tube_gender,
                "AT_PrincipalMerchandiseHierarchyL4": prod_type,
                # "AT_Collection1":                   program,
                "AT_Brand":                         brand_lov_id,
                "AT_CountryOrigin":                 country_raw,
                "AT_Material":                      material_id,
                "AT_CareInstructionEN":             care_en,
                "AT_SAPGender":                     sap_gender,
                "AT_BYGender":                      by_gender,
                "AT_SAPAge":                        "AD",
                "AT_PricingDistributionChannel":    "01",
                "AT_ArticleStatus":                 "A",
                "AT_SeasonCode":                    season_code,
                "AT_ProductType":                   product_type,
                "AT_BU":                            bu,
                "AT_FOB":                           unit_price,
                "AT_HeelHeight":                    heel_height_lov_id,
                "AT_LaunchingDate":                 launch_month,
                "AT_Style":                         style_value,
                "AT_Collection1":                   collection1_lov_id,
                "AT_TechnologyUsed":                technology,

                # Variants dict: sap_color → variant data
                "variants": {},
            }

        # ── Variant-level attributes ─────────────────────────────
        result[generic_key]["variants"][sap_color] = {
            "variant_key":           variant_key,
            "sap_color":             sap_color,
            "medium_colour":         medium_colour,
            "colour_id":             _s(colour_id),
            "ean":                   ean,
            # AT_PrincipalColorName is per colour variant
            "AT_PrincipalColorName": medium_colour,
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
    Build the <Classifications> block (mirrors Crocs/Lotto pattern):
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
    brand_group: str = "",
    mdd: "MDDLoader | None" = None,
) -> str:
    """Build XML for one Generic article with its Colour Variants."""
    if not generic.get("product_number"):
        return ""

    # ── Generic Article ──────────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_F-TempSubCat")  # Footwear

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_key"]

    # Name
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        generic["AT_PrincipalStyleDescription"] or generic["generic_key"]
    )

    # ── Classification References ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Generic Values ───────────────────────────────────────────
    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # Principal attributes (from row data)
    _val_text(gv, "AT_InboundGenericCode",          generic["AT_InboundGenericCode"])
    _val_text(gv, "AT_PrincipalStyleCode",             generic["AT_PrincipalStyleCode"])
    _val_text(gv, "AT_PrincipalStyleDescription",      generic["AT_PrincipalStyleDescription"])
    _val_text(gv, "AT_PrincipalColorName",             generic["AT_PrincipalColorName"])
    _val_text(gv, "AT_PrincipalGenderDescription",     generic["AT_PrincipalGenderDescription"])
    _val_text(gv, "AT_PrincipalSize",                  generic["AT_PrincipalSize"])
    _val_text(gv, "AT_PrincipalMerchandiseHierarchyL1", generic["AT_PrincipalMerchandiseHierarchyL1"])
    _val_text(gv, "AT_PrincipalMerchandiseHierarchyL2", generic["AT_PrincipalMerchandiseHierarchyL2"])
    _val_text(gv, "AT_PrincipalMerchandiseHierarchyL3", generic["AT_PrincipalMerchandiseHierarchyL3"])
    _val_text(gv, "AT_PrincipalMerchandiseHierarchyL4", generic["AT_PrincipalMerchandiseHierarchyL4"])
    if generic.get("AT_Collection1"):
        # LOV-backed attr, but sent as the LOV *value* (element text), not an ID
        _val_text(gv, "AT_Collection1",                    generic["AT_Collection1"])
    if generic.get("AT_CareInstructionEN"):
        _val_text(gv, "AT_CareInstructionEN", generic["AT_CareInstructionEN"])

    # LOV attributes
    # AT_Brand: extract brand name from input filename → look up LOV ID in MDD Brand LOV
    _brand_lov = mdd.lovs.get("AT_Brand", {}) if mdd else {}
    _brand_lov_id = _brand_lov.get(brand.upper(), brand_code)
    _val_lov(gv, "AT_Brand", _brand_lov_id)
    if generic["AT_CountryOrigin"]:
        _val_lov(gv, "AT_CountryOrigin", generic["AT_CountryOrigin"])
    if generic["AT_Material"]:
        _val_lov(gv, "AT_Material",      generic["AT_Material"])
    if generic.get("AT_SAPGender"):
        _val_lov(gv, "AT_Gender",        generic["AT_SAPGender"])
    if generic.get("AT_BYGender"):
        _val_lov(gv, "AT_BYGender",      generic["AT_BYGender"])
    if generic.get("AT_SAPAge"):
        _val_lov(gv, "AT_SAPAge",           generic["AT_SAPAge"])
    if generic.get("AT_PricingDistributionChannel"):
        _val_lov(gv, "AT_PricingDistributionChannel", generic["AT_PricingDistributionChannel"])
    if generic.get("AT_ArticleStatus"):
        _val_lov(gv, "AT_ArticleStatus",    generic["AT_ArticleStatus"])
    if generic.get("AT_HeelHeight"):
        _val_lov(gv, "AT_HeelHeight",       generic["AT_HeelHeight"])
    _val_lov(gv, "AT_BYArticleType", "Inline")
    _val_lov(gv, "AT_ImagesSource", "DI")
    if generic.get("AT_LaunchingDate"):
        _val_text(gv, "AT_LaunchingDate",   generic["AT_LaunchingDate"])
    if generic.get("AT_Style"):
        _val_text(gv, "AT_Style",           generic["AT_Style"])
    if generic.get("AT_TechnologyUsed"):
        _val_text(gv, "AT_TechnologyUsed",  generic["AT_TechnologyUsed"])

    # ── System-level attributes (brand/season/country/comp/sbu) ──
    log.info(
        "[XML] Brand attrs — brand_type='%s'  brand_status='%s'  brand_group='%s'",
        brand_type, brand_status, brand_group,
    )

    # ── Brand Type ───────────────────────────────────────────────
    # brand_type comes from RNA BRANDTYPE_DETAIL column
    # Look up its LOV ID from the Brand Type LOV sheet
    if brand_type:
        bt_lov_id = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type.upper(), "")
        if bt_lov_id:
            _val_lov(gv, "AT_BrandType", bt_lov_id)
        else:
            _val_text(gv, "AT_BrandType", brand_type)

    # ── Brand Status ─────────────────────────────────────────────
    # brand_status comes from RNA BRAND STATUS column
    # Look up its LOV ID from the Brand Status LOV sheet
    if brand_status:
        bs_lov_id = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status.upper(), "")
        if bs_lov_id:
            _val_lov(gv, "AT_BrandStatus", bs_lov_id)
        else:
            _val_text(gv, "AT_BrandStatus", brand_status)

    # ── Brand Group ──────────────────────────────────────────────
    # brand_group comes from RNA BRANDGROUP column
    # Look up its LOV ID from the Brand Group LOV sheet (Col A = LOV ID, Col B = display value)
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
    _SEASON_LOV = {
        "SS": "Spring-Summer", "FW": "Fall-Winter",
        "AW": "Autumn-Winter", "HO": "Holiday",
        "SP": "Spring",        "SM": "Summer",
    }
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

    # ── FOB Currency (default USD) ───────────────────────────────
    _val_lov(gv, "AT_FOBCurrency", "USD")

    # ── FOB (from "ID UNIT PRICE (USD)" column) ──────────────────
    if generic.get("AT_FOB"):
        _val_text(gv, "AT_FOB", generic["AT_FOB"])

    # ── SAP Product Flag (default A) ─────────────────────────────
    _val_lov(gv, "AT_SAPProductFlag", "A")

    # ── Material Type (default ZINA) ─────────────────────────────
    _val_lov(gv, "AT_MaterialType", "ZINA")

    # ── SAP Article Category (default 01) ────────────────────────
    _val_lov(gv, "AT_SAPArticleCategory", "01")

    # ── UOM (default EA) ──────────────────────────────────────────
    _val_lov(gv, "AT_UOM", "EA")

    # ── BY Age (default ADULT) ────────────────────────────────────
    _val_lov(gv, "AT_BYAge", "ADULT")

    # ── Country Size (default UK) ─────────────────────────────────
    _val_lov(gv, "AT_CountrySize", "UK")

    # ── Sports Category EN (default 16) ───────────────────────────
    _val_lov(gv, "AT_SportsCategoryEN", "16")

    # ── Variants disabled: Generic articles only ─────────────────
    # (No variant sub-products generated)

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """
    Main entry point for Clarks Order Form Footwear ETL.

    Called by clarks/lambda_function.py after files are downloaded.
    """
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SS27")
    country_code = getattr(args, "country_code", "")

    log.info(
        "[Clarks-OrderFormFW] Starting: brand=%s  code=%s  season=%s",
        brand, brand_code, season,
    )

    # ── Load MDD ─────────────────────────────────────────────────
    mdd      = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Clarks-OrderFormFW] No MDD file found in %s — LOV lookups disabled", MDD_DIR)

    # ── Validate brand mapping file + load RNA ───────────────────
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    rna = None
    if mapping_file:
        log.info("[Clarks-OrderFormFW] Brand mapping file: %s", mapping_file.name)
        rna = RNALoader(mapping_file)
    else:
        log.warning("[Clarks-OrderFormFW] Brand mapping file not found in %s", ATTR_DIR)

    # ── Load input Order Form Footwear file ──────────────────────────
    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        raise FileNotFoundError(f"No Order Form Footwear file found in {INPUT_DIR}")

    # ── Extract brand name from filename ──────────────────────────
    # Pattern: 0888-FF-CLARKS-RAW Order Form...-SP2029-ID-1.xlsx
    #   parts[0]  = comp code   (e.g. '0888')
    #   parts[1]  = SBU         (e.g. 'FF')
    #   parts[2]  = brand name  (e.g. 'CLARKS')
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
            brand_from_file = brand_token.title()  # e.g. 'Clarks'
            log.info(
                "[Clarks-OrderFormFW] Brand extracted from filename: '%s' → '%s'",
                brand_token, brand_from_file,
            )
    if len(parts) >= 1 and parts[0].strip():
        raw_comp = parts[0].strip().lstrip("0") or "0"  # '0888' → '888'
        filename_comp_code = raw_comp
        log.info("[Clarks-OrderFormFW] Comp code from filename: '%s' → '%s'", parts[0].strip(), filename_comp_code)
    if len(parts) >= 2 and parts[1].strip():
        filename_sbu = parts[1].strip().upper()   # e.g. 'FF'
        log.info("[Clarks-OrderFormFW] SBU from filename: '%s'", filename_sbu)
    if len(parts) >= 2:
        # country id is the 2nd-to-last dash token (e.g. '...SP2029-ID-1')
        candidate = parts[-2].strip().upper()
        if re.match(r'^[A-Z]{2}$', candidate):   # 2-letter ISO code
            filename_country_id = candidate
            log.info("[Clarks-OrderFormFW] Country ID from filename: '%s'", filename_country_id)
    
    # ── Look up Brand LOV ID from MDD ────────────────────────────
    brand_lov_id = brand_code  # fallback to original brand_code if not found
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            brand_key    = brand_from_file.upper()   # e.g. "CLARKS"
            looked_up    = brand_lov.get(brand_key)  # e.g. "CKS"
            brand_lov_id = looked_up if looked_up else brand_code
            log.info(
                "[Clarks-OrderFormFW] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_key, brand_lov_id,
                "(FOUND)" if looked_up else f"(NOT FOUND — fallback to '{brand_code}')",
            )
            if not looked_up:
                log.warning(
                    "[Clarks-OrderFormFW] '%s' not in Brand LOV. "
                    "Available keys: %s",
                    brand_key, list(brand_lov.keys()),
                )
        else:
            log.warning("[Clarks-OrderFormFW] Brand LOV not found in MDD — using fallback brand_code='%s'", brand_code)
    else:
        log.warning("[Clarks-OrderFormFW] No MDD loaded — brand_lov_id defaulting to '%s'", brand_code)
    
    # ── Resolve country name from country ID via MDD ─────────────
    rna_country = ""
    if filename_country_id and mdd:
        rna_country = mdd.country_id_to_name(filename_country_id)
        log.info(
            "[Clarks-OrderFormFW] Country: ID='%s' → name='%s' (for RNA lookup)",
            filename_country_id, rna_country,
        )
        if not rna_country:
            log.warning("[Clarks-OrderFormFW] Country ID '%s' not found in Country LOV", filename_country_id)
    if not rna_country:
        rna_country = country_code   # fallback to args

    # ── RNA lookup: brand_type, brand_status, brand_group ─────
    brand_type = brand_status = brand_group = ""
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
        if not rna_result.get("brand_type"):
            rna_result = rna.get_fuzzy(rna_country, filename_sbu, brand_lov_id)
            log.info("[RNA] Pass 2 result: %s", rna_result)
        # Pass 3: fuzzy with original brand_code fallback
        if not rna_result.get("brand_type"):
            rna_result = rna.get_fuzzy(rna_country, filename_sbu, brand_code)
            log.info("[RNA] Pass 3 result: %s", rna_result)
        # Pass 4: brand code only — last resort
        if not rna_result.get("brand_type"):
            rna_result = rna.get_by_brand_only(brand_lov_id) or rna.get_by_brand_only(brand_code)
            log.info("[RNA] Pass 4 (brand-only) result: %s", rna_result)
        brand_type     = rna_result.get("brand_type",     "")
        brand_status   = rna_result.get("brand_status",   "")
        brand_group    = rna_result.get("brand_group",    "")
        log.info(
            "[RNA] Final resolved — brand_type='%s'  brand_status='%s'  brand_group='%s'",
            brand_type, brand_status, brand_group,
        )
        if not brand_type:
            # Show all unique country values in RNA to help diagnose mismatch
            rna_countries = sorted({k[0] for k in rna.lookup})
            rna_brands    = sorted({k[3] for k in rna.lookup})
            log.warning("[RNA] ✗ All 3 passes failed — no match found")
            log.warning("[RNA] Countries in RNA: %s", rna_countries)
            log.warning("[RNA] Brands in RNA:    %s", rna_brands)
    else:
        log.warning("[RNA] RNALoader not initialised — brand attributes will be empty")

    loader = ClarksOrderFormFWLoader(input_file)

    if loader.df.empty:
        log.warning("[Clarks-OrderFormFW] No data rows found — nothing to process")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_rows(raw_rows, loader, brand_lov_id, mdd, rna)

    if not generics:
        log.warning("[Clarks-OrderFormFW] No valid generics produced")
        return

    # ── Build season classification ID ───────────────────────────
    sp_code  = season[:2].upper() if len(season) >= 2 else season
    sys_part = season[2:] if len(season) > 2 else ""
    sy       = f"20{sys_part}" if len(sys_part) == 2 else sys_part
    season_id = f"CLH_{brand_lov_id}_{sp_code}{sy}"

    # ── Generate XML ─────────────────────────────────────────────
    xml_filename = f"{input_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    # ── Build Classifications block ──────────────────────────────
    cls_el  = build_classifications(brand_from_file, brand_lov_id, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))

    log.info("[Clarks-OrderFormFW] Writing XML → %s", xml_filename)
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

        log.info(
            "[Clarks-OrderFormFW] Processing %d generics",
            len(generics),
        )

        for generic in generics.values():
            px = build_product_xml(
                generic, brand_from_file, brand_lov_id, comp_code, sbu,
                season, season_id, country_code,
                brand_type=brand_type,
                brand_status=brand_status,
                brand_group=brand_group,
                mdd=mdd,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info(
        "[Clarks-OrderFormFW] XML written: %s  (%d KB)",
        xml_path.name, file_kb,
    )
    print(
        f"=== CLARKS ORDER FORM FOOTWEAR SUMMARY ===\n"
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
    parser = argparse.ArgumentParser(description="Clarks Order Form Footwear → Stibo XML")
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
