"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — 2XU Product Bible  v1.0      ║
║  Product Bible (H1 27 SKU List_by Garment Type)                  ║
║                                  → Stibo STEP XML                ║
╚══════════════════════════════════════════════════════════════════╝

Input file  : "H1 27 Product Bible - Wholesale.xlsx"
SKU master tab: "H1 27 SKU List_by Garment Type" (1st tab).
Catalogue tab : "H1 27 Catalogue" (Country of Manufacture).
Country LOV   : "Country Origin Lov" tab (Col A = LOV ID, Col B = Country Name).

  • Column names:
        Product Code
        Product Name
        Hierarchy          (e.g. Accessories / Apparel)
        Field of Play
        Gender
        Category           (e.g. "(a) Arm Cover")
        Colour Code / Colour Description
        Story              (sub-collection name — Aero / Flex / etc.)
        Drops
        New Status         (New / Seasonal / Carry Over)
        HS Code
        RRP (AUD) Inc / RRP (NZD) Inc / MSRP (USD) ex / MSRP (CAD) ex /
        RRP (GPB) Inc / RRP (EUR) Inc

Mapping (from Attributes List_complete_v6 — 2XU tab):
  AT_PrincipalStyleCode          ← Product Code
  AT_PrincipalStyleDescription   ← Product Name
  AT_PrincipalColorCode          ← Colour Code
  AT_PrincipalColorDescription   ← Colour Description
  AT_Gender / AT_BYGender        ← Gender → M/F/U
  AT_SAPAge                      ← default Adults
  AT_EComAgesCategory            ← default Ages 18+ years (id=18+Y)
  AT_PrincipalMerchandiseHierarchyL1 ← Hierarchy
  AT_PrincipalMerchandiseHierarchyL2 ← Field of Play
  AT_PrincipalMerchandiseHierarchyL3 ← Category
  AT_SportsCategoryEN                ← (Hierarchy + Field of Play) → SAP Mapping (col A+B → col H "Sports Category") → MDD "Sports Category" LOV (id_val = LOV ID)
  AT_Collection1                 ← Product Name
  AT_Collection2                 → Story (not sent — no mapping logic in attribute sheet)
  AT_EComProductNameEN           ← Brand + Product Name + SAP Gender + SAP Product Category (SAP Color excluded — no mapping yet)
  AT_FOB                         ← MSRP (USD) ex × 0.74
  AT_FOBCurrency                 ← USD
  AT_OriginalPrice/CurrentPrice  ← RRP per country (default AUD)
  AT_RetailPriceCurrency         ← country_code → MDD CountryIdToName → MDD RetailPriceCurrencyByName
  AT_CountryOrigin               ← Catalogue: Country of Manufacture → Country Origin Lov
  AT_MaterialType                ← Intercompany (default)
  AT_SAPProductFlag              ← intercompany-a (default, id=A)
  AT_BYArticleType               ← Inline (default)
  AT_UOM                         ← EA
  AT_SAPArticleCategory          ← 1 (Generic)
  AT_Country / AT_CountrySize    ← from filename / US default

  One Generic per (Product Code + Colour Code) row.
"""

import re
import sys
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

import os

BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR          = BASE_DIR / "input"
PRODUCTBIBLE_DIR   = INPUT_DIR / "productbible"
MDD_DIR            = INPUT_DIR / "mdd"
ATTR_DIR           = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PRODUCTBIBLE_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_productbible_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# LOV TABLES
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK", "2XU": "2XU",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter", "H1": "Half Year 1", "H2": "Half Year 2",
}
LOV_COUNTRY_ORIGIN = {
    "CN": "China", "VN": "Vietnam", "ID": "Indonesia",
    "KH": "Cambodia", "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia", "TH": "Thailand", "TW": "Taiwan",
}
LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
    "SM": "Smiggle",
}

TWOXU_GENDER_TO_SAP: dict[str, str] = {
    "MENS":   "M", "MEN":    "M",
    "WOMENS": "F", "WOMEN":  "F", "LADIES": "F",
    "UNISEX": "U", "YOUTH":  "U", "KIDS":   "U",
}

TWOXU_HIERARCHY_TO_PROD_GROUP: dict[str, str] = {
    "APPAREL":     "AP",
    "ACCESSORIES": "AC",
    "COMPRESSION": "CP",
    "SPECIALISED": "SP",
}

TWOXU_HIERARCHY_TO_DIVISION: dict[str, str] = {
    "APPAREL":     "A",
    "COMPRESSION": "A",
    "SPECIALISED": "A",
    "ACCESSORIES": "E",
}

# (Hierarchy, Field of Play) → Sports Category, hardcoded from
# "2XU - Stibo Mapping (upd. 21 April 2026).xlsx" → sheet "SAP Mapping"
# (col A "Hierarchy" + col B "Field Of Play" → col H "Sports Category").
# The resolved value is then looked up in the MDD "Sports Category" LOV for id_val.
# Where a (Hierarchy, Field of Play) pair produced more than one Sports Category
# in the sheet (driven by Category col D), the dominant value is used — raw
# sheet counts are shown in the trailing comment.
TWOXU_HIER_FOP_TO_SPORTS_CATEGORY: dict[tuple[str, str], str] = {
    ("ACCESSORIES", "RECOVERY"):    "Running",   # {'Running': 3}
    ("ACCESSORIES", "RUN"):         "Running",   # {'Running': 8, 'Others Accessories': 2}
    ("ACCESSORIES", "SPORTSWEAR"):  "Fitness",   # {'Running': 1, 'Fitness': 2, 'Others Accessories': 1}
    ("ACCESSORIES", "SWIM"):        "Swimming",  # {'Swimming': 2}
    ("ACCESSORIES", "TEAM SPORTS"): "Running",   # {'Running': 1}
    ("ACCESSORIES", "TRAIN"):       "Fitness",   # {'Running': 1, 'Fitness': 4, 'Others Accessories': 1}
    ("ACCESSORIES", "TRIATHLON"):   "Running",   # {'Running': 2}
    ("APPAREL", "RUN"):             "Running",   # {'Running': 15}
    ("APPAREL", "SPORTSWEAR"):      "Fitness",   # {'Fitness': 23, 'Others Apparel': 2}
    ("APPAREL", "TRAIN"):           "Fitness",   # {'Fitness': 11}
    ("COMPRESSION", "RECOVERY"):    "Running",   # {'Running': 4}
    ("COMPRESSION", "RUN"):         "Running",   # {'Running': 6}
    ("COMPRESSION", "TEAM SPORTS"): "Running",   # {'Running': 10}
    ("COMPRESSION", "TRAIN"):       "Fitness",   # {'Fitness': 6}
    ("SPECIALISED", "SWIM"):        "Swimming",  # {'Swimming': 6, 'Others Accessories': 2}
    ("SPECIALISED", "TRIATHLON"):   "Running",   # {'Running': 9}
}

TWOXU_COUNTRY_TO_PRICE: dict[str, tuple[str, str]] = {
    "AU": ("AUD", "RRP (AUD) Inc"),
    "NZ": ("NZD", "RRP (NZD) Inc"),
    "US": ("USD", "MSRP (USD) ex"),
    "CA": ("CAD", "MSRP (CAD) ex"),
    "GB": ("GBP", "RRP (GPB) Inc"),
    "UK": ("GBP", "RRP (GPB) Inc"),
    "EU": ("EUR", "RRP (EUR) Inc"),
    "ID": ("AUD", "RRP (AUD) Inc"),
    "PH": ("AUD", "RRP (AUD) Inc"),
    "MY": ("AUD", "RRP (AUD) Inc"),
    "SG": ("AUD", "RRP (AUD) Inc"),
    "VN": ("AUD", "RRP (AUD) Inc"),
    "TH": ("AUD", "RRP (AUD) Inc"),
    "KH": ("AUD", "RRP (AUD) Inc"),
}
DEFAULT_PRICE_CURRENCY = "AUD"
DEFAULT_PRICE_COLUMN   = "RRP (AUD) Inc"


def _short_colour_code(colour_code: str) -> str:
    """First letter of each '/'-separated colour token (BLK/SRF → BS)."""
    if not colour_code:
        return "XX"
    tokens = re.split(r"[\/\-\s]+", str(colour_code).strip().upper())
    tokens = [t for t in tokens if t]
    if len(tokens) >= 2:
        return (tokens[0][:1] + tokens[1][:1]).ljust(2, "X")
    if len(tokens) == 1:
        return tokens[0][:2].ljust(2, "X")
    return "XX"


def _strip_category_prefix(cat: str) -> str:
    """Strip leading "(x) " prefix from a Category cell, e.g. "(a) Arm Cover" → "Arm Cover"."""
    if not cat:
        return ""
    return re.sub(r"^\s*\([A-Za-z0-9]+\)\s*", "", cat).strip()


def _parse_fabric_composition(fabric_composition: str, fabric_lov: dict = None, content_lov: dict = None, material_lov: dict = None) -> tuple[str, str, str]:
    """
    Parse Fabric Composition to extract AT_Fabric, AT_Content, and AT_Material LOV IDs.
    
    Rules:
    1. Use the first material section on the top line (e.g., "Upper Main Body")
    2. If the top line says "TBD", use the following material section (e.g., "Outer")
    3. Choose the material with the LARGEST percentage and map to respective LOV IDs
    
    Examples:
        "Main:\n76% RECYCLED POLYESTER, 16% POLYESTER, \n8% ELASTANE" 
        → AT_Fabric=LOV_ID for "RECYCLED POLYESTER" from Fabric LOV (76% is largest)
        → AT_Content=LOV_ID for "RECYCLED POLYESTER" from Content LOV (76% is largest)
        → AT_Material=LOV_ID for "RECYCLED POLYESTER" from Material LOV (76% is largest)
        
        "25% Nylon, 85% PU"
        → AT_Fabric=LOV_ID for "PU" from Fabric LOV (85% is largest)
        → AT_Content=LOV_ID for "PU" from Content LOV (85% is largest)
        → AT_Material=LOV_ID for "PU" from Material LOV (85% is largest)
    
    Args:
        fabric_composition: The fabric composition text from Catalogue
        fabric_lov: Dict mapping fabric names to Fabric LOV IDs {"POLYESTER": "FAB001", ...}
        content_lov: Dict mapping content names to Content LOV IDs {"POLYESTER": "CON001", ...}
        material_lov: Dict mapping material names to Material LOV IDs {"POLYESTER": "MAT001", ...}
    
    Returns:
        (AT_Fabric, AT_Content, AT_Material) tuple of LOV IDs
    """
    if not fabric_composition:
        return "", "", ""
    
    fabric_str = str(fabric_composition).strip()
    if not fabric_str or fabric_str.upper() in ("NONE", "NAN", "N/A"):
        return "", "", ""
    
    # Split by newlines to get lines
    lines = [line.strip() for line in fabric_str.split('\n') if line.strip()]
    
    if not lines:
        return "", "", ""
    
    # Pattern to detect a fabric type line (e.g., "Main:", "Outer:", "Mesh:")
    fabric_type_pattern = re.compile(r'^([A-Za-z\s&]+):$')
    
    at_fabric = ""
    at_content = ""
    found_first = False
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Check if this line is a fabric type (ends with ":")
        match = fabric_type_pattern.match(line)
        
        if match:
            # If we already found the first fabric section, stop
            if found_first:
                break
            
            fabric_name = match.group(1).strip()
            
            # Skip if it's TBD
            if fabric_name.upper() != "TBD":
                # Collect all content lines until next fabric type or end
                content_parts = []
                i += 1
                while i < len(lines):
                    next_line = lines[i]
                    # Check if this is another fabric type
                    if fabric_type_pattern.match(next_line):
                        break
                    # Skip TBD lines
                    if next_line.upper() != "TBD":
                        content_parts.append(next_line)
                    i += 1
                
                # Set the result if we have content
                if content_parts:
                    at_fabric = fabric_name
                    at_content = " ".join(content_parts)
                    found_first = True
                    break
            else:
                # Skip TBD section
                i += 1
        else:
            # This is a content line without a fabric type header
            # Only use if we haven't found a fabric section yet
            if not found_first and line.upper() != "TBD":
                # Collect all non-TBD lines until we hit a fabric type
                content_parts = [line]
                i += 1
                while i < len(lines):
                    next_line = lines[i]
                    if fabric_type_pattern.match(next_line):
                        break
                    if next_line.upper() != "TBD":
                        content_parts.append(next_line)
                    i += 1
                
                at_content = " ".join(content_parts)
                found_first = True
                break
            else:
                i += 1
    
    # Extract material with LARGEST percentage and map to LOVs
    at_fabric_lov_id = ""
    at_content_lov_id = ""
    at_material_lov_id = ""
    largest_material = ""
    
    if at_content:
        # Find all materials with percentages: "XX% MATERIAL_NAME"
        # Pattern handles: "25% Nylon. 85% PU" or "76% POLYESTER, 16% ELASTANE"
        # - Case-insensitive for material names (Nylon, POLYESTER, PU)
        # - Separators: comma (,), period (.), or end of string
        pattern = r'(\d+(?:\.\d+)?)\s*%\s+([A-Za-z\s\-]+?)(?:[,.]|$|\n)'
        matches = re.findall(pattern, at_content)
        
        if matches:
            # Find the material with the largest percentage
            max_percentage = 0.0
            for percentage_str, material_name in matches:
                percentage = float(percentage_str)
                if percentage > max_percentage:
                    max_percentage = percentage
                    largest_material = material_name.strip()
        else:
            # Fallback: if no percentage found, use first material
            parts = at_content.split(',')
            if parts:
                largest_material = re.sub(r'\d+%\s*', '', parts[0]).strip()
        
        # Map largest material to Fabric LOV ID (exact match only, case-insensitive)
        if largest_material and fabric_lov:
            at_fabric_lov_id = fabric_lov.get(largest_material, "")
            
            # If no exact match, try case-insensitive
            if not at_fabric_lov_id:
                largest_upper = largest_material.upper()
                for fabric_name, lov_id in fabric_lov.items():
                    if fabric_name.upper() == largest_upper:
                        at_fabric_lov_id = lov_id
                        break
        
        # Map largest material to Content LOV ID (exact match only, case-insensitive)
        if largest_material and content_lov:
            at_content_lov_id = content_lov.get(largest_material, "")
            
            # If no exact match, try case-insensitive
            if not at_content_lov_id:
                largest_upper = largest_material.upper()
                for content_name, lov_id in content_lov.items():
                    if content_name.upper() == largest_upper:
                        at_content_lov_id = lov_id
                        break
        
        # Map largest material to Material LOV ID (exact match only, case-insensitive)
        if largest_material and material_lov:
            at_material_lov_id = material_lov.get(largest_material, "")
            
            # If no exact match, try case-insensitive
            if not at_material_lov_id:
                largest_upper = largest_material.upper()
                for material_name, lov_id in material_lov.items():
                    if material_name.upper() == largest_upper:
                        at_material_lov_id = lov_id
                        break
    
    return at_fabric_lov_id, at_content_lov_id, at_material_lov_id


def _parse_season(season: str) -> tuple[str, str]:
    """
    Split a season token into (sea_code, sea_year).

    Examples:
        SM8008  → ("SM", "8008")   ← 4-digit non-calendar batch number, used as-is
        H226    → ("H2", "2026")   ← 2-digit year, century prepended
        FW2026  → ("FW", "2026")   ← already full 4-digit year
        SS26    → ("SS", "2026")
    """
    s = (season or "").strip().upper()
    if len(s) < 3:
        return s, ""
    # Split: first 2 alpha chars = code, rest = numeric year part
    if s[:2].isalpha() and s[2:].isdigit():
        code    = s[:2]
        num_str = s[2:]
        if len(num_str) == 4:
            # Either a full calendar year (2026) or a batch number (8008) — use as-is
            year = num_str
        elif len(num_str) == 2:
            year = f"20{num_str}"
        else:
            year = num_str
        return code, year
    return s[:2], ""


# ══════════════════════════════════════════════════════════════════
# STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """Loads Core Attributes + LOVs from the MDD Excel."""

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws = wb["Core Attributes"]
        rows = list(ws.iter_rows(values_only=True))
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
                "cardinality":  _v("Cardinality"),
                "validation":   _v("Validation Base Type"),
                "multi_valued": _v("Multi Valued"),
                "lov_name":     _v("Name of LOV"),
                "max_chars":    _v("Max Characters"),
                "group":        _v("PIM Attribute Group"),
            }

        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        self._load_age_lov(wb)
        self._load_retail_price_currency_lovs(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = str(val_id).strip() if val_id else str(val_name).strip()

    def _load_age_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return
        sap_age_map, by_age_map = {}, {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display  = str(row[0]).strip()
            sap_age_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            by_age_val   = str(row[5]).strip() if len(row) > 5 and row[5] else ""
            if age_display and sap_age_code:
                sap_age_map[age_display] = sap_age_code
            if age_display and by_age_val:
                by_age_map[age_display]  = by_age_val
        self.lovs["SAPAge"] = sap_age_map
        self.lovs["BYAge"]  = by_age_map

    def _load_retail_price_currency_lovs(self, wb):
        # CountryIdToName: {country_code.upper(): country_name}  (from "Country LOV", col A=code, col B=name)
        country_id_to_name: dict[str, str] = {}
        if "Country LOV" in wb.sheetnames:
            for r in wb["Country LOV"].iter_rows(min_row=2, values_only=True):
                if r and r[0] and len(r) > 1 and r[1]:
                    country_id_to_name[str(r[0]).strip().upper()] = str(r[1]).strip()
        self.lovs["CountryIdToName"] = country_id_to_name

        # RetailPriceCurrencyByName: {display_name.upper(): currency_code}
        # (from "Retail Price Currency LOV", col A=display name, col B=currency code; data starts row 3)
        retail_price_currency_lov: dict[str, str] = {}
        if "Retail Price Currency LOV" in wb.sheetnames:
            for r in wb["Retail Price Currency LOV"].iter_rows(min_row=3, values_only=True):
                if r and r[0] and len(r) > 1 and r[1]:
                    retail_price_currency_lov[str(r[0]).strip().upper()] = str(r[1]).strip()
        self.lovs["RetailPriceCurrencyByName"] = retail_price_currency_lov

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                code, name = row[0], row[1]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )


class AttributesListLoader:
    """Loads the 2XU tab from the Attributes List workbook (for reference)."""

    def __init__(self, path: Path, brand: str = "2XU"):
        self.path      = path
        self.brand     = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()),
            wb.sheetnames[0],
        )
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"),
            None,
        )
        if hdr_idx is None:
            log.warning("[AttrList] Cannot find header row in sheet '%s'", sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}

        for row in rows[hdr_idx + 1:]:
            attr = row[1] if len(row) > 1 else None
            if not attr:
                continue
            self.attr_map.append({
                "attribute":     str(attr).strip(),
                "cluster":       self._s(row, col, "Cluster"),
                "indicator":     self._s(row, col, "Indicator"),
                "mapping_logic": self._s(row, col, "Field Name / Mapping Logic"),
            })

        wb.close()
        log.info("[AttrList] %d attributes loaded from '%s'",
                 len(self.attr_map), sheet_name)

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class RNALoader:
    """
    Loads the 'Source Mapping Related RNA' tab from the Attributes List workbook.
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
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": ""}


class TwoXUProductBibleLoader:
    """
    Loads the 2XU Product Bible workbook.

    Uses the 1st sheet (SKU List) and "H1 27 Catalogue" sheet.
    Header row is auto-detected by looking for the 'Product Code' column.

    Each DataFrame row represents one (Product Code + Colour Code)
    combination → one Generic article in STEP.
    """

    HEADER_KEY_COLUMN = "PRODUCT CODE"

    def __init__(self, path: Path):
        self.path         = path
        self.df           = pd.DataFrame()
        self.df_catalogue = pd.DataFrame()
        self.sheet        = ""
        self.country_origin_lov: dict[str, str] = {}
        self._load()

    def _load(self):
        log.info("[ProductBible-2XU] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # ── Load SKU List sheet (1st sheet) ──────────────────────
        target = wb.sheetnames[0]
        log.info("[ProductBible-2XU] Using 1st sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = None
        for i, r in enumerate(rows):
            if any(
                isinstance(v, str) and v.strip().upper() == self.HEADER_KEY_COLUMN
                for v in r if v
            ):
                hdr_idx = i
                break

        if hdr_idx is None:
            log.error("[ProductBible-2XU] Cannot find header row "
                      "(no '%s' column)", self.HEADER_KEY_COLUMN)
            wb.close()
            return

        log.info("[ProductBible-2XU] Header at row %d (0-indexed)", hdr_idx)
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        if "Product Code" in df.columns:
            df = df[
                df["Product Code"].notna()
                & (~df["Product Code"].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        df = df.reset_index(drop=True)
        # ══════════════════════════════════════════════════════════════════
        # TEST LIMITER — Set to None for production, or a number to limit rows
        # ══════════════════════════════════════════════════════════════════
        # TEST_ROW_LIMIT = 10  # Set to None to process all rows
        # if TEST_ROW_LIMIT is not None:
        #     df = df.head(TEST_ROW_LIMIT)  # Use .head() to get FIRST N rows
        #     log.info("⚠️  TEST MODE: Limited to first %d rows", TEST_ROW_LIMIT)
        # ══════════════════════════════════════════════════════════════════
        self.df    = df
        self.sheet = target
        log.info("[ProductBible-2XU] %d SKU rows loaded", len(self.df))

        # ── Load Country Origin LOV ───────────────────────────────
        self._load_country_origin_lov(wb)

        # ── Load Catalogue sheet ──────────────────────────────────
        self._load_catalogue_sheet(wb)

        wb.close()

    def _load_country_origin_lov(self, wb):
        """Load Country Origin LOV mapping from 'Country Origin Lov' tab."""
        sheet_name = next(
            (s for s in wb.sheetnames if "COUNTRY" in s.upper() and "ORIGIN" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.debug("[ProductBible-2XU] Country Origin Lov sheet not in product bible — MDD LOV will be used")
            return

        log.info("[ProductBible-2XU] Loading Country Origin LOV from: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Skip header row, column A = LOV ID, column B = Country Name
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id = str(row[0]).strip() if row[0] else ""
            country_name = str(row[1]).strip() if row[1] else ""
            if lov_id and country_name:
                self.country_origin_lov[country_name.upper()] = lov_id

        log.info("[ProductBible-2XU] %d Country Origin LOV entries loaded", len(self.country_origin_lov))

    def _load_catalogue_sheet(self, wb):
        """Load H1 27 Catalogue sheet for Country of Manufacture."""
        sheet_name = next(
            (s for s in wb.sheetnames if "CATALOGUE" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[ProductBible-2XU] Catalogue sheet not found")
            return

        log.info("[ProductBible-2XU] Loading Catalogue from: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Find header row — Catalogue sheet uses "Style Code" as key column
        CATALOGUE_HEADER_KEY = "STYLE CODE"
        hdr_idx = None
        for i, r in enumerate(rows):
            if any(
                isinstance(v, str) and v.strip().upper() == CATALOGUE_HEADER_KEY
                for v in r if v
            ):
                hdr_idx = i
                break

        if hdr_idx is None:
            log.warning("[ProductBible-2XU] Cannot find header in Catalogue sheet")
            return

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df_cat = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        if "Product Code" in df_cat.columns:
            df_cat = df_cat[
                df_cat["Product Code"].notna()
                & (~df_cat["Product Code"].astype(str).str.strip().isin(["", "None", "nan"]))
            ]

        self.df_catalogue = df_cat.reset_index(drop=True)
        log.info("[ProductBible-2XU] %d Catalogue rows loaded", len(self.df_catalogue))


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw:
        return raw
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _num(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("none", "nan"):
        return ""
    m = re.search(r"[\d.]+", s)
    return m.group() if m else ""


def map_article_twoxu(pb_row: dict, brand_code: str = "2XU",
                      country_code: str = "", catalogue_row: dict = None,
                      fabric_lov: dict = None, content_lov: dict = None,
                      material_lov: dict = None) -> dict:
    """Map one Product Bible row → unified article dict."""
    normalized_row = {
        re.sub(r"\s+", " ", str(k).strip()).lower(): v
        for k, v in pb_row.items()
    }

    def _get_by_aliases(*aliases: str) -> str:
        """Resolve a value by trying multiple header aliases (case/whitespace/newline tolerant)."""
        for alias in aliases:
            key = re.sub(r"\s+", " ", alias.strip()).lower()
            if key in normalized_row:
                val = _s(normalized_row.get(key))
                if val:
                    return val
        return ""

    def _first_non_empty(*keys: str) -> str:
        for k in keys:
            v = _s(pb_row.get(k))
            if v:
                return v
        return ""

    style_code  = _s(pb_row.get("Product Code"))
    desc        = _s(pb_row.get("Product Name"))
    story       = _s(pb_row.get("Story"))

    colour_code = _s(pb_row.get("Colour Code"))
    colour_desc = _s(pb_row.get("Colour Description"))

    gender_raw    = _s(pb_row.get("Gender"))
    hierarchy_raw = _get_by_aliases(
        "Hierarchy",
        "App/Acc Product Bibble : 'Hierarchy'",
        "App/Acc\nProduct Bibble : 'Hierarchy'",
    ) or _first_non_empty("Hierarchy", "App/Acc\nProduct Bibble : 'Hierarchy'")
    fop_raw       = _get_by_aliases(
        "Field of Play",
        "Field Of Play",
        "App/Acc Product Bibble : 'Field of Play'",
        "App/Acc\nProduct Bibble : 'Field of Play'",
    ) or _first_non_empty("Field of Play", "Field Of Play", "App/Acc\nProduct Bibble : 'Field of Play'")
    category_raw  = _get_by_aliases(
        "Category",
        "App/Acc Product Bibble : 'Category'",
        "App/Acc\nProduct Bibble : 'Category'",
    ) or _first_non_empty("Category", "App/Acc\nProduct Bibble : 'Category'")
    category      = _strip_category_prefix(category_raw)
    new_status    = _s(pb_row.get("New Status"))
    drops_val     = _s(pb_row.get("Drops")) or _s(pb_row.get("Drops "))

    hs_code = _s(pb_row.get("HS Code"))
    if hs_code.endswith(".0"):
        hs_code = hs_code[:-2]

    msrp_usd = _num(pb_row.get("MSRP (USD) ex"))

    # ── Derived: gender ─────────────────────────────────────────
    gender_code = TWOXU_GENDER_TO_SAP.get(gender_raw.upper(), "U")

    # ── Derived: division / product group ───────────────────────
    div_letter = TWOXU_HIERARCHY_TO_DIVISION.get(hierarchy_raw.upper(), "A")
    prod_group = TWOXU_HIERARCHY_TO_PROD_GROUP.get(hierarchy_raw.upper(), "AP")

    # ── Derived: pricing ────────────────────────────────────────
    country_key = (country_code or "").strip().upper()
    currency, price_col = TWOXU_COUNTRY_TO_PRICE.get(
        country_key, (DEFAULT_PRICE_CURRENCY, DEFAULT_PRICE_COLUMN)
    )
    rrp_str = _num(pb_row.get(price_col))

    # ── Derived: FOB = MSRP (USD) × 0.74 ─────────────────────────
    fob_str = ""
    if msrp_usd:
        try:
            fob_str = f"{float(msrp_usd) * 0.74:.2f}"
        except (TypeError, ValueError):
            fob_str = ""

    # ── Derived: SAP / Generic / Variant codes ──────────────────
    sap_color_short = _short_colour_code(colour_code)
    # Principal Color Code: clean alphanumeric version of colour_code for AT_InboundGenericCode
    principal_color_code = re.sub(r"[^A-Z0-9]", "", colour_code.upper())
    sap_style_code  = f"{style_code.upper()}{sap_color_short}"
    # AT_InboundGenericCode = Brand Code + Principal Style Code + Principal Color Code
    # Each unique (style_code, colour_code) combination creates a new generic
    generic_code    = f"{brand_code}{style_code.upper()}{principal_color_code}"

    sap_color_3 = re.sub(r"[^A-Z0-9]", "", colour_code.upper())[:3].ljust(3, "0")
    variant_code = f"{generic_code}{sap_color_3}000"

    # ── Country of Origin from Catalogue ────────────────────────
    # Get raw country name from Catalogue sheet; LOV lookup done in _add_generic_values using MDD
    coo_country = ""
    if catalogue_row:
        coo_country = _s(catalogue_row.get("Country of Manufacture"))
    
    # ── Fabric Composition from Catalogue ────────────────────────
    at_fabric = ""
    at_content = ""
    at_material = ""
    if catalogue_row:
        fabric_composition = _s(catalogue_row.get("Fabric Composition")) or _s(catalogue_row.get("Fabric Composition "))
        at_fabric, at_content, at_material = _parse_fabric_composition(fabric_composition, fabric_lov, content_lov, material_lov)

    return {
        # Core identifiers
        "article_no":       style_code,
        "model_name":       desc,
        "model_name_short": desc,
        "brand_code":       brand_code,

        # Colour
        "colour":               colour_desc,
        "colour_code":          colour_code,
        "principal_color_code": principal_color_code,
        "sap_color_short":      sap_color_short,
        "sap_color_3":          sap_color_3,

        # Gender / Age
        "gender_code":      gender_code,
        "gender_raw":       gender_raw,
        "age_code":         "AD",
        "age_raw":          "Adults",

        # Classification (note: hierarchy ≈ what the prior file called "range")
        "hierarchy":        hierarchy_raw,
        "range":            hierarchy_raw,
        "field_of_play":    fop_raw,
        "category":         category,
        "story":            story,
        "div_letter":       div_letter,
        "prod_group":       prod_group,
        "nsc":              new_status,
        "drop":             drops_val,

        # Commercial
        "article_type":     "Inline",
        "art_category":     "1",
        "collection1":      desc,
        "collection2":      story,
        "franchise":        "",

        # Pricing
        "rrp":              rrp_str,
        "currency":         currency,
        "fob":              fob_str,

        # Trade
        "hs_code":          hs_code,

        # Origin - from Catalogue sheet "Country of Manufacture" column (LOV lookup via MDD)
        "coo":              coo_country,

        # Fabric Composition - from Catalogue sheet (LOV IDs)
        "at_fabric":        at_fabric,
        "at_content":       at_content,
        "at_material":      at_material,

        # Derived codes
        "sap_style_code":   sap_style_code,
        "generic_code":     generic_code,
        "variant_code":     variant_code,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    warns = []
    art   = mapped["article_no"]
    field_checks = {
        "article_no":   "AT_PrincipalStyleCode",
        "gender_code":  "AT_Gender",
        "age_code":     "AT_SAPAge",
        "brand_code":   "AT_Brand",
    }
    for field, at_id in field_checks.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id}")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(parent, attr_id, value="", id_val="", derived=False):
    if derived:
        return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id = str(id_val).strip() if id_val else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
        return el
    clean_val = str(value).strip() if value else ""
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent, attr_id, id_val, label=""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code, lookup, default=""):
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def resolve_brand_lov_id(mdd, brand_code):
    src_code = (brand_code or "").strip().upper()
    _, brand_label = _lov(src_code, LOV_BRAND, src_code)
    brand_lov_id = src_code
    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})
        for display_name, lov_id in brand_lov.items():
            if (display_name or "").strip().upper() == (brand_label or "").strip().upper():
                brand_lov_id = str(lov_id).strip()
                break
    return brand_lov_id, brand_label


def _resolve_retail_price_currency(country_code: str, mdd: MDDLoader = None) -> str:
    """
    Resolve retail price currency from country code using MDD.

    Steps:
    1. Look up country_code (e.g. "ID") in MDD CountryIdToName → get country name (e.g. "Indonesia")
    2. Look up country name in MDD RetailPriceCurrencyByName → get currency code (e.g. "IDR")

    Uses partial/fuzzy matching for country name lookup.
    """
    if not country_code or not mdd:
        return ""

    code_upper = country_code.strip().upper()

    # Step 1: country code → country name
    country_id_to_name = mdd.lovs.get("CountryIdToName", {})
    country_name = country_id_to_name.get(code_upper, "")
    if not country_name:
        return ""

    # Step 2: country name → currency code (exact then partial match)
    retail_currency_lov = mdd.lovs.get("RetailPriceCurrencyByName", {})
    country_name_upper  = country_name.strip().upper()

    if country_name_upper in retail_currency_lov:
        return retail_currency_lov[country_name_upper]

    for lov_key, currency in retail_currency_lov.items():
        if country_name_upper in lov_key or lov_key in country_name_upper:
            return currency

    return ""


def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None):
    """Write all Generic-level <Value> elements for a 2XU Product Bible article."""
    written: set[str] = set()
    def _w(attr_id, value="", id_val="", derived=False):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val, derived=derived)
            written.add(attr_id)

    def _mw(attr_id, id_val):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    sbu_code, _ = _lov(sbu, LOV_SBU, sbu)
    _mw("AT_SBU", sbu_code)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    brand_lov_id = art["brand_code"]
    _w("AT_Brand",      id_val=brand_lov_id)
    _w("AT_BrandGroup", id_val=brand_name.upper())

    # ── Principal identifiers ────────────────────────────────────
    # AT_PrincipalStyleCode: mapped from 'Product Code' column
    _w("AT_PrincipalStyleCode",        art["article_no"])
    # AT_PrincipalStyleDescription: mapped from 'Product Name' column
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["colour_code"])
    # AT_PrincipalColorName: mapped from 'Colour Description' column
    _w("AT_PrincipalColorName",        art["colour"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
    _w("AT_Color", value="")

    # ── SAP Color ────────────────────────────────────────────────
    # colour_id = ""
    # if mdd:
    #     color_lov = mdd.lovs.get("Color Code", {})
    #     colour_id = (
    #         color_lov.get(art["colour"], "")
    #         or color_lov.get(art["colour"].upper(), "")
    #         or color_lov.get(art["colour_code"], "")
    #         or color_lov.get(art["colour_code"].upper(), "")
    #     )

    # if colour_id:
    #     if colour_id.isdigit():
    #         colour_id = colour_id.zfill(3)
    #     _w("AT_Color", id_val=colour_id)
    #     sap_color_code = colour_id
    # else:
    #     _w("AT_Color", value=art["colour_code"].upper())
    #     sap_color_code = art["sap_color_short"]
    sap_color_code = art["sap_color_short"]

    # AT_InboundGenericCode = Brand Code + Principal Style Code + Principal Color Code
    # Keep this identical to KEY_InboundArticle to satisfy STEP unique key definition.
    at_generic_val = art.get("generic_code", "")
    if not at_generic_val:
        principal_color_code = re.sub(r"[^A-Z0-9]", "", art["colour_code"].upper())
        at_generic_val = f"{brand_lov_id}{art['article_no'].upper()}{principal_color_code}"
    art["at_generic_val"] = at_generic_val
    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",                    g_label, id_val=g_code)
    _w("AT_BYGender",                  g_label, id_val=g_code)
    # AT_PrincipalGenderDescription: mapped from 'Gender' column in Product Bible
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age — 2XU default Adults ─────────────────────────────────
    sap_age_map = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map  = mdd.lovs.get("BYAge",  {}) if mdd else {}
    sap_age_code = sap_age_map.get("Adults", art["age_code"])
    by_age_val   = by_age_map.get("Adults", "Adult")
    _w("AT_SAPAge", "AD", id_val="AD")
    _w("AT_BYAge",  by_age_val.upper(), id_val=by_age_val.upper())

    # ── E-com Ages Category — default Ages 18+ years ─────────────
    _w("AT_EComAgesCategory", id_val="18+Y")

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    # sea_year is pre-computed upstream in run() via _parse_season()
    season_year = art.get("season_year", "")
    _w("AT_SeasonYear", season_year)

    # ── Country of Origin ────────────────────────────────────────
    # Use MDD's Country Origin LOV to resolve country name → LOV code
    coo_name = art.get("coo", "")
    if coo_name and mdd:
        # MDD Country Origin LOV: {country_name → code} e.g. {"China": "CN", "Vietnam": "VN"}
        country_origin_lov = mdd.lovs.get("Country Origin", {})
        # Try exact match first, then case-insensitive
        coo_code = country_origin_lov.get(coo_name, "")
        if not coo_code:
            coo_name_upper = coo_name.upper()
            for name, code in country_origin_lov.items():
                if name.upper() == coo_name_upper:
                    coo_code = code
                    break
        if coo_code:
            _w("AT_CountryOrigin", "", id_val=coo_code)
    elif coo_name:
        # If no MDD, write raw value
        _w("AT_CountryOrigin", coo_name)

    # ── SAP Article Category — Single (00) ───────────────────────
    _w("AT_SAPArticleCategory", id_val=art["art_category"])

    # ── BY Article Type — Inline (default) ───────────────────────
    at_code = art["article_type"]
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = ("License" if at_from_filename.lower().startswith("lic")
                   else at_from_filename.capitalize())
    _w("AT_BYArticleType", at_code, id_val=at_code)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   "Yes", id_val="Y")
    _w("AT_SAPIndicator",  "No",  id_val="N")

    # ── UOM ──────────────────────────────────────────────────────
    _w("AT_UOM", "Each", id_val="EA")

    # ── Pricing ──────────────────────────────────────────────────
    if art["rrp"]:
        _w("AT_OriginalPrice",       art["rrp"])
        _w("AT_CurrentPrice",        art["rrp"])
    # AT_RetailPriceCurrency: dynamic MDD lookup (country_code → country_name → currency)
    _rpc = _resolve_retail_price_currency(art.get("country_code", ""), mdd)
    if _rpc:
        _w("AT_RetailPriceCurrency", "", id_val=_rpc)
    else:
        log.warning(
            "[RetailPriceCurrency] No MDD LOV match for country_code='%s' — skipping attribute",
            art.get("country_code", ""),
        )
    if art["fob"]:
        _w("AT_FOB",        art["fob"])
    _w("AT_FOBCurrency", "USD", id_val="USD")

    # ── Material Type — ZINA ─────────────────────────────────────
    _w("AT_MaterialType", id_val="ZINA")

    # ── SAP Product Flag — intercompany-a ────────────────────────
    _w("AT_SAPProductFlag", id_val="A")

    # ── 2XU Merchandise hierarchy (L1=Hierarchy, L2=FoP, L3=Category)
    if art["hierarchy"]:
        _w("AT_PrincipalMerchandiseHierarchyL1", art["hierarchy"])
    if art["field_of_play"]:
        _w("AT_PrincipalMerchandiseHierarchyL2", art["field_of_play"])
    if art["category"]:
        _w("AT_PrincipalMerchandiseHierarchyL3", art["category"])

    # ── Sports Category EN — (Hierarchy, Field of Play) → SAP Mapping → MDD LOV ──
    # 1. Hierarchy + Field of Play (source columns)
    # 2. Map to Sports Category via TWOXU_HIER_FOP_TO_SPORTS_CATEGORY
    #    (hardcoded from "2XU - Stibo Mapping (upd. 21 April 2026).xlsx" → "SAP Mapping",
    #     col A "Hierarchy" + col B "Field Of Play" → col H "Sports Category")
    # 3. Look that value up in the MDD "Sports Category" LOV for the id_val
    _hier_raw = (art.get("hierarchy") or "").strip()
    _fop_raw  = (art.get("field_of_play") or "").strip()
    _sc_source = None
    if _hier_raw and _fop_raw:
        _sc_source = TWOXU_HIER_FOP_TO_SPORTS_CATEGORY.get(
            (_hier_raw.upper(), _fop_raw.upper())
        )
        if not _sc_source:
            log.warning(
                "[SportsCategoryEN] No SAP Mapping row for (Hierarchy '%s', Field of Play '%s') "
                "— skipping attribute",
                _hier_raw, _fop_raw,
            )
    if _sc_source:
        _sc_lov = mdd.lovs.get("Sports Category", {}) if mdd else {}
        # MDD sheet: {id_str: display_name} → invert to {display_name.upper(): id_str}
        _sc_by_name = {v.strip().upper(): k for k, v in _sc_lov.items()}
        _sc_id      = _sc_by_name.get(_sc_source.upper())
        # Resolve display name from LOV (preserves original casing from MDD)
        _sc_label   = _sc_lov.get(_sc_id, _sc_source) if _sc_id else None
        if _sc_id and _sc_label:
            # LOV IDs are 2-digit (01..16); pad single digits e.g. 7 -> 07
            if _sc_id.isdigit():
                _sc_id = _sc_id.zfill(2)
            _w("AT_SportsCategoryEN", _sc_label, id_val=_sc_id)
        else:
            log.warning(
                "[SportsCategoryEN] No MDD LOV match for Sports Category '%s' "
                "(from Hierarchy '%s' / Field of Play '%s') — skipping attribute",
                _sc_source, _hier_raw, _fop_raw,
            )

    # ── Collections ──────────────────────────────────────────────
    if art["collection1"]:
        _w("AT_Collection1", art["collection1"])
    # AT_Collection2 (Story) not sent — no mapping logic in attribute sheet

    # ── Ecom Product Name EN ──────────────────────────────────────
    # Formula: Brand + Product Name + SAP Gender + SAP Product Category
    # (SAP Color excluded — no mapping available yet)
    _ecom_parts = [
        brand_name,
        art.get("model_name", ""),
        g_label,
        art.get("category", ""),
    ]
    _ecom_name_en = " ".join(p.strip() for p in _ecom_parts if p and p.strip())
    if _ecom_name_en:
        _w("AT_EComProductNameEN", _ecom_name_en)

    # ── Country / Size defaults ──────────────────────────────────
    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)
    
    # ── Brand Type / Brand Category ──────────────────────────────
    _w("AT_BrandType",     art.get("brand_type",     ""))
    _w("AT_BrandCategory", art.get("brand_category", ""))
    
    # ── Fabric / Content / Material ───────────────────────────────
    _w("AT_Fabric",   art.get("at_fabric",   ""), id_val=art.get("at_fabric",   ""))
    _w("AT_Content",  art.get("at_content",  ""), id_val=art.get("at_content",  ""))
    _w("AT_Material", art.get("at_material", ""), id_val=art.get("at_material", ""))
    
    _w("AT_CountrySize", "", id_val="US")


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(brand, brand_code, comp_code, sbu, season_code):
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix, sea_year = _parse_season(season_code)
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
        "H1": "Half 1",         "H2": "Half 2",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
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
    ET.SubElement(confirmed,    f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed,    f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


def build_product_xml(art, brand, brand_code, comp_code, sbu, season_id, mdd=None):
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = art.get("brand_code", brand_code) or brand_code
    div_letter = art.get("div_letter", "A")
    parent_id  = f"PPH_{div_letter}-TempSubCat"

    # KEY_InboundArticle must match AT_InboundGenericCode exactly.
    key_article = art.get("generic_code", "")
    if not key_article:
        principal_color_code = art.get("principal_color_code", "")
        if not principal_color_code:
            # Fallback: derive from colour_code if not pre-computed
            principal_color_code = re.sub(r"[^A-Z0-9]", "", art.get("colour_code", "").upper())
        key_article = f"{b_code}{art['article_no'].upper()}{principal_color_code}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["model_name"] or article_no
    )

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    row, brand_code, mdd, season, season_year, country_code, article_type_from_filename, catalogue_row, fabric_lov, content_lov, material_lov, brand_type, brand_category = row_tuple
    mapped = map_article_twoxu(
        row, brand_code=brand_code, country_code=country_code,
        catalogue_row=catalogue_row,
        fabric_lov=fabric_lov, content_lov=content_lov, material_lov=material_lov,
    )
    mapped["season"]                     = season
    mapped["season_year"]                = season_year
    mapped["country_code"]               = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    mapped["brand_type"]                 = brand_type
    mapped["brand_category"]             = brand_category
    warns = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

    args attributes:
        brand       str   e.g. "2XU"
        brand_code  str   e.g. "2XU"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SP"
        season      str   e.g. "H127"
        seq         int   e.g. 1
        country_code               (optional)  e.g. "AU"
        article_type_from_filename (optional)
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    pb_files = list(PRODUCTBIBLE_DIR.glob("*.xlsx"))

    for label, val in [
        ("MDD", mdd_f),
        ("Attributes List", attr_f),
        ("Product Bible", pb_files),
    ]:
        if not val:
            raise FileNotFoundError(f"No {label} file found in expected directory")

    mdd = MDDLoader(mdd_f)
    _al = AttributesListLoader(attr_f, brand=args.brand)
    rna = RNALoader(attr_f)

    log.info(
        "Pipeline args → brand=%s brand_code=%s comp_code=%s sbu=%s season=%s seq=%s "
        "country=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
        getattr(args, "country_code", ""),
    )

    brand_lov_id, brand_lov_label = resolve_brand_lov_id(mdd, args.brand_code)
    log.info(
        "Resolved brand LOV → brand_code=%s  label=%s  lov_id=%s",
        args.brand_code, brand_lov_label, brand_lov_id,
    )

    # ── Parse season once, upstream (same pattern as crocs/aldo) ──
    sea_code, sea_year = _parse_season(args.season)
    season_id = f"CLH_{brand_lov_id}_{sea_code}{sea_year}"
    log.info("Season parsed → code=%s  year=%s  season_id=%s",
             sea_code, sea_year, season_id)

    # ── Resolve AT_BrandType / AT_BrandCategory ──────────────────
    _country_code = (getattr(args, "country_code", "") or "").strip().upper() or "AU"
    
    # Get country name from MDD Country LOV (LOV has: Name → Code, need to invert)
    country_lov = mdd.lovs.get("Country", {}) if mdd else {}
    # Invert: Code → Name
    country_by_code = {v: k for k, v in country_lov.items()}
    _country_name = country_by_code.get(_country_code, _country_code)
    if _country_name:
        _country_name = _country_name.upper()

    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        _country_name, args.comp_code, args.sbu, brand_lov_id,
    )

    # Try multiple SBU variations (SP → SPORTS, SPORTS DIRECT, etc.)
    _sbu_variants = [args.sbu]
    if args.sbu == "SP":
        _sbu_variants.extend(["SPORTS", "SPORTS DIRECT"])
    
    _rna_result = {"brand_type": "", "brand_category": ""}
    
    # Try all combinations with SBU variants
    for sbu_var in _sbu_variants:
        if _rna_result.get("brand_type") or _rna_result.get("brand_category"):
            break
        _rna_result = rna.get(_country_name, args.comp_code, sbu_var, brand_lov_id)
        if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
            _rna_result = rna.get(_country_name, args.comp_code, sbu_var, args.brand_code)
        if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
            _rna_result = rna.get(_country_name, str(int(args.comp_code)), sbu_var, brand_lov_id)
        if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
            _rna_result = rna.get(_country_name, str(int(args.comp_code)), sbu_var, args.brand_code)
    
    # Final fallback: fuzzy match (ignores comp_code)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        log.warning("[RNA] Exact match failed — trying fuzzy (country+sbu+brand)")
        for sbu_var in _sbu_variants:
            _rna_result = rna.get_fuzzy(_country_name, sbu_var, brand_lov_id)
            if _rna_result.get("brand_type") or _rna_result.get("brand_category"):
                break
            _rna_result = rna.get_fuzzy(_country_name, sbu_var, args.brand_code)
            if _rna_result.get("brand_type") or _rna_result.get("brand_category"):
                break

    _brand_type     = _rna_result.get("brand_type",     "")
    _brand_category = _rna_result.get("brand_category", "")

    log.info(
        "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
        _country_name, _brand_type, _brand_category,
    )

    all_warnings: list[str] = []

    for pb_path in pb_files:
        log.info("─── Processing Product Bible: %s ───", pb_path.name)

        pb = TwoXUProductBibleLoader(pb_path)
        if pb.df.empty:
            log.warning("[ProductBible-2XU] Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in pb.df.iterrows()]
        total_rows = len(rows)
        log.info("Total valid rows: %d", total_rows)

        out_name = pb_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Create lookup for catalogue data by Style Code + Colour Code ──
        # Note: Catalogue sheet uses "Style Code" (not "Product Code") as key column
        catalogue_lookup = {}
        if not pb.df_catalogue.empty:
            for _, cat_row in pb.df_catalogue.iterrows():
                prod_code = _s(cat_row.get("Style Code") or cat_row.get("Product Code"))
                colour_code = _s(cat_row.get("Colour Code"))
                if prod_code and colour_code:
                    key = f"{prod_code}|{colour_code}"
                    catalogue_lookup[key] = cat_row.to_dict()

        # Get Fabric LOV, Content LOV, and Material LOV from MDD
        fabric_lov = mdd.lovs.get("Fabric", {}) if mdd else {}
        content_lov = mdd.lovs.get("Content", {}) if mdd else {}
        material_lov = mdd.lovs.get("Material", {}) if mdd else {}
        
        num_workers = min(8, max(1, total_rows))
        task_args   = []
        for row in rows:
            prod_code = _s(row.get("Product Code"))
            colour_code = _s(row.get("Colour Code"))
            key = f"{prod_code}|{colour_code}"
            catalogue_row = catalogue_lookup.get(key, {})
            
            task_args.append((
                row, brand_lov_id, mdd, args.season, sea_year,
                getattr(args, "country_code", ""),
                getattr(args, "article_type_from_filename", ""),
                catalogue_row,
                fabric_lov,
                content_lov,
                material_lov,
                _brand_type,
                _brand_category
            ))
        
        ordered: list[tuple[int, dict, list]] = []

        log.info("Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
                 total_rows, num_workers)

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_article, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx           = futures[fut]
                mapped, warns = fut.result()
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered]
        del ordered

        log.info("Pass 1 done — mapped=%d articles", len(mapped_articles))

        log.info("Pass 2/2 — streaming XML to %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cls_el  = build_classifications(
            args.brand, brand_lov_id, args.comp_code, args.sbu, args.season,
        )
        cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
        del cls_el

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
            del cls_str

            f.write("  <Products>\n")
            for art in mapped_articles:
                if not art.get("article_no"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, brand_lov_id,
                    args.comp_code, args.sbu, season_id, mdd=mdd,
                )
                product_xml = _XMLNS_RE.sub("", product_xml)
                f.write(f"    {product_xml}\n")
                del product_xml
                written_count += 1

            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        del mapped_articles

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%dKB)", out_path, file_kb)

        print("═══ ARTICLE SUMMARY (2XU PRODUCT BIBLE) ════════════════", flush=True)
        print(f"  Product Bible rows  : {total_rows}",   flush=True)
        print(f"  Articles written    : {written_count}", flush=True)
        print(f"  XML file size       : {file_kb}KB",     flush=True)
        print("════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    rpt_path = LOG_DIR / f"validation_productbible_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — CLI (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — 2XU Product Bible v1.0"
    )
    p.add_argument("--brand",        default="2XU")
    p.add_argument("--brand-code",   default="2XU")
    p.add_argument("--comp-code",    default="0888")
    p.add_argument("--sbu",          default="SP")
    p.add_argument("--season",       default="H127")
    p.add_argument("--seq",          default=1, type=int)
    p.add_argument("--country-code", default="AU")
    args = p.parse_args()
    args.article_type_from_filename = ""
    run(args)


if __name__ == "__main__":
    main()
