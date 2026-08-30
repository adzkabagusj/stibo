"""
╔══════════════════════════════════════════════════════════════════╗
║   K-SWISS — GLOBAL LINE ORDER FORM GENERATOR + STIBO XML v1.0    ║
║  Reads brand mapping file dynamically for attribute mappings     ║
║  and MDD file for LOV lookups                                    ║
╚══════════════════════════════════════════════════════════════════╝

Lambda entry point: run(args, auditor)

Source input : "Buy2 update_KSWISS FW26 Global Line Order Form_Lifestyle-
                feedback 23.12.2025.xlsx" (sheets: "Performance Linelist",
                "Lifestyle Linelist")
Brand mapping : "NEW - Brand mapping files Template.xlsx" → sheet
                "KSwiss(Inline)"

Key Features:
- Dynamically loads attribute mappings from "NEW - Brand mapping files
  Template.xlsx" → "KSwiss(Inline)" sheet
- Fetches LOV IDs from the MDD file sheets dynamically
- Supports both "Direct from Principal" and "Mapping from Principal" with
  column mapping (same dynamic engine used by the other Inline brands)
- Each row in the order form is already one Material # + one colorway
  (unlike Pazzion, colors are not split across sub-rows), so one input
  row → one <Product>
"""

from __future__ import annotations

import logging
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
BASE_DIR      = Path(os.environ.get("LAMBDA_TMP_DIR", "/tmp/stibo_workdir_kswiss"))
XML_OUT_DIR   = BASE_DIR / "output" / "xml"
INPUT_DIR     = BASE_DIR / "input"
ORDERFORM_DIR = INPUT_DIR / "orderform"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

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

# Excel formula-error tokens that sometimes leak into cells (e.g. broken
# IMG lookups) — must be treated as empty, never as real values.
_EXCEL_ERROR_TOKENS = {"#VALUE!", "#N/A", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}

# GENDER column → SAP/BY Gender translation (per KSwiss(Inline) mapping:
# "Women : Female / Men : Male / GS : Unisex"). Values outside this table
# (e.g. "Kids", "Grade School", "Preschool" — age categories that leaked
# into the GENDER column of the Lifestyle sheet) are intentionally left
# unmapped rather than guessed.
GENDER_MAP = {
    "WOMEN": "Female",
    "MEN":   "Male",
    "GS":    "Unisex",
}

# BY Age display → SAP Age ID rollup. SAP Age only has three buckets
# (Children="CH" / Adults="AD" / All Ages="AA" — see MDD "Age LOV" cols
# A-C) while BY Age is finer-grained (see cols F-G: ADULT, ALL AGES,
# GRADE SCHOOL, INFANT, KIDS, PRESCHOOL). Anything child-shaped rolls up
# to Children; anything not listed here defaults to Adults.
_BY_AGE_TO_SAP_AGE = {
    "INFANT":       "CH",
    "PRESCHOOL":    "CH",
    "GRADE SCHOOL": "CH",
    "KIDS":         "CH",
    "ALL AGES":     "AA",
    "ADULT":        "AD",
}

# COLLECTION → Sports Category EN, per KSwiss(Inline) mapping note (row
# "Sports Category EN"). Matched by keyword contained in the raw
# Collection value (case-insensitive). Anything not recognized here is
# skipped rather than guessed — see also Sports Category LOV in the MDD.
#
# "Court style" is checked ahead of bare "court" — the mapping note lists
# them as two distinct keywords ("Court" → Tennis/Padel, "Court style" →
# Lifestyle/Casual), and "Court style" contains "court" as a substring, so
# order matters here.
_SPORTS_CATEGORY_KEYWORDS_LIFESTYLE_CASUAL = ("court style", "sport style", "essential", "lifestyle")
_SPORTS_CATEGORY_KEYWORDS_TENNIS_PADEL     = ("court", "tennis", "padel", "performance")

# Attributes to exclude from the *generic* dynamic-mapping loop only —
# each has bespoke handling below because its brand-mapping field name
# doesn't line up 1:1 with a plain column value.
EXCLUDED_ATTRIBUTES = {
    "AT_Gender",     # Needs Women/Men/GS → Female/Male/Unisex translation first
    "AT_BYGender",   # Same translation as AT_Gender
    "AT_Width",      # Derived from the last character of MATERIAL #, not a direct column
    "AT_SportsCategoryEN",  # Keyword-mapped from COLLECTION, not a direct LOV value
    "AT_Franchise",  # Hardcoded below (Model Family, as-is) — the shared
                     # KSwiss(Inline) mapping sheet's "Franchise" row has no
                     # Stibo Attribute ID filled in (VLOOKUP → #N/A) as of
                     # 2026-08-17, so the generic loop can't pick it up.
                     # AT_Franchise is a real, valid MDD attribute (List Of
                     # Values, "Franchise LOV") — bespoke-written here so
                     # this doesn't silently break again if that mapping
                     # sheet regresses.
}

# ══════════════════════════════════════════════════════════════════
# TEST LIMIT
# ══════════════════════════════════════════════════════════════════
# 1-indexed inclusive data-row range for a quick smoke test — set to None
# (or pass --test-row-range on the CLI / args.test_row_range) for a full
# production run. Same convention as PTP/order_form_usd_fob.py.
TEST_ROW_RANGE: tuple[int, int] | None = None


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — DYNAMIC LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """
    Loads Core Attributes + LOVs from the MDD Excel ("Master Data
    Dictionary (MAA).xlsx" — shared across brands, provided at runtime).
    Dynamically reads all LOV sheets for flexible lookup.
    """

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Load Core Attributes
        if "Core Attributes" in wb.sheetnames:
            ws   = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 1:
                hdr = rows[1]
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
                        "id":          aid,
                        "name":        _v("PIM Attribute Name"),
                        "source_name": _v("Source Attribute Name"),
                        "validation":  _v("Validation Base Type"),
                        "lov_name":    _v("Name of LOV"),
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
        self._load_brand_category_lov(wb)
        self._load_brand_group_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_width_lov(wb)
        self._load_sports_category_lov(wb)

        wb.close()
        log.info("[MDD] %d attributes | %d LOV types loaded", len(self.attributes), len(self.lovs))

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

        sap_age_lov_map = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display = str(row[0]).strip()
            age_lov_id  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            if age_display and age_lov_id:
                sap_age_lov_map[age_display.upper()] = age_lov_id

        by_age_lov_map = {}
        for row in rows[1:]:
            if not row or len(row) < 7:
                continue
            age_display = str(row[6]).strip() if row[6] else ""
            age_lov_id  = str(row[5]).strip() if row[5] else ""
            if age_display and age_lov_id:
                by_age_lov_map[age_display.upper()] = age_lov_id
                if age_display.upper() == "ADULTS":
                    by_age_lov_map["ADULT"] = age_lov_id
                elif age_display.upper() == "ADULT":
                    by_age_lov_map["ADULTS"] = age_lov_id

        self.lovs["AgeLOV"]    = sap_age_lov_map
        self.lovs["AT_SAPAge"] = sap_age_lov_map
        self.lovs["AT_BYAge"]  = by_age_lov_map
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

        self.lovs["GenderLOV"]  = gender_lov
        self.lovs["AT_Gender"]  = gender_lov
        self.lovs["AT_BYGender"] = gender_lov
        log.info("[MDD] Gender LOV loaded — %d entries", len(gender_lov))

    def _load_size_lov(self, wb):
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
        self.lovs["AT_Size"]     = size_lov
        log.info("[MDD] Size Code LOV loaded — %d entries", len(size_lov))

    def _load_color_lov(self, wb):
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
        self.lovs["AT_Color"]     = color_lov
        log.info("[MDD] Color Code LOV loaded — %d entries", len(color_lov))

    def _load_country_lov(self, wb):
        """Col A = ID (e.g. KH, ID, SG), Col B = Country Name."""
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

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        country_lov: dict[str, str] = {}
        country_id_to_name: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            country_id   = str(row[0]).strip() if row[0] else ""
            country_name = str(row[1]).strip() if row[1] else ""
            if country_id.upper() in ("ID", "CODE") and country_name.upper() in ("VALUE", "NAME", "COUNTRY"):
                continue
            if country_id and country_name:
                country_id_to_name[country_id.upper()] = country_name
                country_lov[country_name.upper()] = country_id

        self.lovs["CountryLOV"]      = country_lov
        self.lovs["AT_Country"]      = country_lov
        self.lovs["CountryIDToName"] = country_id_to_name
        log.info("[MDD] Country LOV loaded — %d entries", len(country_id_to_name))

    def _load_country_origin_lov(self, wb):
        """Col A = LOV ID, Col B = LOV value."""
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
        sheet_name = next((s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()
                            and "GROUP" not in s.upper() and "TYPE" not in s.upper() and "STATUS" not in s.upper()
                            and "CATEGORY" not in s.upper()), None)
        if not sheet_name:
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        brand_lov: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                brand_lov[display.upper()] = lov_id
        self.lovs["BrandLOV"] = brand_lov
        self.lovs["AT_Brand"] = brand_lov
        log.info("[MDD] Brand LOV loaded — %d entries", len(brand_lov))

    def _load_brand_type_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()), None,
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

    def _load_brand_category_lov(self, wb):
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
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper()), None,
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

    def _load_brand_status_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper()), None,
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

    def _load_width_lov(self, wb):
        """Width LOV sheet: Col A = Code (e.g. REG), Col B = Width name (e.g. Regular)."""
        sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == "WIDTH LOV"), None)
        if not sheet_name:
            log.warning("[MDD] Width LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            code = str(row[0]).strip() if row[0] else ""
            name = str(row[1]).strip() if row[1] else ""
            if code and name:
                lov_map[name.upper()] = code
        self.lovs["WidthLOV"] = lov_map
        self.lovs["AT_Width"] = lov_map
        log.info("[MDD] Width LOV loaded — %d entries", len(lov_map))

    def _load_sports_category_lov(self, wb):
        """Sports Category LOV sheet: Col A = Name, Col B = ID."""
        sheet_name = next((s for s in wb.sheetnames if "SPORTS" in s.upper() and "CATEGORY" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Sports Category LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        lov_map: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] is not None else ""
            if display and lov_id:
                lov_map[display.upper()] = lov_id
        self.lovs["SportsCategoryLOV"]  = lov_map
        self.lovs["AT_SportsCategoryEN"] = lov_map
        log.info("[MDD] Sports Category LOV loaded — %d entries", len(lov_map))

    def _load_named_lov_sheets(self, wb):
        """Load any sheet with 'LOV' in name that hasn't been loaded yet."""
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            if any(kw in sn.upper() for kw in [
                "AGE", "GENDER", "SIZE CODE", "COLOR CODE", "COUNTRY", "BRAND",
                "SIMPLE", "WIDTH", "SPORTS CATEGORY",
            ]):
                continue  # Already loaded by a dedicated method above

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
        """Get LOV ID for a given attribute and display value."""
        if not display_value:
            return ""
        search_val = str(display_value).strip().upper()

        if attr_id in self.lovs and search_val in self.lovs[attr_id]:
            return self.lovs[attr_id][search_val]

        attr_info = self.attributes.get(attr_id, {})
        lov_name  = attr_info.get("lov_name")
        if lov_name and lov_name in self.lovs and search_val in self.lovs[lov_name]:
            return self.lovs[lov_name][search_val]

        for lov_map in self.lovs.values():
            if search_val in lov_map:
                return lov_map[search_val]

        return display_value

    def get_country_name_by_id(self, country_id: str) -> str:
        if not country_id:
            return ""
        search_id = country_id.strip().upper()
        if "CountryIDToName" in self.lovs and search_id in self.lovs["CountryIDToName"]:
            return self.lovs["CountryIDToName"][search_id]
        return country_id


class BrandMappingLoader:
    """
    Loads the "NEW - Brand mapping files Template.xlsx" file.
    Reads the "KSwiss(Inline)" sheet to get dynamic attribute mappings.

    Column Layout (same structure across all brand mapping sheets):
    - A: Stibo Attribute Name (display name)
    - B: Stibo Attribute ID (e.g., AT_PrincipalStyleCode)
    - C: Stibo Validation (LOV, Text, Number, etc.)
    - G: Field Mapping Type
    - J: Field Name in the Brand File
    - K: Mapping Logic (transformations like "last 3 digits")

    Inclusion Rules (Col G):
    - "Direct from Principal" → INCLUDE
    - "Mapping from Principal" → INCLUDE
    - "Formula in system" + has field name (Col J) → INCLUDE
    - Otherwise → SKIP (system-level / manual-input / AI-analysis only)
    """

    # KSwiss(Inline) sheet has one known field-name typo: "Mapping Family"
    # should read "Model Family" (confirmed by the sheet's own "Inline"
    # notes column, which explicitly says "Inline : Model Family").
    _FIELD_NAME_ALIASES = {
        "MAPPING FAMILY": "Model Family",
    }

    @staticmethod
    def _clean_field_name(actual_field: str) -> str:
        """
        A couple of cells in "Field Name in the Brand File" append a
        free-text note after a newline (e.g. AT_PrincipalSize's field is
        literally 'Size Range\\nSize range is following "no fill color"
        area'). The real column name is always just the first line.
        """
        if "\n" in actual_field:
            actual_field = actual_field.split("\n", 1)[0].strip()
        return actual_field

    def __init__(self, path: Path, sheet_name: str = "KSwiss(Inline)"):
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

        ws   = wb[self.sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Row 2 (index 1) is the header; data starts at row 3 (index 2)
        for row_idx in range(2, len(rows)):
            row = rows[row_idx]
            if not row or len(row) < 10:
                continue

            col_a = row[0]  # Stibo Attribute Name
            col_b = row[1]  # Stibo Attribute ID
            col_c = row[2]  # Stibo Validation
            col_g = row[6]  # Field Mapping Type
            col_j = row[9]  # Field Name in the Brand File

            if not col_b or str(col_b).startswith("#N/A"):
                continue

            col_g_str = str(col_g).strip() if col_g else ""
            col_j_str = str(col_j).strip() if col_j else ""
            col_k_str = str(row[10]).strip() if len(row) > 10 and row[10] else ""

            col_g_lower = col_g_str.lower()

            include = False
            if "direct from principal" in col_g_lower:
                include = True
            elif "formula in system" in col_g_lower and col_j_str:
                include = True
            elif "mapping from" in col_g_lower and "principal" in col_g_lower and col_j_str:
                include = True

            if include:
                actual_field = col_j_str
                if ':' in col_j_str:
                    parts = col_j_str.split(':')
                    if len(parts) > 1:
                        actual_field = parts[-1].strip()

                actual_field = self._clean_field_name(actual_field)
                actual_field = self._FIELD_NAME_ALIASES.get(actual_field.upper(), actual_field)

                mapping = {
                    'attribute_name': str(col_a).strip() if col_a else "",
                    'attribute_id':   str(col_b).strip(),
                    'value_type':     str(col_c).strip().upper() if col_c else "TEXT",
                    'mapping_type':   col_g_str,
                    'field_name':     actual_field,
                    'mapping_logic':  col_k_str,
                }
                self.mappings.append(mapping)

        wb.close()
        log.info("[BrandMapping] Loaded %d attribute mappings", len(self.mappings))
        for i, m in enumerate(self.mappings[:5]):
            log.info("  [%d] %s ← %s (Type: %s)", i + 1, m['attribute_id'], m['field_name'], m['value_type'])

    def get_all_mappings(self) -> list[dict]:
        return self.mappings


class RNALoader:
    """
    Loads RNA (Source Mapping) data to resolve Brand Type, Brand Category,
    Brand Group and Brand Status. Reads from "Source Mapping related RNA"
    sheet in the brand mapping file (shared across brands).
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
            (s for s in wb.sheetnames if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)), None,
        )
        if not sheet_name:
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found in %s", self.path.name)
            wb.close()
            return

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

        def _find(*candidates):
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
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""})

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
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — ORDER FORM LOADER
# ══════════════════════════════════════════════════════════════════

class KSwissOrderFormLoader:
    """
    Loads the K-Swiss "Global Line Order Form" Excel file.

    Reads the "Performance Linelist" and "Lifestyle Linelist" sheets
    (and any other sheet that happens to share the same header layout —
    detection is dynamic, not sheet-name-hardcoded), auto-detecting the
    header row per sheet.

    Unlike Pazzion's order form, colors are NOT split across sub-rows —
    each data row is already exactly one MATERIAL # + one colorway, so no
    post-processing / color-grouping step is required. The sheets do,
    however, repeat the header row mid-table wherever a new size-run /
    category block starts (e.g. Kids block), so those repeated header
    rows must be filtered out before use.
    """

    HEADER_ANCHORS = {"material #", "model name", "color description", "gender"}

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheets_loaded: list[str] = []
        self._load()

    def _find_header_row(self, ws) -> int | None:
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            cells   = [str(v).strip().lower() for v in row if v and str(v).strip()]
            matches = sum(1 for anchor in self.HEADER_ANCHORS if anchor in cells)
            if matches >= 3:
                return i
        return None

    def _load(self):
        log.info("[OrderForm] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[OrderForm] Found %d sheets: %s", len(wb.sheetnames), wb.sheetnames)

        all_dfs = []

        for sheet_name in wb.sheetnames:
            try:
                ws      = wb[sheet_name]
                hdr_idx = self._find_header_row(ws)

                if hdr_idx is None:
                    log.info("[OrderForm] Sheet '%s': no linelist header found — skipping", sheet_name)
                    continue

                log.info("[OrderForm] Sheet '%s': header at row %d (0-indexed)", sheet_name, hdr_idx)

                df_sheet = pd.read_excel(self.path, sheet_name=sheet_name, header=hdr_idx)
                df_sheet.dropna(how="all", inplace=True)
                df_sheet.columns = [str(c) if not isinstance(c, str) else c for c in df_sheet.columns]

                if df_sheet.empty:
                    log.warning("[OrderForm] Sheet '%s' has no data rows after header", sheet_name)
                    continue

                df_sheet["_source_sheet"] = sheet_name

                # Drop repeated header rows (e.g. a new "Kids" block restates
                # the header mid-sheet) and fully-blank article rows.
                material_col = next(
                    (c for c in df_sheet.columns if str(c).strip().upper() == "MATERIAL #"), None,
                )
                if material_col is not None:
                    df_sheet = df_sheet[
                        df_sheet[material_col].notna()
                        & (df_sheet[material_col].astype(str).str.strip().str.upper() != "MATERIAL #")
                    ]

                df_sheet = df_sheet.loc[:, ~df_sheet.columns.astype(str).str.startswith("Unnamed:")]

                if df_sheet.empty:
                    log.warning("[OrderForm] Sheet '%s': no article rows remain after filtering", sheet_name)
                    continue

                all_dfs.append(df_sheet)
                self.sheets_loaded.append(sheet_name)
                log.info("[OrderForm] Sheet '%s': %d rows loaded, columns: %s",
                         sheet_name, len(df_sheet), list(df_sheet.columns[:10]))

            except Exception as e:
                log.error("[OrderForm] Failed to read sheet '%s': %s", sheet_name, e)

        wb.close()

        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            log.info("[OrderForm] Combined total: %d rows from %d sheets", len(self.df), len(self.sheets_loaded))
        else:
            log.warning("[OrderForm] No data loaded from any sheet")


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element | None:
    """Append <Value AttributeID="..."> to parent."""
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip()  if value  else ""
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
    """Convert any value to a clean string; skip NaN/None/Excel error tokens.

    Whole-number floats (e.g. 57.0) are rendered without the trailing
    ".0" — pandas upcasts a column to float64 the moment ANY row across
    the concatenated sheets has a NaN/decimal in it, even though the
    source Excel cell itself held a clean int (57). Left alone, that
    silently violates "populated from principal data as is" for
    fields like AT_FOB/DP.
    """
    if v is None:
        return ""
    import math
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    if s.upper() in _EXCEL_ERROR_TOKENS:
        return ""
    return s


def _get_field(row_data: dict, field_name: str) -> str:
    """
    Case-insensitive column lookup against row_data.
    Tries exact match first, then falls back to a normalized (whitespace-
    and case-insensitive) comparison — needed because the Performance
    sheet uses "COLLECTION" while the Lifestyle sheet uses "Collection".
    """
    v = row_data.get(field_name)
    if v is not None:
        s = _safe_str(v)
        if s:
            return s

    def normalize(s):
        return " ".join(str(s).replace("\n", " ").split()).upper()

    field_normalized = normalize(field_name)
    for k, val in row_data.items():
        if normalize(k) == field_normalized:
            s = _safe_str(val)
            if s:
                return s
    return ""


def _resolve_currency_from_country(country_code: str) -> str:
    """Resolve Retail Price Currency LOV ID from country code."""
    country_to_currency = {
        "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD",
        "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


def _derive_season(season_code: str) -> tuple[str, str]:
    """'FW26' → ('FW', '2026'). Handles bare prefix or already 4-digit year."""
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
    """Build classification ID like 'CLH_KSW_FW2026'."""
    pfx, yr = _derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"

def _normalize_brand_display(brand: str) -> str:
    """
    Normalize brand token used inside Classification IDs (e.g. CLH_{brand}Batches,
    CLH_{brand}Articles) so K SWISS always renders as 'KSwiss', regardless of the
    casing/spacing passed in via args.brand (e.g. 'Kswiss', 'K_SWISS', 'K SWISS').
    """
    key = re.sub(r"[^A-Z0-9]", "", (brand or "").upper())
    if key == "KSWISS":
        return "KSwiss"
    return brand

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    brand_display = _normalize_brand_display(brand)

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand_display}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
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

def _determine_parent_id(row_data: dict, source_sheet: str) -> str:
    """K-Swiss Global Line Order Form is footwear-only (Performance +
    Lifestyle linelists), so default to PPH_F-TempSubCat."""
    sheet = (source_sheet or "").lower()
    if "acc" in sheet or "accessories" in sheet:
        return "PPH_A-TempSubCat"
    if "app" in sheet or "apparel" in sheet:
        return "PPH_AP-TempSubCat"
    return "PPH_F-TempSubCat"


def _resolve_gender(raw_gender: str) -> str:
    """Women→Female, Men→Male, GS→Unisex. Anything else → '' (unmapped)."""
    if not raw_gender:
        return ""
    return GENDER_MAP.get(raw_gender.strip().upper(), "")


def _resolve_age(raw_gender: str, mdd: MDDLoader) -> tuple[str, str]:
    """
    Resolve (by_age_id, sap_age_id) for a row.

    The Lifestyle Linelist's "Kids" block doesn't put a gender in the
    GENDER column at all — it puts an age value there instead (Grade
    School / Preschool / Kids), which is confirmed by cross-checking the
    MDD's own BY Age LOV: its display values (ADULT / ALL AGES / GRADE
    SCHOOL / INFANT / KIDS / PRESCHOOL) match those leaked values
    verbatim. So rather than defaulting every row to Adults, look the raw
    GENDER cell up against the BY Age LOV first; only fall back to the
    documented "Default: Adults" (KSwiss(Inline) mapping) when it isn't
    an age value (i.e. it's a real Women/Men/GS gender row).

    SAP Age only has three buckets (Children / Adults / All Ages), so the
    finer-grained BY Age categories roll up into them.
    """
    by_age_lov = mdd.lovs.get("AT_BYAge", {})
    search = (raw_gender or "").strip().upper()

    if search in by_age_lov:
        by_age_display = search
        by_age_id       = by_age_lov[search]
    else:
        by_age_display = "ADULT"
        by_age_id       = by_age_lov.get("ADULT", "ADULT")

    sap_age_id = _BY_AGE_TO_SAP_AGE.get(by_age_display, "AD")
    return by_age_id, sap_age_id


def _resolve_width(material_number: str) -> str:
    """
    Last character of MATERIAL # → Regular/Wide, per KSwiss(Inline)
    mapping note: "Mapping from Last digit material number: M: Regular,
    W: Wide". Anything else (no trailing M/W) → '' (unmapped).
    """
    if not material_number:
        return ""
    last = material_number.strip()[-1:].upper()
    if last == "M":
        return "Regular"
    if last == "W":
        return "Wide"
    return ""


def _resolve_sports_category(collection_raw: str) -> str:
    """
    COLLECTION → Sports Category EN, per KSwiss(Inline) mapping note.
    Keyword-matched (case-insensitive); returns '' if no keyword matches
    rather than guessing.
    """
    if not collection_raw:
        return ""
    c = collection_raw.strip().lower()
    # Lifestyle/Casual checked first: "court style" (Lifestyle/Casual) must
    # win over the bare "court" keyword (Tennis/Padel) it contains.
    if any(kw in c for kw in _SPORTS_CATEGORY_KEYWORDS_LIFESTYLE_CASUAL):
        return "Lifestyle / Casual"
    if any(kw in c for kw in _SPORTS_CATEGORY_KEYWORDS_TENNIS_PADEL):
        return "Tennis / Padel"
    return ""


def build_product_xml(
    row_data:       dict,
    brand:          str,
    brand_code:     str,
    comp_code:      str,
    sbu:            str,
    season_id:      str,
    season:         str,
    country_code:   str,
    brand_type:     str,
    brand_category: str,
    brand_group:    str,
    brand_status:   str,
    brand_mapping:  BrandMappingLoader,
    mdd:            MDDLoader,
) -> str:
    """
    Build one <Product> XML string:
      • UserTypeID = PRD_GenericArticle
      • ParentID   = PPH_F-TempSubCat
      • KeyValue   KEY_InboundArticle
      • Name
      • ClassificationReference  ×2  (Merchandiser + UnConfirmedForSeason)
      • Values  (dynamic from brand mapping + fixed system-level values)

    One input row == one Product — MATERIAL # already identifies a single
    colorway (unlike Pazzion, colors are not grouped across sub-rows).
    """
    article_code = _safe_str(row_data.get("MATERIAL #") or row_data.get("Material #"))
    source_sheet = _safe_str(row_data.get("_source_sheet", ""))

    if not article_code:
        return ""  # skip rows with no identifying code

    generic_code = f"{brand_code[:3].upper()}{article_code.upper()}"
    parent_id    = _determine_parent_id(row_data, source_sheet)
    model_name   = _get_field(row_data, "Model Name")

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = model_name

    # ── Classification references ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{_normalize_brand_display(brand)}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    row_data["_generic_code"] = generic_code

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_product_values(
        vals_el, row_data, brand, brand_code, comp_code, sbu, season,
        country_code, brand_type, brand_category, brand_group, brand_status,
        brand_mapping, mdd,
    )

    return ET.tostring(g_el, encoding="unicode")


def _add_product_values(
    vals_el:        ET.Element,
    row_data:       dict,
    brand:          str,
    brand_code:     str,
    comp_code:      str,
    sbu:            str,
    season:         str,
    country_code:   str,
    brand_type:     str,
    brand_category: str,
    brand_group:    str,
    brand_status:   str,
    brand_mapping:  BrandMappingLoader,
    mdd:            MDDLoader,
) -> None:
    """
    Write all <Value> elements.

    Values are RESOLVED dynamically (from the KSwiss(Inline) brand mapping
    sheet + a handful of bespoke rules for Gender/Width/Sports Category EN
    whose source data doesn't line up 1:1 with the generic engine), but
    WRITTEN in the same fixed sequence Birkenstock's reference generator
    uses (system/org → brand classification → currencies → SAP/article
    flags → BCI → age → gender → country of origin → product
    characteristics → country size/article status → merchandise hierarchy
    → remaining principal-data fields), so the two brands' XML files read
    the same way attribute-for-attribute. See
    birkenstock/OFS_FC_Mitra_ss27.py's per-article Values block for the
    reference ordering this mirrors.
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

    # ── Resolve every KSwiss(Inline)-driven value up front (order-
    #    independent) — writing happens later, in the fixed sequence
    #    below, not in whatever order the mapping sheet rows happen to
    #    be in. Keeps the "read the sheet dynamically" design intact
    #    while still producing a deterministic, Birkenstock-matching
    #    attribute order.
    dynamic: dict[str, dict] = {}
    for mapping in brand_mapping.get_all_mappings():
        attr_id    = mapping["attribute_id"]
        field_name = mapping["field_name"]
        value_type = mapping["value_type"]

        if attr_id in EXCLUDED_ATTRIBUTES or not field_name or attr_id in dynamic:
            continue

        raw_value = _get_field(row_data, field_name)
        if not raw_value:
            continue

        if value_type == "LOV":
            lov_id = mdd.get_lov_id(attr_id, raw_value)
            if lov_id:
                dynamic[attr_id] = {"id_val": lov_id}
        else:
            dynamic[attr_id] = {"value": raw_value}

    def _w_dynamic(attr_id: str):
        """Write a previously-resolved dynamic value, if any."""
        v = dynamic.get(attr_id)
        if v:
            _w(attr_id, **v)

    # Bespoke: Gender — Women→Female, Men→Male, GS→Unisex
    gender_raw = _get_field(row_data, "GENDER") or _get_field(row_data, "Gender")
    gender_resolved = _resolve_gender(gender_raw)
    gender_lov_id = ""
    if gender_resolved:
        gender_lov_id = mdd.lovs.get("AT_Gender", {}).get(gender_resolved.upper(), "")
    elif gender_raw:
        log.debug("[Gender] Unmapped GENDER value '%s' — AT_Gender/AT_BYGender left unset", gender_raw)

    # AT_Width is intentionally left blank for Inline (per request) — not
    # derived/written at all. _resolve_width() is kept below (unused) in
    # case Width is reinstated later.

    # Bespoke: Sports Category EN — keyword-mapped from COLLECTION
    collection_raw = _get_field(row_data, "COLLECTION") or _get_field(row_data, "Collection")
    sports_category = _resolve_sports_category(collection_raw)
    sports_cat_lov_id = mdd.lovs.get("AT_SportsCategoryEN", {}).get(sports_category.upper(), "") if sports_category else ""
    if collection_raw and not sports_category:
        log.debug("[Sports Category] No keyword match for COLLECTION='%s' — AT_SportsCategoryEN left unset", collection_raw)

    # Bespoke: Franchise — Model Family, populated as-is (per KSwiss(Inline)
    # mapping note: "Direct from Principal" ← "Model Family"). Hardcoded
    # rather than routed through the mapping sheet — see EXCLUDED_ATTRIBUTES.
    franchise_raw = _get_field(row_data, "Model Family") or _get_field(row_data, "MODEL FAMILY")

    # ══════════════════════════════════════════════════════════════
    # 1. System / org
    # ══════════════════════════════════════════════════════════════
    generic_code = _safe_str(row_data.get("_generic_code", ""))
    if generic_code:
        _w("AT_InboundGenericCode", value=generic_code)
    _mw("AT_CompanyCode", comp_code)
    _mw("AT_SBU",        sbu)
    if country_code:
        _w("AT_Country", id_val=country_code)
    _w("AT_Brand", id_val=brand_code)

    sea_raw    = (season or "").strip().upper()
    sea_prefix = sea_raw[:2] if len(sea_raw) >= 2 else sea_raw
    sea_label  = LOV_SEASON.get(sea_prefix, sea_prefix)
    _w("AT_Season", sea_label, id_val=sea_prefix)

    # ══════════════════════════════════════════════════════════════
    # 2. Brand Group / Type / Category / Status (RNA)
    # ══════════════════════════════════════════════════════════════
    if brand_group:
        bg_lov_id = mdd.lovs.get("AT_BrandGroup", {}).get(brand_group.upper(), "")
        _w("AT_BrandGroup", id_val=bg_lov_id) if bg_lov_id else _w("AT_BrandGroup", value=brand_group)

    if brand_type:
        bt_lov_id = mdd.lovs.get("AT_BrandType", {}).get(brand_type.upper(), "")
        _w("AT_BrandType", id_val=bt_lov_id) if bt_lov_id else _w("AT_BrandType", value=brand_type)

    if brand_category:
        bc_lov_id = mdd.lovs.get("AT_BrandCategory", {}).get(brand_category.upper(), "")
        _w("AT_BrandCategory", id_val=bc_lov_id) if bc_lov_id else _w("AT_BrandCategory", value=brand_category)

    if brand_status:
        bs_lov_id = mdd.lovs.get("AT_BrandStatus", {}).get(brand_status.upper(), "")
        _w("AT_BrandStatus", id_val=bs_lov_id) if bs_lov_id else _w("AT_BrandStatus", value=brand_status)

    # ══════════════════════════════════════════════════════════════
    # 3. Currencies
    # ══════════════════════════════════════════════════════════════
    _w("AT_FOBCurrency", id_val="USD")   # Default: USD per KSwiss(Inline) mapping
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _w("AT_RetailPriceCurrency", id_val=currency_lov_id)

    # ══════════════════════════════════════════════════════════════
    # 4. SAP / article flags
    # ══════════════════════════════════════════════════════════════
    _w("AT_SAPProductFlag",     id_val="A")     # Intercompany
    _w("AT_MaterialType",       id_val="ZINA")  # Intercompany
    _w("AT_SAPArticleCategory", id_val="1")     # Generic / Variant
    _w("AT_UOM",                id_val="EA")
    _w("AT_ImagesSource",       id_val="PHO")


    # ══════════════════════════════════════════════════════════════
    # 6. BY Age / SAP Age — default Adults, but honor the age data that
    #    leaks into the GENDER column for Kids-block rows (Grade School /
    #    Preschool / Kids) instead of blindly defaulting every row.
    # ══════════════════════════════════════════════════════════════
    by_age_id, sap_age_id = _resolve_age(gender_raw, mdd)
    _w("AT_BYAge",  id_val=by_age_id)
    _w("AT_SAPAge", id_val=sap_age_id)

    # ══════════════════════════════════════════════════════════════
    # 7. BY Gender / Gender / Principal Gender Description
    # ══════════════════════════════════════════════════════════════
    if gender_lov_id:
        _w("AT_BYGender", id_val=gender_lov_id)
        _w("AT_Gender",   id_val=gender_lov_id)
    _w_dynamic("AT_PrincipalGenderDescription")

    # ══════════════════════════════════════════════════════════════
    # 8. Country of Origin
    # ══════════════════════════════════════════════════════════════
    _w_dynamic("AT_CountryOrigin")

    # ══════════════════════════════════════════════════════════════
    # 9. Product characteristics (Sports Category EN — K-Swiss'
    #    equivalent of Birkenstock's Material/PatternPrint/Silhouette slot)
    #    NOTE: AT_Width is intentionally NOT written for Inline (per
    #    request) — see _resolve_width() above, now unused.
    # ══════════════════════════════════════════════════════════════
    _w("AT_PricingDistributionChannel", id_val="01")   # Retailer

    if sports_category:
        _w("AT_SportsCategoryEN", id_val=sports_cat_lov_id) if sports_cat_lov_id else _w("AT_SportsCategoryEN", value=sports_category)

    # ══════════════════════════════════════════════════════════════
    # 10. Country Size / Article Status (footwear-only, per Birkenstock's
    #     "is_footwear" gating — K-Swiss Global Line is footwear-only)
    # ══════════════════════════════════════════════════════════════
    _w("AT_CountrySize", id_val="EU")
    _w("AT_ArticleStatus", id_val="A")

    # K-Swiss-specific defaults with no Birkenstock analog in this slot
    _w("AT_EComAgesCategory", id_val="18+Y")
    
    # Set BY Article Type default
    _w_dynamic("AT_BYArticleType")
    if "AT_BYArticleType" not in written:
        _w("AT_BYArticleType", value="Inline")

    # ══════════════════════════════════════════════════════════════
    # 11. Principal Merchandise Hierarchy L1–L2
    # ══════════════════════════════════════════════════════════════
    _w_dynamic("AT_PrincipalMerchandiseHierarchyL1")
    _w_dynamic("AT_PrincipalMerchandiseHierarchyL2")

    # Bespoke: Franchise — Model Family, as-is (see resolution above).
    # Written as free text, not ID — Franchise LOV is an open ~1,918-value
    # list with no separate code column, so there's nothing to look up.
    if franchise_raw:
        _w("AT_Franchise", value=franchise_raw)

    # ══════════════════════════════════════════════════════════════
    # 12. Remaining principal-data fields (style code/description/color/
    #     size/FOB) + Season Year — same trailing position as Birkenstock's
    #     StyleCode/StyleDescription/ColorName/Size/FOB/SeasonYear block.
    # ══════════════════════════════════════════════════════════════
    for attr_id in (
        "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
        "AT_PrincipalColorName", "AT_PrincipalSize", "AT_FOB",
    ):
        _w_dynamic(attr_id)

    season_year = ""
    if len(sea_raw) >= 4:
        year_part = sea_raw[2:]
        if len(year_part) == 2 and year_part.isdigit():
            season_year = f"20{year_part}"
        elif len(year_part) == 4 and year_part.isdigit():
            season_year = year_part
    if season_year:
        _w("AT_SeasonYear", value=season_year)

    # ══════════════════════════════════════════════════════════════
    # 13. Catch-all — anything else the KSwiss(Inline) sheet maps that
    #     wasn't explicitly placed above (forward-compatible with future
    #     mapping-sheet additions; empty today).
    # ══════════════════════════════════════════════════════════════
    for attr_id in dynamic:
        _w_dynamic(attr_id)


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — MAIN RUN FUNCTION
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """Main entry point for K-Swiss Global Line Order Form processing."""
    log.info("[K-Swiss Order Form] Starting ...")

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

    mdd           = MDDLoader(mdd_file)
    brand_mapping = BrandMappingLoader(brand_mapping_file)
    rna           = RNALoader(brand_mapping_file)
    orderform     = KSwissOrderFormLoader(orderform_file)

    if orderform.df.empty:
        log.warning("[K-Swiss Order Form] No data rows found")
        return

    # ── Pipeline args ────────────────────────────────────────────
    brand      = getattr(args, "brand",      "KSwiss")
    brand_code = getattr(args, "brand_code", "KSW").upper()
    comp_code  = getattr(args, "comp_code",  "0888")
    sbu        = getattr(args, "sbu",        "SP").upper()
    season     = getattr(args, "season",     "FW26").upper()

    # ── Country code: try to parse a 2-letter segment from filename,
    #    else fall back to args / default (K-Swiss's order form filename
    #    doesn't embed a country code, so this normally falls through) ──
    _fname_stem  = orderform_file.stem
    _fname_parts = _fname_stem.replace(" ", "-").split("-")
    _country_from_file = ""
    for _part in reversed(_fname_parts):
        _p = _part.strip().upper()
        if len(_p) == 2 and _p.isalpha():
            _country_from_file = _p
            break

    if _country_from_file:
        log.info("[K-Swiss] Country code parsed from filename: '%s'", _country_from_file)
        country_code = _country_from_file
    else:
        country_code = getattr(args, "country_code", "ID").upper()
        log.info("[K-Swiss] Country code from args (not found in filename): '%s'", country_code)

    log.info(
        "Pipeline → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  country=%s",
        brand, brand_code, comp_code, sbu, season, country_code,
    )

    # ── Resolve Brand Type / Category / Group / Status from RNA ────
    country_name = mdd.get_country_name_by_id(country_code)

    rna_result = rna.get(country_name, comp_code, sbu, brand_code)
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        try:
            comp_code_stripped = str(int(comp_code))
        except (ValueError, TypeError):
            comp_code_stripped = comp_code
        rna_result = rna.get(country_name, comp_code_stripped, sbu, brand_code)
    if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
        rna_result = rna.get_fuzzy(country_name, sbu, brand_code)

    brand_type     = rna_result.get("brand_type",     "")
    brand_category = rna_result.get("brand_category", "")
    brand_group    = rna_result.get("brand_group",    "")
    brand_status    = rna_result.get("brand_status",   "")

    if brand_type or brand_category:
        log.info("[RNA] ✓ Resolved — brand_type='%s'  brand_category='%s'  brand_group='%s'  brand_status='%s'",
                 brand_type, brand_category, brand_group, brand_status)
    else:
        log.warning("[RNA] ✗ All passes failed — brand_type/brand_category will be EMPTY for country='%s' comp='%s' sbu='%s' brand='%s'",
                    country_name, comp_code, sbu, brand_code)

    season_id = _season_id_full(brand_code, season)

    log.info("[K-Swiss Order Form] Sheets loaded: %s", orderform.sheets_loaded)
    log.info("[K-Swiss Order Form] Total rows before filtering: %d", len(orderform.df))

    filtered_df = orderform.df[
        orderform.df.apply(
            lambda row: bool(_safe_str(row.get("MATERIAL #") or row.get("Material #"))),
            axis=1,
        )
    ].copy()

    log.info("[K-Swiss Order Form] Total rows to process after filtering: %d", len(filtered_df))

    test_row_range = getattr(args, "test_row_range", None) or TEST_ROW_RANGE
    if test_row_range:
        start, end = test_row_range
        filtered_df = filtered_df.iloc[start - 1:end].copy()
        log.warning("[TEST MODE] Limited to data rows %d–%d (%d rows) — clear "
                    "TEST_ROW_RANGE (or drop --test-row-range) for a full production run",
                    start, end, len(filtered_df))

    cls_el  = build_classifications(brand, brand_code, comp_code, sbu, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    del cls_el

    out_name = orderform_file.stem + ".xml"
    out_path = XML_OUT_DIR / out_name
    XML_OUT_DIR.mkdir(parents=True, exist_ok=True)
    export_time   = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    written_count = 0

    log.info("[K-Swiss Order Form] Streaming XML → %s ...", out_name)

    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
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

        f.write(f"{cls_str}\n")
        del cls_str

        f.write("  <Products>\n")

        for _, row in filtered_df.iterrows():
            row_dict = row.to_dict()

            product_xml = build_product_xml(
                row_dict, brand, brand_code, comp_code, sbu,
                season_id, season, country_code, brand_type, brand_category,
                brand_group, brand_status, brand_mapping, mdd,
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

    print("═══ K-SWISS GLOBAL LINE ORDER FORM SUMMARY ═════════════════", flush=True)
    print(f"  Input file      : {orderform_file.name}",               flush=True)
    print(f"  Sheets          : {', '.join(orderform.sheets_loaded)}", flush=True)
    print(f"  Products written: {written_count}",                     flush=True)
    print(f"  Output XML      : {out_path.name}",                     flush=True)
    print(f"  File size       : {file_kb}KB",                         flush=True)
    if test_row_range:
        print(f"  TEST MODE       : ON (rows {test_row_range[0]}–{test_row_range[1]})", flush=True)
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
    parser.add_argument("--brand", default="KSwiss")
    parser.add_argument("--brand-code", dest="brand_code", default="KSW")
    parser.add_argument("--season", default="FW26")
    parser.add_argument("--country-code", dest="country_code", default="ID")
    parser.add_argument("--test-row-range", dest="test_row_range", type=int, nargs=2,
                         metavar=("START", "END"),
                         help="1-indexed inclusive data-row range for a quick smoke test, e.g. --test-row-range 1 10")

    parsed_args = parser.parse_args()
    args = types.SimpleNamespace(**vars(parsed_args))
    if args.test_row_range:
        args.test_row_range = tuple(args.test_row_range)

    run(args)
