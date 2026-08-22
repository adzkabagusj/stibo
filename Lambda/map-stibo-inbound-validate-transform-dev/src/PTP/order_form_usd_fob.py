"""
╔══════════════════════════════════════════════════════════════════╗
║   PTP — DYNAMIC ORDER FORM (USD FOB) GENERATOR + STIBO XML v1.1  ║
║  Reads brand mapping file dynamically for attribute mappings     ║
║  and MDD file for LOV lookups                                    ║
╚══════════════════════════════════════════════════════════════════╝

Lambda entry point: run(args, auditor)

v1.1 change: _add_product_values() attribute WRITE ORDER now matches
Birkenstock's XMLGenerator.build_products() sequence exactly — system/org
fields (InboundGenericCode → CompanyCode/SBU → Country → Brand → Season →
BrandGroup/Type/Category/Status → currencies → SAP flags → UOM → BCI →
Age → Gender → CountryOrigin → PricingDistributionChannel → ArticleStatus)
are written FIRST, and the dynamic/mapped fields (Principal Hierarchy,
style code/description/barcode, sports category, etc.) are written LAST.
Previously PTP wrote dynamic fields first and system fields last — same
data, different <Value> element order in the output XML. No attribute
values changed, only ordering.
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

# force=True: see PTP/lambda_function.py's logging setup comment — without
# it this is a no-op under Lambda (pre-existing root handler at WARNING),
# which silently swallows every INFO log in this file.
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', force=True)
log = logging.getLogger(__name__)

# ── Paths (Lambda-aware) ─────────────────────────────────────────
BASE_DIR      = Path(os.environ.get("LAMBDA_TMP_DIR", "/tmp/stibo_workdir_ptp"))
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

LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter", "SP": "Spring",
    "SM": "Summer",
}

SPORTS_CATEGORY_BY_BRAND = {
    "PTP":  "Fitness / Training",
    "BAHE": "Yoga / Pilates",
}

# ── Reporting brand code (generic/style-code prefix + MDM lookups) ────
# Confirmed with the team (2026-08-19): AT_InboundGenericCode /
# KEY_InboundArticle always use "PTP" as the brand code — BAHE rows
# included, NOT a "BH" prefix (the mapping sheet's "Generic"/"SAP Style
# Code" mapping_logic notes suggested "BAHE BH +Type B.2", but that's
# not what's actually wanted). Same "PTP" value also covers AT_Brand LOV
# / RNA lookups, since BAHE has no Brand LOV entry and no RNA row of its
# own in the Master Data Dictionary / brand mapping sheet either — it
# reports under PTP everywhere. This does NOT affect
# SPORTS_CATEGORY_BY_BRAND, which stays per-brand.
REPORTING_BRAND_CODE_OVERRIDES = {
    "BAHE": "PTP",
}


def _reporting_brand_code(brand_code: str) -> str:
    """Brand code to use for the generic/style-code prefix and for
    AT_Brand LOV / RNA lookups (see REPORTING_BRAND_CODE_OVERRIDES
    above)."""
    return REPORTING_BRAND_CODE_OVERRIDES.get((brand_code or "").strip().upper(), (brand_code or "").strip().upper())

# ── ITEM RANGE → Division letter (PPH_{letter}-TempSubCat) ────────────
# Mirrors Birkenstock's DIVISION_PARENT_MAP / _get_division_code()
# pattern exactly. PTP has no column literally named "Division" (only
# ITEM RANGE / SUB CATEGORY), so this is a keyword-substring match
# against ITEM RANGE. Covers all letters from the confirmed legend:
#   Accessories = E, Footwear = F, Apparel = A, Equipment = Q, Toys = T
# Order matters — more specific keywords first so e.g. "FITNESS
# ACCESSORIES" matches ACCESSORIES (E) rather than falling through.
ITEM_RANGE_DIVISION_KEYWORDS: list[tuple[str, str]] = [
    ("ACCESSOR", "E"),   # ACCESSORIES, FITNESS ACCESSORIES
    ("FOOTWEAR", "F"),
    ("SHOE",     "F"),
    ("APPAREL",  "A"),
    ("CLOTHING", "A"),
    ("TOY",      "T"),
]


def _get_ptp_division_code(item_range: str) -> str:
    """Return the ParentID letter for a PTP ITEM RANGE value, or 'X' if
    unrecognized (same 'X' convention as Birkenstock's _get_division_code
    — an internal signal for the caller to apply a business fallback,
    not a literal value written into the XML)."""
    ir_upper = (item_range or "").strip().upper()
    for keyword, letter in ITEM_RANGE_DIVISION_KEYWORDS:
        if keyword in ir_upper:
            return letter
    if ir_upper in KNOWN_EQUIPMENT_ITEM_RANGES:
        return "Q"
    return "X"


# ITEM RANGE values confirmed (from the actual order form) to be
# legitimate Equipment-division items — routed to Q silently, no
# warning. Anything NOT in this set and NOT matching a keyword above is
# genuinely unrecognized and gets logged (see build_product_xml()), so
# a brand-new ITEM RANGE value in a future file surfaces for review
# instead of being silently misfiled OR silently spamming warnings for
# values we already know are fine.
KNOWN_EQUIPMENT_ITEM_RANGES = {
    "RESISTANCE", "PILATES", "RECOVERY", "YOGA", "CORE", "CARDIO",
    "HOT & COLD THERAPY", "BODYWEIGHT", "POSTURE", "RETAIL STANDS",
}


EXCLUDED_ATTRIBUTES = {
    "AT_SAPStyleCode", "AT_Generic", "AT_GenericDescription",
    "AT_Variant", "AT_VariantDescription", "AT_Material",
    # AT_PrincipalBarcode: dropped from this file entirely — barcode data
    # is now owned by ean_source.py (AT_Barcode / DC_Barcode). Excluded
    # here so the dynamic mapping loop never writes it even if the brand
    # mapping sheet's row for it changes shape.
    "AT_PrincipalBarcode",
}

_SPECIAL_CASED_ATTRIBUTES = {
    "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
    "AT_SportsCategoryEN", "AT_FOB",
}

# ══════════════════════════════════════════════════════════════════
# TEST LIMIT
# ══════════════════════════════════════════════════════════════════
# TEST_ROW_RANGE: tuple[int, int] | None = (230, 250)


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — DYNAMIC LOADERS  (unchanged from v1.0)
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb  = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

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
        self._load_brand_category_lov(wb)
        self._load_brand_status_lov(wb)

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
        self.lovs["AgeLOV"] = sap_age_lov_map
        self.lovs["AT_SAPAge"] = sap_age_lov_map
        self.lovs["AT_BYAge"] = by_age_lov_map
        log.info("[MDD] Age LOV loaded — SAP: %d entries, BY: %d entries", len(sap_age_lov_map), len(by_age_lov_map))

    def _load_gender_lov(self, wb):
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
        sheet_name = next((s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        # Sheet header is ('Value ID of LOV', 'Values of LOV') — ID first,
        # display second, same column order as every other *_LOV sheet in
        # the MDD (e.g. Brand Type/Group/Category/Status use 'Code' then
        # display). Previously read backwards (row[0] as display, row[1]
        # as id), which built an ID→display dict instead of the
        # ID/display→ID dict lookup sites expect — silently broken for
        # any brand whose ID differs from its display name. Keyed by
        # both the ID and the display name (upper-cased) since source
        # files sometimes carry the brand code (e.g. "PTP") and
        # sometimes the display text (e.g. "NEW BALANCE") in their BRAND
        # column.
        brand_lov: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                brand_lov[lov_id.upper()] = lov_id
                brand_lov[display.upper()] = lov_id
        self.lovs["BrandLOV"] = brand_lov
        self.lovs["AT_Brand"] = brand_lov
        log.info("[MDD] Brand LOV loaded — %d entries", len(brand_lov))

    def _load_brand_type_lov(self, wb):
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

    def _load_brand_group_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP"), None)
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

    def _load_brand_status_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS LOV"),
            next((s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet_name:
            sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == "BRAND STATUS"), None)
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

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            if any(kw in sn.upper() for kw in ["AGE", "GENDER", "SIZE CODE", "COLOR CODE", "COUNTRY", "BRAND", "SIMPLE"]):
                continue
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
        if not display_value:
            return ""
        search_val = str(display_value).strip().upper()
        if attr_id in self.lovs:
            if search_val in self.lovs[attr_id]:
                return self.lovs[attr_id][search_val]
        attr_info = self.attributes.get(attr_id, {})
        lov_name = attr_info.get("lov_name")
        if lov_name and lov_name in self.lovs:
            if search_val in self.lovs[lov_name]:
                return self.lovs[lov_name][search_val]
        for lov_map in self.lovs.values():
            if search_val in lov_map:
                return lov_map[search_val]
        return display_value

    def get_country_name_by_id(self, country_id: str) -> str:
        if not country_id:
            return ""
        search_id = country_id.strip().upper()
        if "CountryIDToName" in self.lovs:
            if search_id in self.lovs["CountryIDToName"]:
                result = self.lovs["CountryIDToName"][search_id]
                log.info("[MDD] Country ID '%s' resolved to '%s'", country_id, result)
                return result
            log.warning("[MDD] Country ID '%s' NOT found in CountryIDToName map", country_id)
        else:
            log.warning("[MDD] CountryIDToName map NOT loaded — MDD Country LOV sheet missing?")
        return country_id


class BrandMappingLoader:
    def __init__(self, path: Path, sheet_name: str = "PTP"):
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
        for row_idx in range(2, len(rows)):
            row = rows[row_idx]
            if not row or len(row) < 10:
                continue
            col_a = row[0]
            col_b = row[1]
            col_c = row[2]
            col_g = row[6]
            col_j = row[9]
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
        for i, m in enumerate(self.mappings[:5]):
            log.info("  [%d] %s ← %s (Type: %s)", i + 1, m['attribute_id'], m['field_name'], m['value_type'])

    def get_all_mappings(self) -> list[dict]:
        return self.mappings


class RNALoader:
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

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu        or "").strip().upper(),
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


def _resolve_rna_for_brand(rna: "RNALoader", country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
    log.info("[RNA] Pass 1 — exact match key: ('%s', '%s', '%s', '%s')", country_name, comp_code, sbu, brand_code)
    result = rna.get(country_name, comp_code, sbu, brand_code)
    if not result.get("brand_type") and not result.get("brand_category"):
        try:
            comp_code_stripped = str(int(comp_code))
        except (ValueError, TypeError):
            comp_code_stripped = comp_code
        log.info("[RNA] Pass 2 — stripped comp_code key: ('%s', '%s', '%s', '%s')",
                 country_name, comp_code_stripped, sbu, brand_code)
        result = rna.get(country_name, comp_code_stripped, sbu, brand_code)
    if not result.get("brand_type") and not result.get("brand_category"):
        log.warning("[RNA] Pass 3 — fuzzy (country+sbu+brand), brand_code='%s'", brand_code)
        result = rna.get_fuzzy(country_name, sbu, brand_code)
    if result.get("brand_type") or result.get("brand_category"):
        log.info("[RNA] ✓ brand_code='%s' → brand_type='%s' brand_category='%s' brand_group='%s' brand_status='%s'",
                 brand_code, result.get("brand_type"), result.get("brand_category"),
                 result.get("brand_group"), result.get("brand_status"))
    else:
        log.warning("[RNA] ✗ No RNA match for brand_code='%s' (country='%s' comp='%s' sbu='%s') — "
                    "Brand Type/Category/Group/Status will be EMPTY for this brand's rows",
                    brand_code, country_name, comp_code, sbu)
    return result


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — ORDER FORM LOADER  (unchanged)
# ══════════════════════════════════════════════════════════════════

class PTPOrderFormLoader:
    HEADER_ANCHORS = {"display name", "colour", "upc code", "customer buy price", "hs code"}

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheets_loaded: list[str] = []
        self._load()

    def _find_header_row(self, ws) -> int | None:
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
                df_sheet = pd.read_excel(self.path, sheet_name=sheet_name, header=hdr_idx)
                df_sheet.dropna(how="all", inplace=True)
                df_sheet.columns = [str(c).strip() if not isinstance(c, str) else c.strip() for c in df_sheet.columns]
                df_sheet = df_sheet.loc[:, ~df_sheet.columns.str.startswith("Unnamed:")]
                if df_sheet.empty:
                    log.warning("[OrderForm] Sheet '%s' has no data rows after header", sheet_name)
                    continue
                df_sheet["_source_sheet"] = sheet_name
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


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — XML HELPERS  (unchanged)
# ══════════════════════════════════════════════════════════════════

_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


def _val(parent: ET.Element, attr_id: str, value: str = "", id_val: str = "") -> ET.Element | None:
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
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _safe_str(v) -> str:
    if v is None:
        return ""
    import math
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _get_field(row_data: dict, field_name: str) -> str:
    v = row_data.get(field_name)
    if v is not None:
        return _safe_str(v)

    def normalize(s):
        return " ".join(str(s).replace("\n", " ").split()).upper()

    field_normalized = normalize(field_name)
    for k, val in row_data.items():
        if normalize(k) == field_normalized:
            return _safe_str(val)
    return ""


def _resolve_currency_from_country(country_code: str) -> str:
    country_to_currency = {
        "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD",
        "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


def _derive_season(season_code: str) -> tuple[str, str]:
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
    pfx, yr = _derive_season(season_code)
    return f"CLH_{brand_code.upper()}_{pfx}{yr}"


# Fixed principal brand name for ALL classification IDs in this file —
# confirmed by the team as "PtpBahe" (exact casing, not all-caps). This
# covers every classification touchpoint uniformly:
#   Batches parent (ParentID)        → CLH_PtpBaheBatches
#   Season classification            → CLH_PtpBahe_{season}{year}
#   Confirmed/Unconfirmed Articles   → CLH_PtpBahe_{season}{year}CA / UA
#   Merchandiser reference           → CLH_PtpBaheArticles
# One single classification tree covers every product regardless of that
# row's own BRAND column value (PTP or BAHE) — BRAND is a product-line
# label, not a separate classification principal (see module docstring).
PRINCIPAL_BRAND_NAME = "PtpBahe"


def _principal_season_id(season_code: str) -> str:
    pfx, yr = _derive_season(season_code)
    return f"CLH_{PRINCIPAL_BRAND_NAME}_{pfx}{yr}"


def build_classifications(season_code: str) -> ET.Element:
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    pfx, yr  = _derive_season(season_code)
    short    = f"{pfx} {yr}".strip()
    sea_name = LOV_SEASON.get(pfx, pfx)

    batches_par = f"CLH_{PRINCIPAL_BRAND_NAME}Batches"
    season_id   = _principal_season_id(season_code)
    display     = f"{PRINCIPAL_BRAND_NAME} {sea_name} {yr}".strip()

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


def build_product_xml(
    row_data:            dict,
    brand_code:          str,
    comp_code:           str,
    sbu:                 str,
    season:              str,
    country_code:        str,
    rna_result:          dict,
    brand_mapping:       "BrandMappingLoader",
    mdd:                 "MDDLoader",
    color:               str = "",
) -> str:
    style_code = _safe_str(row_data.get("NAME"))
    if not style_code:
        return ""

    # Per PTP mapping sheet rows 13 (AT_PrincipalStyleCode) and 23
    # (AT_Generic — "MAA Generic Code (Max 12 Digits - 3 Digits brand code
    # + 9 Digits Principal Style)"): both give field name "UPC Code", NOT
    # NAME. Confirmed with the team (2026-08-17) — the generic/key code
    # and AT_PrincipalStyleCode are built from UPC CODE, matching
    # ean_source.py's KEY_InboundArticle formula exactly (previously
    # flagged there as a mismatch against this file's NAME-based key —
    # that FLAG is now resolved by this change). UPC codes are unique per
    # row already, so no truncation/collision tradeoff applies here the
    # way it did for NAME-based style codes. No fallback to NAME if UPC
    # CODE is blank — the row is simply skipped (matches the existing
    # NAME-blank early return above).
    principal_style_code = _get_field(row_data, "UPC CODE")
    if not principal_style_code:
        return ""
    style_clean  = re.sub(r"[^A-Z0-9]", "", principal_style_code.upper())
    # Always the reporting brand code ("PTP", even for BAHE rows) — see
    # REPORTING_BRAND_CODE_OVERRIDES. NOT a naive brand_code[:3] slice.
    brand_prefix = _reporting_brand_code(brand_code)
    generic_code = f"{brand_prefix}{style_clean}"

    # ParentID letter: generic mapping via _get_ptp_division_code() —
    # covers all six divisions from the confirmed legend (Accessories=E,
    # Footwear=F, Apparel=A, Equipment=Q, Toys=T), not just a two-way
    # Accessories/Equipment patch. PTP's own mapping sheet (row 74,
    # "SAP Product Division") marks this "Manual input" with no source
    # column, so ITEM RANGE is used as the practical proxy (confirmed
    # necessary — live import behavior showed 66 of 238 rows tagged
    # ACCESSORIES/FITNESS ACCESSORIES were wrongly landing under
    # Equipment when this was a single hardcoded default). Any ITEM
    # RANGE that doesn't match a known keyword falls back to Q
    # (Equipment) — PTP's core business — with a warning logged so
    # unrecognized values are visible rather than silently misfiled.
        # ParentID is fixed — all PTP/BAHE products route to Accessories (E),
    # confirmed by the team. ITEM_RANGE_DIVISION_KEYWORDS /
    # _get_ptp_division_code() are no longer used for parent routing.
    parent_id = "PPH_E-TempSubCat"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    name_text = _get_field(row_data, "DISPLAY NAME")
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name_text

    # Both classification references use the fixed unified "PtpBahe"
    # principal (see PRINCIPAL_BRAND_NAME / build_classifications()) —
    # not the row's own BRAND value. Confirmed by team: the correct
    # classification name for this file is "PtpBahe", matching the tree
    # built once in build_classifications(), not a per-brand-token one.
    season_id = _principal_season_id(season)

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{PRINCIPAL_BRAND_NAME}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    row_data["_generic_code"] = generic_code
    row_data["_color"] = color
    row_data["_principal_style_code"] = principal_style_code

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_product_values(vals_el, row_data, brand_code, comp_code, sbu, season, country_code,
                         rna_result, brand_mapping, mdd)

    return ET.tostring(g_el, encoding="unicode")


def _add_product_values(
    vals_el:       ET.Element,
    row_data:      dict,
    brand_code:    str,
    comp_code:     str,
    sbu:           str,
    season:        str,
    country_code:  str,
    rna_result:    dict,
    brand_mapping: "BrandMappingLoader",
    mdd:           "MDDLoader",
) -> None:
    """
    Write order now matches Birkenstock's XMLGenerator.build_products()
    sequence exactly:

    1. System / org      — InboundGenericCode, CompanyCode, SBU, Country,
                            Brand, Season, BrandGroup/Type/Category/Status,
                            currencies, SAP flags, UOM, BCI, Age, Gender,
                            CountryOrigin, PricingDistributionChannel,
                            ArticleStatus
    2. Special-cased      — StyleCode, StyleDescription, SportsCategoryEN
    3. Dynamic / mapped   — everything else from the brand mapping sheet
                            (Principal Hierarchy L1-L5, ColorName, Size,
                            HSCode, Collection, etc.)
    4. SeasonYear         — written last, same as Birkenstock's ordering
                            (FOB/SeasonYear trail the article in both files)
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

    # ── STEP 1: System / org values (Birkenstock order) ────────────
    generic_code = _safe_str(row_data.get("_generic_code", ""))
    if generic_code:
        _w("AT_InboundGenericCode", value=generic_code)

    _mw("AT_CompanyCode", comp_code)
    _mw("AT_SBU",        sbu)

    if country_code:
        _w("AT_Country", id_val=country_code)

    # AT_Brand must be a valid Brand LOV ID, not the raw order-form BRAND
    # text — those happened to be identical for "PTP" but not for every
    # brand, and BAHE has no Brand LOV entry of its own at all (reports
    # under PTP — see REPORTING_BRAND_CODE_OVERRIDES).
    reporting_brand_code = _reporting_brand_code(brand_code)
    brand_lov_id = mdd.lovs.get("AT_Brand", {}).get(reporting_brand_code.upper(), "")
    _w("AT_Brand", id_val=brand_lov_id or reporting_brand_code)

    sea_raw = (season or "").strip().upper()
    sea_prefix = sea_raw[:2] if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_prefix, sea_prefix)
    _w("AT_Season", sea_label, id_val=sea_prefix)

    brand_type     = rna_result.get("brand_type", "")
    brand_group    = rna_result.get("brand_group", "")
    brand_category = rna_result.get("brand_category", "")
    brand_status   = rna_result.get("brand_status", "")

    if brand_group:
        bg_lov_id = mdd.lovs.get("AT_BrandGroup", {}).get(brand_group.upper(), "")
        _w("AT_BrandGroup", value=brand_group if not bg_lov_id else "", id_val=bg_lov_id)

    if brand_type:
        bt_lov_id = mdd.lovs.get("AT_BrandType", {}).get(brand_type.upper(), "")
        _w("AT_BrandType", value=brand_type if not bt_lov_id else "", id_val=bt_lov_id)

    if brand_category:
        bc_lov_id = mdd.lovs.get("AT_BrandCategory", {}).get(brand_category.upper(), "")
        _w("AT_BrandCategory", value=brand_category if not bc_lov_id else "", id_val=bc_lov_id)

    if brand_status:
        bs_lov_id = mdd.lovs.get("AT_BrandStatus", {}).get(brand_status.upper(), "")
        _w("AT_BrandStatus", value=brand_status if not bs_lov_id else "", id_val=bs_lov_id)

    _w("AT_FOBCurrency", id_val="USD")

    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _w("AT_RetailPriceCurrency", id_val=currency_lov_id)

    _w("AT_SAPProductFlag", id_val="A")
    _w("AT_MaterialType", id_val="ZINA")
    _w("AT_SAPArticleCategory", id_val="1")
    _w("AT_UOM", id_val="EA")
    _w("AT_BCI", id_val="COMMERCIAL")
    _w("AT_BYAge", id_val="ADULT")
    _w("AT_SAPAge", id_val="AD")
    _w("AT_BYGender", id_val="U")
    _w("AT_Gender", id_val="U")
    _w("AT_PricingDistributionChannel", id_val="01")
    _w("AT_ArticleStatus", id_val="A")
    # AT_BYArticleType: mapping sheet row's attribute_id resolves to
    # "#N/A" (a broken formula in the source template), so
    # BrandMappingLoader silently skips it and this was never written.
    # Per MDD Core Attributes ("BY" → AT_BYArticleType) and the mapping
    # sheet's own mapping_logic note ("Default : Inline"), this file is
    # inline-only, so it's hardcoded here like AT_BYAge/AT_BYGender above.
    _w("AT_BYArticleType", id_val="Inline")

    # ── STEP 2: Special-cased attributes ────────────────────────────
    # AT_PrincipalStyleCode: per mapping sheet row 13, field name is
    # "UPC Code" (not NAME — confirmed with the team 2026-08-17). Reuses
    # the exact value build_product_xml() already resolved for the
    # generic key (including its NAME fallback for rows with a blank UPC
    # CODE), so this attribute and KEY_InboundArticle never disagree.
    style_code_raw = _safe_str(row_data.get("_principal_style_code", ""))
    if style_code_raw:
        _w("AT_PrincipalStyleCode", value=style_code_raw)

    # AT_PrincipalStyleDescription: mapping sheet row 15 (row index 14)
    # gives field name "Brand + DISPLAY NAME" — a concatenation formula,
    # not a literal order-form column, so the dynamic mapping loop below
    # (which does a straight column lookup) can never resolve it and was
    # silently skipping this attribute entirely. Built here instead as
    # BRAND + " " + DISPLAY NAME, per that row's own field-name formula
    # (e.g. brand_code="PTP", DISPLAY NAME="TOTAL RESISTANCE GYM" →
    # "PTP TOTAL RESISTANCE GYM"). Uses the row's own resolved brand_code
    # (PTP or BAHE), not a fixed default — same per-row brand handling as
    # the rest of this file.
    display_name_raw = _get_field(row_data, "DISPLAY NAME")
    if display_name_raw:
        brand_upper = brand_code.strip().upper()
        # Avoid double-prefixing when DISPLAY NAME already leads with the
        # brand token (e.g. BRAND="BAHE", DISPLAY NAME="BAHE Resistance
        # Band" → "BAHE Resistance Band", not "BAHE BAHE Resistance Band").
        if display_name_raw.strip().upper().startswith(brand_upper):
            style_description = display_name_raw.strip()
        else:
            style_description = f"{brand_upper} {display_name_raw}".strip()
        _w("AT_PrincipalStyleDescription", value=style_description)

    # AT_FOB: CUSTOMER BUY PRICE sometimes carries a long division-derived
    # decimal (e.g. 2.388235294117647), which Stibo rejects as an
    # oversized value for this currency field. Round to 2 decimal places
    # — the dynamic mapping loop below would otherwise write it as-is
    # (raw str(), no rounding), so AT_FOB is excluded from that loop
    # (see _SPECIAL_CASED_ATTRIBUTES) and handled here instead.
    fob_raw = _get_field(row_data, "CUSTOMER BUY PRICE")
    if fob_raw:
        try:
            _w("AT_FOB", value=f"{round(float(fob_raw), 2):g}")
        except ValueError:
            _w("AT_FOB", value=fob_raw)

    # AT_PrincipalBarcode is intentionally NOT written here. Barcode data
    # (UPC CODE + INNERPACK BARCODE) is now owned exclusively by the
    # ean_source.py companion, which writes AT_Barcode inside a DC_Barcode
    # DataContainer on the Variant — dropped from this file to avoid two
    # different generators both claiming to own barcode data for the same
    # article. See ean_source.py's module docstring for the barcode
    # mapping and the KEY_InboundArticle mismatch FLAG.
    sports_category = SPORTS_CATEGORY_BY_BRAND.get(brand_code.strip().upper())
    if sports_category:
        sc_lov_id = mdd.get_lov_id("AT_SportsCategoryEN", sports_category)
        # Sports Category LOV IDs in the MDD are stored as plain numbers
        # (1, 2, 3...) but Stibo's LOV expects zero-padded 2-digit codes
        # ("01", "02", "03"...) — confirmed with the team (2026-08-19).
        if sc_lov_id.isdigit():
            sc_lov_id = sc_lov_id.zfill(2)
        if sc_lov_id:
            _w("AT_SportsCategoryEN", value=sports_category, id_val=sc_lov_id)

    # ── STEP 3: Dynamic attributes from brand mapping sheet ─────────
    # (Principal Hierarchy L1-L5, StyleDescription, ColorName, Size,
    # HSCode, Collection, etc. — whatever "Direct from Principal" /
    # "Formula in system" rows resolve from the order form.)
    for mapping in brand_mapping.get_all_mappings():
        attr_id    = mapping["attribute_id"]
        field_name = mapping["field_name"]
        value_type = mapping["value_type"]

        if attr_id in EXCLUDED_ATTRIBUTES or attr_id in _SPECIAL_CASED_ATTRIBUTES:
            continue
        if not field_name or attr_id in written:
            continue

        raw_value = _get_field(row_data, field_name)
        if not raw_value:
            continue

        if value_type == "LOV":
            lov_id = mdd.get_lov_id(attr_id, raw_value)
            if lov_id:
                _w(attr_id, id_val=lov_id)
        else:
            _w(attr_id, value=raw_value)

    # ── STEP 4: AT_SeasonYear — written last (Birkenstock parity) ──
    # AT_FOB is now written in STEP 2 above (special-cased for rounding);
    # SeasonYear is derived here since it depends on the season string,
    # not a source column.
    season_year = ""
    if len(sea_raw) >= 4:
        year_part = sea_raw[2:]
        if len(year_part) == 2 and year_part.isdigit():
            season_year = f"20{year_part}"
        elif len(year_part) == 4 and year_part.isdigit():
            season_year = year_part
        else:
            year_match = re.search(r"20\d{2}", sea_raw)
            if not year_match:
                short_match = re.search(r"\d{2}$", sea_raw)
                season_year = f"20{short_match.group()}" if short_match else ""
            else:
                season_year = year_match.group()
    if season_year:
        _w("AT_SeasonYear", value=season_year)


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — MAIN RUN FUNCTION  (unchanged from v1.0)
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    log.info("[PTP Order Form] Starting ...")

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
    orderform     = PTPOrderFormLoader(orderform_file)

    if orderform.df.empty:
        log.warning("[PTP Order Form] No data rows found")
        return

    brand_code_default = getattr(args, "brand_code", "PTP").upper()
    comp_code          = getattr(args, "comp_code",  "0888")
    sbu                = getattr(args, "sbu",        "SP").upper()
    season             = getattr(args, "season",     "SS26").upper()
    country_code       = getattr(args, "country_code", "ID").upper()

    log.info(
        "Pipeline → brand_code(default)=%s  comp_code=%s  sbu=%s  season=%s  country=%s",
        brand_code_default, comp_code, sbu, season, country_code,
    )

    country_name = mdd.get_country_name_by_id(country_code)
    log.info("[RNA] Country code '%s' resolved to '%s' from MDD Country LOV", country_code, country_name)

    log.info("[PTP Order Form] Sheets loaded: %s", orderform.sheets_loaded)
    log.info("[PTP Order Form] Total rows before filtering: %d", len(orderform.df))

    filtered_df = orderform.df[
        orderform.df.apply(lambda row: bool(_safe_str(row.get("NAME"))), axis=1)
    ].copy()

    log.info("[PTP Order Form] Total rows to process after filtering: %d", len(filtered_df))

    # test_row_range = getattr(args, "test_row_range", None) or TEST_ROW_RANGE
    # if test_row_range:
    #     start, end = test_row_range
    #     filtered_df = filtered_df.iloc[start - 1:end].copy()
    #     log.warning("[TEST MODE] Limited to data rows %d–%d (%d rows) — clear "
    #                 "TEST_ROW_RANGE (or drop --test-row-range) for a full production run",
    #                 start, end, len(filtered_df))

    brand_codes_in_file = sorted({
        (_safe_str(row.get("BRAND")).strip().upper() or brand_code_default)
        for _, row in filtered_df.iterrows()
    })
    if not brand_codes_in_file:
        brand_codes_in_file = [brand_code_default]
    log.info("[PTP Order Form] Distinct brand codes found in file: %s", brand_codes_in_file)

    rna_cache: dict[str, dict] = {}

    def _rna_for_brand(brand_code: str) -> dict:
        key = brand_code.strip().upper()
        if key not in rna_cache:
            rna_cache[key] = _resolve_rna_for_brand(rna, country_name, comp_code, sbu, key)
        return rna_cache[key]

    cls_el  = build_classifications(season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    del cls_el

    out_name = orderform_file.stem + ".xml"
    out_path = XML_OUT_DIR / out_name
    export_time   = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    written_count = 0

    log.info("[PTP Order Form] Streaming XML → %s ...", out_name)

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

            colors_raw = _safe_str(row_dict.get("COLOUR"))
            if colors_raw:
                colors = [c.strip() for c in colors_raw.replace('\n', ',').split(',') if c.strip()]
            else:
                colors = [""]

            row_brand_code = _safe_str(row_dict.get("BRAND")).strip().upper() or brand_code_default
            # RNA has no rows for BAHE (only PTP) — resolve via the
            # reporting brand code (see REPORTING_BRAND_CODE_OVERRIDES),
            # while row_brand_code itself stays BAHE for the generic
            # prefix / sports category / style description below.
            row_rna_result = _rna_for_brand(_reporting_brand_code(row_brand_code))

            for color in colors:
                product_xml = build_product_xml(
                    row_dict, row_brand_code, comp_code, sbu, season, country_code,
                    row_rna_result, brand_mapping, mdd, color,
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

    print("═══ PTP ORDER FORM SUMMARY ═════════════════════════════════", flush=True)
    print(f"  Input file      : {orderform_file.name}",              flush=True)
    print(f"  Sheets          : {', '.join(orderform.sheets_loaded)}", flush=True)
    print(f"  Brand codes     : {', '.join(brand_codes_in_file)}",   flush=True)
    print(f"  Products written: {written_count}",                    flush=True)
    print(f"  Output XML      : {out_path.name}",                    flush=True)
    print(f"  File size       : {file_kb}KB",                        flush=True)
    # if test_row_range:
    #     print(f"  TEST MODE       : ON (rows {test_row_range[0]}–{test_row_range[1]})", flush=True)
    # print("═══════════════════════════════════════════════════════════", flush=True)

    if auditor:
        try:
            auditor.set_xml_uploads([str(out_path)])
        except AttributeError:
            pass


if __name__ == "__main__":
    import argparse
    import types

    parser = argparse.ArgumentParser()
    parser.add_argument("--brand", default="PTP")
    parser.add_argument("--brand-code", dest="brand_code", default="PTP")
    parser.add_argument("--comp-code", dest="comp_code", default="0888")
    parser.add_argument("--sbu", default="SP")
    parser.add_argument("--season", default="SS26")
    parser.add_argument("--country-code", dest="country_code", default="ID")
    parser.add_argument("--test-row-range", dest="test_row_range", type=int, nargs=2,
                         metavar=("START", "END"),
                         help="1-indexed inclusive data-row range for a quick smoke test, e.g. --test-row-range 1 10")

    parsed_args = parser.parse_args()
    run_args = types.SimpleNamespace(**vars(parsed_args))
    if run_args.test_row_range:
        run_args.test_row_range = tuple(run_args.test_row_range)
    run(run_args)