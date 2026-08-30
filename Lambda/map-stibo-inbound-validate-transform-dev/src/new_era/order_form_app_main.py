"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — New Era Order Form App v1.0      ║
║   Fully mapping-driven: no per-attribute Python code              ║
╚══════════════════════════════════════════════════════════════════╝

Input file : "...Order form - APAC FW26 (APP)..." (sheet "APAC FW26-APP")

Structural twin of order_form_acc_main.py (New Era's Accessories ETL) —
same "New Era(Inline)" mapping tab, same dynamic attribute engine, just
targeting the Apparel file/sheet and the "File (APP) Apparel" Col J block
instead of "Bags and Other Accessories". See order_form_acc_main.py's
module docstring for the full column contract (Col A/B/C/G/J/K); it is
identical here.

Same convention as anta/linelist_main.py — this ETL does NOT hardcode
which attributes to send or where they come from. Instead it reads the
"New Era(Inline)" tab of the shared brand-mapping workbook
("NEW - Brand mapping files Template.xlsx") at runtime and builds the list
of attributes to emit from these columns:

  A  Stibo Attribute            — human label (log/debug only)
  B  Stibo Attribute ID         — AT_xxx id written into the XML. Rows with
                                   a missing/#N/A id are always skipped.
  C  Stibo Validation           — "LOV"  -> value is written as the Value
                                   element's ID attribute (id-type).
                                   anything else (text/regexp/date/number/
                                   legacyisodatetime/...) -> written as
                                   element text (value-type).
  G  Field Mapping Type         — decides whether the attribute is included:
                                     "Direct from Principal"            -> include
                                     "Formula in System" + Col J value  -> include
                                     otherwise                          -> excluded
  J  Field Name in the Brand File — column to read on the "APAC FW26-APP"
                                     sheet of the New Era input file.
  K  Mapping Logic              — if present, parsed as newline-separated
                                   "source : id - label" entries
                                   (e.g. "M : M - Male") into a lookup
                                   table applied to the raw value. A raw
                                   value with no match in the table is
                                   skipped for that attribute (no default
                                   guess). If Col K is empty, the raw value
                                   is sent as-is (passthrough).

Adding/removing/re-pointing an attribute is a spreadsheet edit on the
"New Era(Inline)" tab — this file does not need to change.

System-level fields (Company, SBU, Country, Season, Season Year) and
AT_Brand / AT_BrandType / AT_BrandStatus / AT_BrandGroup are supplied
structurally from run() args / filename metadata + MDD/RNA lookups —
same convention as every other brand module (see anta/linelist_main.py),
not from the dynamic per-row engine below.

Grouping: one Generic Article per unique Style Code (the source column of
AT_PrincipalStyleCode) — Generic articles only, no Variant sub-products.
Same simplification as anta/linelist_main.py (no confirmed SAP Color/Size
source in this input file yet).

ASSUMPTION flagged for confirmation: the Generic Article ParentID defaults
to the Apparel hierarchy (PPH_A-TempSubCat below — same convention as
asics/app_acc_main.py, reebok/article_master_apparel.py, new_balance's
apparel modules). Change DEFAULT_PRODUCT_PARENT_ID if New Era Order Form
App articles route to a different division.

ASSUMPTION flagged for confirmation: any attribute the mapping sheet marks
as a flat system default with NO Col J source (e.g. "Default : EA" for
AT_UOM) is currently NOT sent — the dynamic engine only includes rows with
"Direct from Principal" or "Formula in System" + a Col J value, per this
file's exact spec. Add such defaults as hardcoded _val_lov()/_val_text()
calls in build_generic_product() once confirmed, same as Anta's own
AT_SAPAge/AT_CountryOrigin/etc. defaults.

FORCED_DIRECT_COLUMNS mostly reuses the Accessories file's column names
(Description/Silhouette/Style Name), except AT_PrincipalColorName, which
uses "Color" here (Accessories uses the British spelling "Colour") —
confirmed against the real Apparel sheet, per user direction (2026-08-19).
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

log = logging.getLogger(__name__)

# ======================================================================
# PATHS
# ======================================================================
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", Path(__file__).resolve().parent))

INPUT_DIR   = BASE_DIR / "input" / "orderform_app"
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "New Era"
BRAND_CODE = "NRA"

MAPPING_SHEET_NAME = "New Era(Inline)"   # tab in the brand-mapping workbook
INPUT_SHEET_NAME   = "APAC FW26-APP"     # sheet in the New Era input file

# The "New Era(Inline)" mapping tab's Col J is shared across THREE physical
# brand files (Apparel / Bags and Other Accessories / Headwear) — for many
# rows it packs all three column names into one cell, e.g.:
#     File (APP) Apparel
#     Column : MI#
#
#     File (ACC) Bags and Other Accessories
#     Column : MI#
#
#     File (ACC) Headwear
#     Column : MI
# "APAC FW26-APP" (this file) is confirmed to be the "Apparel" file
# specifically — TARGET_FILE_TYPE_LABEL below is matched case-insensitively
# as a substring against each block's "File (...) LABEL" line to pick out
# just that file type's column name. Col J values with no "File (...)"
# blocks at all (plain single-value cells, e.g. "EAN/UPC") are used as-is.
TARGET_FILE_TYPE_LABEL = "Apparel"

# See "ASSUMPTION flagged" note in the module docstring.
DEFAULT_PRODUCT_PARENT_ID = "PPH_A-TempSubCat"

# Attributes whose Col-J source column (from the mapping tab) has a known
# mismatch against the actual input sheet's header — the mapping sheet's
# Col J is always tried first; if that column isn't found in the input
# file, the matching fallback column name here is tried instead. Currently
# empty — the attributes that used to live here (style code/description,
# color, merch hierarchy L1/L2, collection, SPU grouping) were moved to
# FORCED_DIRECT_COLUMNS below because the mapping tab's Col J/K parsing
# proved unreliable for them in production (see FORCED_DIRECT_COLUMNS'
# comment).
COLUMN_FALLBACKS: dict[str, str] = {}

# Attributes that bypass the mapping tab ENTIRELY — read directly from a
# hardcoded input column, never via Col J/K. Added per user direction
# (2026-08-19) after production runs showed the mapping tab's Col J/K text
# (packed "File (...) / Column : ..." blocks reused for BOTH source-column
# hints and, apparently, spurious Col K "mapping logic" — see
# _RESERVED_MAPPING_LOGIC_KEYS) was resolving these attributes to the
# wrong column, or silently skipping them, even after the Col J/K parsing
# fixes. AT_PrincipalStyleCode / AT_PrincipalStyleDescription are handled
# directly in build_generic_product() (they feed KEY_InboundArticle/<Name>
# too); the rest are applied via the loop right after it.
FORCED_DIRECT_COLUMNS: dict[str, str] = {
    "AT_PrincipalColorName":             "Color",
    "AT_PrincipalMerchandiseHierarchyL1": "Silhouette",
    "AT_PrincipalMerchandiseHierarchyL2": "Style Name",
    "AT_Collection1":                    "Silhouette",
    "AT_SPUGroupingName":                "Style Name",
}

# Attribute IDs handled by dedicated logic in build_generic_product()
# instead of the generic per-row passthrough/mapped engine — excluded from
# the dynamic loop to avoid double-emission.
#   AT_Silhouette: input "Silhouette" -> attributes file sheet
#   "New Era MD Mappings" (Col B match + Col C="Apparel" -> Col H)
#   -> MDD "Silhouette LOV" (Col B match -> Col A LOV ID); if LOV ID is
#   single-digit numeric, zero-pad to 2 digits.
#   AT_Fit: input "Silhouette" -> attributes file sheet
#   "New Era MD Mappings" (Col B match + Col C="Apparel" -> Col I)
#   -> MDD "Fit LOV" (Col B match -> Col A LOV ID).
#   AT_PrincipalAgeDescription: fixed default "Adults", not sourced from
#   any input column.
#   AT_SAPAge: fixed default LOV id "AD".
#   AT_Gender: derived from "Silhouette" — if the first word starts with
#   "W" (case-insensitive, e.g. "W Crew Tee"), Women -> LOV id "F";
#   otherwise Men -> LOV id "M".
#   AT_CountryOrigin: fixed default LOV id "CN".
#   AT_PricingDistributionChannel: fixed default LOV id "01".
#   AT_FOB: "Price" column value minus 33% (not a straight passthrough —
#   the dynamic engine only supports raw/Col-K-table values, not
#   arithmetic).
#   AT_FOBCurrency: fixed default LOV id "USD".
#   AT_MaterialType: fixed default LOV id "ZINA".
#   AT_SAPArticleCategory: fixed default LOV id "01".
#   AT_UOM: fixed default LOV id "EA".
#   AT_BYArticleType: fixed default LOV id "Inline".
#   AT_BYAge: fixed default LOV id "ADULT".
#   AT_BYGender: same Silhouette-derived logic as AT_Gender (see above).
#   AT_Content: "Fabric Content" column's highest-percentage material
#   (e.g. "100% POLYESTER RIPSTOP 200D" -> "POLYESTER RIPSTOP 200D"),
#   matched against the MDD "Content LOV" tab (Col B display name
#   substring-matched -> Col A LOV id). Replaces the earlier AT_Material
#   mapping, which has been removed (per user direction, 2026-08-19).
#   AT_CountrySize: fixed default LOV id "INT".
#   AT_EComAgesCategory: fixed default LOV id "18+Y".
#   AT_ArticleStatus: fixed default LOV id "A".
#   AT_SAPProductFlag: fixed default LOV id "A".
#   AT_SportsCategoryEN: fixed default LOV id "06".
#   AT_PrincipalSize: NOT a single named column — mapping tab's Col J for
#   Apparel is "Column : O - AN (No Fill Color)": columns O..AN of the
#   input sheet are the size grid (each column's own header IS a size
#   label, e.g. "OSFA"/"S"/"634"), and a column counts as offered for a
#   given row when that cell has NO fill (grayed-out = not applicable to
#   that item). Resolved by NewEraOrderFormLoader into each row dict
#   under _SIZE_ROW_KEY (see _read_size_grid_fills) since the normal
#   Col-J-name lookup can't address a column range or read cell styling.
#   (all per user direction, 2026-08-19)
SPECIAL_ATTRIBUTE_IDS: set[str] = {
    "AT_PrincipalAgeDescription", "AT_SAPAge", "AT_Gender", "AT_CountryOrigin",
    "AT_PricingDistributionChannel", "AT_FOB", "AT_FOBCurrency", "AT_MaterialType",
    "AT_SAPArticleCategory", "AT_UOM", "AT_BYArticleType", "AT_BYAge", "AT_BYGender",
    "AT_Content", "AT_Silhouette", "AT_Fit", "AT_CountrySize", "AT_EComAgesCategory", "AT_ArticleStatus",
    "AT_SAPProductFlag", "AT_SportsCategoryEN", "AT_PrincipalSize",
    "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
    *FORCED_DIRECT_COLUMNS.keys(),
}

# AT_PrincipalSize size-grid columns (see SPECIAL_ATTRIBUTE_IDS comment above).
SIZE_GRID_FIRST_COL = "N"
SIZE_GRID_LAST_COL  = "X"
# Synthetic key NewEraOrderFormLoader stashes the resolved, comma-joined
# offered-size string under in each row dict (not a real input column).
_SIZE_ROW_KEY = "__AT_PrincipalSize__"


def _read_size_grid_fills(path: Path, sheet_name: str, header_row_idx: int,
                           raw_header: tuple) -> dict[int, str]:
    """
    Second, non-read-only pass over `sheet_name` limited to the
    SIZE_GRID_FIRST_COL:SIZE_GRID_LAST_COL columns, for AT_PrincipalSize.
    openpyxl's read_only/values_only mode (used for everything else in
    NewEraOrderFormLoader) doesn't expose cell styling at all, so the
    "no fill = offered" convention from the mapping tab's "Column : O - AN
    (No Fill Color)" note has to be read separately, here.

    Returns {1-based sheet row number: comma-joined offered size labels},
    one entry per data row that has at least one offered size. Labels come
    from `raw_header` at each size-grid column position — those header
    cells ARE the size labels (e.g. "OSFA", "S", "634"), not a shared
    "Size" column name.
    """
    first_idx0 = openpyxl.utils.column_index_from_string(SIZE_GRID_FIRST_COL) - 1
    last_idx0  = openpyxl.utils.column_index_from_string(SIZE_GRID_LAST_COL) - 1
    size_labels = {
        i: _s(raw_header[i]) for i in range(first_idx0, last_idx0 + 1)
        if i < len(raw_header) and _s(raw_header[i])
    }
    if not size_labels:
        log.warning(
            "[NewEra-OrderFormApp] Size grid columns %s:%s have no header labels — "
            "AT_PrincipalSize will not be sent",
            SIZE_GRID_FIRST_COL, SIZE_GRID_LAST_COL,
        )
        return {}

    result: dict[int, str] = {}
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    try:
        ws = wb[sheet_name]
        for row_cells in ws.iter_rows(
            min_row=header_row_idx + 2, min_col=first_idx0 + 1, max_col=last_idx0 + 1,
        ):
            offered = []
            for cell in row_cells:
                label = size_labels.get(cell.column - 1)
                if not label:
                    continue
                pattern = cell.fill.patternType if cell.fill else None
                if pattern is None:
                    offered.append(label)
            if offered:
                result[row_cells[0].row] = ",".join(offered)
    finally:
        wb.close()
    return result


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
    return "" if s in ("None", "nan", "0", "NaT", "#N/A", "#VALUE!") else s


def _norm_header(s: str) -> str:
    """Case/whitespace-insensitive header key for column matching."""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _resolve_currency_from_country(country_code: str) -> str:
    """
    Resolve Retail Price Currency LOV ID from country code.
    Same country-to-currency mapping used across every other brand module
    (e.g. anta/linelist_main.py, clarks/order_form_footwear_main.py).
    """
    country_to_currency = {
        "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD",
        "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


_PCT_MATERIAL_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*([^%,/]+)")


def _top_material_text(fabric_content: str) -> str:
    """
    Extract the highest-percentage material name from a raw "Fabric
    Content" cell, for AT_Content.

    e.g. "100% POLYESTER RIPSTOP 200D" -> "POLYESTER RIPSTOP 200D"
         "60% Cotton, 40% Polyester"   -> "Cotton" (60 > 40)
    Same percentage-extraction technique as
    reebok/article_master_footwear.py's _top_material().
    """
    s = _s(fabric_content)
    if not s:
        return ""
    matches = [(pct, name) for pct, name in _PCT_MATERIAL_RE.findall(s) if name.strip()]
    if not matches:
        return ""
    _, name = max(matches, key=lambda m: float(m[0]))
    return name.strip().rstrip(",/.:").strip()


def _resolve_material_lov_id(material_text: str, material_lov: list[tuple[str, str]] | None) -> str:
    """
    Match extracted material text (e.g. "POLYESTER RIPSTOP 200D") against
    the MDD Material LOV's display names via substring search (an exact
    match is unlikely — the raw text usually carries extra descriptive
    words the LOV entry doesn't). `material_lov` is pre-sorted longest
    display name first, so the most specific match wins.
    """
    if not material_lov:
        return ""
    text_upper = material_text.strip().upper()
    if not text_upper:
        return ""
    for display_upper, lov_id in material_lov:
        if display_upper in text_upper:
            return lov_id
    return ""


# ======================================================================
# XML HELPERS
# ======================================================================

def _val(parent: ET.Element, attr_id: str, text: str = "", lov_id: str = "") -> None:
    if not text and not lov_id:
        return
    v = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    v.set("AttributeID", attr_id)
    if lov_id:
        v.set("ID", lov_id)
    if text:
        v.text = text


def _val_text(parent: ET.Element, attr_id: str, text: str) -> None:
    t = _s(text)
    if t:
        _val(parent, attr_id, text=t)


def _val_lov(parent: ET.Element, attr_id: str, lov_id: str) -> None:
    lid = _s(lov_id)
    if lid:
        _val(parent, attr_id, lov_id=lid)


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    v = _s(id_val)
    if not v:
        return
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    ET.SubElement(mv, f"{{{STIBO_NS}}}Value").set("ID", v)


# ======================================================================
# MDD LOADER  (Brand / Brand Type / Brand Status / Brand Group / Country
# LOVs — same sheet-layout convention as anta/linelist_main.py)
# ======================================================================

class MDDLoader:
    """Loads the LOV tables needed to resolve AT_Brand / AT_BrandType /
    AT_BrandStatus / AT_BrandGroup, plus Country (id -> name, used to key
    the RNA lookup)."""

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        self._load_brand_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_brand_group_lov(wb)
        self._load_country_lov(wb)
        self._load_silhouette_lov(wb)
        self._load_fit_lov(wb)
        self._load_content_lov(wb)
        wb.close()

    def _load_brand_lov(self, wb) -> None:
        """Brand LOV: col A = LOV ID, col B = display name."""
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()
             and "TYPE" not in s.upper() and "GROUP" not in s.upper() and "STATUS" not in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Brand LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_Brand"] = lov

    def _load_brand_type_lov(self, wb) -> None:
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Brand Type LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_BrandType"] = lov

    def _load_brand_status_lov(self, wb) -> None:
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Brand Status LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_BrandStatus"] = lov

    def _load_brand_group_lov(self, wb) -> None:
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet:
            sheet = next((s for s in wb.sheetnames if s.strip().upper() == "BRAND GROUP"), None)
        if not sheet:
            log.warning("[MDD] Brand Group LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_BrandGroup"] = lov

    def _load_country_lov(self, wb) -> None:
        """Country LOV: col A = LOV ID (e.g. 'ID'), col B = display name (e.g. 'Indonesia')."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper()
                  and "ORIGIN" not in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Country LOV sheet not found")
            return
        id_to_name: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                id_to_name[lov_id.upper()] = display
        self.lovs["CountryLOV"] = id_to_name

    def country_id_to_name(self, country_id: str) -> str:
        return self.lovs.get("CountryLOV", {}).get((country_id or "").strip().upper(), "")

    def _load_silhouette_lov(self, wb) -> None:
        """Silhouette LOV: col A = LOV ID (Code), col B = display value."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "SILHOUETTE LOV"),
            next((s for s in wb.sheetnames if "SILHOUETTE" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Silhouette LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_Silhouette"] = lov

    def _load_fit_lov(self, wb) -> None:
        """Fit LOV: col A = LOV ID (Code), col B = display value."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "FIT LOV"),
            next((s for s in wb.sheetnames if s.strip().upper() == "FIT"), None),
        )
        if not sheet:
            log.warning("[MDD] Fit LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_Fit"] = lov

    def _load_content_lov(self, wb) -> None:
        """Content LOV: col A = LOV ID (Code), col B = display name (Content)."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "CONTENT LOV"),
            next((s for s in wb.sheetnames if "CONTENT" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Content LOV sheet not found")
            return
        entries: list[tuple[str, str]] = []
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                entries.append((display.upper(), lov_id))
        # Longest display name first, so a substring match prefers the most
        # specific entry (e.g. "COTTON SPANDEX" over "COTTON" if both
        # happened to match — see _resolve_material_lov_id()).
        entries.sort(key=lambda e: len(e[0]), reverse=True)
        self.lovs["ContentLOV"] = entries


# ======================================================================
# RNA LOADER  (Brand Type / Brand Status / Brand Group — same
# "Source Mapping related RNA" sheet + 4-pass lookup used by every brand)
# ======================================================================

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
        self.path = path
        self.lookup: dict[tuple, dict] = {}
        self._load()

    def _load(self) -> None:
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
        wb.close()

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(isinstance(v, str) and "COUNTRY" in v.upper() for v in r if v)),
            None,
        )
        if hdr_idx is None:
            log.warning("[RNA] Cannot find header row in sheet '%s'", sheet_name)
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
                "brand_type":   _cell(row, c_btype),
                "brand_group":  _cell(row, c_bgroup),
                "brand_status": _cell(row, c_bstatus),
            }
            count += 1

    def get(self, country_name: str, comp_code: str, sbu: str, brand_code: str) -> dict:
        key = (
            self._norm_country(country_name),
            self._norm_comp_code(comp_code),
            (sbu or "").strip().upper(),
            (brand_code or "").strip().upper(),
        )
        return self.lookup.get(key, {"brand_type": "", "brand_group": "", "brand_status": ""})

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
        return {"brand_type": "", "brand_group": "", "brand_status": ""}

    def get_by_brand_only(self, brand_code: str) -> dict:
        b = (brand_code or "").strip().upper()
        for (_, _, _, kb), val in self.lookup.items():
            if kb == b:
                return val
        return {"brand_type": "", "brand_group": "", "brand_status": ""}


# ======================================================================
# ATTRIBUTE MAPPING ENGINE  (drives everything — see module docstring)
# ======================================================================

# Meta-note labels that share Col J's "File : ... / Column : ..." authoring
# convention — these sometimes leak into Col K as documentation about which
# file/column the mapping applies to, NOT a genuine source->target
# translation rule. A line like "Column : MI#" would otherwise be
# misparsed as a real mapping entry with key "COLUMN", which then makes
# every real input value fail to match (since raw data is never literally
# "File" or "Column") and the whole attribute gets silently skipped.
_RESERVED_MAPPING_LOGIC_KEYS = {"FILE", "COLUMN", "SOURCE"}


def _parse_mapping_logic(raw_logic) -> dict[str, str]:
    """
    Parse Col K "Mapping Logic" text into {SOURCE_UPPER: resolved_id_or_text}.

    Expected line format: "<source> : <target_id> - <target_label>"
      e.g. "M : M - Male"   -> {"M": "M"}
           "W : F - Female" -> {"W": "F"}
    Lines that don't contain ":" (e.g. free-text notes), or whose key is a
    reserved meta-note label (see _RESERVED_MAPPING_LOGIC_KEYS), are
    ignored — callers fall back to raw passthrough when the resulting
    table is empty.
    """
    mapping: dict[str, str] = {}
    text = _s(raw_logic)
    if not text:
        return mapping
    for line in text.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        src, _, tgt = line.partition(":")
        src = src.strip()
        tgt = tgt.strip()
        if not src or not tgt:
            continue
        if src.upper() in _RESERVED_MAPPING_LOGIC_KEYS:
            continue
        # "F - Female" -> take the id token before " - "
        tgt_id = tgt.split(" - ", 1)[0].strip() or tgt
        mapping[src.upper()] = tgt_id
    return mapping


# Matches one "File (CODE) Label\nColumn : value" block within a Col J cell
# — see TARGET_FILE_TYPE_LABEL note above for the raw text shape.
_FILE_BLOCK_RE = re.compile(
    r"File\s*\(([^)]*)\)\s*([^\n]*)\n\s*Column\s*:\s*([^\n]*)",
    re.IGNORECASE,
)


def _resolve_col_j_for_file_type(raw_col_j, file_type_label: str) -> str:
    """
    Resolve a Col J cell to the single column name that applies to
    `file_type_label`.

    - No "File (...)" blocks found at all -> plain single-value cell
      (e.g. "EAN/UPC") -> returned as-is.
    - "File (...)" blocks found -> return the "Column :" value from the
      block whose label contains `file_type_label` (case-insensitive). If
      blocks exist but none match, returns "" (this attribute has no
      source for our file type -> excluded upstream, same as an empty
      Col J).
    """
    text = _s(raw_col_j)
    if not text:
        return ""
    blocks = list(_FILE_BLOCK_RE.finditer(text))
    if not blocks:
        return text
    for m in blocks:
        label = m.group(2).strip()
        if file_type_label.upper() in label.upper():
            return m.group(3).strip()
    return ""


class AttributeRule:
    def __init__(self, name: str, attr_id: str, validation: str,
                 mapping_type: str, source_column: str, mapping_logic_raw):
        self.name          = name
        self.attr_id       = attr_id
        self.is_lov        = validation.strip().upper() == "LOV"
        self.mapping_type  = mapping_type
        self.source_column = source_column
        self.mode          = "direct"   # "direct" | "mapped"
        self.value_map: dict[str, str] = {}

        # Col K "mapping logic if any" — applied whenever it parses to a
        # non-empty table, independent of the Col G mapping type (per this
        # file's exact spec; unlike anta/linelist_main.py, which only
        # applies Col K for a specific "Mapping from Principal Data" type).
        parsed = _parse_mapping_logic(mapping_logic_raw)
        if parsed:
            self.mode = "mapped"
            self.value_map = parsed

    def resolve(self, raw_value: str) -> str | None:
        """Return the value to write for this attribute, or None to skip it."""
        raw = _s(raw_value)
        if not raw:
            return None
        if self.mode == "direct" or not self.value_map:
            return raw
        # "mapped" mode with a real lookup table — no default guess.
        return self.value_map.get(raw.upper())


class AttributeMappingLoader:
    """
    Reads the brand tab (MAPPING_SHEET_NAME) from the shared brand-mapping
    workbook and turns it into the list of AttributeRule objects that
    drive XML generation. See module docstring for the column contract.
    """

    def __init__(self, path: Path, sheet_name: str = MAPPING_SHEET_NAME):
        self.path = path
        self.sheet_name = sheet_name
        self.rules: list[AttributeRule] = []
        self._load()

    def _load(self) -> None:
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            target = next(
                (s for s in wb.sheetnames if s.strip().upper() == self.sheet_name.upper()),
                None,
            )
            if not target:
                raise ValueError(
                    f"[AttrMap] Sheet '{self.sheet_name}' not found in {self.path.name}. "
                    f"Available: {wb.sheetnames}"
                )
            rows = list(wb[target].iter_rows(values_only=True))
        finally:
            wb.close()

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if r and _s(r[0]).upper() == "STIBO ATTRIBUTE"),
            None,
        )
        if hdr_idx is None:
            raise ValueError(f"[AttrMap] Cannot find 'Stibo Attribute' header row in '{target}'")

        skipped_no_id       = 0
        skipped_type        = 0
        skipped_no_file_col = 0
        for row in rows[hdr_idx + 1:]:
            if not row or not row[0]:
                continue
            name          = _s(row[0])
            attr_id       = _s(row[1]) if len(row) > 1 else ""
            validation    = _s(row[2]) if len(row) > 2 else ""
            mapping_type  = _s(row[6]) if len(row) > 6 else ""
            source_column_raw = row[9] if len(row) > 9 else ""
            source_column = _resolve_col_j_for_file_type(source_column_raw, TARGET_FILE_TYPE_LABEL)
            mapping_logic = row[10] if len(row) > 10 else ""

            if not attr_id:
                skipped_no_id += 1
                continue

            if not source_column and _FILE_BLOCK_RE.search(_s(source_column_raw)):
                # Col J had "File (...)" blocks, but none matched our
                # TARGET_FILE_TYPE_LABEL — this attribute genuinely has no
                # source for the Apparel file.
                skipped_no_file_col += 1

            mt = mapping_type.strip()
            include = (
                (mt == "Direct from Principal" and source_column)
                or (mt.upper().startswith("FORMULA IN SYSTEM") and source_column)
            )
            if not include:
                skipped_type += 1
                continue

            self.rules.append(AttributeRule(
                name=name, attr_id=attr_id, validation=validation,
                mapping_type=mt, source_column=source_column,
                mapping_logic_raw=mapping_logic,
            ))

        log.info(
            "[AttrMap] Loaded %d active attribute rule(s) from '%s' tab "
            "(%d skipped: no valid attribute ID, %d skipped: mapping type not eligible, "
            "%d skipped: no '%s' column in Col J)",
            len(self.rules), self.sheet_name, skipped_no_id, skipped_type,
            skipped_no_file_col, TARGET_FILE_TYPE_LABEL,
        )

    def find(self, attr_id: str) -> AttributeRule | None:
        return next((r for r in self.rules if r.attr_id == attr_id), None)


class NewEraMDMappingLoader:
    """
        Loads silhouette/fit mappings from 'New Era MD Mappings' in the brand mapping file.
    Required mapping path for AT_Silhouette:
      Col B (source silhouette) + Col C == 'Apparel' -> Col H (mapped silhouette value)
        Required mapping path for AT_Fit:
            Col B (source silhouette) + Col C == 'Apparel' -> Col I (mapped fit value)
    """

    def __init__(self, path: Path):
        self.path = path
        self.silhouette_map: dict[str, str] = {}
        self.fit_map: dict[str, str] = {}
        self._load()

    def _load(self) -> None:
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet_name = next(
                (s for s in wb.sheetnames if s.strip().upper() == "NEW ERA MD MAPPINGS"),
                None,
            )
            if not sheet_name:
                log.warning("[AttrMap] 'New Era MD Mappings' sheet not found")
                return
            rows = list(wb[sheet_name].iter_rows(values_only=True))
        finally:
            wb.close()

        for row in rows[1:]:
            if not row:
                continue
            src_silhouette = _s(row[1]) if len(row) > 1 else ""  # Col B
            category       = _s(row[2]) if len(row) > 2 else ""  # Col C
            silhouette_value = _s(row[7]) if len(row) > 7 else ""  # Col H
            fit_value        = _s(row[8]) if len(row) > 8 else ""  # Col I

            if not src_silhouette:
                continue
            if category.upper() != "APPAREL":
                continue

            src_key = src_silhouette.upper()
            if silhouette_value:
                self.silhouette_map[src_key] = silhouette_value
            if fit_value:
                self.fit_map[src_key] = fit_value

    def resolve_silhouette_value(self, src_silhouette: str) -> str:
        return self.silhouette_map.get(_s(src_silhouette).upper(), "")

    def resolve_fit_value(self, src_silhouette: str) -> str:
        return self.fit_map.get(_s(src_silhouette).upper(), "")


# ======================================================================
# INPUT FILE LOADER  ("APAC FW26-APP" sheet — header row/columns
# discovered dynamically from the attribute rules' Col J source-column
# names, same convention as anta/linelist_main.py's AntaLinelistLoader)
# ======================================================================

class NewEraOrderFormLoader:
    def __init__(self, path: Path, needed_columns: list[str], sheet_name: str = INPUT_SHEET_NAME):
        self.path = path
        self.sheet_name = sheet_name
        self.headers: list[str] = []
        self.rows: list[dict] = []
        self._needed = {_norm_header(c) for c in needed_columns if c}
        self._load()

    def _load(self) -> None:
        log.info("[NewEra-OrderFormApp] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        target = next(
            (s for s in wb.sheetnames if s.strip().upper() == self.sheet_name.upper()),
            None,
        )
        if not target:
            wb.close()
            raise ValueError(
                f"[NewEra-OrderFormApp] Sheet '{self.sheet_name}' not found. "
                f"Available: {wb.sheetnames}"
            )
        log.info("[NewEra-OrderFormApp] Using sheet: '%s'", target)

        raw_rows = list(wb[target].iter_rows(values_only=True))
        wb.close()

        if not raw_rows:
            log.warning("[NewEra-OrderFormApp] Sheet '%s' is empty", target)
            return

        # ── Locate header row: the row (within the first 20) whose cells
        #    match the most Col-J source-column names from the mapping sheet.
        best_idx, best_score = 0, -1
        for i in range(min(20, len(raw_rows))):
            r = raw_rows[i]
            if not r:
                continue
            cells = {_norm_header(c) for c in r if c is not None}
            score = len(cells & self._needed)
            if score > best_score:
                best_idx, best_score = i, score
        if best_score <= 0:
            log.warning(
                "[NewEra-OrderFormApp] Could not confidently locate header row "
                "(no cell matched any expected column name) — defaulting to row 0"
            )
        log.info(
            "[NewEra-OrderFormApp] Header at row %d (0-indexed), matched %d/%d expected columns",
            best_idx, max(best_score, 0), len(self._needed),
        )

        raw_header = raw_rows[best_idx]
        header = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(raw_header)]
        seen: dict[str, int] = {}
        for i, col in enumerate(header):
            if col in seen:
                seen[col] += 1
                header[i] = f"{col}_{seen[col]}"
            else:
                seen[col] = 0
        self.headers = header

        # AT_PrincipalSize: resolved from the size-grid columns' fill state
        # (see _read_size_grid_fills), keyed by 1-based sheet row number so
        # it can be matched up to each data row below.
        size_by_row = _read_size_grid_fills(self.path, target, best_idx, raw_header)

        for sheet_row, r in enumerate(raw_rows[best_idx + 1:], start=best_idx + 2):
            if not r or all(c is None for c in r):
                continue
            row_dict = {header[i]: (r[i] if i < len(r) else None) for i in range(len(header))}
            row_dict[_SIZE_ROW_KEY] = size_by_row.get(sheet_row, "")
            self.rows.append(row_dict)

        log.info("[NewEra-OrderFormApp] Loaded %d data rows", len(self.rows))

    def find_col(self, name: str) -> str | None:
        """Case/whitespace-insensitive header lookup."""
        target = _norm_header(name)
        for h in self.headers:
            if _norm_header(h) == target:
                return h
        return None


def _resolve_source_col(loader: "NewEraOrderFormLoader", rule: "AttributeRule") -> str | None:
    """
    Resolve the actual input-sheet column for a rule: the mapping sheet's
    Col J is always tried first; if that column isn't present in the input
    file, fall back to COLUMN_FALLBACKS[rule.attr_id] (if any).
    """
    col = loader.find_col(rule.source_column)
    if col:
        return col
    fallback_name = COLUMN_FALLBACKS.get(rule.attr_id)
    if not fallback_name:
        return None
    col = loader.find_col(fallback_name)
    return col


def _resolve_at_silhouette_lov_id(
    raw_silhouette: str,
    md_mapping: "NewEraMDMappingLoader | None",
    mdd: "MDDLoader | None",
) -> str:
    """
    Resolve AT_Silhouette LOV ID via:
      input Silhouette -> New Era MD Mappings (B/C/H, C='Apparel')
      -> MDD Silhouette LOV (B->A)
    """
    src = _s(raw_silhouette)
    if not src or not md_mapping or not mdd:
        return ""

    mapped_silhouette = md_mapping.resolve_silhouette_value(src)
    if not mapped_silhouette:
        return ""

    silhouette_lov = mdd.lovs.get("AT_Silhouette", {})
    lov_id = _s(silhouette_lov.get(mapped_silhouette.upper(), ""))
    if lov_id.isdigit() and len(lov_id) == 1:
        lov_id = lov_id.zfill(2)
    return lov_id


def _resolve_at_fit_lov_id(
    raw_silhouette: str,
    md_mapping: "NewEraMDMappingLoader | None",
    mdd: "MDDLoader | None",
) -> str:
    """
    Resolve AT_Fit LOV ID via:
      input Silhouette -> New Era MD Mappings (B/C/I, C='Apparel')
      -> MDD Fit LOV (B->A)
    """
    src = _s(raw_silhouette)
    if not src or not md_mapping or not mdd:
        return ""

    mapped_fit = md_mapping.resolve_fit_value(src)
    if not mapped_fit:
        return ""

    fit_lov = mdd.lovs.get("AT_Fit", {})
    return _s(fit_lov.get(mapped_fit.upper(), ""))


# ======================================================================
# GROUPER — one Generic Article per unique Style Code
# ======================================================================

def group_rows(rows: list[dict], style_col: str) -> dict[str, dict]:
    """First-row-wins grouping by Style Code (no colour/size variants — see
    module docstring)."""
    groups: dict[str, dict] = {}
    for row in rows:
        style = _s(row.get(style_col, ""))
        if not style:
            continue
        if style not in groups:
            groups[style] = row
    log.info("[Grouper] %d rows -> %d generic(s) (grouped by Style Code)", len(rows), len(groups))
    return groups


# ======================================================================
# CLASSIFICATIONS BUILDER  (same 3-tier pattern used across every brand)
# ======================================================================

_SEASON_LABELS = {
    "SS": "Spring Summer", "FW": "Fall Winter",
    "AW": "Autumn Winter", "HO": "Holiday",
    "SP": "Spring",        "SM": "Summer",
}


def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_tail   = season_code[2:]
    sea_year   = f"20{sea_tail}" if len(sea_tail) == 2 else sea_tail
    full_season = f"{sea_prefix}{sea_year}"

    brand_id       = brand.replace(" ", "")   # classification IDs can't contain spaces (e.g. "New Era" -> "NewEra")
    season_id      = f"CLH_{brand_code}_{full_season}"
    batches_parent = f"CLH_{brand_id}Batches"
    sea_name       = _SEASON_LABELS.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short    = f"{sea_prefix} {sea_year}".strip()

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


# ======================================================================
# XML BUILDER
# ======================================================================

def build_generic_product(
    row: dict,
    loader: NewEraOrderFormLoader,
    rules: list[AttributeRule],
    brand: str,
    brand_code: str,
    comp_code: str,
    sbu: str,
    season: str,
    season_id: str,
    country_code: str,
    brand_lov_id: str = "",
    brand_type: str = "",
    brand_status: str = "",
    brand_group: str = "",
    mdd: "MDDLoader | None" = None,
    md_mapping: "NewEraMDMappingLoader | None" = None,
) -> str:
    # AT_PrincipalStyleCode: read directly from "MI#" — bypasses the
    # mapping tab entirely (per user direction, 2026-08-19: the tab's Col
    # J/K parsing proved unreliable for this attribute, sometimes
    # resolving to the wrong input column). Formula: KEY_InboundArticle =
    # AT_InboundGenericCode = brand_code + style_code.
    style_col  = loader.find_col("MI#")
    style_code = _s(row.get(style_col, "")) if style_col else ""
    if not style_code:
        return ""

    generic_code = f"{brand_code}{style_code}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", DEFAULT_PRODUCT_PARENT_ID)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    # AT_PrincipalStyleDescription: read directly from "Description" —
    # same bypass as above. Feeds <Name> directly; falls back to the
    # generic code if the column is missing/empty for this row.
    name_text = generic_code
    desc_col = loader.find_col("Description")
    desc_val = _s(row.get(desc_col, "")) if desc_col else ""
    if desc_val:
        name_text = desc_val
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name_text

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # AT_InboundGenericCode: same value as the KEY_InboundArticle KeyValue
    # above — not in the mapping tab (no Col J source), populated
    # structurally the same way anta/linelist_main.py does.
    _val_text(gv, "AT_InboundGenericCode", generic_code)

    # AT_PrincipalStyleCode / AT_PrincipalStyleDescription: write the
    # values already resolved above (bypassing the mapping tab — see
    # FORCED_DIRECT_COLUMNS' comment) as their own attribute entries too.
    _val_text(gv, "AT_PrincipalStyleCode", style_code)
    if desc_val:
        _val_text(gv, "AT_PrincipalStyleDescription", desc_val)

    # FORCED_DIRECT_COLUMNS: attributes read directly from a hardcoded
    # input column, bypassing the mapping tab entirely (see its comment).
    for attr_id, col_name in FORCED_DIRECT_COLUMNS.items():
        col = loader.find_col(col_name)
        if not col:
            continue
        val = _s(row.get(col, ""))
        if val:
            _val_text(gv, attr_id, val)

    # AT_PrincipalAgeDescription: fixed default "Adults" — not sourced from
    # any input column (per user direction, 2026-08-19).
    _val_text(gv, "AT_PrincipalAgeDescription", "Adults")

    # AT_SAPAge / AT_CountryOrigin: fixed LOV defaults, not sourced from
    # any input column (per user direction, 2026-08-19).
    _val_lov(gv, "AT_SAPAge", "AD")
    _val_lov(gv, "AT_CountryOrigin", "CN")
    _val_lov(gv, "AT_PricingDistributionChannel", "01")

    # AT_Gender / AT_BYGender: both derived from "Silhouette" — if the
    # first word starts with "W" (case-insensitive), Women -> "F";
    # otherwise Men -> "M" (per user direction, 2026-08-19: AT_BYGender
    # matches AT_Gender's logic exactly, no longer a flat "U" default).
    silhouette_col = loader.find_col("Silhouette")
    if silhouette_col:
        raw_silhouette = _s(row.get(silhouette_col, ""))
        if raw_silhouette:
            first_word = raw_silhouette.split()[0] if raw_silhouette.split() else ""
            gender_lov_id = "F" if first_word.upper().startswith("W") else "M"
            _val_lov(gv, "AT_Gender", gender_lov_id)
            _val_lov(gv, "AT_BYGender", gender_lov_id)

            # AT_Silhouette: Silhouette -> New Era MD Mappings (B/C/H, C=Apparel)
            # -> MDD Silhouette LOV (B->A), with single-digit ID zero-padding.
            silhouette_lov_id = _resolve_at_silhouette_lov_id(raw_silhouette, md_mapping, mdd)
            if silhouette_lov_id:
                _val_lov(gv, "AT_Silhouette", silhouette_lov_id)

            # AT_Fit: Silhouette -> New Era MD Mappings (B/C/I, C=Apparel)
            # -> MDD Fit LOV (B->A).
            fit_lov_id = _resolve_at_fit_lov_id(raw_silhouette, md_mapping, mdd)
            if fit_lov_id:
                _val_lov(gv, "AT_Fit", fit_lov_id)

    # AT_FOB: "Price" column value minus 33% (per user direction,
    # 2026-08-19). No fallback — skipped if the raw value isn't numeric.
    price_col = loader.find_col("Price")
    if price_col:
        raw_price = _s(row.get(price_col, ""))
        if raw_price:
            try:
                fob_val = float(raw_price) * (1 - 0.33)
                _val_text(gv, "AT_FOB", f"{fob_val:.2f}")
            except ValueError:
                pass

    # AT_FOBCurrency / AT_MaterialType / AT_SAPArticleCategory / AT_UOM /
    # AT_BYArticleType / AT_BYAge: fixed LOV defaults, not sourced from any
    # input column (per user direction, 2026-08-19).
    _val_lov(gv, "AT_FOBCurrency", "USD")
    _val_lov(gv, "AT_MaterialType", "ZINA")
    _val_lov(gv, "AT_SAPArticleCategory", "01")
    _val_lov(gv, "AT_UOM", "EA")
    _val_lov(gv, "AT_BYArticleType", "Inline")
    _val_lov(gv, "AT_BYAge", "ADULT")

    # AT_Content: "Fabric Content" column's highest-percentage material,
    # matched against the MDD "Content LOV" tab (per user direction,
    # 2026-08-19 — replaces the earlier AT_Material mapping). No fallback —
    # skipped if no match is found.
    fabric_col = loader.find_col("Fabric Content")
    if fabric_col:
        raw_fabric = _s(row.get(fabric_col, ""))
        if raw_fabric:
            top_material = _top_material_text(raw_fabric)
            if top_material:
                content_lov_id = _resolve_material_lov_id(
                    top_material, mdd.lovs.get("ContentLOV") if mdd else None,
                )
                if content_lov_id:
                    _val_lov(gv, "AT_Content", content_lov_id)

    # AT_CountrySize / AT_EComAgesCategory / AT_ArticleStatus /
    # AT_SAPProductFlag / AT_SportsCategoryEN / AT_PackDetails: fixed LOV
    # defaults, not sourced from any input column (per user direction,
    # 2026-08-19).
    _val_lov(gv, "AT_CountrySize", "INT")
    _val_lov(gv, "AT_EComAgesCategory", "18+Y")
    _val_lov(gv, "AT_ArticleStatus", "A")
    _val_lov(gv, "AT_SAPProductFlag", "A")
    _val_lov(gv, "AT_SportsCategoryEN", "06")
    # AT_PrincipalSize: comma-joined offered-size labels resolved by
    # _read_size_grid_fills (columns O–AN, no fill = offered).
    _val_text(gv, "AT_PrincipalSize", row.get(_SIZE_ROW_KEY, ""))

    # ── System-level fields: supplied from run() args, not from the
    #    dynamic per-row engine — same convention as every other brand.
    _multival(gv, "AT_CompanyCode", comp_code)
    _multival(gv, "AT_SBU", sbu)
    if country_code:
        _val_lov(gv, "AT_Country", country_code.upper())

    # AT_RetailPriceCurrency: derived from country_code, not any input column.
    currency_lov_id = _resolve_currency_from_country(country_code)
    if currency_lov_id:
        _val_lov(gv, "AT_RetailPriceCurrency", currency_lov_id)

    sea_code = season[:2].upper() if len(season) >= 2 else season
    _val_lov(gv, "AT_Season", sea_code)
    year_match = re.search(r"20\d{2}", season)
    season_year = year_match.group() if year_match else ""
    if not season_year:
        tail = re.search(r"\d{2}$", season)
        season_year = f"20{tail.group()}" if tail else ""
    if season_year:
        _val_text(gv, "AT_SeasonYear", season_year)

    # ── Brand / Brand Type / Brand Status / Brand Group ──────────
    # Resolved via MDD Brand LOV + RNA (Source Mapping related RNA tab),
    # same convention as anta/linelist_main.py.
    _val_lov(gv, "AT_Brand", brand_lov_id or brand_code)

    if brand_type:
        bt_lov_id = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type.upper(), "")
        if bt_lov_id:
            _val_lov(gv, "AT_BrandType", bt_lov_id)
        else:
            _val_text(gv, "AT_BrandType", brand_type)

    if brand_status:
        bs_lov_id = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status.upper(), "")
        if bs_lov_id:
            _val_lov(gv, "AT_BrandStatus", bs_lov_id)
        else:
            _val_text(gv, "AT_BrandStatus", brand_status)

    if brand_group:
        bg_lov_id = (mdd.lovs.get("AT_BrandGroup", {}) if mdd else {}).get(brand_group.upper(), "")
        if bg_lov_id:
            _val_lov(gv, "AT_BrandGroup", bg_lov_id)
        else:
            _val_text(gv, "AT_BrandGroup", brand_group)

    # ── Dynamic per-row attributes — driven entirely by the "New Era(Inline)"
    #    mapping tab.
    for rule in rules:
        if rule.attr_id in SPECIAL_ATTRIBUTE_IDS:
            continue
        col = _resolve_source_col(loader, rule)
        if not col:
            continue
        value = rule.resolve(row.get(col, ""))
        if value is None:
            continue
        if rule.is_lov:
            _val_lov(gv, rule.attr_id, value)
        else:
            _val_text(gv, rule.attr_id, value)

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    brand        = getattr(args, "brand", BRAND_NAME) or BRAND_NAME
    brand_code   = getattr(args, "brand_code", BRAND_CODE) or BRAND_CODE
    comp_code    = getattr(args, "comp_code", "0888") or "0888"
    sbu          = getattr(args, "sbu", "SP") or "SP"
    season       = getattr(args, "season", "SS27") or "SS27"
    country_code = getattr(args, "country_code", "") or ""

    log.info("[NewEra-OrderFormApp] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    # ── Load MDD (Brand / BrandType / BrandStatus / BrandGroup / Country LOVs) ──
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[NewEra-OrderFormApp] No MDD file found in %s — LOV lookups disabled", MDD_DIR)

    # ── Load attribute mapping rules from the "New Era(Inline)" tab ───
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    if not mapping_file:
        raise FileNotFoundError(
            f"No brand-mapping workbook found in {ATTR_DIR} — cannot load "
            f"the '{MAPPING_SHEET_NAME}' attribute rules"
        )
    attr_map = AttributeMappingLoader(mapping_file, MAPPING_SHEET_NAME)
    if not attr_map.rules:
        log.warning("[NewEra-OrderFormApp] No active attribute rules found — nothing will be sent")

    # ── RNA lookup: brand_type, brand_status, brand_group ─────────────
    rna = RNALoader(mapping_file)
    md_mapping = NewEraMDMappingLoader(mapping_file)

    brand_lov_id = brand_code
    if mdd:
        looked_up = mdd.lovs.get("AT_Brand", {}).get(brand.upper())
        brand_lov_id = looked_up if looked_up else brand_code

    rna_country = mdd.country_id_to_name(country_code) if mdd and country_code else ""
    if not rna_country:
        rna_country = country_code

    brand_type = brand_status = brand_group = ""
    rna_result = rna.get(rna_country, comp_code, sbu, brand_lov_id)
    if not rna_result.get("brand_type"):
        rna_result = rna.get_fuzzy(rna_country, sbu, brand_lov_id)
    if not rna_result.get("brand_type"):
        rna_result = rna.get_fuzzy(rna_country, sbu, brand_code)
    if not rna_result.get("brand_type"):
        rna_result = rna.get_by_brand_only(brand_lov_id) or rna.get_by_brand_only(brand_code)
    brand_type   = rna_result.get("brand_type", "")
    brand_status = rna_result.get("brand_status", "")
    brand_group  = rna_result.get("brand_group", "")

    # ── Load the New Era Order Form App input file ─────────────────────
    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        raise FileNotFoundError(f"No New Era Order Form App file found in {INPUT_DIR}")

    needed_columns = (
        [r.source_column for r in attr_map.rules]
        + list(COLUMN_FALLBACKS.values())
        + list(FORCED_DIRECT_COLUMNS.values())
        + ["MI#", "Description", "Price", "Fabric Content"]
    )
    loader = NewEraOrderFormLoader(input_file, needed_columns, INPUT_SHEET_NAME)
    if not loader.rows:
        log.warning("[NewEra-OrderFormApp] No data rows found — nothing to process")
        return

    # AT_PrincipalStyleCode: grouping key comes straight from "MI#" —
    # bypasses the mapping tab entirely, same as build_generic_product()
    # (see FORCED_DIRECT_COLUMNS' comment).
    style_col = loader.find_col("MI#")
    if not style_col:
        raise ValueError(
            f"[NewEra-OrderFormApp] Could not find column 'MI#' "
            f"(AT_PrincipalStyleCode source) in sheet '{INPUT_SHEET_NAME}'. "
            f"Headers found: {loader.headers}"
        )

    generics = group_rows(loader.rows, style_col)
    if not generics:
        log.warning("[NewEra-OrderFormApp] No valid generics produced")
        return

    sp_code  = season[:2].upper() if len(season) >= 2 else season
    sys_part = season[2:] if len(season) > 2 else ""
    sy       = f"20{sys_part}" if len(sys_part) == 2 else sys_part
    season_id = f"CLH_{brand_code}_{sp_code}{sy}"

    xml_filename = f"{input_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = 0

    cls_el  = build_classifications(brand, brand_code, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))

    log.info("[NewEra-OrderFormApp] Writing XML -> %s", xml_filename)
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
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")

        for row in generics.values():
            px = build_generic_product(
                row, loader, attr_map.rules, brand, brand_code,
                comp_code, sbu, season, season_id, country_code,
                brand_lov_id=brand_lov_id, brand_type=brand_type,
                brand_status=brand_status, brand_group=brand_group,
                mdd=mdd, md_mapping=md_mapping,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[NewEra-OrderFormApp] XML written: %s (%d KB)", xml_path.name, file_kb)
    print(
        f"=== NEW ERA ORDER FORM APP SUMMARY ===\n"
        f"  Input rows : {len(loader.rows)}\n"
        f"  Generics   : {generic_count}\n"
        f"  XML size   : {file_kb} KB",
        flush=True,
    )

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return xml_path


# ======================================================================
# FILENAME METADATA PARSER
# ======================================================================

def _parse_metadata_from_filename(dirs: dict, triggered_file_type: str = None) -> dict:
    """
    Parse comp_code, sbu, brand, brand_code, season, seq from the
    triggering file's filename.

    Expected format (dash-separated), same convention as Anta/Crocs:
        {CompanyCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Country}-{Seq}
    e.g. "0888-SP-NEW ERA-Copy of Order form - APAC FW26 (APP) 1
          (MAA Sport) Inline-Multi-SP2029-ID-1.xlsx"

    NOTE: the FileType segment above itself contains " - " (a dash), so a
    plain split can yield MORE than 8 tokens — season is located by regex
    search rather than by a fixed positional index, same robust approach
    used by reebok/ean_main.py's parser. brand_code is always the fixed
    BRAND_CODE constant (never extracted from the filename): "New Era"'s
    own brand token would otherwise false-match the "embedded 3-letter
    brand code" pattern other brands' parsers use (parts[2]="NEW ERA" ->
    "ERA" looks like a 3-letter code, but isn't one).
    """
    candidate_dirs = []
    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs.append(dirs[triggered_file_type])
    candidate_dirs.append(dirs.get("orderform_app"))

    seen: set = set()
    candidate_dirs = [
        d for d in candidate_dirs
        if d is not None and str(d) not in seen and not seen.add(str(d))
    ]

    for folder in candidate_dirs:
        if not folder or not folder.exists():
            continue
        xls_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm"))
        for xlsx_file in xls_files:
            stem = xlsx_file.stem

            parts = re.split(r"\s*-\s*", stem)
            if len(parts) < 6:
                log.warning(
                    "  Filename '%s' has only %d '-'-separated parts (need ≥6) — skipping",
                    stem, len(parts),
                )
                continue

            comp_code = parts[0].strip()
            sbu       = parts[1].strip()
            if not comp_code:
                log.warning("  comp_code empty — skipping: '%s'", stem)
                continue

            brand = parts[2].strip().title() if len(parts) > 2 else BRAND_NAME

            # ── Locate season token by regex ─────────────────────
            season     = None
            season_idx = None
            for i, p in enumerate(parts):
                if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
                    season     = p.strip().upper()
                    season_idx = i
                    break

            if not season or season_idx is None:
                log.warning("  No valid season token in '%s' — skipping", stem)
                continue

            trailing = parts[season_idx + 1:]

            seq = 1
            for p in reversed(trailing):
                if p.strip().isdigit():
                    seq = int(p.strip())
                    break

            country_code = ""
            for p in trailing:
                p = p.strip()
                if re.match(r"^[A-Z]{2,3}$", p, re.IGNORECASE) and not p.isdigit():
                    country_code = p.upper()
                    break

            return {
                "comp_code":    comp_code,
                "sbu":          sbu,
                "brand":        brand,
                "brand_code":   BRAND_CODE,
                "season":       season,
                "seq":          seq,
                "country_code": country_code,
            }

    log.warning("Could not parse metadata from filename — using defaults")
    return {
        "comp_code": "0888", "sbu": "SP", "brand": BRAND_NAME,
        "brand_code": BRAND_CODE, "season": "SS27", "seq": 1,
        "country_code": "",
    }


# ======================================================================
# CLI  (local test)
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="New Era Order Form App -> Stibo XML")
    parser.add_argument("--brand",        default=BRAND_NAME)
    parser.add_argument("--brand_code",   default=BRAND_CODE)
    parser.add_argument("--comp_code",    default="0888")
    parser.add_argument("--sbu",          default="SP")
    parser.add_argument("--season",       default="SS27")
    parser.add_argument("--country_code", default="")
    cli_args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    run(cli_args)
