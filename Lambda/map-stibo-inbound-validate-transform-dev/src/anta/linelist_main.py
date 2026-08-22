"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Anta Line List (Inline) v1.0     ║
║   Fully mapping-driven: no per-attribute Python code              ║
╚══════════════════════════════════════════════════════════════════╝

Input file : "Stibo Line List Template - Anta.xlsx"  (sheet "1. Template")

Unlike every other brand module in this repo, this ETL does NOT hardcode
which attributes to send or where they come from. Instead it reads the
"ANTA" tab of the shared brand-mapping workbook
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
                                     "Direct from Principal"       + Col J -> include, raw passthrough
                                     "Formula in System*"          + Col J -> include, raw passthrough
                                     "Mapping from Principal Data" + Col J -> include, value translated via Col K
                                     anything else (Manual Input, N/A,
                                     AI Translation, AI Image Analysis,
                                     Mapping from RNA/MD, or "Formula in
                                     System" with no Col J)              -> excluded
  J  Field Name in the Brand File — column to read on the "1. Template"
                                     sheet of the Anta input file.
  K  Mapping Logic              — for "Mapping from Principal Data" rows,
                                   parsed as newline-separated
                                   "source : id - label" entries
                                   (e.g. "M : M - Male") into a lookup
                                   table. A raw value with no match in the
                                   table is skipped for that attribute
                                   (no default guess — same no-fallback
                                   convention used for Clarks AT_Gender).

Adding/removing/re-pointing an attribute is a spreadsheet edit on the ANTA
tab — this file does not need to change.

System-level fields (Company, SBU, Country, Season, Season Year) are marked
"Manual Input in Portal" in the mapping sheet — i.e. not sourced from the
brand file at all — so they are supplied structurally from run() args /
filename metadata, exactly like every other brand module, not from the
dynamic per-row engine below.

AT_Brand / AT_BrandType / AT_BrandStatus / AT_BrandGroup are marked
"Formula in System" / "Mapping from RNA" with no Col J source, so they're
also outside the dynamic engine — resolved the same way as
clarks/order_form_footwear_main.py: AT_Brand from the MDD Brand LOV
(brand name -> LOV ID, e.g. "Anta" -> "ATA"), and Type/Status/Group from the
attributes workbook's "Source Mapping related RNA" tab (keyed by country +
comp_code + sbu + brand code, with the same 4-pass fallback: exact match ->
fuzzy without comp_code -> fuzzy with the raw brand_code -> brand-code-only),
each then translated through its own MDD LOV with a raw-text fallback if
unmapped. See MDDLoader / RNALoader below.

Grouping: one Generic Article per unique Style Code (the source column of
AT_PrincipalStyleCode). SAP Color ("AI Images Analysis") and SAP Size
("Formula in System" / "BY Feedback") have no source column in the mapping
sheet — there's currently no way to derive a 3-digit SAP color/size code
from the input file, so this module emits Generic Articles only, no
colour/size Variant sub-products (same simplification already used by
clarks/order_form_footwear_main.py).

ASSUMPTION flagged for confirmation: the Generic Article ParentID
(division routing — Footwear/Apparel/Accessories/Equipment) defaults to
Footwear (PPH_F-TempSubCat) below, since "SAP Product Division" is excluded
from the dynamic engine (its Stibo Attribute ID is #N/A in the mapping
sheet). Change DEFAULT_PRODUCT_PARENT_ID if Anta linelists are not
Footwear-only.
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

INPUT_DIR   = BASE_DIR / "input" / "linelist"
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Anta"
BRAND_CODE = "ATA"

MAPPING_SHEET_NAME = "ANTA"          # tab in the brand-mapping workbook
INPUT_SHEET_NAME   = "1. Template"   # sheet in the Anta linelist file

# See "ASSUMPTION flagged" note in the module docstring.
DEFAULT_PRODUCT_PARENT_ID = "PPH_F-TempSubCat"

# ParentID by Category column value.
CATEGORY_PARENT_IDS: dict[str, str] = {
    "FTW": "PPH_F-TempSubCat",
    "APP": "PPH_A-TempSubCat",
    "ACC": "PPH_E-TempSubCat",
}

# Attributes whose Col-J source column (from the ANTA mapping tab) has a
# known mismatch against the actual "1. Template" header — the mapping
# sheet's Col J is always tried first; if that column isn't found in the
# input file, this fallback column name is tried instead.
#   AT_PrincipalMerchandiseHierarchyL3: mapping sheet Col J says
#   "Sport Category", but the real input file's header is "Sports Category".
COLUMN_FALLBACKS: dict[str, str] = {
    "AT_PrincipalMerchandiseHierarchyL3": "Sports Category",
}

# Attribute IDs handled by dedicated logic in build_generic_product()
# instead of the generic per-row passthrough/mapped engine — excluded from
# the dynamic loop to avoid double-emission.
#   AT_IncomingMonth: needs a date reformat (input M/D/YYYY -> DD-MM-YYYY)
#   before being sent as a text value.
#   AT_EComAgesCategory: raw value looked up in the MDD "E-com Ages
#   Category LOV" tab (Col B display name -> Col A LOV id) instead of the
#   mapping sheet's raw Direct-from-Principal passthrough.
#   AT_SportsCategoryEN: raw value from "Sports Category" looked up in the
#   MDD "Sports Category LOV" tab (Col A display name -> Col B LOV id),
#   zero-padded to 2 digits when the id is a single digit (e.g. "2" -> "02").
#   AT_LaunchingDate: needs a date reformat (input M/D/YY(YY) -> DD-Mon-YYYY,
#   e.g. "9/23/26" -> "23-Sep-2026") before being sent as a text value.
#   AT_CountrySize: derived from "Category" column with fixed LOV IDs:
#     FTW -> US, APP -> ASIA, ACC -> skipped.
SPECIAL_ATTRIBUTE_IDS: set[str] = {
    "AT_IncomingMonth", "AT_EComAgesCategory", "AT_SportsCategoryEN", "AT_LaunchingDate",
    "AT_CountrySize",
}

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
    Same country-to-currency mapping as clarks/order_form_footwear_main.py.
    """
    country_to_currency = {
        "ID": "IDR",  # Indonesia -> Indonesian Rupiah
        "PH": "PHP",  # Philippines -> Philippine Peso
        "TH": "THB",  # Thailand -> Thailand Baht
        "SG": "SGD",  # Singapore -> Singapore Dollar
        "MY": "MYR",  # Malaysia -> Malaysian Ringgit
        "VN": "VND",  # Vietnam -> Vietnamese Dong
        "KH": "USD",  # Cambodia -> United States Dollar
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


def _format_date_ddmmyyyy(raw_value) -> str:
    """
    Format a date value as DD-MM-YYYY, for AT_IncomingMonth.

    Accepts a datetime/date object (openpyxl returns these for date-typed
    cells) or a string in M/D/YYYY format (e.g. "9/23/2026" or "9/23/2026."
    with a trailing period, as the Anta input file uses). Returns "" if the
    value can't be parsed as a date — no guessing.
    """
    if raw_value is None:
        return ""
    if hasattr(raw_value, "day") and hasattr(raw_value, "month") and hasattr(raw_value, "year"):
        return f"{raw_value.day:02d}-{raw_value.month:02d}-{raw_value.year}"
    s = str(raw_value).strip().rstrip(".").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if not m:
        return ""
    month, day, year = int(m.group(1)), int(m.group(2)), m.group(3)
    return f"{day:02d}-{month:02d}-{year}"


_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def _format_date_ddmonyyyy(raw_value) -> str:
    """
    Format a date value as DD-Mon-YYYY (e.g. "23-Sep-2026"), for
    AT_LaunchingDate.

    Accepts a datetime/date object (openpyxl returns these for date-typed
    cells) or a string in M/D/YYYY or M/D/YY format (e.g. "9/23/2026" or
    "9/23/26", optionally with a trailing period). Returns "" if the value
    can't be parsed as a date — no guessing.
    """
    if raw_value is None:
        return ""
    if hasattr(raw_value, "day") and hasattr(raw_value, "month") and hasattr(raw_value, "year"):
        return f"{raw_value.day:02d}-{_MONTH_ABBR[raw_value.month]}-{raw_value.year}"
    s = str(raw_value).strip().rstrip(".").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})$", s)
    if not m:
        return ""
    month, day, year = int(m.group(1)), int(m.group(2)), m.group(3)
    if len(year) == 2:
        year = f"20{year}"
    if month not in _MONTH_ABBR:
        return ""
    return f"{day:02d}-{_MONTH_ABBR[month]}-{year}"


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
# LOVs — same sheet-layout convention as clarks/order_form_footwear_main.py)
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
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        self._load_brand_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_brand_group_lov(wb)
        self._load_country_lov(wb)
        self._load_ecom_ages_category_lov(wb)
        self._load_sports_category_lov(wb)
        wb.close()
        log.info("[MDD] %d LOV types loaded", len(self.lovs))

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
        log.info("[MDD] Brand LOV: %d entries", len(lov))

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
        log.info("[MDD] Brand Type LOV: %d entries", len(lov))

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
        log.info("[MDD] Brand Status LOV: %d entries", len(lov))

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
        log.info("[MDD] Brand Group LOV: %d entries", len(lov))

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
        log.info("[MDD] Country LOV: %d entries", len(id_to_name))

    def country_id_to_name(self, country_id: str) -> str:
        return self.lovs.get("CountryLOV", {}).get((country_id or "").strip().upper(), "")

    def _load_ecom_ages_category_lov(self, wb) -> None:
        """E-com Ages Category LOV: col A = LOV ID (Code), col B = display name."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "E-COM AGES CATEGORY LOV"),
            next(
                (s for s in wb.sheetnames
                 if "AGES" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()),
                None,
            ),
        )
        if not sheet:
            log.warning("[MDD] E-com Ages Category LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_EComAgesCategory"] = lov
        log.info("[MDD] E-com Ages Category LOV: %d entries", len(lov))

    def _load_sports_category_lov(self, wb) -> None:
        """Sports Category LOV: col A = display name, col B = LOV ID."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "SPORTS CATEGORY LOV"),
            next(
                (s for s in wb.sheetnames
                 if "SPORTS" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()),
                None,
            ),
        )
        if not sheet:
            log.warning("[MDD] Sports Category LOV sheet not found")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            if lov_id.endswith(".0"):
                lov_id = lov_id[:-2]
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["AT_SportsCategoryEN"] = lov
        log.info("[MDD] Sports Category LOV: %d entries", len(lov))


# ======================================================================
# RNA LOADER  (Brand Type / Brand Status / Brand Group — same
# "Source Mapping related RNA" sheet + 4-pass lookup used by Clarks)
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
        log.info("[RNA] %d entries loaded from '%s'", count, sheet_name)

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

def _parse_mapping_logic(raw_logic) -> dict[str, str]:
    """
    Parse Col K "Mapping Logic" text into {SOURCE_UPPER: resolved_id_or_text}.

    Expected line format: "<source> : <target_id> - <target_label>"
      e.g. "M : M - Male"   -> {"M": "M"}
           "W : F - Female" -> {"W": "F"}
    Lines that don't contain ":" (e.g. free-text notes like
    "Mapping from Principal Data") are ignored — callers fall back to raw
    passthrough when the resulting table is empty.
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
        # "F - Female" -> take the id token before " - "
        tgt_id = tgt.split(" - ", 1)[0].strip() or tgt
        mapping[src.upper()] = tgt_id
    return mapping


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

        mt = mapping_type.strip()
        if mt == "Mapping from Principal Data":
            self.mode = "mapped"
            self.value_map = _parse_mapping_logic(mapping_logic_raw)

    def resolve(self, raw_value: str) -> str | None:
        """Return the value to write for this attribute, or None to skip it."""
        raw = _s(raw_value)
        if not raw:
            return None
        if self.mode == "direct" or not self.value_map:
            return raw
        # "mapped" mode with a real lookup table — no default guess.
        resolved = self.value_map.get(raw.upper())
        if resolved is None:
            log.info(
                "[AttrMap] '%s' (%s): raw value '%s' not in Col K mapping table "
                "%s — attribute skipped for this row",
                self.name, self.attr_id, raw, sorted(self.value_map.keys()),
            )
        return resolved


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

        skipped_no_id = 0
        skipped_type  = 0
        for row in rows[hdr_idx + 1:]:
            if not row or not row[0]:
                continue
            name          = _s(row[0])
            attr_id       = _s(row[1]) if len(row) > 1 else ""
            validation    = _s(row[2]) if len(row) > 2 else ""
            mapping_type  = _s(row[6]) if len(row) > 6 else ""
            source_column = _s(row[9]) if len(row) > 9 else ""
            mapping_logic = row[10] if len(row) > 10 else ""

            if not attr_id:
                skipped_no_id += 1
                continue

            mt = mapping_type.strip()
            include = (
                (mt == "Direct from Principal" and source_column)
                or (mt.startswith("Formula in System") and source_column)
                or (mt == "Mapping from Principal Data" and source_column)
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
            "(%d skipped: no valid attribute ID, %d skipped: mapping type not eligible)",
            len(self.rules), self.sheet_name, skipped_no_id, skipped_type,
        )
        for r in self.rules:
            log.info(
                "[AttrMap]   %-38s -> %-32s  col='%s'  %s  mode=%s",
                r.name, r.attr_id, r.source_column,
                "LOV(id)" if r.is_lov else "text(value)", r.mode,
            )

    def find(self, attr_id: str) -> AttributeRule | None:
        return next((r for r in self.rules if r.attr_id == attr_id), None)


# ======================================================================
# INPUT FILE LOADER  ("1. Template" sheet — header row/columns discovered
# dynamically from the attribute rules' Col J source-column names)
# ======================================================================

class AntaLinelistLoader:
    def __init__(self, path: Path, needed_columns: list[str], sheet_name: str = INPUT_SHEET_NAME):
        self.path = path
        self.sheet_name = sheet_name
        self.headers: list[str] = []
        self.rows: list[dict] = []
        self._needed = {_norm_header(c) for c in needed_columns if c}
        self._load()

    def _load(self) -> None:
        log.info("[Anta-Linelist] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        target = next(
            (s for s in wb.sheetnames if s.strip().upper() == self.sheet_name.upper()),
            None,
        )
        if not target:
            target = next((s for s in wb.sheetnames if "template" in s.lower()), None)
        if not target:
            wb.close()
            raise ValueError(
                f"[Anta-Linelist] Sheet '{self.sheet_name}' not found. "
                f"Available: {wb.sheetnames}"
            )
        log.info("[Anta-Linelist] Using sheet: '%s'", target)

        raw_rows = list(wb[target].iter_rows(values_only=True))
        wb.close()

        if not raw_rows:
            log.warning("[Anta-Linelist] Sheet '%s' is empty", target)
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
                "[Anta-Linelist] Could not confidently locate header row "
                "(no cell matched any expected column name) — defaulting to row 0"
            )
        log.info(
            "[Anta-Linelist] Header at row %d (0-indexed), matched %d/%d expected columns",
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

        for r in raw_rows[best_idx + 1:]:
            if not r or all(c is None for c in r):
                continue
            self.rows.append({header[i]: (r[i] if i < len(r) else None) for i in range(len(header))})

        log.info("[Anta-Linelist] Loaded %d data rows", len(self.rows))

    def find_col(self, name: str) -> str | None:
        """Case/whitespace-insensitive header lookup."""
        target = _norm_header(name)
        for h in self.headers:
            if _norm_header(h) == target:
                return h
        return None


def _resolve_source_col(loader: "AntaLinelistLoader", rule: "AttributeRule") -> str | None:
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
    if col:
        log.info(
            "[AttrMap] '%s': Col-J column '%s' not found in input file — "
            "using fallback column '%s'",
            rule.attr_id, rule.source_column, fallback_name,
        )
    return col


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

    season_id      = f"CLH_{brand_code}_{full_season}"
    batches_parent = f"CLH_{brand}Batches"
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
    loader: AntaLinelistLoader,
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
) -> str:
    style_rule = next((r for r in rules if r.attr_id == "AT_PrincipalStyleCode"), None)
    style_col  = _resolve_source_col(loader, style_rule) if style_rule else None
    style_code = _s(row.get(style_col, "")) if style_col else ""
    if not style_code:
        return ""

    generic_code = f"{brand_code}{style_code}"

    category_col = loader.find_col("Category")
    category_val = _s(row.get(category_col, "")).upper() if category_col else ""
    parent_id = CATEGORY_PARENT_IDS.get(category_val, DEFAULT_PRODUCT_PARENT_ID)

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code

    # Name: prefer AT_PrincipalStyleDescription's resolved value, else style code.
    name_text = generic_code
    desc_rule = next((r for r in rules if r.attr_id == "AT_PrincipalStyleDescription"), None)
    if desc_rule:
        desc_col = _resolve_source_col(loader, desc_rule)
        desc_val = desc_rule.resolve(row.get(desc_col, "")) if desc_col else None
        if desc_val:
            name_text = desc_val
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name_text

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # AT_InboundGenericCode: same value as the KEY_InboundArticle KeyValue
    # above — not in the ANTA mapping tab (no Col J source), populated
    # structurally the same way clarks/order_form_footwear_main.py does
    # (generic["AT_InboundGenericCode"] = generic_key).
    _val_text(gv, "AT_InboundGenericCode", generic_code)

    # AT_SAPAge: mapping sheet says "Default : Adults" with no Col J source
    # — same fixed default clarks/order_form_footwear_main.py sends ("AD").
    _val_lov(gv, "AT_SAPAge", "AD")

    # AT_CountryOrigin: mapping sheet says "Default : China" with no Col J
    # source — fixed default, not derived from any input column.
    _val_lov(gv, "AT_CountryOrigin", "CN")

    # AT_PricingDistributionChannel: mapping sheet says "Default : Retailer"
    # with no Col J source — fixed default, same "01" every other brand uses.
    _val_lov(gv, "AT_PricingDistributionChannel", "01")

    # AT_SAPProductFlag: mapping sheet says "Default : A - Intercompany" with
    # no Col J source — fixed default, same "A" every other brand uses.
    _val_lov(gv, "AT_SAPProductFlag", "A")

    # AT_MaterialType: mapping sheet says "Default : Intercompany" with no
    # Col J source — fixed default, same "ZINA" every other brand uses.
    _val_lov(gv, "AT_MaterialType", "ZINA")

    # AT_SAPArticleCategory: mapping sheet says "Default : Generic" with no
    # Col J source — fixed default, same "01" every other brand uses.
    _val_lov(gv, "AT_SAPArticleCategory", "01")

    # AT_UOM: mapping sheet says "Default : EA" with no Col J source —
    # fixed default, same "EA" every other brand uses.
    _val_lov(gv, "AT_UOM", "EA")

    # AT_BYArticleType: not on the ANTA tab under a valid attribute ID
    # ("Article Type" row has attr_id #N/A) — same unconditional "Inline"
    # default clarks/order_form_footwear_main.py sends.
    _val_lov(gv, "AT_BYArticleType", "Inline")

    # AT_BYAge: mapping sheet says "Default : Adults" with no Col J source
    # — same fixed default clarks/order_form_footwear_main.py sends ("ADULT").
    _val_lov(gv, "AT_BYAge", "ADULT")

    # AT_ArticleStatus: mapping sheet says "Default when created: Active"
    # with no Col J source — fixed default, same "A" every other brand uses.
    _val_lov(gv, "AT_ArticleStatus", "A")

    # ── System-level fields: "Manual Input in Portal" per the mapping sheet
    #    — supplied from run() args, not from the dynamic per-row engine.
    _multival(gv, "AT_CompanyCode", comp_code)
    _multival(gv, "AT_SBU", sbu)
    if country_code:
        _val_lov(gv, "AT_Country", country_code.upper())

    # AT_RetailPriceCurrency: derived from country_code, not any input
    # column — same country-to-currency map as Clarks footwear.
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
    # "Formula in System" / "Mapping from RNA" per the mapping sheet (no Col J
    # source) — resolved the same way as clarks/order_form_footwear_main.py:
    # AT_Brand from the MDD Brand LOV; Type/Status/Group from RNA, with each
    # RNA value looked up in its own MDD LOV and a raw-text fallback if the
    # LOV has no match.
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

    # AT_CountrySize: derived from input "Category" column.
    #   FTW -> US
    #   APP -> ASIA
    #   ACC -> skip attribute
    category_col = loader.find_col("Category")
    if category_col:
        category_val = _s(row.get(category_col, "")).upper()
        if category_val == "FTW":
            _val_lov(gv, "AT_CountrySize", "US")
        elif category_val == "APP":
            _val_lov(gv, "AT_CountrySize", "ASIA")

    # ── Dynamic per-row attributes — driven entirely by the ANTA mapping tab.
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

    # AT_IncomingMonth: same source column as AT_LaunchingDate ("Launching
    # Date"), reformatted from the input's M/D/YYYY (e.g. "9/23/2026.") to
    # DD-MM-YYYY and sent as a text value.
    incoming_month_rule = next((r for r in rules if r.attr_id == "AT_IncomingMonth"), None)
    if incoming_month_rule:
        im_col = _resolve_source_col(loader, incoming_month_rule)
        if im_col:
            formatted = _format_date_ddmmyyyy(row.get(im_col))
            if formatted:
                _val_text(gv, "AT_IncomingMonth", formatted)

    # AT_LaunchingDate: reformatted from the input's M/D/YY(YY) (e.g.
    # "9/23/26") to DD-Mon-YYYY (e.g. "23-Sep-2026") and sent as a text
    # value, instead of the mapping sheet's raw Direct-from-Principal
    # passthrough (which would send the input's raw datetime string).
    launching_date_rule = next((r for r in rules if r.attr_id == "AT_LaunchingDate"), None)
    if launching_date_rule:
        ld_col = _resolve_source_col(loader, launching_date_rule)
        if ld_col:
            formatted = _format_date_ddmonyyyy(row.get(ld_col))
            if formatted:
                _val_text(gv, "AT_LaunchingDate", formatted)

    # AT_EComAgesCategory: raw value from "Ages Category" looked up in the
    # MDD "E-com Ages Category LOV" tab (Col B display name -> Col A LOV
    # id). No fallback — skipped if the raw value isn't in the LOV.
    ecom_ages_rule = next((r for r in rules if r.attr_id == "AT_EComAgesCategory"), None)
    if ecom_ages_rule:
        ea_col = _resolve_source_col(loader, ecom_ages_rule)
        if ea_col:
            raw_val = _s(row.get(ea_col, ""))
            if raw_val:
                lov_id = (mdd.lovs.get("AT_EComAgesCategory", {}) if mdd else {}).get(raw_val.upper(), "")
                if lov_id:
                    _val_lov(gv, "AT_EComAgesCategory", lov_id)
                else:
                    log.info(
                        "[EComAgesCategory] raw value '%s' not found in MDD LOV — attribute skipped",
                        raw_val,
                    )

    # AT_SportsCategoryEN: raw value from "Sports Category" looked up in the
    # MDD "Sports Category LOV" tab (Col A display name -> Col B LOV id),
    # zero-padded to 2 digits when the id is a single digit. No fallback —
    # skipped if the raw value isn't in the LOV.
    sports_cat_rule = next((r for r in rules if r.attr_id == "AT_SportsCategoryEN"), None)
    if sports_cat_rule:
        sc_col = _resolve_source_col(loader, sports_cat_rule)
        if sc_col:
            raw_val = _s(row.get(sc_col, ""))
            if raw_val:
                lov_id = (mdd.lovs.get("AT_SportsCategoryEN", {}) if mdd else {}).get(raw_val.upper(), "")
                if lov_id:
                    if len(lov_id) == 1 and lov_id.isdigit():
                        lov_id = f"0{lov_id}"
                    _val_lov(gv, "AT_SportsCategoryEN", lov_id)
                else:
                    log.info(
                        "[SportsCategoryEN] raw value '%s' not found in MDD LOV — attribute skipped",
                        raw_val,
                    )

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

    log.info("[Anta-Linelist] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    # ── Load MDD (Brand / BrandType / BrandStatus / BrandGroup / Country LOVs) ──
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Anta-Linelist] No MDD file found in %s — LOV lookups disabled", MDD_DIR)

    # ── Load attribute mapping rules from the ANTA tab ────────────────
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    if not mapping_file:
        raise FileNotFoundError(
            f"No brand-mapping workbook found in {ATTR_DIR} — cannot load "
            f"the '{MAPPING_SHEET_NAME}' attribute rules"
        )
    attr_map = AttributeMappingLoader(mapping_file, MAPPING_SHEET_NAME)
    if not attr_map.rules:
        log.warning("[Anta-Linelist] No active attribute rules found — nothing will be sent")

    # ── RNA lookup: brand_type, brand_status, brand_group ─────────────
    # Same file as the attribute mapping (mapping_file) — "Source Mapping
    # related RNA" is a separate tab in the same brand-mapping workbook.
    rna = RNALoader(mapping_file)

    brand_lov_id = brand_code
    if mdd:
        looked_up = mdd.lovs.get("AT_Brand", {}).get(brand.upper())
        brand_lov_id = looked_up if looked_up else brand_code
        log.info(
            "[Anta-Linelist] Brand LOV lookup: key='%s' -> LOV ID='%s' %s",
            brand.upper(), brand_lov_id,
            "(FOUND)" if looked_up else f"(NOT FOUND -- fallback to '{brand_code}')",
        )

    rna_country = mdd.country_id_to_name(country_code) if mdd and country_code else ""
    if not rna_country:
        rna_country = country_code

    brand_type = brand_status = brand_group = ""
    log.info(
        "[RNA] Searching with: country='%s' comp='%s' sbu='%s' brand='%s'",
        rna_country, comp_code, sbu, brand_lov_id,
    )
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
    log.info(
        "[RNA] Resolved -- brand_type='%s' brand_status='%s' brand_group='%s'",
        brand_type, brand_status, brand_group,
    )

    # ── Load the Anta linelist input file ──────────────────────────────
    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        raise FileNotFoundError(f"No Anta linelist file found in {INPUT_DIR}")

    needed_columns = [r.source_column for r in attr_map.rules] + list(COLUMN_FALLBACKS.values())
    loader = AntaLinelistLoader(input_file, needed_columns, INPUT_SHEET_NAME)
    if not loader.rows:
        log.warning("[Anta-Linelist] No data rows found — nothing to process")
        return

    style_rule = attr_map.find("AT_PrincipalStyleCode")
    if not style_rule:
        raise ValueError(
            "[Anta-Linelist] ANTA mapping tab has no active 'AT_PrincipalStyleCode' "
            "rule — cannot group rows into Generic Articles"
        )
    style_col = _resolve_source_col(loader, style_rule)
    if not style_col:
        raise ValueError(
            f"[Anta-Linelist] Could not find column '{style_rule.source_column}' "
            f"(AT_PrincipalStyleCode source) in sheet '{INPUT_SHEET_NAME}'. "
            f"Headers found: {loader.headers}"
        )

    generics = group_rows(loader.rows, style_col)
    if not generics:
        log.warning("[Anta-Linelist] No valid generics produced")
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

    log.info("[Anta-Linelist] Writing XML -> %s", xml_filename)
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
                brand_status=brand_status, brand_group=brand_group, mdd=mdd,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[Anta-Linelist] XML written: %s (%d KB)", xml_path.name, file_kb)
    print(
        f"=== ANTA LINELIST SUMMARY ===\n"
        f"  Input rows : {len(loader.rows)}\n"
        f"  Generics   : {generic_count}\n"
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
    parser = argparse.ArgumentParser(description="Anta Line List -> Stibo XML")
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
