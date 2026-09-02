"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Reebok Apparel v1.0             ║
║   Article Master (NuORDER export) → Stibo STEP XML              ║
║   (Generic Articles only — variant explosion tracked for audit) ║
╚══════════════════════════════════════════════════════════════════╝

Source file
    "Reebok article master data with image - APPAREL (SS27).xlsx"
    Sheet : "NuORDER Order Data"  (header on row 1; one row per SKU/size)

This file is a flat per-SKU export (Article Number + Color Name + Size +
UPC per row), unlike the order-form style inputs used by other brands.
Rows are grouped back into Generic articles keyed by
`{brand_code}{Article Number}` — matching the "Reebok(Inline)" tab of the
shared brand-mapping workbook, which maps:

    Stibo Attribute              ← Brand File Column
    ────────────────────────────────────────────────
    Principal Style Code         ← Article Number
    Principal Style Description  ← Name  (strip "Reebok"/"REE" tokens)
    Principal Color Description  ← Color Name
    Principal Gender Description ← Gender
    Principal Age Description    ← Age Group
    Principal Size               ← Size Range
    Country of Origin            ← T1 Supplier Country Description
    Principal Merch Hierarchy L1 ← Department
    Principal Merch Hierarchy L2 ← Sports Category
    Principal Merch Hierarchy L3 ← Article Business Segment
    Principal Merch Hierarchy L4 ← Category Marketing Line
    Sports Category EN           ← Sports Category  (via "Reebok MD
                                   Mappings sheet 2" → MDD "Sports Category LOV")
    Content                      ← Material Composition (fibre families →
                                   MDD "Content LOV")
    Fabric                       ← Material Composition (construction
                                   keyword → MDD "Fabric LOV")
    Country Size                 ← "Formula in System": FW → US, APP → Asia
    Launching Date               ← Retail Intro Date
    FOB                          ← Customer Price (USD) (USD)
    Original / Current Price     ← M.S.R.P (USD)
    SAP Age / BY Age             ← Age Group   (via MDD "Age LOV")
    SAP Gender / BY Gender       ← Gender      (via MDD "Gender LOV")

Fields that require inputs not present in this repo (SAP 103 color code
via "AI Image Analyst", SAP size code via BY feedback, the brand-specific
"Mapping Attributes Reebok.xlsx" for SAP Product Division/Group/Category
and Silhouette/Fabric) are intentionally left out of the XML rather than
guessed — extend `build_product_xml()` once those lookup tables exist.

Scope note (matches current Asics-footwear precedent in this codebase):
    Only one <Product UserTypeID="PRD_GenericArticle"> element is emitted
    per generic article. Per-SKU variant rows (color+size / UPC) are
    grouped and counted for audit purposes but not written as individual
    <Product UserTypeID="PRD_VariantArticle"> elements — AT_EComSize was
    the one variant-level attribute this feed carried and MD has since
    dropped it from scope, and the rest (SAP 103 Color via "AI Image
    Analyst", AT_Size/AT_SAPSize via BY feedback) have no input here.
    Barcodes/variants for Reebok come from reebok/ean_main.py instead.
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
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", Path(__file__).resolve().parent))

INPUT_DIR   = BASE_DIR / "input" / "apparel"
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (INPUT_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Reebok"
BRAND_CODE = "REE"   # 3-letter SAP/MDD brand code for Reebok (Brand LOV)

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')

# ======================================================================
# STATIC LOOKUP TABLES (best-effort — used when the MDD sheet lacks a match)
# ======================================================================

# Age Group (Reebok file, e.g. "Adult", "Kids", "Junior") → (AT_SAPAge, AT_BYAge)
# codes. NOT sourced from the MDD "Age LOV" sheet — those columns don't hold
# valid live-Stibo LOV IDs. Confirmed against the brand mapping template's own
# "Reebok MD Mappings sheet 1" tab (Reebok raw Age Group → Stibo codes:
# Adult→AD-Adults, Infant→CH-Children/Infant, CHILDRENS→CH-Children/Kids,
# Junior→CH-Children/Grade School, Kids→CH-Children/Kids), cross-checked
# against nike/orderConfirmation_main.py (NIKE_AGE_DESC_RULES, same MDD,
# confirmed against live Stibo): AT_SAPAge (LOV_Age) only accepts AD / CH / AA
# in all supported environments; AT_BYAge (LOV_BYAge) only accepts ADULT /
# ALL AGES / GRADE SCHOOL / INFANT / KIDS / PRESCHOOL.
AGE_GROUP_RULES: dict[str, tuple[str, str]] = {
    "ADULT":        ("AD", "ADULT"),
    "ADULTS":       ("AD", "ADULT"),
    "KID":          ("CH", "KIDS"),
    "KIDS":         ("CH", "KIDS"),
    "CHILD":        ("CH", "KIDS"),
    "CHILDREN":     ("CH", "KIDS"),
    "CHILDRENS":    ("CH", "KIDS"),
    "JUNIOR":       ("CH", "GRADE SCHOOL"),
    "YOUTH":        ("CH", "KIDS"),
    "INFANT":       ("CH", "INFANT"),
    "INFANTS":      ("CH", "INFANT"),
    "TODDLER":      ("CH", "INFANT"),
    "BABY":         ("CH", "INFANT"),
    "PRESCHOOL":    ("CH", "PRESCHOOL"),
    "GRADE SCHOOL": ("CH", "GRADE SCHOOL"),
    "ALL AGES":     ("AA", "ALL AGES"),
}

# Age Group (Reebok file) → AT_EComAgesCategory LOV code. Per the brand
# mapping template's "Reebok MD Mappings sheet 3" tab (Age Group → E-com
# Ages Category display label), resolved against the MDD "E-com Ages
# Category LOV" sheet's codes. Only covers the Age Group values Reebok's
# own sheet actually maps — anything else is intentionally left unmapped
# rather than guessed.
EC_AGES_CATEGORY_MAP: dict[str, str] = {
    "CHILDRENS": "6-8Y",    # Ages 6-8 years
    "JUNIOR":    "13-17Y",  # Ages 13-17 years
    "ADULT":     "18+Y",    # Ages 18+ years
    "ADULTS":    "18+Y",    # Ages 18+ years
    "KIDS":      "13-17Y",  # Ages 13-17 years
    "KID":       "13-17Y",  # Ages 13-17 years
    "INFANT":    "18-24M",  # Ages 18-24 months
    "INFANTS":   "18-24M",  # Ages 18-24 months
}

# Reebok raw "Sports Category" → Stibo "Sports Category EN" display value,
# per the brand mapping template's "Reebok MD Mappings sheet 2" tab
# (Reebok Inline, source column "Sports Category"). The display value is
# then resolved to its AT_SportsCategoryEN LOV ID through the MDD
# "Sports Category LOV" sheet — see MDDLoader._load_sports_category_lov()
# and _sports_category_en_lov_id() below. Only the values Reebok's own
# sheet maps are listed; anything else is left unmapped rather than
# guessed (the LOV rejects unknown IDs on import).
SPORTS_CATEGORY_EN_MAP: dict[str, str] = {
    "RUNNING":         "Running",
    "TRAINING":        "Fitness / Training",
    "WALKING":         "Outdoor / Trail / Hiking",
    "TENNIS":          "Tennis / Padel",
    "BASKETBALL":      "Basketball",
    "CASUAL":          "Lifestyle / Casual",
    "SWIM":            "Swimming",
    "OUTDOOR":         "Outdoor / Trail / Hiking",
    "FOOTBALL/SOCCER": "Soccer",
    "GOLF":            "Golf",
    "SKATE":           "Lifestyle / Casual",
}

# AT_CountrySize — "Formula in System" per the Reebok(Inline) sheet, which
# spells the rule out per product type: "FW: US / APP: Asia / ACC: Manual
# Input". Apparel is a flat "Asia" → MDD "Country Size LOV" ID "ASIA"
# (the LOV's display value is literally "ASIA"). The ID is validated
# against that LOV before it's written. ("Reebok MD Mappings sheet 1" still
# shows an empty APP cell — the attribute row above is the current one.)
COUNTRY_SIZE = "ASIA"

# ── AT_Content (Apparel) ──────────────────────────────────────────────
# "Direct from Principal" ← "Material Composition" per the Reebok(Inline)
# sheet. AT_Content is an LOV, not free text (writing the raw composition
# string is what threw "Illegal LOV" on import before), and the MDD
# "Content LOV" sheet is a fibre-family list:
#     Cotton / Cotton Blend / Cotton Spandex / Nylon / Nylon Blend /
#     Poly Spandex / Polyester / Polyester Blend / Spandex /
#     Viscose / Viscose Blend
# So the composition is parsed into fibre percentages and reduced to the
# family that matches. Reebok writes both full names ("60% Cotton 40%
# Polyester") and ISO abbreviations ("87% PA, 13% EL"), so both are keyed.
CONTENT_FIBER_CODES: dict[str, str] = {
    "CO":  "COTTON",
    "CT":  "COTTON",
    "PES": "POLYESTER",
    "PL":  "POLYESTER",
    "PET": "POLYESTER",
    "PA":  "NYLON",
    "NY":  "NYLON",
    "EL":  "SPANDEX",
    "EA":  "SPANDEX",
    "SP":  "SPANDEX",
    "VI":  "VISCOSE",
    "CV":  "VISCOSE",
    "RY":  "VISCOSE",
}

# Checked in order — the first substring hit wins, so longer/more specific
# names must come first ("POLYESTER" before any bare "POLY" form).
CONTENT_FIBER_NAMES: list[tuple[str, str]] = [
    ("COTTON",    "COTTON"),
    ("POLYESTER", "POLYESTER"),
    ("POLYAMIDE", "NYLON"),
    ("NYLON",     "NYLON"),
    ("ELASTANE",  "SPANDEX"),
    ("SPANDEX",   "SPANDEX"),
    ("LYCRA",     "SPANDEX"),
    ("VISCOSE",   "VISCOSE"),
    ("RAYON",     "VISCOSE"),
]

# Two-fibre compositions where the minor fibre is spandex/elastane get
# their own Content LOV entries; the families without a dedicated
# "<fibre> Spandex" code fall back to "<fibre> Blend".
CONTENT_SPANDEX_PAIRS: dict[str, str] = {
    "COTTON":    "COTTONSPANDEX",
    "POLYESTER": "POLYSPANDEX",
    "NYLON":     "NYLONBLEND",
    "VISCOSE":   "VISCOSEBLEND",
}

# ── AT_Fabric (Apparel) ───────────────────────────────────────────────
# "Mapping from Principal" ← "Material Composition" per the Reebok(Inline)
# sheet. The MDD "Fabric LOV" lists knit/weave constructions rather than
# fibres, so the composition text is scanned for a construction keyword.
# Ordered longest-first so "Heavy Jersey"/"French Terry" win over the bare
# "Jersey"/"Terry" entries. No keyword → attribute skipped (a composition
# that names no construction has no defensible Fabric value).
FABRIC_KEYWORDS: list[tuple[str, str]] = [
    ("HEAVY JERSEY", "HJ"),
    ("FRENCH TERRY", "TE"),
    ("CORDUROY",     "CO"),
    ("INTERLOCK",    "IL"),
    ("MICROFIBER",   "MF"),
    ("MICROFIBRE",   "MF"),
    ("SUSTAINABLE",  "SU"),
    ("NEOPRENE",     "NP"),
    ("RIPSTOP",      "RS"),
    ("RIBSTOP",      "RS"),
    ("JERSEY",       "JE"),
    ("FLEECE",       "FL"),
    ("DENIM",        "DE"),
    ("SCUBA",        "SC"),
    ("TERRY",        "TE"),
    ("TWILL",        "TW"),
    ("PIQUE",        "PI"),
    ("PIQUÉ",        "PI"),
]


def _s(v) -> str:
    """Clean string value — returns empty string for None/nan/empty."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT", "#VALUE!") else s


def _clean_name(name: str) -> str:
    """Strip standalone 'Reebok' / 'REE' tokens from a product name (per mapping notes)."""
    if not name:
        return name
    cleaned = re.sub(r"\b(REEBOK|REE)\b", "", name, flags=re.IGNORECASE)
    return re.sub(r"\s{2,}", " ", cleaned).strip(" -")


def _fmt_date(v) -> str:
    """Normalise a date cell (datetime, or 'MM/DD/YYYY' string) to
    DD-Mon-YYYY (e.g. "01-Aug-2027"), for AT_LaunchingDate."""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%d-%b-%Y")
        except Exception:
            return ""
    s = _s(v)
    if not s:
        return ""
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%b-%Y")
        except ValueError:
            continue
    return s


def _fmt_date_ddmmyyyy(v) -> str:
    """Normalise a date cell (datetime, or 'MM/DD/YYYY' string) to
    DD-MM-YYYY (e.g. "01-08-2027"), for AT_IncomingMonth."""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%d-%m-%Y")
        except Exception:
            return ""
    s = _s(v)
    if not s:
        return ""
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return s


def _resolve_currency_from_country(country_code: str) -> str:
    """Resolve Retail Price Currency LOV ID from destination country code."""
    country_to_currency = {
        "ID": "IDR", "PH": "PHP", "TH": "THB", "SG": "SGD",
        "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


_PCT_MATERIAL_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*([^%,/]+)")


def _top_material(material_composition: str) -> str:
    """Extract the highest-percentage material name from a raw "Material
    Composition" cell, per the brand mapping template's AT_Material rule
    ("Material Composition (Highest Percentage)").

    Apparel cells commonly stack multiple labeled sections separated by "/"
    (e.g. "Shell: 87% PA, 13% EL / Panel: 92% PES, 8% EL") — real Reebok
    data uses "/" as often as "," to separate entries, sometimes with no
    delimiter at all ("60% Cotton 40% Polyester"). Footwear cells carry an
    "Upper Composition: ... Bottom Unit Composition: ..." split — only the
    "Upper Composition" segment is considered there. The name for each
    entry is captured up to the next "%", "," or "/" so entries don't
    swallow their neighbors; a stray trailing number (from a delimiter-less
    "40% Polyester" run) is trimmed off afterwards.
    """
    s = _s(material_composition)
    if not s:
        return ""
    if "Upper Composition" in s:
        s = s.split("Upper Composition:", 1)[-1]
        s = s.split("Bottom Unit Composition")[0]
    matches = [(pct, name) for pct, name in _PCT_MATERIAL_RE.findall(s) if name.strip()]
    if not matches:
        return ""
    _, name = max(matches, key=lambda m: float(m[0]))
    name = name.strip().rstrip(",/.:").strip()
    return re.sub(r"\s+\d+\s*$", "", name).strip()


def _lov_codes(mdd, lov_name: str) -> set[str]:
    """Set of valid LOV IDs for an MDD sheet-backed LOV (values of the
    display→ID map loaded by MDDLoader)."""
    lov = (mdd.lovs.get(lov_name, {}) if mdd else {})
    return {str(v).strip() for v in lov.values() if str(v).strip()}


def _sports_category_en_lov_id(sports_category: str, mdd=None) -> str:
    """Reebok "Sports Category" → AT_SportsCategoryEN LOV ID.

    Two hops, both taken from the shared mapping workbook:
      1. raw value → Stibo display value   (SPORTS_CATEGORY_EN_MAP,
         "Reebok MD Mappings sheet 2")
      2. display value → numeric LOV ID    (MDD "Sports Category LOV")
    Single-digit IDs are zero-padded to two chars, matching the LOV IDs
    Stibo accepts (same rule as asics/footwear_main.py). Returns "" when
    either hop misses, so the caller skips the attribute.
    """
    raw = re.sub(r"\s+", " ", _s(sports_category)).strip().upper()
    if not raw:
        return ""
    display = SPORTS_CATEGORY_EN_MAP.get(raw)
    if not display:
        return ""
    lov = (mdd.lovs.get("SportsCategoryLOV", {}) if mdd else {})
    lov_id = _s(lov.get(display.upper(), ""))
    if len(lov_id) == 1 and lov_id.isdigit():
        lov_id = f"0{lov_id}"
    return lov_id


def _country_size_lov_id(mdd=None) -> str:
    """COUNTRY_SIZE validated against the MDD "Country Size LOV". Returns
    "" when the module has no default (Apparel) or the ID isn't in the LOV."""
    if not COUNTRY_SIZE:
        return ""
    return COUNTRY_SIZE if COUNTRY_SIZE in _lov_codes(mdd, "CountrySizeLOV") else ""


# Fibre entries look like "60% Cotton 40% Polyester" or "87% PA, 13% EL".
# The name is captured as letters-only so a delimiter-less run doesn't let
# one entry swallow the next entry's percentage digits.
_FIBER_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*([A-Za-z][A-Za-z .\-]*)")


def _fiber_family(name: str) -> str:
    """Normalise one fibre name/abbreviation to a Content LOV family."""
    n = re.sub(r"[^A-Z]", "", _s(name).upper())
    if not n:
        return ""
    if n in CONTENT_FIBER_CODES:
        return CONTENT_FIBER_CODES[n]
    for token, family in CONTENT_FIBER_NAMES:
        if token in n:
            return family
    return "OTHER"


def _fiber_totals(material_composition: str) -> dict[str, float]:
    """Sum the declared percentages per fibre family. Multi-section cells
    ("Shell: ... / Panel: ...") are summed across sections — the input
    carries no section weighting, so the overall dominant fibre is used."""
    s = _s(material_composition)
    if not s:
        return {}
    if "Upper Composition" in s:
        s = s.split("Upper Composition:", 1)[-1].split("Bottom Unit Composition")[0]
    totals: dict[str, float] = {}
    for pct, name in _FIBER_PCT_RE.findall(s):
        family = _fiber_family(name)
        if not family:
            continue
        totals[family] = totals.get(family, 0.0) + float(pct)
    return totals


def _content_lov_code(material_composition: str, mdd=None) -> str:
    """"Material Composition" → AT_Content LOV ID, validated against the
    MDD "Content LOV". Returns "" when the dominant fibre has no Content
    family (e.g. wool, acrylic) or the derived ID isn't in the LOV."""
    totals = _fiber_totals(material_composition)
    if not totals:
        return ""
    ranked   = sorted(totals.items(), key=lambda kv: -kv[1])
    dominant = ranked[0][0]
    if dominant == "OTHER":
        return ""
    families = [f for f, _ in ranked]
    if len(families) == 1:
        code = dominant
    elif len(families) == 2 and families[1] == "SPANDEX":
        code = CONTENT_SPANDEX_PAIRS.get(dominant, f"{dominant}BLEND")
    elif dominant == "SPANDEX":
        code = "SPANDEX"
    else:
        code = f"{dominant}BLEND"
    return code if code in _lov_codes(mdd, "Content") else ""


def _fabric_lov_code(material_composition: str, mdd=None) -> str:
    """"Material Composition" → AT_Fabric LOV ID, validated against the MDD
    "Fabric LOV". Returns "" when the text names no known construction."""
    s = _s(material_composition).upper()
    if not s:
        return ""
    valid = _lov_codes(mdd, "Fabric")
    for keyword, code in FABRIC_KEYWORDS:
        if keyword in s and code in valid:
            return code
    return ""


# ======================================================================
# MDD LOADER
# ======================================================================

class MDDLoader:
    """Loads LOV tables from the shared Master Data Dictionary Excel."""

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
        self._load_country_lov(wb)
        self._load_gender_lov(wb)
        self._load_age_lov(wb)
        self._load_brand_lov(wb)
        self._load_brand_type_lov(wb)
        self._load_brand_status_lov(wb)
        self._load_brand_category_lov(wb)
        self._load_brand_group_lov(wb)
        self._load_sports_category_lov(wb)
        self._load_country_size_lov(wb)

        wb.close()
        log.info("[MDD] %d LOV types loaded", len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip().upper()
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
                    self.lovs.setdefault(display_name, {})[str(name).strip().upper()] = (
                        str(code).strip() if code else str(name).strip()
                    )

    def _load_color_lov(self, wb):
        """Color Code LOV: display name → LOV ID (best-effort match on Color Name text)."""
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
        log.info("[MDD] Color Code LOV: %d entries", len(lov))

    def _load_country_origin_lov(self, wb):
        """Country Origin LOV: full country name → ISO code."""
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
        self.lovs["CountryOriginLOV"] = lov
        log.info("[MDD] Country Origin LOV: %d entries", len(lov))

    def _load_country_lov(self, wb):
        """Country LOV: Col A = ISO code, Col B = display name."""
        sheet_name = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY LOV"),
            next((s for s in wb.sheetnames if "COUNTRY" in s.upper() and "LOV" in s.upper()
                  and "ORIGIN" not in s.upper()), None),
        )
        if not sheet_name:
            log.warning("[MDD] Country LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        id_to_name: dict[str, str] = {}
        for row in rows[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                id_to_name[lov_id.upper()] = display
        self.lovs["CountryLOV"] = id_to_name
        log.info("[MDD] Country LOV loaded — %d entries", len(id_to_name))

    def country_id_to_name(self, country_id: str) -> str:
        return self.lovs.get("CountryLOV", {}).get((country_id or "").strip().upper(), "")

    def _load_gender_lov(self, wb):
        """Gender LOV: col0=display name, col2=SAP Gender Code."""
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
        log.info("[MDD] Gender LOV: %d entries", len(lov))

    def _load_age_lov(self, wb):
        """Age LOV: col0=Age name, col1=Article Description (BY code), col2=SAP Age Code."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "AGE LOV"),
            next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()
                  and "GENDER" not in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Age LOV sheet not found")
            return
        sap_age: dict[str, str] = {}
        by_age: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 1 or not row[0]:
                continue
            display = str(row[0]).strip()
            by_code  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            sap_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if by_code:
                by_age[display.upper()] = by_code
            if sap_code:
                sap_age[display.upper()] = sap_code
        self.lovs["AgeLOV_BY"]  = by_age
        self.lovs["AgeLOV_SAP"] = sap_age
        log.info("[MDD] Age LOV: %d entries", len(sap_age))

    def _load_brand_lov(self, wb):
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()
             and "TYPE" not in s.upper() and "GROUP" not in s.upper()
             and "STATUS" not in s.upper() and "CATEGORY" not in s.upper()),
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
        self.lovs["BrandLOV"] = lov
        log.info("[MDD] Brand LOV: %d entries loaded", len(lov))

    def _load_brand_type_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "TYPE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Brand Type LOV sheet not found")
            return
        lov_map: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["AT_BrandType"] = lov_map
        log.info("[MDD] Brand Type LOV — %d entries", len(lov_map))

    def _load_brand_status_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "STATUS" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Brand Status LOV sheet not found")
            return
        lov_map: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["AT_BrandStatus"] = lov_map
        log.info("[MDD] Brand Status LOV — %d entries", len(lov_map))

    def _load_brand_category_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Brand Category LOV sheet not found")
            return
        lov_map: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["AT_BrandCategory"] = lov_map
        log.info("[MDD] Brand Category LOV — %d entries", len(lov_map))

    def _load_brand_group_lov(self, wb):
        sheet_name = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "GROUP" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Brand Group LOV sheet not found")
            return
        lov_map: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""
            display = str(row[1]).strip() if row[1] else ""
            if lov_id and display:
                lov_map[display.upper()] = lov_id
        self.lovs["AT_BrandGroup"] = lov_map
        log.info("[MDD] Brand Group LOV — %d entries", len(lov_map))

    def _load_sports_category_lov(self, wb):
        """Sports Category LOV: Col A = display value, Col B = LOV ID.

        Loaded explicitly because the sheet's column order is the reverse
        of the ID-then-name layout _load_named_lov_sheets() assumes.
        """
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "SPORTS CATEGORY LOV"),
            next((s for s in wb.sheetnames
                  if "SPORTS" in s.upper() and "CATEGORY" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Sports Category LOV sheet not found — AT_SportsCategoryEN will be skipped")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["SportsCategoryLOV"] = lov
        log.info("[MDD] Sports Category LOV: %d entries", len(lov))

    def _load_country_size_lov(self, wb):
        """Country Size LOV: Col A = display value, Col B = LOV ID (same
        reversed layout as Sports Category LOV)."""
        sheet = next(
            (s for s in wb.sheetnames if s.strip().upper() == "COUNTRY SIZE LOV"),
            next((s for s in wb.sheetnames
                  if "COUNTRY" in s.upper() and "SIZE" in s.upper() and "LOV" in s.upper()), None),
        )
        if not sheet:
            log.warning("[MDD] Country Size LOV sheet not found — AT_CountrySize will be skipped")
            return
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            display = str(row[0]).strip() if row[0] else ""
            lov_id  = str(row[1]).strip() if row[1] else ""
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["CountrySizeLOV"] = lov
        log.info("[MDD] Country Size LOV: %d entries", len(lov))

    def lookup(self, lov_key: str, display_value: str) -> str:
        """Look up a LOV ID. Returns display_value unchanged if not found."""
        lov = self.lovs.get(lov_key, {})
        if not lov or not display_value:
            return display_value
        return lov.get(display_value.strip().upper(), display_value)


# ======================================================================
# RNA LOADER — Brand Type / Category / Group / Status
# (identical technique to the other brand ETLs — reads the shared
#  "NEW - Brand mapping files Template.xlsx" attributes file)
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

    def get_by_brand_only(self, brand_code: str) -> dict:
        b = (brand_code or "").strip().upper()
        for (_, _, _, kb), val in self.lookup.items():
            if kb == b:
                return val
        return {"brand_type": "", "brand_category": "", "brand_group": "", "brand_status": ""}


# ======================================================================
# INPUT FILE LOADER
# ======================================================================

class ArticleMasterLoader:
    """
    Loads the Reebok Apparel Article Master (NuORDER export) from Excel.
    """

    PREFERRED_SHEETS = ["NuORDER Order Data"]
    HEADER_SIGNAL = "Article Number"

    def __init__(self, path: Path):
        self.path    = path
        self.df      = pd.DataFrame()
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[ArticleMaster] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[ArticleMaster] Available sheets: %s", wb.sheetnames)

        target = next(
            (s for s in wb.sheetnames if s.strip() in self.PREFERRED_SHEETS),
            wb.sheetnames[0],
        )
        log.info("[ArticleMaster] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.warning("[ArticleMaster] Sheet '%s' is empty", target)
            return

        hdr_idx = 0
        for i in range(min(20, len(rows))):
            if rows[i] and any(
                str(v).strip() == self.HEADER_SIGNAL for v in rows[i] if v is not None
            ):
                hdr_idx = i
                log.info("[ArticleMaster] Header detected at row %d (found '%s')", hdr_idx, self.HEADER_SIGNAL)
                break
        else:
            log.warning(
                "[ArticleMaster] '%s' not found in first 20 rows — defaulting to row 0",
                self.HEADER_SIGNAL,
            )

        raw_header = rows[hdr_idx]
        header = [
            str(h).strip() if h is not None else f"col_{i}"
            for i, h in enumerate(raw_header)
        ]

        seen: dict[str, int] = {}
        for ci, col in enumerate(header):
            if col in seen:
                seen[col] += 1
                header[ci] = f"{col}_{seen[col]}"
            else:
                seen[col] = 0

        self.headers = header
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
        df = df.dropna(how="all")

        self.df = df.reset_index(drop=True)
        log.info(
            "[ArticleMaster] Loaded %d data rows | %d columns",
            len(self.df), len(header),
        )

    @staticmethod
    def _find_col(headers: list[str], candidates: list[str]) -> str | None:
        for c in candidates:
            if c in headers:
                return c
            for h in headers:
                if h.lower() == c.lower():
                    return h
        return None

    def find_col(self, *candidates: str) -> str | None:
        return self._find_col(self.headers, list(candidates))


# ======================================================================
# GROUPER — flat SKU rows → Generic articles + variant counts
# ======================================================================

def group_rows(raw_rows: list[dict], loader: ArticleMasterLoader, brand_code: str) -> dict[str, dict]:
    """Group flat per-SKU rows into Generic articles keyed by brand_code + Article Number."""
    col_article  = loader.find_col("Article Number")
    col_style_no = loader.find_col("Style Number")
    col_name     = loader.find_col("Name")
    col_color    = loader.find_col("Color Name")
    col_gender   = loader.find_col("Gender")
    col_age      = loader.find_col("Age Group")
    col_size     = loader.find_col("Size")
    col_size_rng = loader.find_col("Size Range")
    col_upc      = loader.find_col("UPC")
    col_sku      = loader.find_col("SKU")
    col_coo      = loader.find_col("T1 Supplier Country Description")
    col_dept     = loader.find_col("Department")
    col_sports   = loader.find_col("Sports Category")
    col_segment  = loader.find_col("Article Business Segment")
    col_mkt_line = loader.find_col("Category Marketing Line")
    col_material = loader.find_col("Material Composition")
    col_status   = loader.find_col("Article Status")
    col_fob      = loader.find_col("Customer Price (USD) (USD)", "Customer Price (USD)")
    col_msrp     = loader.find_col("M.S.R.P (USD)", "MSRP (USD)")
    col_launch   = loader.find_col("Retail Intro Date")
    col_uom      = loader.find_col("Sales Unit")

    log.info(
        "[Grouper] Key columns → article=%s name=%s color=%s gender=%s age=%s size=%s "
        "upc=%s coo=%s dept=%s sports=%s segment=%s mkt_line=%s material=%s status=%s "
        "fob=%s msrp=%s launch=%s uom=%s",
        col_article, col_name, col_color, col_gender, col_age, col_size,
        col_upc, col_coo, col_dept, col_sports, col_segment, col_mkt_line, col_material,
        col_status, col_fob, col_msrp, col_launch, col_uom,
    )

    if col_article is None:
        log.warning(
            "[Grouper] Could not find 'Article Number' column. Available columns: %s",
            loader.headers,
        )
        return {}

    result: dict[str, dict] = {}

    for row in raw_rows:
        article_number = _s(row.get(col_article, "")) if col_article else ""
        if not article_number:
            continue

        generic_key = f"{brand_code}{article_number}"

        if generic_key not in result:
            result[generic_key] = {
                "generic_key":      generic_key,
                "article_number":   article_number,
                "style_number":     _s(row.get(col_style_no, "")) if col_style_no else "",
                "name":             _clean_name(_s(row.get(col_name, ""))) if col_name else "",
                "color_name":       _s(row.get(col_color, "")) if col_color else "",
                "gender":           _s(row.get(col_gender, "")) if col_gender else "",
                "age_group":        _s(row.get(col_age, "")) if col_age else "",
                # AT_PrincipalSize ← "Size Range" per Reebok(Inline) mapping
                # sheet (NOT the "Size" column, which feeds the variant/
                # Principal Barcode size code below).
                "principal_size":   _s(row.get(col_size_rng, "")) if col_size_rng else "",
                "country_origin":   _s(row.get(col_coo, "")) if col_coo else "",
                "department":       _s(row.get(col_dept, "")) if col_dept else "",
                "sports_category":  _s(row.get(col_sports, "")) if col_sports else "",
                "segment":          _s(row.get(col_segment, "")) if col_segment else "",
                "mkt_line":         _s(row.get(col_mkt_line, "")) if col_mkt_line else "",
                "material":         _s(row.get(col_material, "")) if col_material else "",
                "status_raw":       _s(row.get(col_status, "")) if col_status else "",
                "fob":              _s(row.get(col_fob, "")) if col_fob else "",
                "msrp":             _s(row.get(col_msrp, "")) if col_msrp else "",
                "launch_date":      _fmt_date(row.get(col_launch)) if col_launch else "",
                "uom":              _s(row.get(col_uom, "")) if col_uom else "",
                "brand_code":       brand_code,
                "variants":         {},
            }

        color_code = _s(row.get(col_color, "")) if col_color else ""
        size_code  = _s(row.get(col_size, ""))  if col_size  else ""
        upc        = _s(row.get(col_upc, ""))   if col_upc   else ""
        sku        = _s(row.get(col_sku, ""))   if col_sku   else ""
        variant_key = f"{generic_key}|{color_code}|{size_code}"
        result[generic_key]["variants"][variant_key] = {
            "color_name": color_code,
            "size":       size_code,
            "upc":        upc,
            "sku":        sku,
        }

    log.info(
        "[Grouper] %d rows → %d generics, %d total variants",
        len(raw_rows), len(result), sum(len(g["variants"]) for g in result.values()),
    )
    return result


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


# ======================================================================
# CLASSIFICATIONS BUILDER
# ======================================================================

def build_classifications(brand: str, brand_code: str, season_code: str) -> ET.Element:
    """Build the <Classifications> block (Season / Confirmed / Unconfirmed)."""
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
    brand_category: str = "",
    brand_group: str = "",
    mdd: "MDDLoader | None" = None,
) -> str:
    """Build XML for one Generic article."""
    if not generic.get("article_number"):
        return ""

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_A-TempSubCat")  # Apparel

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_key"]

    name = generic.get("name") or generic["generic_key"]
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = name

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    gv = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # ── Principal fields (Direct from Principal) ──────────────────
    _val_text(gv, "AT_PrincipalStyleCode",        generic["article_number"])
    _val_text(gv, "AT_InboundGenericCode",        generic["generic_key"])
    if generic.get("name"):
        _val_text(gv, "AT_PrincipalStyleDescription", generic["name"])
    if generic.get("color_name"):
        # Canonical Stibo attribute ID for "Principal Color Description" is
        # AT_PrincipalColorName (per Template sheet + K-Swiss's real output) —
        # NOT AT_PrincipalColorDescription, which doesn't exist in the schema.
        _val_text(gv, "AT_PrincipalColorName", generic["color_name"])
    if generic.get("gender"):
        _val_text(gv, "AT_PrincipalGenderDescription", generic["gender"])
    if generic.get("age_group"):
        _val_text(gv, "AT_PrincipalAgeDescription", generic["age_group"])
    if generic.get("principal_size"):
        _val_text(gv, "AT_PrincipalSize", generic["principal_size"])

    # ── Merchandise hierarchy (Direct from Principal) ──────────────
    if generic.get("department"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL1", generic["department"])
    if generic.get("sports_category"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL2", generic["sports_category"])
        # AT_SportsCategoryEN is an LOV attribute in Stibo, not free text —
        # writing the raw "Sports Category" string as-is threw "Illegal LOV"
        # on import. It now goes through SPORTS_CATEGORY_EN_MAP ("Reebok MD
        # Mappings sheet 2") and the MDD "Sports Category LOV".
        sports_cat_lov_id = _sports_category_en_lov_id(generic["sports_category"], mdd)
        if sports_cat_lov_id:
            _val_lov(gv, "AT_SportsCategoryEN", sports_cat_lov_id)
        else:
            log.warning(
                "[XML] Sports Category '%s' has no Sports Category EN LOV ID — "
                "AT_SportsCategoryEN skipped for %s",
                generic["sports_category"], generic["generic_key"],
            )
    if generic.get("segment"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL3", generic["segment"])
    if generic.get("mkt_line"):
        _val_text(gv, "AT_PrincipalMerchandiseHierarchyL4", generic["mkt_line"])

    # ── Content / Fabric (Direct & Mapping from Principal — both read the
    # single "Material Composition" col per the Reebok(Inline) sheet) ──
    # AT_Content and AT_Fabric are LOV attributes in Stibo, not free text —
    # writing the raw composition string is what threw "Illegal LOV" on
    # import. Both now resolve to real LOV IDs (MDD "Content LOV" /
    # "Fabric LOV") and are skipped when the composition supports neither.
    if generic.get("material"):
        content_code = _content_lov_code(generic["material"], mdd)
        if content_code:
            _val_lov(gv, "AT_Content", content_code)
        else:
            log.warning(
                "[XML] Material Composition %r → no Content LOV match — "
                "AT_Content skipped for %s",
                generic["material"][:80], generic["generic_key"],
            )

        fabric_code = _fabric_lov_code(generic["material"], mdd)
        if fabric_code:
            _val_lov(gv, "AT_Fabric", fabric_code)
        else:
            log.info(
                "[XML] Material Composition %r names no Fabric construction — "
                "AT_Fabric skipped for %s",
                generic["material"][:80], generic["generic_key"],
            )

        # AT_Material stays unmapped: the MDD "Material LOV" is a separate
        # table the mapping workbook doesn't tie to Material Composition,
        # so the extracted top fibre would still be an illegal LOV ID.
        # top_material = _top_material(generic["material"])
        # if top_material:
        #     _val_text(gv, "AT_Material", top_material)

    # ── Pricing (Direct from Principal) ────────────────────────────
    if generic.get("fob"):
        _val_text(gv, "AT_FOB", generic["fob"])
    if generic.get("msrp"):
        _val_text(gv, "AT_OriginalPrice", generic["msrp"])
        _val_text(gv, "AT_CurrentPrice",  generic["msrp"])

    # ── Launching Date (Direct from Principal ← Retail Intro Date) ──
    if generic.get("launch_date"):
        _val_text(gv, "AT_LaunchingDate", generic["launch_date"])

    # ── UOM: "Formula in System / Default: EA" per mapping sheet — a flat
    # system default, NOT sourced from the brand file's "Sales Unit" column.
    # (Every row in the real Reebok file happens to say "EA" there anyway,
    # but reading it was a deviation from the documented mapping logic.)
    _val_lov(gv, "AT_UOM", "EA")

    # ── Country Size ("Formula in System": FW → US, APP/ACC → blank) ──
    country_size_id = _country_size_lov_id(mdd)
    if country_size_id:
        _val_lov(gv, "AT_CountrySize", country_size_id)
    elif COUNTRY_SIZE:
        log.warning(
            "[XML] Country Size '%s' not found in MDD Country Size LOV — "
            "AT_CountrySize skipped for %s",
            COUNTRY_SIZE, generic["generic_key"],
        )

    # ── Country of Origin: T1 Supplier Country Description → ISO code ──
    if generic.get("country_origin"):
        coo_lov = mdd.lovs.get("CountryOriginLOV", {}) if mdd else {}
        coo_code = coo_lov.get(generic["country_origin"].upper(), "")
        if coo_code:
            _val_lov(gv, "AT_CountryOrigin", coo_code)
        else:
            log.warning(
                "[XML] Country Origin '%s' not found in MDD — attribute skipped for %s",
                generic["country_origin"], generic["generic_key"],
            )

    # ── Article Status: "Formula in System / Default when created: Active.
    # Will be updated by SAP (deactivation flag)" per mapping sheet — always
    # "A" at creation time; the brand file's own "Article Status" column
    # (e.g. "Buy Ready") is not the source per spec, SAP's own deactivation
    # flag is what changes this later, not this ETL reading brand text.
    _val_lov(gv, "AT_ArticleStatus", "A")

    # ── SAP Age / BY Age (static rule table — see AGE_GROUP_RULES) ──
    if generic.get("age_group"):
        rule = AGE_GROUP_RULES.get(generic["age_group"].upper())
        if rule:
            sap_age, by_age = rule
            _val_lov(gv, "AT_SAPAge", sap_age)
            _val_lov(gv, "AT_BYAge", by_age)
        else:
            log.warning(
                "[XML] Age Group '%s' not found in AGE_GROUP_RULES — SAP/BY Age skipped for %s",
                generic["age_group"], generic["generic_key"],
            )
        # AT_EComAgesCategory is disabled — the Reebok(Inline) spec sheet
        # documents its real source as the "Gender" column (via the missing
        # "Mapping Attributes Reebok.xlsx"), not Age Group. EC_AGES_CATEGORY_MAP
        # below is an Age-Group-based mapping that produces valid LOV codes but
        # reads the wrong source column, so it's left unused until the real
        # Gender→LOV mapping is provided.
        # ec_age_code = EC_AGES_CATEGORY_MAP.get(generic["age_group"].upper())
        # if ec_age_code:
        #     _val_lov(gv, "AT_EComAgesCategory", ec_age_code)
        # else:
        #     log.warning(
        #         "[XML] Age Group '%s' not found in EC_AGES_CATEGORY_MAP — "
        #         "AT_EComAgesCategory skipped for %s",
        #         generic["age_group"], generic["generic_key"],
        #     )

    # ── SAP Gender (AT_Gender) / BY Gender (via MDD "Gender LOV") ──
    # NOTE: "SAP Gender" maps to AT_Gender per the mapping template (Template
    # sheet row: "SAP Gender" -> AT_Gender) -- there is no separate
    # "AT_SAPGender" attribute in the schema; a prior version of this code
    # wrote a bogus AT_SAPGender value alongside AT_Gender, confirmed absent
    # from the canonical attribute list and from K-Swiss's real production XML.
    if generic.get("gender"):
        gender_lov = mdd.lovs.get("GenderLOV", {}) if mdd else {}
        gender_id  = gender_lov.get(generic["gender"].upper(), "")
        if gender_id:
            _val_lov(gv, "AT_Gender",    gender_id)
            _val_lov(gv, "AT_BYGender",  gender_id)
        else:
            log.warning(
                "[XML] Gender '%s' not found in Gender LOV — Gender attributes skipped for %s",
                generic["gender"], generic["generic_key"],
            )

    # ── AT_Brand ─────────────────────────────────────────────────
    brand_lov    = mdd.lovs.get("BrandLOV", {}) if mdd else {}
    brand_lov_id = brand_lov.get(brand.upper(), brand_code)
    _val_lov(gv, "AT_Brand", brand_lov_id)

    # ── Brand Type / Status / Category / Group (via RNA) ───────────
    if brand_type:
        bt = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type.upper(), "")
        _val_lov(gv, "AT_BrandType", bt) if bt else _val_text(gv, "AT_BrandType", brand_type)
    if brand_status:
        bs = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status.upper(), "")
        _val_lov(gv, "AT_BrandStatus", bs) if bs else _val_text(gv, "AT_BrandStatus", brand_status)
    if brand_category:
        bc = (mdd.lovs.get("AT_BrandCategory", {}) if mdd else {}).get(brand_category.upper(), "")
        _val_lov(gv, "AT_BrandCategory", bc) if bc else _val_text(gv, "AT_BrandCategory", brand_category)
    if brand_group:
        bg = (mdd.lovs.get("AT_BrandGroup", {}) if mdd else {}).get(brand_group.upper(), "")
        _val_lov(gv, "AT_BrandGroup", bg) if bg else _val_text(gv, "AT_BrandGroup", brand_group)

    # ── Company + SBU (MultiValue) ──────────────────────────────────
    mv_comp = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_comp.set("AttributeID", "AT_CompanyCode")
    ET.SubElement(mv_comp, f"{{{STIBO_NS}}}Value").set("ID", comp_code)

    mv_sbu = ET.SubElement(gv, f"{{{STIBO_NS}}}MultiValue")
    mv_sbu.set("AttributeID", "AT_SBU")
    ET.SubElement(mv_sbu, f"{{{STIBO_NS}}}Value").set("ID", sbu)

    # ── Season + Season Year ────────────────────────────────────────
    sea_code = season[:2].upper() if len(season) >= 2 else season
    _val_lov(gv, "AT_Season", sea_code)

    year_match = re.search(r"20\d{2}", season)
    if not year_match:
        short_match = re.search(r"\d{2}$", season)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    if season_year:
        _val_text(gv, "AT_SeasonYear", season_year)

    # ── Country (destination) + Currency ────────────────────────────
    if country_code:
        _val_lov(gv, "AT_Country", country_code.upper())
    currency_id = _resolve_currency_from_country(country_code)
    if currency_id:
        _val_lov(gv, "AT_RetailPriceCurrency", currency_id)

    # ── System defaults (all "Formula in System" per mapping sheet) ───
    _val_lov(gv, "AT_BYArticleType", "Inline")
    _val_lov(gv, "AT_SAPProductFlag", "A")
    _val_lov(gv, "AT_PricingDistributionChannel", "01")
    _val_lov(gv, "AT_FOBCurrency", "USD")            # Default: USD
    _val_lov(gv, "AT_MaterialType", "ZINA")           # Default: Intercompany
    _val_lov(gv, "AT_SAPArticleCategory", "1")        # Default: Generic

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# FILENAME METADATA PARSER
# ======================================================================

def _parse_metadata_from_filename(dirs: dict, triggered_file_type: str = None) -> dict:
    """
    Parse comp_code, sbu, brand, brand_code, season, seq from the
    triggered brand file's filename.

    Supports two formats:
      1. Strict dash-separated: {CompanyCode}-{SBU}-{Brand}-{FileType}-{Season}-{Seq}.xlsx
      2. Loose (as used for Reebok's NuORDER export):
         "Reebok article master data with image - APPAREL (SS27).xlsx"
         → season is pulled from anywhere in the filename (e.g. "SS27").
    """
    candidate_dirs = []
    if triggered_file_type and triggered_file_type in dirs:
        candidate_dirs.append(dirs[triggered_file_type])

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
            log.info("Attempting metadata parse from filename: '%s'", stem)

            parts = re.split(r"\s*-\s*", stem)
            if len(parts) >= 6:
                comp_code = parts[0].strip()
                sbu       = parts[1].strip()
                brand     = parts[2].strip().title()

                season = None
                season_idx = None
                for i, p in enumerate(parts):
                    if re.match(r"^[A-Z]{2}\d{2,4}$", p.strip(), re.IGNORECASE):
                        season     = p.strip().upper()
                        season_idx = i
                        break

                if season and season_idx:
                    trailing     = parts[season_idx + 1:]
                    seq          = next((int(p) for p in reversed(trailing) if p.strip().isdigit()), 1)
                    country_code = next(
                        (p.strip().upper() for p in trailing
                         if re.match(r"^[A-Z]{2,3}$", p.strip(), re.IGNORECASE) and not p.strip().isdigit()),
                        "",
                    )
                    log.info(
                        "Parsed (strict): comp=%s sbu=%s brand=%s season=%s seq=%s country=%s",
                        comp_code, sbu, brand, season, seq, country_code,
                    )
                    return {
                        "comp_code": comp_code, "sbu": sbu, "brand": brand,
                        "brand_code": BRAND_CODE, "season": season, "seq": seq,
                        "country_code": country_code,
                    }

            # ── Fallback: extract season from anywhere in filename ────
            season_match = re.search(r"\b([A-Z]{2})(\d{2,4})\b", stem, re.IGNORECASE)
            if season_match:
                season_prefix = season_match.group(1).upper()
                season_year   = season_match.group(2)
                if len(season_year) == 2:
                    season_year = f"20{season_year}"
                season = f"{season_prefix}{season_year}"

                log.info(
                    "Parsed (fallback): brand=Reebok code=%s season=%s (defaults applied)",
                    BRAND_CODE, season,
                )
                return {
                    "comp_code": "0888", "sbu": "SP", "brand": "Reebok",
                    "brand_code": BRAND_CODE, "season": season, "seq": 1,
                    "country_code": "",
                }

    log.warning("Could not parse metadata from filename — using defaults")
    return {
        "comp_code": "0888", "sbu": "SP", "brand": "Reebok",
        "brand_code": BRAND_CODE, "season": "SS27", "seq": 1,
        "country_code": "",
    }


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """Main entry point for Reebok Apparel Article Master ETL."""
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SS27")
    country_code = getattr(args, "country_code", "")

    log.info(
        "[ReebokApparel] Starting: brand=%s code=%s season=%s",
        brand, brand_code, season,
    )

    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[ReebokApparel] No MDD file found in %s — LOV lookups disabled", MDD_DIR)

    rna = None
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None) or next(ATTR_DIR.glob("*.xlsm"), None)
    if mapping_file:
        rna = RNALoader(mapping_file)
        log.info("[ReebokApparel] RNA loaded: %d entries", len(rna.lookup))
    else:
        log.warning("[ReebokApparel] No attributes/brand-mapping file found in %s — RNA disabled", ATTR_DIR)

    input_file = next(INPUT_DIR.glob("*.xls*"), None)
    if not input_file:
        raise FileNotFoundError(f"No Reebok apparel article master file found in {INPUT_DIR}")

    # ── Resolve country name from country ID via MDD (for RNA lookup) ──
    rna_country = ""
    if country_code and mdd:
        rna_country = mdd.country_id_to_name(country_code)
        if not rna_country:
            log.warning("[ReebokApparel] Country ID '%s' not found in Country LOV", country_code)
    if not rna_country:
        rna_country = country_code

    # ── Look up Brand LOV ID from MDD ────────────────────────────
    brand_lov_id = brand_code
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV", {})
        looked_up = brand_lov.get(brand.upper())
        if looked_up:
            brand_lov_id = looked_up
        else:
            log.warning(
                "[ReebokApparel] '%s' not in Brand LOV — using fallback brand_code='%s'",
                brand.upper(), brand_code,
            )

    # ── RNA lookup: brand_type, brand_status, brand_category, brand_group ──
    brand_type = brand_status = brand_category = brand_group = ""
    if rna:
        rna_result = rna.get(rna_country, comp_code, sbu, brand_lov_id)
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_fuzzy(rna_country, sbu, brand_lov_id)
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_fuzzy(rna_country, sbu, brand_code)
        if not rna_result.get("brand_type") and not rna_result.get("brand_category"):
            rna_result = rna.get_by_brand_only(brand_lov_id) or rna.get_by_brand_only(brand_code)
        brand_type     = rna_result.get("brand_type", "")
        brand_status   = rna_result.get("brand_status", "")
        brand_category = rna_result.get("brand_category", "")
        brand_group    = rna_result.get("brand_group", "")
        log.info(
            "[RNA] Resolved — brand_type='%s' brand_status='%s' brand_category='%s' brand_group='%s'",
            brand_type, brand_status, brand_category, brand_group,
        )

    loader = ArticleMasterLoader(input_file)
    if loader.df.empty:
        log.warning("[ReebokApparel] No data rows found — nothing to process")
        return None

    raw_rows = loader.df.to_dict("records")
    generics = group_rows(raw_rows, loader, brand_lov_id)
    # generics = dict(list(generics.items())[:5])  # TEST LIMIT — uncomment to process only first 5 products

    if not generics:
        log.warning("[ReebokApparel] No valid generics produced")
        return None

    sp_code  = season[:2].upper() if len(season) >= 2 else season
    sys_part = season[2:] if len(season) > 2 else ""
    sy       = f"20{sys_part}" if len(sys_part) == 2 else sys_part
    season_id = f"CLH_{brand_lov_id}_{sp_code}{sy}"

    xml_filename = f"{input_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    cls_el  = build_classifications(brand, brand_lov_id, season)
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))

    log.info("[ReebokApparel] Writing XML → %s", xml_filename)
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

        for generic in generics.values():
            px = build_product_xml(
                generic, brand, brand_lov_id, comp_code, sbu,
                season, season_id, country_code,
                brand_type=brand_type, brand_status=brand_status,
                brand_category=brand_category, brand_group=brand_group,
                mdd=mdd,
            )
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[ReebokApparel] XML written: %s (%d KB)", xml_path.name, file_kb)
    print(
        f"=== REEBOK APPAREL SUMMARY ===\n"
        f"  Input rows : {len(raw_rows)}\n"
        f"  Generics   : {generic_count}\n"
        f"  Variants   : {variant_count} (tracked only — not emitted as XML)\n"
        f"  XML size   : {file_kb} KB",
        flush=True,
    )

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return xml_path


# ======================================================================
# CLI (local test)
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Reebok Apparel Article Master → Stibo XML")
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
