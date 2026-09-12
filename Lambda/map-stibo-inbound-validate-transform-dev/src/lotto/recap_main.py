"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — LOTTO Licensed Recap v2.0     ║
║  Licensed Recap Sample Development → Stibo STEP XML              ║
╚══════════════════════════════════════════════════════════════════╝

Source file
  • Single input file  : "RECAP SAMPLE DEVELOPMENT-LOT-{SEASON}-FOOTWEAR.xlsx"
  • Sheet              : "LOT" / "LOTTO" (first sheet as a last resort)
  • Header row         : located by scoring rows against RECAP_COLUMNS
  • Filter             : rows with a blank "Supp Art #" are dropped

Two ingestions, one Lambda
──────────────────────────
The same article reaches the MAP Portal twice:

  1st ingestion  The principal's own recap.  Only the columns the principal
                 fills are populated; the Lambda completes the rest from
                 formulas and mapping tables.  The user then enriches the
                 record in Smartsheet (Stibo) and in MDTools.
  2nd ingestion  The completed recap exported back out of MDTools and
                 re-uploaded.  Same layout, plus the finalised columns
                 ("Updated Image", "Final FOB", "proposed_retail_price", …).

detect_ingestion_phase() tells the two apart by looking for those
2nd-ingestion-only columns; RECAP_COLUMNS then decides which column wins per
field, always falling back to the other when a cell is blank.

Column mapping (Lotto Mapping Issues and References.xlsx → "Mapping to STIBO")
    Supp Art #      → AT_PrincipalStyleCode (mirrored to the generic code)
    Age Group       → AT_PrincipalAgeDescription, AT_SAPAge, AT_BYAge
    Gender          → AT_Gender, AT_BYGender, AT_PrincipalGenderDescription
    Division        → AT_PrincipalMerchandiseHierarchyL1
    MD Category     → AT_PrincipalMerchandiseHierarchyL2, AT_SportsCategoryEN
    FOB Price       → AT_FOB          (2nd ingestion: Final FOB)
    FOB Currency    → AT_FOBCurrency
    ETA DATE        → AT_IncomingMonth
    Image           → AT_ThumbnailImage (2nd ingestion: Updated Image)
    BCI             → AT_BCI          (1st ingestion: Manual Input → not sent;
                                        2nd ingestion: recap "BCI" column)
    Color / Code    → AT_PrincipalColorName / AT_PrincipalColorCode
    Size Range      → AT_PrincipalSize (2nd ingestion: ProductSize)
    Season          → AT_Season + AT_SeasonYear

Age & Gender come from sheet "BY Age & Gender":
    Gender    Male/Men/Boys → Male · Female/Women/Girls → Female · Unisex
    Age Group Adult→(Adults, Adult)  Kids→(Children, Kids)
              All Ages→(All Ages, All Ages)  Infant→(Children, Infant)
              Preschool→(Children, Preschool)
              Grade School→(Children, Grade School)
  SAP Age and BY Age are *different* value sets — "Children" is a SAP Age and
  must never be sent as a BY Age.

Every LOV attribute is written with an ID resolved by _lov_id() (MDD →
documented code table → display value).  A <Value> carrying text but no ID is
silently dropped by STIBO; that was the cause of the "Not Populated" results
for SAP/BY Age, SAP/BY Gender and BCI in UAT.

Article structure:
  • Article Type    : "License" (1st ingestion) or the recap "Article Type"
  • Article Category: Generic (1) — has size variants
  • Each row = one Generic (Supp Art # + Color Code combination)

Generic code : brand(3) + article-type(1) + season-year digit(1)
               + code category(1) + last 4 of Supp Art # + gender(1) + colour(1)
"""
from __future__ import annotations

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

INPUT_DIR    = BASE_DIR / "input"
RECAP_DIR    = INPUT_DIR / "recap"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [RECAP_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# NORMALISATION HELPERS
# ══════════════════════════════════════════════════════════════════

def _norm_key(v) -> str:
    """Normalise any label / header / LOV display value to an A-Z0-9 key.

    Recap workbooks and MDD sheets both suffer from cosmetic drift - trailing
    spaces, embedded newlines ("FOB \nCurrency"), "#" vs "No.", "Grade School"
    vs "GRADE-SCHOOL".  Comparing on this key removes that whole class of
    "Not Populated" defects.
    """
    return re.sub(r"[^A-Z0-9]", "", str(v or "").upper())


def _nkeyed(d: dict) -> dict:
    """Re-key a lookup table by _norm_key so lookups ignore case/punctuation."""
    return {_norm_key(k): v for k, v in d.items()}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior", "K": "Kids",
}
# BY Age is its own value set - "Children" is a *SAP* Age value and must never
# be sent as a BY Age (sheet "BY Age & Gender", rows 20-25).
LOV_BY_AGE = _nkeyed({
    "ADULT": "Adult", "ADULTS": "Adult", "AD": "Adult",
    "KIDS": "Kids", "KID": "Kids", "K": "Kids",
    "CHILDREN": "Kids", "CHILD": "Kids", "CH": "Kids",
    "ALL AGES": "All Ages", "AA": "All Ages",
    "INFANT": "Infant", "IN": "Infant",
    "PRESCHOOL": "Preschool",
    "GRADE SCHOOL": "Grade School",
})
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter",
}
LOV_COUNTRY_ORIGIN = {
    "CN": "China", "VN": "Vietnam", "ID": "Indonesia",
    "KH": "Cambodia", "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia", "TH": "Thailand",
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

# ── Country code → MDD COUNTRY name ──────────────────────────────
_COUNTRY_MAP: dict[str, str] = {
    "ID": "INDONESIA",
    "MY": "MALAYSIA",
    "PH": "PHILIPPINES",
    "SG": "SINGAPORE",
    "TH": "THAILAND",
    "VN": "VIETNAM",
    "KH": "CAMBODIA",
}

_COUNTRY_NAME_TO_CODE: dict[str, str] = {
    "INDONESIA": "ID", "ID": "ID",
    "PHILIPPINES": "PH", "PHILIPPINE": "PH", "PH": "PH",
    "THAILAND": "TH", "TH": "TH",
    "SINGAPORE": "SG", "SG": "SG",
    "MALAYSIA": "MY", "MY": "MY",
    "VIETNAM": "VN", "VN": "VN",
    "CAMBODIA": "KH", "CAMBODIAN": "KH", "KH": "KH",
}

def _parse_country_code_from_file_or_args(filename: str, cli_country: str = "") -> str:
    # 1. First check if filename has an explicit country token (e.g. PH, KH, TH, MY, SG, VN, ID)
    stem_upper = Path(filename).stem.upper().replace("_", " ").replace("-", " ")
    tokens = stem_upper.split()
    for token in tokens:
        if token in _COUNTRY_NAME_TO_CODE and token != "SP":
            return _COUNTRY_NAME_TO_CODE[token]

    # 2. Fallback to CLI argument if provided
    cli_c = (cli_country or "").strip().upper()
    if cli_c:
        return cli_c
            
    return "ID"

# ── LOTTO-specific lookup tables ───────────────────────────────
#
# Source of truth: "Lotto Mapping Issues and References.xlsx"
#   • sheet "BY Age & Gender"  → recap Gender / Age Group → SAP + BY values
#   • sheet "Mapping to STIBO" → per-attribute source column and LOV examples
#   • sheet "UAT Result"       → the defects these tables exist to fix
#
# Each table maps a recap value to the pair of *display* values STIBO expects:
# (SAP display, BY display).  LOV ids are resolved separately by _lov_id() so
# the MDD stays authoritative and we never guess an id we could look up.

# recap "Gender" → (SAP Gender display, BY Gender display)
LOTTO_GENDER_MAP: dict[str, tuple[str, str]] = _nkeyed({
    "Male":    ("Male",   "Male"),
    "Men":     ("Male",   "Male"),
    "Mens":    ("Male",   "Male"),
    "Men's":   ("Male",   "Male"),
    "Boy":     ("Male",   "Male"),
    "Boys":    ("Male",   "Male"),
    "M":       ("Male",   "Male"),
    "Female":   ("Female", "Female"),
    "Women":    ("Female", "Female"),
    "Womens":   ("Female", "Female"),
    "Women's":  ("Female", "Female"),
    "Girl":     ("Female", "Female"),
    "Girls":    ("Female", "Female"),
    "F":        ("Female", "Female"),
    "W":        ("Female", "Female"),
    "Unisex":  ("Unisex", "Unisex"),
    "Uni":     ("Unisex", "Unisex"),
    "U":       ("Unisex", "Unisex"),
})

# recap "Age Group" → (SAP Age display, BY Age display)
LOTTO_AGE_GROUP_MAP: dict[str, tuple[str, str]] = _nkeyed({
    "Adult":        ("Adults",   "Adult"),
    "Adults":       ("Adults",   "Adult"),
    "AD":           ("Adults",   "Adult"),
    "Kids":         ("Children", "Kids"),
    "Kid":          ("Children", "Kids"),
    "Children":     ("Children", "Kids"),
    "Child":        ("Children", "Kids"),
    "CH":           ("Children", "Kids"),
    "All Ages":     ("All Ages", "All Ages"),
    "AA":           ("All Ages", "All Ages"),
    "Infant":       ("Children", "Infant"),
    "Preschool":    ("Children", "Preschool"),
    "Pre School":   ("Children", "Preschool"),
    "Grade School": ("Children", "Grade School"),
})

# Fallback only: recap Gender → Age Group, used when the recap carries no
# "Age Group" column at all.  Anything not listed here is an Adult.
LOTTO_GENDER_TO_AGE_GROUP: dict[str, str] = _nkeyed({
    "Boy": "Kids", "Boys": "Kids", "Girl": "Kids", "Girls": "Kids",
    "Kids": "Kids", "Kid": "Kids", "Children": "Kids",
})

# LOV ids documented in "Mapping to STIBO" rows 81-90.  BY Age / BY Gender / BCI
# have no documented id list; for those _lov_id() falls back to the display
# value itself - the convention the UAT "Result" column confirms for BCI
# ("Commercial", *not* "COMMERCIAL").
SAP_AGE_LOV_ID: dict[str, str] = {
    "Adults": "AD", "Children": "CH", "All Ages": "AA",
    "Infant": "IN", "Junior": "JR",
}
SAP_GENDER_LOV_ID: dict[str, str] = {"Male": "M", "Female": "F", "Unisex": "U"}

# BY Age LOV ids are the upper-cased display values — Stibo LOV_BYAge accepts
# exactly ADULT / ALL AGES / GRADE SCHOOL / INFANT / KIDS / PRESCHOOL (see the
# Nike and Reebok modules).  The MDD "Age LOV" sheet lists the BY Age display
# values (col G) without an id column, so the lookup there always misses and
# the mixed-case display ("Preschool") went out as the id — that is the
# "BY Age still wrong" UAT feedback.
BY_AGE_LOV_ID: dict[str, str] = {
    v: v.upper() for v in ("Adult", "Kids", "All Ages", "Infant", "Preschool", "Grade School")
}

# BCI has no 1st-ingestion default: "Mapping to STIBO" row 219 says
# "1st ingestion : Manual Input", so the attribute is left blank until the
# 2nd ingestion supplies the recap "BCI" column.

# LOTTO Category/Code Category to Product Division
LOTTO_CATEGORY_TO_DIVISION: dict[str, str] = {
    "OUTDOOR":      "F",   # Footwear
    "CASUAL":       "F",
    "KIDS":         "F",
    "RUNNING":      "F",
    "SPORT":        "F",
    "FOOTWEAR":     "F",
    "FW":           "F",
    "APPAREL":      "A",
    "APP":          "A",
    "ACCESSORIES":  "E",
    "ACC":          "E",
}

# LOTTO Article Type → 1-char code for InboundGenericCode
LOTTO_ARTICLE_TYPE_CODE: dict[str, str] = {
    "LICENSE":   "R",
    "SSE":       "X",
    "WHOLESALE": "W",
    "SAMPLE":    "S",
}

# LOTTO Size mapping to SAP 3-char codes (footwear sizes)
LOTTO_SIZE_TO_SAP: dict[str, str] = {
    "36":   "036",
    "37":   "037",
    "38":   "038",
    "39":   "039",
    "40":   "040",
    "41":   "041",
    "42":   "042",
    "43":   "043",
    "44":   "044",
    "45":   "045",
    "46":   "046",
    "47":   "047",
    "48":   "048",
    "49":   "049",
    "50":   "050",
    # Youth/Kids sizes
    "28":   "028",
    "29":   "029",
    "30":   "030",
    "31":   "031",
    "32":   "032",
    "33":   "033",
    "34":   "034",
    "35":   "035",
    # Apparel sizes
    "XS":   "XSX",
    "S":    "SXX",
    "M":    "MXX",
    "L":    "LXX",
    "XL":   "XLX",
    "XXL":  "2XL",
    "2XL":  "2XL",
    "3XL":  "3XL",
}

# LOTTO Material Type LOV
LOTTO_MATERIAL_TYPE_LOV: dict[str, str] = {
    "ZHAW": "Direct Articles (QVB)",
    "ZINA": "Intercompany Articles",
    "ZUNB": "Nonvaluated Articles (QB)",
    "ZDIN": "Nonstock Articles (NQVB)",
}

# Default Material Type for LOTTO Licensed
LOTTO_DEFAULT_MATERIAL_TYPE = "ZINA"

# ── recap "MD Category" → Sports Category EN display value ───────
# Licensed scope, Footwear only.  Values transcribed from "Mapping to STIBO"
# rows 288-298 (Category = Sports category) plus the UAT Result rows 58-66 list.
LOV_SPORTS_CATEGORY_EN: dict[str, str] = _nkeyed({
    "Badminton":      "Tennis / Padel",
    "Casual":         "Lifestyle / Casual",
    "Court":          "Tennis / Padel",
    "Court style":    "Lifestyle / Casual",
    "Essential pack": "Lifestyle / Casual",
    "Five-a-side":    "Soccer",
    "Futsal":         "Soccer",
    "Hiking":         "Outdoor / Trail / Hiking",
    "Kids":           "Lifestyle / Casual",
    "Lifestyle":      "Lifestyle / Casual",
    "Outdoor":        "Outdoor / Trail / Hiking",
    "Outdoor shoe":   "Outdoor / Trail / Hiking",
    "Paddle":         "Tennis / Padel",
    "Padel":          "Tennis / Padel",
    "Performance":    "Tennis / Padel",
    "Running":        "Running",
    "Sandals":        "Other",
    "Soccer":         "Soccer",
    "Sport style":    "Lifestyle / Casual",
    "Tennis":         "Tennis / Padel",
})


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
        wb  = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws  = wb["Core Attributes"]
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
        self._load_gender_lov(wb)
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
        """Load Age LOV sheet: col A = display value, col B = LOV ID. Also loads SAPAge and BYAge."""
        sheet_name = next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return

        sap_age_map = {}
        by_age_map  = {}
        age_lov_map = {}  # col A (display) → col B (LOV ID)
        by_age_lov_map = {}  # col G (BY Age display) → col F (LOV ID)

        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display  = str(row[0]).strip()
            age_lov_id   = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            sap_age_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            by_age_val   = str(row[5]).strip() if len(row) > 5 and row[5] else ""
            by_age_disp  = str(row[6]).strip() if len(row) > 6 and row[6] else ""  # col G
            by_age_id    = str(row[5]).strip() if len(row) > 5 and row[5] else ""  # col F

            if age_display and age_lov_id:
                age_lov_map[age_display] = age_lov_id
            if age_display and sap_age_code:
                sap_age_map[age_display] = sap_age_code
            if age_display and by_age_val:
                by_age_map[age_display] = by_age_val
            if by_age_disp and by_age_id:
                by_age_lov_map[by_age_disp] = by_age_id

        self.lovs["AgeLOV"]    = age_lov_map
        self.lovs["ByAgeLOV"]  = by_age_lov_map
        self.lovs["SAPAge"]    = sap_age_map
        self.lovs["BYAge"]     = by_age_map
        log.info("[MDD] Age LOV loaded — AgeLOV: %d, ByAgeLOV: %d, SAPAge: %d, BYAge: %d",
                 len(age_lov_map), len(by_age_lov_map), len(sap_age_map), len(by_age_map))

    def _load_gender_lov(self, wb):
        """Load Gender LOV sheet: col A = display value, col C = LOV ID."""
        sheet_name = next((s for s in wb.sheetnames if "GENDER" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Gender LOV sheet not found")
            return
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        gender_lov: dict[str, str] = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            display = str(row[0]).strip()
            lov_id  = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if display and lov_id:
                gender_lov[display] = lov_id
        self.lovs["GenderLOV"] = gender_lov
        log.info("[MDD] Gender LOV loaded — %d entries", len(gender_lov))

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
                name, code = row[0], row[1]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )


class LottoMDMappingLoader:
    """
    Loads the 'Lotto MD Mapping' tab from the Brand Mapping workbook.
    Col G (index 6)  = source value from input file.
    Col H (index 7)  = SAP Age LOV display value    → used for AT_SAPAge.
    Col I (index 8)  = SAP Gender LOV display value → used for AT_Gender.
    Col J (index 9)  = BY Age LOV display value     → used for AT_BYAge.
    Col K (index 10) = BY Gender LOV display value  → used for AT_BYGender.
    """

    def __init__(self, path: Path):
        self.path          = path
        self.col_g_to_h: dict[str, str] = {}  # {source_upper: sap_age_display}
        self.col_g_to_i: dict[str, str] = {}  # {source_upper: sap_gender_display}
        self.col_g_to_j: dict[str, str] = {}  # {source_upper: by_age_display}
        self.col_g_to_k: dict[str, str] = {}  # {source_upper: by_gender_display}
        self._load()

    def _load(self):
        log.info("[MDMapping] Loading from: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = "Lotto MD Mapping"
        if sheet_name not in wb.sheetnames:
            log.warning("[MDMapping] '%s' sheet not found", sheet_name)
            wb.close()
            return

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        count = 0
        for row in rows:
            if not row or len(row) <= 8:
                continue
            src   = row[6]                               # Col G
            sap_a = row[7]                               # Col H — SAP Age
            sap_g = row[8]                               # Col I — SAP Gender
            by_a  = row[9]  if len(row) > 9  else None  # Col J — BY Age
            by_g  = row[10] if len(row) > 10 else None  # Col K — BY Gender
            if not src:
                continue
            src_str = str(src).strip().upper()
            if sap_a: self.col_g_to_h[src_str] = str(sap_a).strip()
            if sap_g: self.col_g_to_i[src_str] = str(sap_g).strip()
            if by_a:  self.col_g_to_j[src_str] = str(by_a).strip()
            if by_g:  self.col_g_to_k[src_str] = str(by_g).strip()
            if any([sap_a, sap_g, by_a, by_g]):
                count += 1

        wb.close()
        log.info("[MDMapping] %d mappings loaded from '%s'", count, sheet_name)

    def get_sap_age_lov_value(self, raw_value: str) -> str:
        """Return SAP Age LOV display value (col H) for a raw source value."""
        return self.col_g_to_h.get((raw_value or "").strip().upper(), "")

    def get_lov_value(self, raw_value: str) -> str:
        """Return SAP Gender LOV display value (col I) for a raw source value."""
        return self.col_g_to_i.get((raw_value or "").strip().upper(), "")

    def get_by_age_lov_value(self, raw_value: str) -> str:
        """Return BY Age LOV display value (col J) for a raw source value."""
        return self.col_g_to_j.get((raw_value or "").strip().upper(), "")

    def get_by_gender_lov_value(self, raw_value: str) -> str:
        """Return BY Gender LOV display value (col K) for a raw source value."""
        return self.col_g_to_k.get((raw_value or "").strip().upper(), "")


class AttributesListLoader:
    """Loads the lotto Inline (New Format) tab from the NEW - Brand mapping files Template workbook."""

    def __init__(self, path: Path, brand: str = "LOTTO"):
        self.path      = path
        self.brand     = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Use "Lotto(Inline + Licensed)" tab
        sheet_name = "Lotto(Inline + Licensed)"
        if sheet_name not in wb.sheetnames:
            log.warning("[AttrList] '%s' sheet not found", sheet_name)
            wb.close()
            return

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Find header row that contains "Attributes" in column B (index 1)
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
        log.info("[AttrList] %d attributes loaded from '%s'", len(self.attr_map), sheet_name)

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
        # ── FIX: strip .0 suffix (e.g. "888.0" → "888") ─────────
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
        # pass 1: country + sbu + brand
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and ks == s and kb == b:
                return val
        # pass 2: country + brand only (first match)
        for (kc, _, ks, kb), val in self.lookup.items():
            if kc == c and kb == b:
                return val
        return {"brand_type": "", "brand_category": ""}


class LottoRecapLoader:
    """
    Loads the LOTTO Licensed Recap workbook.

    The active sheet is "LOTTO".
    Header row is at row 2 (1-based).

    Each row represents one (Supp Art # / Color) combination.
    """

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self.phase = 1          # 1st or 2nd ingestion, set by _load()
        self._load()

    def _load(self):
        log.info("[Recap-LOTTO] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Use LOT or LOTTO sheet
        if "LOT" in wb.sheetnames:
            target = "LOT"
        elif "LOTTO" in wb.sheetnames:
            target = "LOTTO"
        else:
            target = wb.sheetnames[0]
        log.info("[Recap-LOTTO] Using sheet: '%s'", target)
        
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Find the header row by scoring the first 10 rows against the columns
        # we actually consume.  The old "first row containing any keyword" rule
        # latched onto title/banner rows such as "RECAP SAMPLE DEVELOPMENT -
        # FOOTWEAR CATEGORY", which shifted every column by one and left the
        # documented mappings reading empty cells.
        expected = {
            _norm_key(name)
            for second, first in RECAP_COLUMNS.values()
            for name in (second + first)
        }

        best_idx, best_score = None, 0
        for i, row in enumerate(rows[:10]):
            if not row:
                continue
            score = sum(1 for c in row if c and _norm_key(c) in expected)
            if score > best_score:
                best_idx, best_score = i, score

        if best_idx is None or best_score < 2:
            log.warning(
                "[Recap-LOTTO] No convincing header row found (best score %d) — "
                "defaulting to row 3 (index 2)", best_score,
            )
            hdr_idx = 2  # Row 3 in Excel
        else:
            hdr_idx = best_idx
            log.info(
                "[Recap-LOTTO] Header at index %d (Excel row %d), matched %d known columns",
                hdr_idx, hdr_idx + 1, best_score,
            )

        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]

        log.info("[Recap-LOTTO] Header row %d: %s", hdr_idx, header)

        unknown = [h for h in rows[hdr_idx] if h and _norm_key(h) not in expected]
        if unknown:
            log.info("[Recap-LOTTO] Columns not consumed by this mapping: %s", unknown)

        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        log.info("[Recap-LOTTO] Found columns: %s", list(df.columns))
        
        # Log first data row for debugging
        if len(df) > 0:
            first_row_dict = df.iloc[0].to_dict()
            log.info("[Recap-LOTTO] First data row sample: %s", {k: v for k, v in list(first_row_dict.items())[:5]})

        # Filter out rows where Supp Art # is blank
        supp_art_col = next(
            (c for c in df.columns if "SUPP ART" in c.upper() or c.strip().upper() == "SUPP ART #"), 
            None
        )
        
        if supp_art_col:
            log.info("[Recap-LOTTO] Found supplier article column: '%s'", supp_art_col)
            
            # Log sample values before filtering
            sample_values = df[supp_art_col].head(3).tolist()
            log.info("[Recap-LOTTO] Sample '%s' values before filtering: %s", supp_art_col, sample_values)
            
            df = df[
                df[supp_art_col].notna()
                & (~df[supp_art_col].astype(str).str.strip().isin(["", "None", "nan"]))
            ]
            df = df.rename(columns={supp_art_col: "Supp Art #"})
            
            # Log sample values after filtering and renaming
            if len(df) > 0:
                sample_after = df["Supp Art #"].head(3).tolist()
                log.info("[Recap-LOTTO] Sample 'Supp Art #' values after filtering: %s", sample_after)
        else:
            log.warning("[Recap-LOTTO] 'Supp Art #' column not found! Available columns: %s", list(df.columns))

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        self.phase = detect_ingestion_phase(self.df.columns)
        log.info("[Recap-LOTTO] %d rows loaded (ingestion %d)", len(self.df), self.phase)

# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT") else s


# ══════════════════════════════════════════════════════════════════
# RECAP SOURCE COLUMNS — 1st vs 2nd ingestion
# ══════════════════════════════════════════════════════════════════
#
# The same article is uploaded to the MAP Portal twice:
#
#   1st ingestion — the principal's own recap.  Only the columns the principal
#                   fills are populated; the rest is completed later by the
#                   user in Smartsheet and then in MDTools.
#   2nd ingestion — the completed recap exported back out of MDTools.  It keeps
#                   the original layout and *adds* the finalised columns
#                   ("Updated Image", "Final FOB", "proposed_retail_price", …).
#
# One Lambda handles both: for every field we list the 2nd-ingestion column
# names first and the 1st-ingestion ones second, then pick according to the
# detected phase.  A blank cell always falls through to the other candidate, so
# a partially-filled 2nd ingestion still keeps the principal's original value.
#
# Column names come from "Lotto Mapping Issues and References.xlsx" →
# "Mapping to STIBO", column "Field Name in the Brand File".

RECAP_COLUMNS: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
    #  field           2nd-ingestion columns                1st-ingestion columns
    "supp_art":      ((),                                  ("Supp Art #", "Supp Art#", "Supplier Article", "Article")),
    "style_desc":    (("Principal Style Description",),    ()),
    "colour":        (("Updated Color", "Updated Colour"), ("Color", "Colour")),
    "colour_code":   (("Updated Color Code",),             ("Color Code", "Colour Code")),
    "gender":        ((),                                  ("Gender",)),
    "gender_code":   ((),                                  ("Gender Code",)),
    "age_group":     ((),                                  ("Age Group", "Age")),
    "division":      ((),                                  ("Division",)),
    "md_category":   ((),                                  ("MD Category", "Category")),
    "code_category": ((),                                  ("Code Category",)),
    "size_range":    ((),                                  ("Size Range",)),
    "product_size":  (("ProductSize", "Product Size"),     ()),
    "fob":           (("Final FOB",),                      ("FOB Price", "FOB")),
    "fob_currency":  ((),                                  ("FOB Currency", "Currency")),
    "landed_cost":   (("landed_cost", "Landed Cost"),      ()),
    "retail_price":  (("proposed_retail_price", "Proposed Retail Price"), ()),
    "price_range":   (("Price Range",),                    ()),
    "image":         (("Updated Image",),                  ("Image", "Thumbnail Image")),
    "bci":           (("BCI",),                            ()),
    "eta_date":      ((),                                  ("ETA DATE", "ETA Date", "ETA")),
    "article_type":  (("Article Type",),                   ()),
    "supplier":      (("Vendor Code",),                    ("Supplier", "Vendor")),
    "coo":           (("Country of Origin",),              ()),
    "noa":           (("NOA", "Nature of Article"),        ()),
    "merch_hier":    (("Merchandise Hierarchy",),          ()),
    "season":        ((),                                  ("Season",)),
    "outsole":       ((),                                  ("Outsole Material",)),
    "upper":         ((),                                  ("Upper Material",)),
}

# Columns that only ever exist in a 2nd ingestion — their presence is what
# tells the two uploads apart.
SECOND_INGESTION_MARKERS: tuple[str, ...] = tuple(
    name for second, _first in RECAP_COLUMNS.values() for name in second
)


def detect_ingestion_phase(columns) -> int:
    """Return 1 or 2 for the ingestion this recap workbook represents."""
    present = {_norm_key(c) for c in columns}
    hits = sorted({m for m in SECOND_INGESTION_MARKERS if _norm_key(m) in present})
    phase = 2 if hits else 1
    log.info(
        "[Recap-LOTTO] Ingestion phase %d — 2nd-ingestion columns present: %s",
        phase, ", ".join(hits) if hits else "none",
    )
    return phase


class RecapRow:
    """Header-insensitive, ingestion-aware accessor over one recap row.

    Recap headers drift between seasons and between the principal's file and
    the MDTools export — trailing spaces, embedded newlines ("FOB \nCurrency"),
    "Supp Art #" vs "Supp Art#", upper vs title case.  Every one of those turned
    a documented mapping into a "Not Populated" defect in UAT (rows 7-18), so
    lookups go through _norm_key() instead of exact dict keys.
    """

    def __init__(self, row, phase: int = 1):
        self.phase = phase
        self._cells: dict[str, object] = {}
        for k, v in row.items():
            nk = _norm_key(k)
            if nk and nk not in self._cells:
                self._cells[nk] = v

    def raw(self, field: str):
        """First non-empty cell among the candidates for *field*, or None."""
        second, first = RECAP_COLUMNS.get(field, ((), (field,)))
        names = (second + first) if self.phase >= 2 else (first + second)
        for name in names:
            v = self._cells.get(_norm_key(name))
            if v is None:
                continue
            if str(v).strip() in ("", "None", "nan", "NaT"):
                continue
            return v
        return None

    def get(self, field: str) -> str:
        """Cleaned string value for *field* ('' when absent or blank)."""
        return _s(self.raw(field))

    def has_column(self, field: str) -> bool:
        second, first = RECAP_COLUMNS.get(field, ((), (field,)))
        return any(_norm_key(n) in self._cells for n in (second + first))


# ══════════════════════════════════════════════════════════════════
# LOV RESOLUTION
# ══════════════════════════════════════════════════════════════════

def _lov_id(mdd, lov_names, display: str, fallback: dict | None = None) -> str:
    """Resolve a LOV *display* value to the id STIBO expects.

    Order: the MDD LOV sheets, then the code table documented in
    "Mapping to STIBO", then the display value itself.

    Never returns "" for a non-empty display.  ``_val`` writes an element with
    either an ``ID`` or text — and STIBO silently drops a LOV attribute that
    arrives as bare text.  That is precisely why SAP Age, BY Age, SAP Gender and
    BY Gender all came back "Not Populated" in UAT: the MDD lookup missed, the
    display value was still truthy, and the fallback branch was skipped.
    """
    disp = (display or "").strip()
    if not disp:
        return ""

    if mdd is not None:
        want = _norm_key(disp)
        for lov_name in lov_names:
            for value_name, value_id in (mdd.lovs.get(lov_name) or {}).items():
                if _norm_key(value_name) == want and str(value_id).strip():
                    return str(value_id).strip()

    if fallback:
        for k, v in fallback.items():
            if _norm_key(k) == _norm_key(disp):
                return v

    return disp


_CURRENCY_ALIASES = _nkeyed({
    "US$": "USD", "$": "USD", "USD$": "USD", "US Dollar": "USD", "US Dollars": "USD",
    "Rp": "IDR", "RMB": "CNY", "Euro": "EUR",
})


def _norm_currency(raw: str) -> str:
    """Normalise a recap FOB currency cell to a 3-letter ISO LOV id."""
    v = (raw or "").strip()
    if not v:
        return ""
    alias = _CURRENCY_ALIASES.get(_norm_key(v))
    if alias:
        return alias
    m = re.search(r"\b([A-Za-z]{3})\b", v)
    if m:
        return m.group(1).upper()
    return re.sub(r"[^A-Z]", "", v.upper())[:3]


def _resolve_gender_age(mapped: dict, md_mapping=None) -> None:
    """Fill the four SAP/BY Gender + Age display values on *mapped*.

    Priority is the brand-mapping workbook ("Lotto MD Mapping" tab) when it has
    a row for the value, then the tables transcribed from the "BY Age & Gender"
    sheet.  Gender drives Gender; the recap "Age Group" column drives Age, and
    only when that column is absent does Gender stand in for it.
    """
    gender_raw = mapped.get("gender_raw", "")
    age_raw    = mapped.get("age_group", "")

    sap_g, by_g = LOTTO_GENDER_MAP.get(_norm_key(gender_raw), ("", ""))
    if md_mapping is not None:
        sap_g = md_mapping.get_lov_value(gender_raw) or sap_g
        by_g  = md_mapping.get_by_gender_lov_value(gender_raw) or by_g
    if not sap_g and gender_raw:
        sap_g = gender_raw.strip().title()
    if not by_g:
        by_g = sap_g

    age_key = age_raw or LOTTO_GENDER_TO_AGE_GROUP.get(_norm_key(gender_raw), "Adult")
    sap_a, by_a = LOTTO_AGE_GROUP_MAP.get(_norm_key(age_key), ("", ""))
    if md_mapping is not None:
        sap_a = (md_mapping.get_sap_age_lov_value(age_key)
                 or md_mapping.get_sap_age_lov_value(gender_raw) or sap_a)
        by_a  = (md_mapping.get_by_age_lov_value(age_key)
                 or md_mapping.get_by_age_lov_value(gender_raw) or by_a)
    if not sap_a:
        sap_a = "Adults"
    if not by_a:
        by_a = LOV_BY_AGE.get(_norm_key(sap_a), sap_a)

    mapped["sap_gender_display"] = sap_g
    mapped["by_gender_display"]  = by_g
    mapped["sap_age_display"]    = sap_a
    mapped["by_age_display"]     = by_a


def _find_currency_mdd_source_file() -> Path | None:
    """Find MDD workbook with Country LOV and Retail Price Currency LOV sheets."""
    candidates: list[Path] = []
    for d in (MDD_DIR, ATTR_DIR):
        if d.exists():
            candidates += list(d.glob("*.xlsx")) + list(d.glob("*.xlsm"))
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for p in candidates:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            s_names = [s.upper() for s in wb.sheetnames]
            has_country = any("COUNTRY" in s and "LOV" in s and "SIZE" not in s and "ORIGIN" not in s for s in s_names)
            has_currency = any("RETAIL PRICE CURRENCY" in s for s in s_names)
            wb.close()
            if has_country and has_currency:
                log.info("[MDD] Selected currency workbook: %s", p.name)
                return p
            log.info("[MDD] Skipping %s — no retail currency LOV sheets", p.name)
        except Exception as e:
            log.warning("[MDD] Cannot inspect %s: %s", p.name, e)
    return None


def _load_retail_currency_mdd(path: Path | None = None) -> dict:
    """Load Country LOV and Retail Price Currency LOV from MDD workbook."""
    empty = {"country_id_to_name": {}, "retail_price_currency_lov": {}}
    if path is None:
        path = _find_currency_mdd_source_file()
    if path is None or not path.exists():
        log.warning("[MDD] No workbook found for retail price currency lookup.")
        return empty

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        log.warning("[MDD] Cannot open retail currency workbook %s: %s", path.name, e)
        return empty

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
                v0 = _s(row[0]).strip()
                v1 = _s(row[1]).strip()
                if v0 and v1:
                    country_id_to_name[v0.upper()] = v1
                    country_id_to_name[v1.upper()] = v0

    retail_price_currency_lov: dict[str, str] = {}
    currency_sheet = next(
        (s for s in wb.sheetnames if "RETAIL PRICE CURRENCY" in s.upper()),
        None,
    )
    if currency_sheet:
        # Data starts at row 3 in "Retail Price Currency LOV" sheet
        for row in wb[currency_sheet].iter_rows(min_row=3, values_only=True):
            if row and row[0] and len(row) > 1 and row[1]:
                retail_price_currency_lov[_s(row[0]).strip().upper()] = _s(row[1]).strip()

    wb.close()
    log.info(
        "[MDD] Retail currency lookup loaded — countries=%d currencies=%d",
        len(country_id_to_name), len(retail_price_currency_lov),
    )
    return {
        "country_id_to_name": country_id_to_name,
        "retail_price_currency_lov": retail_price_currency_lov,
    }


def _resolve_retail_price_currency(country_code: str, mdd: dict | None = None) -> str:
    """country_code (e.g. 'ID', 'KH') → Retail Price Currency LOV ID (e.g. 'IDR', 'USD')."""
    fallback_map = {
        "ID": "IDR", "PH": "PHP", "TH": "THB",
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD"
    }
    code_upper = (country_code or "").strip().upper()
    if not code_upper:
        return ""

    if mdd:
        country_name = mdd.get("country_id_to_name", {}).get(code_upper, "")
        retail_currency_lov = mdd.get("retail_price_currency_lov", {})

        if country_name and retail_currency_lov:
            country_name_upper = country_name.strip().upper()

            # Special case mapping for countries like Cambodia where country name != currency name
            if code_upper == "KH" or "CAMBODIA" in country_name_upper:
                for currency_name, c_code in retail_currency_lov.items():
                    if "UNITED STATES" in currency_name or c_code == "USD":
                        return c_code
                return "USD"

            if country_name_upper in retail_currency_lov:
                return retail_currency_lov[country_name_upper]

            country_root = country_name_upper.rstrip("S").rstrip("N")
            for lov_entry, currency in retail_currency_lov.items():
                lov_entry_upper = lov_entry.upper()
                if (country_name_upper in lov_entry_upper or 
                    country_root in lov_entry_upper or
                    lov_entry_upper.split()[0].rstrip("N") in country_name_upper):
                    return currency

    # Fallback if MDD lookup didn't yield a match
    return fallback_map.get(code_upper, "")


def _fmt_date(v) -> str:
    """Format a date to DD-MM-YYYY (used by AT_IncomingMonth)."""
    if isinstance(v, datetime):
        return v.strftime("%d-%m-%Y")
    if hasattr(v, "strftime"):
        return v.strftime("%d-%m-%Y")
    raw = _s(v)
    if not raw:
        return raw
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            pass
    return raw


def _size_to_sap_code(size_raw: str) -> str:
    """
    Convert a single LOTTO size string to a 3-char SAP size code.
    Falls back to a zero-padded or truncated version of the raw value.
    """
    key = size_raw.strip().upper()
    if key in LOTTO_SIZE_TO_SAP:
        return LOTTO_SIZE_TO_SAP[key]
    
    # Try numeric conversion (e.g., "37" → "037")
    if key.isdigit():
        return key.zfill(3)
    
    # Generic fallback: strip non-alphanumeric, take first 3 chars, pad to 3
    cleaned = re.sub(r"[^A-Z0-9]", "", key)[:3].ljust(3, "X")
    return cleaned


def _parse_season(season: str) -> tuple[str, str]:
    """
    Split a season token into (sea_code, sea_year).
    
    Examples:
        SS26  → ("SS", "2026")
        FW27  → ("FW", "2027")
        SP2028 → ("SP", "2028")
    """
    s = (season or "").strip().upper()
    if len(s) < 3:
        return s, ""
    
    # Match 2-letter code + 2 or 4 digit year
    match = re.match(r"^([A-Z]{2})(\d{2,4})$", s)
    if match:
        code = match.group(1)
        year_part = match.group(2)
        # Convert 2-digit year to 4-digit
        if len(year_part) == 2:
            year = f"20{year_part}"
        else:
            year = year_part
        return code, year
    
    return s[:2], ""


def map_article_lotto(recap_row, brand_code: str = "LOT", phase: int = 1) -> dict:
    """Map one LOTTO recap row -> unified article dict.

    ``phase`` is the ingestion (1 or 2) as returned by detect_ingestion_phase();
    it decides which source column wins for the fields that differ between the
    principal's upload and the MDTools export — see RECAP_COLUMNS.

    Source columns (Mapping to STIBO, "Field Name in the Brand File"):
      Supp Art #    -> AT_PrincipalStyleCode  (mirrored to the generic code)
      Age Group     -> AT_PrincipalAgeDescription, AT_SAPAge, AT_BYAge
      Gender        -> AT_Gender, AT_BYGender, AT_PrincipalGenderDescription
      Division      -> AT_PrincipalMerchandiseHierarchyL1
      MD Category   -> AT_PrincipalMerchandiseHierarchyL2, AT_SportsCategoryEN
      FOB Currency  -> AT_FOBCurrency
      ETA DATE      -> AT_IncomingMonth
      Image         -> AT_ThumbnailImage   (2nd ingestion: Updated Image)
      BCI           -> AT_BCI              (1st ingestion: default Commercial)
    """
    row = recap_row if isinstance(recap_row, RecapRow) else RecapRow(recap_row, phase)

    # ── Extract raw values from row ──────────────────────────────
    supp_art         = row.get("supp_art")
    color            = row.get("colour")
    color_code       = row.get("colour_code")
    gender           = row.get("gender")
    gender_code_raw  = row.get("gender_code")
    code_category    = row.get("code_category")
    article_type_raw = row.get("article_type")
    md_category      = row.get("md_category")
    division_col     = row.get("division")
    size_range       = row.get("size_range")
    outsole          = row.get("outsole")
    upper            = row.get("upper")
    supplier         = row.get("supplier")
    fob_raw          = row.get("fob")
    currency_raw     = row.get("fob_currency")
    season           = row.get("season")
    age_group        = row.get("age_group")
    image            = row.get("image")
    bci              = row.get("bci")
    eta_date         = row.raw("eta_date")

    # ── Derived fields ──────────────────────────────────────────
    # Legacy single-char SAP codes, still used to build the generic code.
    sap_gender_display = LOTTO_GENDER_MAP.get(_norm_key(gender), ("", ""))[0]
    gender_code = SAP_GENDER_LOV_ID.get(sap_gender_display, "U")

    age_key  = age_group or LOTTO_GENDER_TO_AGE_GROUP.get(_norm_key(gender), "Adult")
    sap_age_display = LOTTO_AGE_GROUP_MAP.get(_norm_key(age_key), ("Adults", ""))[0]
    age_code = SAP_AGE_LOV_ID.get(sap_age_display, "AD")

    # Division: prefer the explicit "Code Category", then "Division",
    # then "MD Category".
    div_source = code_category or division_col or md_category
    # PPH parent still falls back to Footwear, as it always has.
    div_letter = LOTTO_CATEGORY_TO_DIVISION.get(div_source.strip().upper(), "") or "F"

    # Footwear gates AT_SportsCategoryEN and AT_CountrySize, so it needs
    # positive evidence — never the "F" default that div_letter carries.
    is_footwear = (
        _norm_key(division_col) in ("FOOTWEAR", "FW")
        or _norm_key(code_category) in ("FOOTWEAR", "FW", "F")
    )

    # Article Type — 1st ingestion has no column and defaults to Licensed.
    by_art_type = article_type_raw or "License"

    # FOB: extract numeric value
    def _extract_price(v: str) -> str:
        if not v:
            return ""
        m = re.search(r"[\d.]+", str(v))
        return m.group() if m else str(v).strip()

    fob_str = _extract_price(fob_raw)

    # FOB Currency -> 3-letter ISO id (UAT Result rows 7-9: "Not Populated")
    currency = _norm_currency(currency_raw)

    # Country of Origin — manual input on the 1st ingestion, a column on the
    # 2nd.  CN stays the default the portal has always used.
    coo = row.get("coo") or "CN"

    # Generate model name (AI-generated format per attributes list)
    # Example: "LOTTO FH240429 MALE BLACK"
    model_name = f"LOTTO {supp_art} {gender.upper()} {color.upper()}".strip()

    # ── InboundGenericCode formula ────────────────────────────────
    # 3-char Brand Code + 1-char Article Type + 1-digit Year
    # + 1-char Code Category + last 4 chars Supp Art #
    # + 1-char Gender Code + 1-char Color Code  (total = 12)

    # Article Type (1 char)
    art_type_char = LOTTO_ARTICLE_TYPE_CODE.get(by_art_type.upper(), "R")

    # Season year digit — last digit of the 2-digit year in season token
    # e.g. SS26 -> "6", FW27 -> "7"
    _season_match = re.match(r"^[A-Z]{1,2}(\d{2,4})$", season.upper())
    season_year_digit = _season_match.group(1)[-1] if _season_match else ""

    # Code Category (1 char, alphanumeric only)
    code_cat_char = re.sub(r"[^A-Z0-9]", "", code_category.upper())[:1]

    # Last 4 alphanumeric chars of Supp Art #
    supp_art_alnum = re.sub(r"[^A-Z0-9]", "", supp_art.upper())
    supp_art_last4 = supp_art_alnum[-4:].rjust(4, "0") if supp_art_alnum else "0000"

    # Gender Code (1 char) — fall back to the derived SAP gender when the recap
    # has no "Gender Code" column, so the code keeps its 12-char shape.
    gender_char = re.sub(r"[^A-Z0-9]", "", (gender_code_raw or gender_code).upper())[:1]

    # Color Code (1 char)
    color_char = re.sub(r"[^A-Z0-9]", "", color_code.upper())[:1]

    generic_code = (
        f"{brand_code[:3]}"
        f"{art_type_char}"
        f"{season_year_digit}"
        f"{code_cat_char}"
        f"{supp_art_last4}"
        f"{gender_char}"
        f"{color_char}"
    )

    # colour_token still derived from color_code for SAP variant use
    color_code_clean = re.sub(r"[^A-Z0-9]", "", (color_code or color).upper())[:3]

    # Colour token (3-char SAP token for variants)
    colour_token = color_code_clean.ljust(3, "X")

    # Parse sizes into a list of dicts with size and sap_size_code.
    # 1st ingestion: the "Size Range" column.  2nd ingestion: "ProductSize"
    # carries the finalised per-article size instead.
    sizes_list: list[dict] = []
    size_source = row.get("product_size") or size_range
    if size_source:
        # Handle different separators: comma, hyphen, etc.
        # Examples: "36-40", "36, 37, 38", "S/M/L"
        raw_sizes = []
        if "-" in size_source and "," not in size_source:
            # Range format like "36-40"
            parts = size_source.split("-")
            if len(parts) == 2:
                try:
                    start = int(parts[0].strip())
                    end = int(parts[1].strip())
                    raw_sizes = [str(i) for i in range(start, end + 1)]
                except ValueError:
                    # Not numeric range, treat as single size
                    raw_sizes = [size_source.strip()]
        else:
            # Comma or slash separated
            separators = [",", "/", ";"]
            raw = size_source
            for sep in separators:
                if sep in raw:
                    raw_sizes = [s.strip() for s in raw.split(sep) if s.strip()]
                    break
            if not raw_sizes:
                # Single size
                raw_sizes = [size_source.strip()]

        # Convert raw sizes to dicts with SAP size codes
        for size in raw_sizes:
            sap_code = _size_to_sap_code(size)
            sizes_list.append({
                "size": size,
                "sap_size_code": sap_code
            })

    return {
        # Core identifiers
        "article_no":        supp_art,
        "sap_style_code":    generic_code[3:],  # generic without the brand code
        "model_name":        model_name,
        "brand_code":        brand_code,
        "style_desc":        row.get("style_desc"),

        # Colour
        "colour":            color,
        "colour_code":       color_code if color_code else color,
        "colour_token":      colour_token,       # 3-char SAP token

        # Gender / Age
        "gender_code":       gender_code,
        "gender_code_raw":   gender_code_raw,
        "gender_raw":        gender,
        "age_code":          age_code,
        "age_group":         age_group,          # Age Group column (drives SAP/BY Age)

        # Classification
        "division":          md_category,        # raw category for hierarchy
        "division_col":      division_col,       # Division column   -> PMH L1
        "md_category":       md_category,        # MD Category column -> PMH L2
        "code_category":     code_category,      # raw code category
        "div_letter":        div_letter,         # SAP division letter
        "category":          md_category,        # kept for backwards compat
        "is_footwear":       is_footwear,
        "sub_category":      "",
        "sports_cat_en":     LOV_SPORTS_CATEGORY_EN.get(_norm_key(md_category), ""),

        # Commercial
        "article_type":      by_art_type,        # License
        "art_category":      "1",                # Generic (has variants)
        "collection1":       "",
        "collection2":       "",
        "franchise":         "Licensed",
        "noa":               row.get("noa"),
        "merch_hier":        row.get("merch_hier"),

        # Materials
        "outsole_material":  outsole,
        "upper_material":    upper,

        # Pricing
        "fob":               fob_str,
        "currency":          currency,
        "rrp":               _extract_price(row.get("retail_price")),
        "landed_cost":       _extract_price(row.get("landed_cost")),
        "price_range":       row.get("price_range"),

        # Origin / Supplier
        "coo":               coo,
        "supplier":          supplier,

        # Size
        "size_range":        size_range,
        "sizes_list":        sizes_list,

        # BY / image / schedule
        "image":             image,              # Image column      -> AT_ThumbnailImage
        "bci":               bci,                # BCI column        -> AT_BCI
        "eta_date":          eta_date,           # ETA DATE column   -> AT_IncomingMonth

        # Derived codes
        "generic_code":      generic_code,
        "season_raw":        season,
        "ingestion_phase":   row.phase,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader | None) -> list[str]:
    """Validate mandatory MDD attributes against the mapped article dict."""
    warns = []
    art   = mapped["article_no"]
    if mdd is None:
        return warns
    field_checks = {
        "article_no":         "AT_PrincipalStyleCode",
        "sap_gender_display": "AT_Gender",
        "sap_age_display":    "AT_SAPAge",
        "brand_code":         "AT_Brand",
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

def _val(
    parent:   ET.Element,
    attr_id:  str,
    value:    str = "",
    id_val:   str = "",
    derived:  bool = False,
) -> ET.Element | None:
    # """Append a <Value AttributeID="..."> element to *parent*."""
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


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = "") -> None:
    """Append a <MultiValue> → <Value ID="..."> element."""
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def resolve_brand_lov_id(mdd: MDDLoader | None, brand_code: str) -> tuple[str, str]:
    """Resolve brand identifiers from source brand code."""
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


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
) -> None:
    """Write all Generic-level <Value> elements for a LOTTO article."""
    written: set[str] = set()

    def _w(attr_id: str, value: str = "", id_val: str = "", derived: bool = False):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val, derived=derived)
            written.add(attr_id)

    def _mw(attr_id: str, id_val: str):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    sbu_code, _ = _lov(sbu, LOV_SBU, sbu)
    _mw("AT_SBU", sbu_code)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    brand_lov_id = art["brand_code"]
    _w("AT_Brand",      id_val=brand_lov_id)
    _w("AT_BrandGroup", id_val=brand_name.upper())

    # ── Principal identifiers ────────────────────────────────────
    # UAT Result rows 2-4: "Revise -> Mirror to generic" — AT_PrincipalStyleCode
    # carries the generic code, not the raw "Supp Art #".
    _w("AT_PrincipalStyleCode",  art["generic_code"])
    _w("AT_PrincipalColorName",  art["colour"])
    _w("AT_PrincipalColorCode",  art["colour_code"])
    _w("AT_SAPStyleCode",        art["sap_style_code"])
    if art.get("style_desc"):
        _w("AT_PrincipalStyleDescription", art["style_desc"])
    
    # AT_PrincipalSize → Size range
    if art.get("size_range"):
        _w("AT_PrincipalSize", art["size_range"])

    # ── Colour (SAP Color) ───────────────────────────────────────
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id = color_lov.get(art["colour"], "") \
                or color_lov.get(art["colour"].upper(), "")

    if colour_id:
        if colour_id.isdigit():
            colour_id = colour_id.zfill(3)
        _w("AT_Color", id_val=colour_id)
    # else:
    #     _w("AT_Color", value=art["colour"].upper())
    # else: skip AT_Color — invalid to send raw color name as LOV value


    # AT_InboundGenericCode — same formula as KEY_InboundArticle
    at_generic_val = art["generic_code"]
    art["at_generic_val"] = at_generic_val
    _w("AT_InboundGenericCode", at_generic_val)
    _w("AT_Generic", at_generic_val)

    # ── Gender (UAT Result rows 36-51: "Not Populated") ──────────
    # Source: recap "Gender"; mapping table: sheet "BY Age & Gender".
    # Both attributes are LOVs, so they must always carry an ID — a <Value>
    # with text but no ID is dropped by STIBO, which is what "Not Populated"
    # meant here.  _lov_id() guarantees a non-empty id.
    sap_gender = art.get("sap_gender_display", "")
    by_gender  = art.get("by_gender_display", "")

    if sap_gender:
        _w("AT_Gender", sap_gender,
           id_val=_lov_id(mdd, ("GenderLOV", "Gender", "SAP Gender"),
                          sap_gender, SAP_GENDER_LOV_ID))
    if by_gender:
        _w("AT_BYGender", by_gender,
           id_val=_lov_id(mdd, ("BYGenderLOV", "BY Gender", "GenderLOV", "Gender"),
                          by_gender))

    _w("AT_PrincipalGenderDescription", art["gender_raw"])
    if art.get("gender_code_raw"):
        _w("AT_PrincipalGenderCode", art["gender_code_raw"])

    # ── Age (UAT Result rows 19-35: "Not Populated") ─────────────
    # Source: recap "Age Group" (Mapping to STIBO rows 81 / 210); the Gender
    # column only stands in when the recap has no Age Group at all.
    sap_age = art.get("sap_age_display", "")
    by_age  = art.get("by_age_display", "")

    if sap_age:
        _w("AT_SAPAge", sap_age,
           id_val=_lov_id(mdd, ("AgeLOV", "Age", "SAP Age"), sap_age, SAP_AGE_LOV_ID))
    if by_age:
        # MDD "BY Age" LOV first (if the MDD ever gains an id column), then
        # BY_AGE_LOV_ID.  "AgeLOV" is deliberately not consulted: it holds the
        # SAP ages and would turn BY Age "All Ages" into the SAP id "AA".
        by_age_id = _lov_id(mdd, ("ByAgeLOV", "BY Age"), by_age, BY_AGE_LOV_ID)
        if by_age_id == by_age:
            by_age_id = by_age.upper()
        _w("AT_BYAge", by_age, id_val=by_age_id)

    # AT_PrincipalAgeDescription — UAT Result rows 5-6 ask for it mirrored onto
    # the generic, so it is written for every article: the raw "Age Group" when
    # the recap has one, otherwise the value the Age mapping resolved to.
    age_description = art.get("age_group") or by_age
    if age_description:
        _w("AT_PrincipalAgeDescription", age_description)

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    if not year_match:
        short_match = re.search(r"\d{2}$", sea_raw)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    _w("AT_SeasonYear", season_year)

    # ── Country of Origin ────────────────────────────────────────
    # coo_code, coo_label = _lov(art["coo"], LOV_COUNTRY_ORIGIN, art["coo"])
    # _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    # ── Article Category ─────────────────────────────────────────
    # Generic (1) for Licensed recap with size variants
    _w("AT_SAPArticleCategory", id_val=art.get("art_category", "1"))

    # ── Merchandise Hierarchy (UAT Result rows 10-13) ────────────
    # L1 → recap "Division", L2 → recap "MD Category".
    if art.get("division_col"):
        _w("AT_PrincipalMerchandiseHierarchyL1", art["division_col"])
    if art.get("md_category"):
        _w("AT_PrincipalMerchandiseHierarchyL2", art["md_category"])

    # ── Thumbnail Image (UAT Result rows 14-15) ──────────────────
    # 1st ingestion: "Image"; 2nd ingestion: "Updated Image".
    if art.get("image"):
        _w("AT_ThumbnailImage", art["image"])

    # ── Incoming Month → recap "ETA DATE" (UAT Result row 18) ────
    if art.get("eta_date"):
        incoming_month = _fmt_date(art["eta_date"])
        if incoming_month:
            _w("AT_IncomingMonth", incoming_month)

    # ── Article Type & BCI (UAT Result rows 16-17) ───────────────
    # 1st ingestion defaults to Licensed; the 2nd carries an "Article Type"
    # column.
    art_type = art.get("article_type") or "License"
    _w("AT_BYArticleType", art_type, id_val=art_type)

    # BCI — "Mapping to STIBO" row 219: 1st ingestion = Manual Input (the user
    # fills it in Smartsheet later), 2nd ingestion = recap "BCI" column.  So
    # nothing is sent unless the recap actually carries a BCI value; no
    # default is invented on the 1st ingestion.
    bci_display = art.get("bci")
    if bci_display:
        _w("AT_BCI", bci_display, id_val=_lov_id(mdd, ("BCI", "BCILOV"), bci_display))

    # ── System indicators ────────────────────────────────────────
    _w("AT_SAPProductFlag", "A",  id_val="A")

    # ── Pricing ──────────────────────────────────────────────────
    fob_value = art.get("fob")
    if fob_value:
        _w("AT_FOB", fob_value)
    
    # AT_FOBCurrency — recap "FOB Currency", normalised to a 3-letter ISO id
    # (UAT Result rows 7-9: "Not Populated").
    currency_value = art.get("currency")
    if currency_value:
        _w("AT_FOBCurrency", currency_value,
           id_val=_lov_id(mdd, ("FOB Currency", "FOBCurrency", "Currency"),
                          currency_value))
    
    rpc_id = _resolve_retail_price_currency(art.get("country_code", ""), art.get("mdd_currency"))
    if rpc_id:
        _w("AT_RetailPriceCurrency", id_val=rpc_id)

    # ── Vendor identification ────────────────────────────────────
    if art.get("supplier"):
        _w("AT_MainVendorIdentification", "1")

    # ── Collections / hierarchy ──────────────────────────────────
    if art.get("collection1"):
        _w("AT_Collection1", art["collection1"])
    if art.get("collection2"):
        _w("AT_Collection2", art["collection2"])

    # ── Material Type ────────────────────────────────────────────
    _w("AT_MaterialType", "", id_val=LOTTO_DEFAULT_MATERIAL_TYPE)

    # ── Pricing Distribution Channel ─────────────────────────────
    _w("AT_PricingDistributionChannel", "", id_val="01")

    # ── Country (destination country) ────────────────────────────
    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)
    
    # ── Brand Type / Brand Category ───────────────────────────────
    _w("AT_BrandType",     art.get("brand_type",     ""))
    # _w("AT_BrandCategory", art.get("brand_category", ""))

    # ── Sports Category EN (UAT Result rows 52-75) ───────────────
    # Licensed scope is Footwear only; the source is the recap "MD Category"
    # column (UAT rows 73-74), mapped through LOV_SPORTS_CATEGORY_EN.
    if art.get("is_footwear"):
        sc_display = art.get("sports_cat_en", "")
        if sc_display:
            sc_id = _lov_id(
                mdd,
                ("Sports Category", "Sports Category EN", "SportsCategoryLOV"),
                sc_display,
            )
            if sc_id != sc_display and sc_id.isdigit():
                sc_id = sc_id.zfill(2)
            _w("AT_SportsCategoryEN", sc_display, id_val=sc_id)
        elif art.get("md_category"):
            log.warning(
                "[SportsCategory] No mapping for MD Category %r (article %s)",
                art.get("md_category"), art.get("article_no"),
            )

    # ── Country Size — default EUR for Footwear (Mapping row 307) ─
    if art.get("is_footwear"):
        _w("AT_CountrySize", "", id_val="EU")

    # ── UOM (Unit of Measure) ────────────────────────────────────
    uom_code = "EA"
    _w("AT_UOM", LOV_UOM.get(uom_code, "Each"), id_val=uom_code)


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def _add_variant_values(
    vals_el: ET.Element,
    art:     dict,
    size:    str,
    sap_size_code: str,
) -> None:
    """Write Variant-level <Value> elements for one size."""

    def _w(attr_id, value="", id_val=""):
        _val(vals_el, attr_id, value=value, id_val=id_val)

    # ── KEY_Variant defining attributes (required for key resolution) ──
    b_code  = art.get("brand_code", "LOT")
    _w("AT_Brand",       id_val=b_code)
    # Mirror to generic here as well, so a variant key resolves to the same
    # style code the Generic carries (UAT Result rows 2-4).
    _w("AT_PrincipalStyleCode", art["generic_code"])
    _w("AT_Size",               sap_size_code,         id_val=sap_size_code)

    # ── Size attributes ────────────────────────────────────────
    _w("AT_PrincipalSizeCode", size)
    _w("AT_SAPSize",           sap_size_code, id_val=sap_size_code)


def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand}Batches"

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


def build_product_xml(
    art:        dict,
    brand:      str,
    brand_code: str,
    comp_code:  str,
    sbu:        str,
    season_id:  str,
    mdd=None
) -> str:
    """Build a <Product> XML fragment for one LOTTO article."""
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code      = art.get("brand_code", brand_code) or brand_code
    div_letter  = art.get("div_letter", "F")
    parent_id   = f"PPH_{div_letter}-TempSubCat"

    # KEY_Article
    colour_id_for_key = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id_for_key = color_lov.get(art["colour"], "") \
                        or color_lov.get(art["colour"].upper(), "")
    if colour_id_for_key and colour_id_for_key.isdigit():
        colour_id_for_key = colour_id_for_key.zfill(3)

    sap_color_for_key = colour_id_for_key if colour_id_for_key else art["colour_token"]
    key_article = f"{b_code}{art['article_no']}{sap_color_for_key}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = art["generic_code"]

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["model_name"] or article_no
    )

    # ── Classification references ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Values ───────────────────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one LOTTO row."""
    (row, brand_code, mdd_loader, season, country_code, md_mapping,
     brand_type, brand_category, mdd_currency, phase) = row_tuple

    mapped = map_article_lotto(row, brand_code=brand_code, phase=phase)
    mapped["season"]         = season
    mapped["country_code"]   = country_code
    mapped["brand_type"]     = brand_type
    mapped["brand_category"] = brand_category
    mapped["mdd_currency"]   = mdd_currency

    # SAP / BY Gender + Age display values.  The brand-mapping workbook wins
    # when it has a row; otherwise the "BY Age & Gender" tables do.  LOV ids
    # are resolved later, in _add_generic_values, against the MDD.
    _resolve_gender_age(mapped, md_mapping)

    log.debug(
        "[Gender/Age] gender=%r age_group=%r → SAP %s / %s   BY %s / %s",
        mapped.get("gender_raw"), mapped.get("age_group"),
        mapped["sap_gender_display"], mapped["sap_age_display"],
        mapped["by_gender_display"],  mapped["by_age_display"],
    )

    warns = validate(mapped, mdd_loader)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point for LOTTO Licensed Recap.

    args must have attributes:
        brand       str   e.g. "LOTTO"
        brand_code  str   e.g. "LOT"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SP"
        season      str   e.g. "SS26"
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    def find_brand_mapping_file(d: Path) -> Path | None:
        """Look for 'NEW - Brand mapping files Template.xlsx' or files with brand mapping keywords.
        Always picks the most recently modified file when multiple matches exist."""
        
        if not d.exists():
            return None
        
        # Search for all matching files (exact name or keyword matches)
        keywords = ["brand mapping", "brand_mapping", "mapping template", "brand template"]
        matches = []

        for file in d.glob("*.xlsx"):
            filename_lower = file.name.lower()
            # Match if it's the exact name OR contains keywords
            if file.name == "NEW - Brand mapping files Template.xlsx" or \
               any(kw in filename_lower for kw in keywords):
                matches.append(file)

        if not matches:
            return None
        
        # Sort by modification time (most recent first)
        matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        selected = matches[0]
        
        if len(matches) > 1:
            log.info("[BrandMapping] Found %d matching file(s), using most recent: %s (modified: %s)", 
                     len(matches), selected.name, 
                     datetime.fromtimestamp(selected.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'))
        else:
            log.info("[BrandMapping] Using file: %s", selected.name)
        
        return selected

    mdd_f      = first(MDD_DIR)
    attr_f     = find_brand_mapping_file(ATTR_DIR)
    recap_files = list(RECAP_DIR.glob("*.xlsx"))

    # ── Mandatory file checks ────────────────────────────────────
    for label, val in [("MDD", mdd_f), ("Brand Mapping", attr_f), ("Recap", recap_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    # ── Log selected files ───────────────────────────────────────
    log.info("═══ SELECTED FILES ═════════════════════════════════")
    log.info("  MDD File           : %s", mdd_f.name if mdd_f else "None")
    log.info("  Brand Mapping File : %s", attr_f.name if attr_f else "None")
    if attr_f:
        log.info("  Brand Mapping Path : %s", str(attr_f))
        log.info("  Brand Mapping Mod  : %s", datetime.fromtimestamp(attr_f.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'))
    log.info("  Recap Files        : %d file(s)", len(recap_files))
    log.info("════════════════════════════════════════════════════")

    mdd    = MDDLoader(mdd_f)

    # The 13 attributes COE reported as failing for Lotto (UAT Result sheet).
    # If one is missing from the MDD it can never be populated no matter what
    # we send, so surface that up front instead of debugging it per article.
    _uat_attrs = [
        "AT_PrincipalStyleCode", "AT_PrincipalAgeDescription", "AT_FOBCurrency",
        "AT_PrincipalMerchandiseHierarchyL1", "AT_PrincipalMerchandiseHierarchyL2",
        "AT_ThumbnailImage", "AT_BCI", "AT_IncomingMonth", "AT_SAPAge",
        "AT_BYAge", "AT_BYGender", "AT_Gender", "AT_SportsCategoryEN", "AT_UOM",
    ]
    _unknown = [a for a in _uat_attrs if a not in mdd.attributes]
    if _unknown:
        log.warning("[MDD] UAT attributes not present in the MDD: %s", ", ".join(_unknown))
    else:
        log.info("[MDD] All %d UAT attributes present in the MDD.", len(_uat_attrs))

    _al    = AttributesListLoader(attr_f, brand=args.brand)
    md_map = LottoMDMappingLoader(attr_f)
    rna    = RNALoader(attr_f)

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Season ID ────────────────────────────────────────────────
    brand_lov_id, brand_lov_label = resolve_brand_lov_id(mdd, args.brand_code)
    log.info(
        "Resolved brand LOV → source_brand_code=%s  brand_label=%s  brand_lov_id=%s",
        args.brand_code, brand_lov_label, brand_lov_id
    )

    # ── Resolve AT_BrandType / AT_BrandCategory ──────────────────
    _country_code = (getattr(args, "country_code", "") or "").strip().upper() or "ID"
    _country_name = _COUNTRY_MAP.get(_country_code, _country_code)

    log.info(
        "[RNA] Querying → country=%r  comp=%r  sbu=%r  brand=%r",
        _country_name, args.comp_code, args.sbu, brand_lov_id,
    )

    _rna_result = rna.get(_country_name, args.comp_code, args.sbu, brand_lov_id)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, args.comp_code, args.sbu, args.brand_code)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, str(int(args.comp_code)), args.sbu, brand_lov_id)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        _rna_result = rna.get(_country_name, str(int(args.comp_code)), args.sbu, args.brand_code)
    if not _rna_result.get("brand_type") and not _rna_result.get("brand_category"):
        log.warning("[RNA] Exact match failed — trying fuzzy (country+sbu+brand)")
        _rna_result = rna.get_fuzzy(_country_name, args.sbu, brand_lov_id)

    _brand_type     = _rna_result.get("brand_type",     "")
    _brand_category = _rna_result.get("brand_category", "")

    log.info(
        "[RNA] country=%s  brand_type='%s'  brand_category='%s'",
        _country_name, _brand_type, _brand_category,
    )

    # ── Load retail currency MDD ─────────────────────────────────
    currency_mdd = _load_retail_currency_mdd(mdd_f)   # reuse already-found MDD file

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for recap_path in recap_files:
        log.info("─── Processing Recap: %s ───", recap_path.name)

        recap = LottoRecapLoader(recap_path)
        if recap.df.empty:
            log.warning("[Recap-LOTTO] Empty dataframe — skipping."); continue

        rows       = [row for _, row in recap.df.iterrows()]
        # rows = rows[:5]  # ← TEST MODE: uncomment to limit to 5 articles
        total_rows = len(rows)
        log.info("Total valid rows: %d", total_rows)

        # ── Output filename ──────────────────────────────────────
        out_name = recap_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        file_country_code = _parse_country_code_from_file_or_args(
            recap_path.name, getattr(args, "country_code", "")
        )
        log.info("[Recap-LOTTO] Country code resolved: '%s' for file: %s", file_country_code, recap_path.name)

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (row, brand_lov_id, mdd, args.season, file_country_code, md_map,
             _brand_type, _brand_category, currency_mdd, recap.phase)
            for row in rows
        ]
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

        # ── Pass 2: stream XML to file ───────────────────────────
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
                    args.comp_code, args.sbu, season_id, mdd=mdd
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

        print("═══ ARTICLE SUMMARY (LOTTO LICENSED RECAP) ═════", flush=True)
        print(f"  Recap rows (valid)     : {total_rows}",   flush=True)
        print(f"  Articles written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb}KB",     flush=True)
        print(f"  Brand Type in XML      : '{_brand_type}'", flush=True)
        print(f"  Brand Category in XML  : '{_brand_category}'", flush=True)
        print("════════════════════════════════════════════════",  flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — LOTTO Licensed Recap v1.0")
    p.add_argument("--brand",      default="LOTTO")
    p.add_argument("--brand-code", default="LOT")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
