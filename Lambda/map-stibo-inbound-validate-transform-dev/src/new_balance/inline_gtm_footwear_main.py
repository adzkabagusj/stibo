"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — NEW BALANCE  (GTM Footwear) v2.0 ║
║   Line List (Footwear GTM) → Stibo STEP XML                      ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : S2'27 GTM 2 APAC Footwear Line List 260722 (Ryan).xlsx
               ("S2 27 GTM 2 APAC Footwear Line List 260722 (Ryan).xlsx")
Primary tab  : FIRST VISIBLE sheet — "S2'27 GTM 2 Footwear Line list"
               (title row index 0, header row index 3, data from index 4;
                the sheet is located dynamically by scanning for a header
                row that contains "Item Number", so a renamed tab still
                loads — same strategy as every NB ETL)

MAPPING AUTHORITY (v2.0 rewrite — every column cross-checked):
  1. "NEW - Brand mapping files Template.xlsx"  → sheet "NEW BALANCE  LINELIST"
     (note: the sheet name carries TWO spaces) — the operative spec for this
     file type: the router fetches THIS file as the "attributes" input for
     linelist_footwear_gtm (see lambda_function._NB_NEW_MAPPING_TRIGGER_TYPES).
  2. "Attributes List_complete_v6_25032026.xlsx" → sheet "NEW BALANCE - LL"
     (v6, 25-03-2026) — cross-check; agrees with (1) on every column below.
  3. "Master Data Dictionary (MAA).xlsx"         → Core Attributes + LOV sheets
     (cardinality / LOV IDs / max chars used at runtime for self-checks).

────────────────────────────────────────────────────────────────────
COLUMN → ATTRIBUTE MAP (v2.0 — every column of the 260722 GTM sheet)
────────────────────────────────────────────────────────────────────
  DIRECT (BM sheet rows 13/14/16/18/21/79/80/41/103/123/233 — "Direct
  from Principal" / "Mapping from Principal", Footwear branch):
    Product Number            → AT_PrincipalStyleCode           [BM row 13]
    Short Product Description → AT_PrincipalStyleDescription    [BM row 14]
                               → <Name> (blank → Item Number fallback)
    Colorway Name             → AT_PrincipalColorName           [BM row 16]
                                 (remove the leading "Item Number - "
                                 prefix when present, per the sheet note)
    Adult Gender/ Kids Closure Type
                              → AT_PrincipalGenderDescription   [BM row 18]
                                 (plain text, truncated to MDD max chars)
    SIZE                      → AT_PrincipalSize                [BM row 21]
                                 (newlines collapsed to spaces)
    Line Plan Business        → AT_PrincipalMerchandiseHierarchyL2 [BM row 79]
                               → AT_SportsCategoryEN (LOV ID)    [BM row 121]
    CATEGORY                  → AT_PrincipalMerchandiseHierarchyL3 [BM row 80]
    Production Factory Name   → AT_CountryOrigin                [BM row 41]
                                 ("PYV - VNM - P00" → middle 3-letter code
                                  VNM → 2-letter LOV ID VN / Vietnam)
    Catalog Material Description
                              → AT_MaterialUpper (LOV close value) [BM row 103]
                               → AT_Collection1                 [BM row 122]
    Series#                   → AT_Collection2                  [BM row 123]
    Intro Date - APAC         → AT_IncomingMonth (DD-MM-YYYY)   [BM row 233]

  DERIVED from Item Number 1st letter [BM rows 37/38/95/96/132/104 —
  "Mapping from 1st digits column 'Item Number'"]:
    M→Men  W→Women  U→Unisex  Y→Youth  G→Grade School  P→Preschool
    I→Infant  N→(NB kids line, grade-school sizing)
      → AT_Gender      (SAP Gender LOV: M/F/U)            [BM row 38]
      → AT_SAPAge      (SAP Age LOV: AD/CH)               [BM row 37]
      → AT_BYGender    (Male / Female / Unisex)           [BM row 96]
      → AT_BYAge       (Adult / Kids / Grade School / …)  [BM row 95]
      → AT_EComAgesCategory (A/G/P/I/Y)                   [BM row 132]
      → AT_Fastening   (see _closure_to_fastening)         [BM row 104]

  CONTEXT (filename/portal metadata — not a source column):
    AT_Country, AT_CompanyCode, AT_SBU, AT_Brand, AT_BrandGroup,
    AT_Season, AT_SeasonYear

  FILE-TYPE / SPEC DEFAULTS (BM rows 78/83/84/77/90/93/130/232):
    AT_PrincipalMerchandiseHierarchyL1 = "Footwear"   [BM row 78 default]
    AT_MaterialType     = ZINA (Intercompany Articles) [BM row 83]
    AT_SAPArticleCategory = 1 / Generic                [BM row 84]
    AT_BYArticleType    = Inline                        [BM row 93: FW default]
    AT_SAPProductFlag   = A / Intercompany              [BM row 77]
    AT_UOM              = EA / Each                     [BM row 90]
    AT_CountrySize      = US  (All US Size — FW default) [BM row 130]
    AT_ArticleStatus    = Active                        [BM row 232]

  PIPELINE KEY (repo convention, every NB ETL):
    KEY_InboundArticle / AT_InboundGenericCode = brand(3) + Item Number
    (row-unique: 2784 distinct Item Numbers in the 260722 sheet).

────────────────────────────────────────────────────────────────────
COLUMNS READ BUT *NOT* MAPPED TO STIBO
────────────────────────────────────────────────────────────────────
  None of the columns below appear as a "Field Name in the Brand File"
  for any attribute row of the two mapping sheets (Footwear branch):

    APAC RANGE?                  (internal range flag: X / DROP / 00)
    TYPE                         (NEW / C-O — duplicates Item - CarryOver/New)
    Primary Image                (BM row 94 "Thumbnail Image" exists but its
                                 Stibo Attribute ID is #N/A — no STEP
                                 attribute is configured yet; images flow
                                 through the Brand Dashboard / BY pipeline,
                                 not the article XML)
    Item - CarryOver/New         (not referenced by any attribute row)
    Region Retail                (retail pricing arrives via BY response)
    Phaseout Date - APAC         (not referenced by any attribute row)
    Region Sales Sample Set      (not referenced by any attribute row)
    GAME PLAN / MARKETING MONTH  (not referenced)
    APAC DTC DA Scale            (not referenced)
    DTC/GKA Exclusive            (not referenced)
    GLOBAL 52 Weeks / Kids NOS   (not referenced)
    Regional Merch Comments      (not referenced)
    Lifecycle                    (Sustain/Launch/Incubate — not referenced)
    MAP                          (not referenced)

  ATTRIBUTES THE SPEC PINS BUT WHOSE SOURCE COLUMN IS ABSENT from the
  GTM sheet (deliberately NOT emitted, same convention as v1.0):
    AT_LaunchingDate  — BM row 166 wants column "SEA Intro Date"; the GTM
                        sheet only carries "Intro Date - APAC" (which feeds
                        AT_IncomingMonth per BM row 233). Not back-filled.
    AT_FOB / AT_FOBCurrency — BM row 60: Footwear branch = N/A.
    AT_BCI            — BM row 97: Footwear = Manual Input.
    AT_Silhouette     — BM row 101: Footwear = Manual Input.
    AT_TechnologyUsed — BM row 158: FW = Manual MD input.
    AT_PrincipalColorCode — BM row 15: Footwear = N/A.
    AT_PrincipalGenderCode / AT_PrincipalAgeCode / AT_PrincipalAgeDescription
                      — BM rows 17/19/20: Not Available.
    AT_PrincipalMerchandiseHierarchyL4 / L5 — BM rows 81/82: Footwear = N/A.

  FORMULA-IN-SYSTEM CODES NOT COMPUTED HERE (BM rows 22-26): AT_SAPStyleCode,
  AT_Generic, AT_GenericDescription, AT_Variant, AT_VariantDescription. Their
  formulas ("Type A" / brand+style+SAP color+SAP size) depend on the SAP
  color/size codes that only arrive later via the BY response / AI image
  analysis — the lambda never guesses them (v1.0 held the same line, then
  labelled "MAA asked not to send it"; v2.0 keeps them out and cites the
  spec's "Formula in System" typing instead).

────────────────────────────────────────────────────────────────────
DEVIATIONS (documented, deliberate)
────────────────────────────────────────────────────────────────────
 1. AT_PrincipalStyleCode now comes from Product Number (BM row 13) — v1.0
    used Item Number because the old GTM export's Product Number was a
    SAP-style code MAA withdrew. The 260722 sheet's Product Number is the
    style-level principal code (470 distinct values) the new spec requires.
    KEY_InboundArticle stays keyed on the row-unique Item Number — the key
    must remain 1:1 with a source row (2784 rows would collapse to 470
    articles if keyed on Product Number).
 2. E-com Ages Category is derived from the Item Number 1st letter (BM row
    132). The v6 "NEW BALANCE - LL" sheet (row 144) instead reads the
    "Adult Gender/ Kids Closure Type" column — but that column mixes closure
    types ("Bungee", "Hook and Loop") into the same field, which has no
    age meaning; the Brand Mapping sheet's rule is deterministic and is the
    one implemented (documented fallback: v6 row-144 mapping).
 3. Unrecognised Item-Number prefixes (e.g. the single "RCVRYSK1" recovery
    slide) fall back to the "Adult Gender/ Kids Closure Type" column and
    then to Unisex/Adult — logged as validation warnings, never guessed.
 4. Sports Category LOV IDs are zero-padded 2-digit strings ("07" for
    Running) — the ID format the live Stibo "Sports Category" LOV accepts
    (same rule as asics/footwear_main.py, anta/linelist_main.py, reebok).
 5. LOV attributes are only emitted when the derived value resolves to a
    real LOV ID (Material - Upper, Fastening, Country Origin); a miss is
    skipped with a warning — writing raw text to a LOV attribute throws
    "Illegal LOV" on STEP import (lesson recorded in reebok/article_master_*).
"""

import re
import os
import sys
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Base directory — overridden by Lambda via LAMBDA_TMP_DIR env var
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist_footwear_gtm"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_footwear_gtm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# SHEET SCOPE — operator instruction: ingest the FIRST VISIBLE sheet
# (the 260722 workbook also carries a hidden "APAC CO Request" tab that
# must never be ingested). Name-based fallbacks cover older exports
# whose tab is called "APAC" / "Line Sheet".
# ══════════════════════════════════════════════════════════════════
SHEET_MODE = "first_visible"          # operator instruction (default)
SHEET_NAME_FALLBACKS = ["APAC", "Line Sheet", "Line List", "Sheet1"]


# ══════════════════════════════════════════════════════════════════
# SOURCE COLUMN CONTRACT — every column this ETL reads, and the
# attributes each one feeds (BM = "NEW BALANCE  LINELIST" sheet row).
# ══════════════════════════════════════════════════════════════════
SOURCE_COLUMNS: dict[str, list[str]] = {
    "Product Number":                    ["AT_PrincipalStyleCode"],
    "Short Product Description":         ["AT_PrincipalStyleDescription", "<Name>"],
    "Colorway Name":                     ["AT_PrincipalColorName"],
    "Adult Gender/ Kids Closure Type":   ["AT_PrincipalGenderDescription",
                                          "AT_Fastening (children branch)",
                                          "AT_Gender/AT_SAPAge fallback"],
    "SIZE":                              ["AT_PrincipalSize"],
    "Item Number":                       ["AT_InboundGenericCode",
                                          "AT_Gender", "AT_SAPAge",
                                          "AT_BYGender", "AT_BYAge",
                                          "AT_EComAgesCategory", "AT_Fastening"],
    "Line Plan Business":                ["AT_SportsCategoryEN",
                                          "AT_PrincipalMerchandiseHierarchyL2"],
    "CATEGORY":                          ["AT_PrincipalMerchandiseHierarchyL3"],
    "Production Factory Name":           ["AT_CountryOrigin"],
    "Catalog Material Description":      ["AT_MaterialUpper", "AT_Collection1"],
    "Series#":                           ["AT_Collection2"],
    "Intro Date - APAC":                 ["AT_IncomingMonth"],
}

# Columns present in the 260722 sheet that no attribute row of the two
# mapping sheets references (see module docstring for the full audit).
UNMAPPED_COLUMNS: dict[str, str] = {
    "APAC RANGE?":               "internal range flag (X / DROP / 00)",
    "TYPE":                      "NEW / C-O — duplicates Item - CarryOver/New",
    "Primary Image":             "BM row 94 has no Stibo Attribute ID (#N/A)",
    "Item - CarryOver/New":      "not referenced by any attribute row",
    "Region Retail":             "retail pricing arrives via BY response",
    "Phaseout Date - APAC":      "not referenced by any attribute row",
    "Region Sales Sample Set":   "not referenced by any attribute row",
    "GAME PLAN":                 "not referenced by any attribute row",
    "MARKETING MONTH":           "not referenced by any attribute row",
    "APAC DTC DA Scale":         "not referenced by any attribute row",
    "DTC/GKA Exclusive":         "not referenced by any attribute row",
    "GLOBAL 52 Weeks / Kids NOS": "not referenced by any attribute row",
    "Regional Merch Comments":   "not referenced by any attribute row",
    "Lifecycle":                 "not referenced by any attribute row",
    "MAP":                       "not referenced by any attribute row",
}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES  (IDs resolved against the MDD at runtime; the hard-coded
# fallbacks below match the MDD LOV sheets — same convention as every
# NB ETL in this repository)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "NEW": "NEW BALANCE", "ADI": "ADIDAS", "NIK": "NIKE",
    "SMI": "SMIGGLE",     "ALD": "ALDO",   "CRO": "CROCS",
    "LOT": "LOTTO",       "BIR": "BIRKENSTOCK",
}

# MDD "Gender LOV" — SAP Gender (value → ID)
LOV_GENDER = {"Male": "M", "Female": "F", "Unisex": "U"}

# MDD "Age LOV" — SAP Age (value → ID)
LOV_SAP_AGE = {"Adults": "AD", "Children": "CH", "All Ages": "AA"}

# BY-side LOV IDs are the upper-cased BY values (MDD "Age LOV" col F/G and
# "Gender LOV" col G — same convention as new_balance/inline_main_accessories)
LOV_BY_AGE = {
    "AD": "ADULT", "CH": "KIDS", "AA": "ALL AGES",
    "IN": "INFANT", "JR": "KIDS",
}

# MDD "Season LOV" — code → season name
LOV_SEASON = {
    "SP": "Spring",        "SM": "Summer",       "FL": "Fall",
    "WN": "Winter",        "CO": "Core",         "SS": "Spring-Summer",
    "FW": "Fall-Winter",   "AL": "All Season",
}

LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
    "FW": "Footwear", "AP": "Apparel", "AC": "Accessories",
}

# MDD "Article Category LOV" — 1 = Generic (BM row 84 default)
LOV_SAP_ARTICLE_CATEGORY = {
    "0": "Single", "1": "Generic", "2": "Variant",
    "10": "Sell set (Hampers)", "11": "Prepack (Musical Box)",
}

# MDD "BY Article Type" — code → name
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
}

# MDD "E-com Ages Category LOV" — value → ID
LOV_ECOM_AGES = {
    "Adult":                  "A",
    "Grade School":           "G",
    "Infant":                 "I",
    "Play School / Pre School": "P",
    "Toddler":                "T",
    "Youth":                  "Y",
}

# ── Item Number 1st letter → gender / age profile ────────────────────────────
# BM rows 37/38/95/96/132: "Mapping from 1st digits column 'Item Number'".
# M = Men, W = Women, U = Unisex (adult); Y = Youth; G = Grade School;
# P = Preschool; I = Infant; N = NB kids line (grade-school sizing,
# e.g. NW574/NW327 — kids Lifestyle/Performance products).
# Fields: sap_gender_id, sap_gender_val, sap_age_id, sap_age_val,
#         by_gender, by_age, ecom_age_val
ITEM_PREFIX_PROFILE: dict[str, dict] = {
    "M": {"gender_id": "M", "gender_val": "Male",   "age_id": "AD", "age_val": "Adults",
          "by_gender": "Male",   "by_age": "Adult",        "ecom_age": "Adult"},
    "W": {"gender_id": "F", "gender_val": "Female", "age_id": "AD", "age_val": "Adults",
          "by_gender": "Female", "by_age": "Adult",        "ecom_age": "Adult"},
    "U": {"gender_id": "U", "gender_val": "Unisex", "age_id": "AD", "age_val": "Adults",
          "by_gender": "Unisex", "by_age": "Adult",        "ecom_age": "Adult"},
    "Y": {"gender_id": "U", "gender_val": "Unisex", "age_id": "CH", "age_val": "Children",
          "by_gender": "Unisex", "by_age": "Kids",         "ecom_age": "Youth"},
    "G": {"gender_id": "U", "gender_val": "Unisex", "age_id": "CH", "age_val": "Children",
          "by_gender": "Unisex", "by_age": "Grade School", "ecom_age": "Grade School"},
    "P": {"gender_id": "U", "gender_val": "Unisex", "age_id": "CH", "age_val": "Children",
          "by_gender": "Unisex", "by_age": "Preschool",    "ecom_age": "Play School / Pre School"},
    "I": {"gender_id": "U", "gender_val": "Unisex", "age_id": "CH", "age_val": "Children",
          "by_gender": "Unisex", "by_age": "Infant",       "ecom_age": "Infant"},
    "N": {"gender_id": "U", "gender_val": "Unisex", "age_id": "CH", "age_val": "Children",
          "by_gender": "Unisex", "by_age": "Kids",         "ecom_age": "Grade School"},
}

# ── Line Plan Business → Sports Category (BM row 121, verbatim) ─────────────
LOV_LINE_PLAN_TO_SPORTS_CAT: dict[str, str] = {
    "RUNNING":               "Running",
    "TRAINING":              "Fitness / Training",
    "BASEBALL AND SOFTBALL": "Other",
    "BASKETBALL":            "Basketball",
    "CRICKET":               "Other",
    "GLOBAL FOOTBALL":       "Soccer",
    "PICKLEBALL":            "Tennis / Padel",
    "TENNIS":                "Tennis / Padel",
    "GOLF":                  "Golf",
    "CORE ATHLETIC":         "Fitness / Training",
    "SKATE":                 "Skateboarding",
    "KIDS PERFORMANCE":      "Running",
    "KIDS LIFESTYLE":        "Lifestyle / Casual",
    "LIFESTYLE":             "Lifestyle / Casual",
    "OTHER FOP":             "Other",
}

# ── Sports Category → MDD "Sports Category LOV" ID ───────────────────────────
# IDs are ZERO-PADDED 2-digit strings — the format the live Stibo LOV accepts
# (same rule as asics/footwear_main.py / anta/linelist_main.py / reebok).
LOV_SPORTS_CATEGORY_ID: dict[str, str] = {
    "Badminton":                             "01",
    "Basketball":                            "02",
    "Cycling":                               "03",
    "Fitness / Training":                    "04",
    "Golf":                                  "05",
    "Lifestyle / Casual":                    "06",
    "Running":                               "07",
    "Soccer":                                "08",
    "Swimming":                              "09",
    "Tennis / Padel":                        "10",
    "Walking":                               "11",
    "Outdoor / Trail / Hiking":              "12",
    "Skateboarding":                         "13",
    "Yoga / Pilates":                        "14",
    "Martial Arts / Boxing / Combat sports": "15",
    "Other":                                 "16",
}

# ── Production Factory Name → Country of Origin ──────────────────────────────
# BM row 41: "Mapping from column 'Production Factory Name'
#             Notes : Use the middle 3-letter country code"
# The 3-letter token (VNM/IDN/CHN…) is converted to the 2-letter
# "Country Origin LOV" ID via the ISO 3166-1 alpha-3 → alpha-2 table below.
ISO_A3_TO_A2: dict[str, str] = {
    "ARE": "AE", "ARG": "AR", "AUS": "AU", "AUT": "AT", "BEL": "BE",
    "BGD": "BD", "BGR": "BG", "BRA": "BR", "BRN": "BN", "CAN": "CA",
    "CHE": "CH", "CHL": "CL", "CHN": "CN", "COL": "CO", "CZE": "CZ",
    "DEU": "DE", "DNK": "DK", "DZA": "DZ", "EGY": "EG", "ESP": "ES",
    "EST": "EE", "FIN": "FI", "FRA": "FR", "GBR": "GB", "GRC": "GR",
    "HKG": "HK", "HRV": "HR", "HUN": "HU", "IDN": "ID", "IND": "IN",
    "IRL": "IE", "ISR": "IL", "ITA": "IT", "JOR": "JO", "JPN": "JP",
    "KHM": "KH", "KOR": "KR", "LKA": "LK", "LTU": "LT", "MEX": "MX",
    "MMR": "MM", "MYS": "MY", "NLD": "NL", "NOR": "NO", "NZL": "NZ",
    "PAK": "PK", "PHL": "PH", "POL": "PL", "PRT": "PT", "ROU": "RO",
    "RUS": "RU", "SAU": "SA", "SGP": "SG", "SVK": "SK", "SVN": "SI",
    "SWE": "SE", "THA": "TH", "TUR": "TR", "TWN": "TW", "UKR": "UA",
    "USA": "US", "VNM": "VN", "ZAF": "ZA",
}

# MDD "Country Origin LOV" (ID → name) — subset covering every country the
# NB factories map to (resolved against the full MDD LOV at runtime).
LOV_COUNTRY_ORIGIN: dict[str, str] = {
    "CN": "China",     "VN": "Vietnam",  "ID": "Indonesia",
    "TH": "Thailand",  "MY": "Malaysia", "SG": "Singapore",
    "US": "USA",       "TW": "Taiwan",   "HK": "Hong Kong",
    "LK": "Sri Lanka", "KH": "Cambodia", "PT": "Portugal",
    "IN": "India",     "BD": "Bangladesh", "PK": "Pakistan",
    "IT": "Italy",     "GB": "United Kingdom", "JP": "Japan",
    "JO": "Jordan",    "AU": "Australia", "DE": "Germany",
    "BR": "Brazil",    "FR": "France",   "MX": "Mexico",
    "TR": "Turkey",    "KR": "Korea, Republic of", "CA": "Canada",
}

# ── Catalog Material Description → "Material -Upper" LOV close value ─────────
# BM row 103: "system to generate close value". The FIRST material token of
# the description is mapped to the closest MDD "Material -Upper LOV" value
# (EVA/LEA/MES/PES/PU/SUE/SUS/SYN/TEX); unknown first tokens fall through to
# the first token that does resolve.
MATERIAL_UPPER_CLOSE_MAP: dict[str, tuple] = {
    "SYNTHETIC":        ("SYN", "Synthetic"),
    "MESH":             ("MES", "Mesh"),
    "SUEDE":            ("SUE", "Suede"),
    "PIG SUEDE":        ("SUE", "Suede"),
    "SPLIT SUEDE":      ("SUE", "Suede"),
    "HAIRY SUEDE":      ("SUE", "Suede"),
    "SYNTHETIC SUEDE":  ("SUE", "Suede"),
    "SYNTHETIC NUBUCK": ("SUE", "Suede"),
    "NUBUCK":           ("SUE", "Suede"),
    "TEXTILE":          ("TEX", "Textile"),
    "ENGINEERED KNIT":  ("TEX", "Textile"),
    "LEATHER":          ("LEA", "Leather"),
    "PIGSKIN":          ("LEA", "Leather"),
    "MICROFIBER":       ("SYN", "Synthetic"),
    "NYLON":            ("SYN", "Synthetic"),
    "POLYESTER":        ("PES", "Polyester"),
    "EVA":              ("EVA", "EVA"),
    "PU":               ("PU",  "PU"),
}

# ── Adult Gender/ Kids Closure Type → Fastening LOV ──────────────────────────
# BM row 104: children closure types map to the MDD "Fastening LOV";
# SAP-Age "Adults" defaults to Lace Up when the column carries no closure.
CLOSURE_TO_FASTENING: dict[str, tuple] = {
    "LACE":                                    ("LU", "Lace Up"),
    "BUNGEE":                                  ("BL", "Bungee Lace"),
    "BUNGEE LACE WITH HOOK AND LOOP TOP STRAP": ("BL", "Bungee Lace"),
    "HOOK AND LOOP":                           ("VE", "Velcro"),
    "SLIP ON":                                 ("NL", "No Lace"),
}
# Values of the column that are gender markers, not closures:
GENDER_TOKENS = {"MENS", "WOMENS", "UNISEX", "KIDS", "BOYS", "GIRLS",
                 "GRADESCHOOL BOYS", "GRADESCHOOL GIRLS",
                 "PRESCHOOL BOYS", "PRESCHOOL GIRLS"}

DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E", "FOOTWEAR": "F",
    "APPAREL": "A", "EQUIPMENT": "Q", "TOYS": "T",
    "SHOES": "F",
}

# The GTM sheet is footwear-only, so the division is fixed by the file type
# (BM row 78: L1 default "Footwear").
DIVISION = "Footwear"

# MDD AT_PrincipalGenderDescription max chars (MDD: 10)
MAX_GENDER_DESC_CHARS = 10

# MDD AT_PrincipalColorName — no max-chars limit in the MDD (n/a).

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

# Serialise STEP-namespace elements with the DEFAULT namespace (clean,
# unprefixed tags after the _XMLNS_RE strip below) — same convention as
# the newest NB ETL modules; guarantees identical output standalone AND
# inside the unified lambda.
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        ws   = wb["Core Attributes"]
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
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[str(val_name).strip()] = (
                    str(val_id).strip() if val_id else str(val_name).strip()
                )

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

    def max_chars(self, attr_id: str, default: int = 0) -> int:
        raw = (self.attributes.get(attr_id, {}) or {}).get("max_chars") or ""
        digits = re.sub(r"[^\d]", "", str(raw).split("\n")[0]) if raw else ""
        try:
            return int(digits) if digits else default
        except ValueError:
            return default

    def lov_id(self, lov_display: str, value: str) -> str:
        """Resolve a display value against an MDD LOV (exact/title/upper)."""
        table = self.lovs.get(lov_display, {})
        if not value:
            return ""
        return (
            table.get(value)
            or table.get(value.title())
            or table.get(value.upper())
            or ""
        )


class BrandMappingLinelistLoader:
    """
    Parses the "NEW BALANCE  LINELIST" sheet (note: TWO spaces in the sheet
    name — matched fuzzily on collapsed whitespace) of the NEW Brand Mapping
    file that the router downloads into input/attributes/.

    The ETL's attribute mapping is hard-coded from that sheet (repo
    convention — same as every NB ETL), but the loader lets the module
    SELF-CHECK at runtime that every attribute ID it emits is declared in
    the live mapping sheet. Undeclared IDs are logged as warnings so a
    mapping-sheet revision surfaces immediately instead of silently
    diverging (same contract as new_balance/inline_sample_*_main.py).
    """

    SHEET_NAME_NORM = "NEW BALANCE LINELIST"

    def __init__(self, path: Path):
        self.path  = path
        self.sheet = ""
        self.declared_ids: set[str] = set()
        self.rows: list[dict] = []
        self._load()

    def _load(self):
        log.info("[BrandMapping] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        target = next(
            (s for s in wb.sheetnames
             if re.sub(r"\s+", " ", s).strip().upper() == self.SHEET_NAME_NORM),
            None,
        )
        if target is None:
            log.warning(
                "[BrandMapping] Sheet '%s' not found in %s — self-check disabled "
                "(available: %s)", self.SHEET_NAME_NORM, self.path.name,
                wb.sheetnames[:8],
            )
            wb.close()
            return
        self.sheet = target
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows[:5])
             if r and len(r) > 1 and r[1] is not None
             and str(r[1]).strip() == "Stibo Attribute ID"),
            1,
        )
        for row in rows[hdr_idx + 1:]:
            if not row or len(row) < 2:
                continue
            attr_id = row[1]
            if attr_id is None:
                continue
            attr_id = str(attr_id).strip()
            if not attr_id or attr_id.startswith("#"):
                continue
            self.declared_ids.add(attr_id)
            self.rows.append({
                "attribute":   str(row[0]).strip() if row[0] else "",
                "id":          attr_id,
                "validation":  str(row[2]).strip() if len(row) > 2 and row[2] else "",
            })
        wb.close()
        log.info("[BrandMapping] %d attribute IDs declared in '%s'",
                 len(self.declared_ids), self.sheet)


class GTMFootwearLineListLoader:
    """
    Loads the New Balance APAC Footwear GTM Line Sheet.

    Sheet selection (SHEET_MODE):
      * "first_visible"  → the first visible sheet ("S2'27 GTM 2 Footwear
                           Line list") — operator instruction; skips the
                           hidden "APAC CO Request" tab.
      * name fallback    → "APAC" / "Line Sheet" / "Line List" (older
                           exports whose tab kept the old name).

    The header row is located dynamically by scanning for "Item Number"
    (the sheet carries a merged title row and blank rows above the header).
    """

    ROW_KEY = "Item Number"

    def __init__(self, path: Path, sheet_mode: str = SHEET_MODE):
        self.path  = path
        self.mode  = sheet_mode
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[GTM-FW] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        if self.mode == "first_visible":
            target = next(
                (s for s in wb.sheetnames if wb[s].sheet_state == "visible"), None
            )
        else:
            target = next(
                (s for s in SHEET_NAME_FALLBACKS if s in wb.sheetnames), None
            )
        if target is None:
            target = max(wb.sheetnames, key=lambda s: wb[s].max_row or 0)
        log.info("[GTM-FW] Using sheet: '%s' (mode=%s)", target, self.mode)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(str(v).strip() == self.ROW_KEY for v in r if v)),
            None,
        )
        if hdr_idx is None:
            hdr_idx = next(
                (i for i, r in enumerate(rows)
                 if sum(1 for v in r[:15] if isinstance(v, str) and v.strip()) >= 8),
                None,
            )
        if hdr_idx is None:
            log.error("[GTM-FW] Cannot find header row in '%s'", target)
            wb.close()
            return

        log.info("[GTM-FW] Header at row index %d (row %d)", hdr_idx, hdr_idx + 1)
        header = [
            str(h).replace("\n", " ").strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
        # True 1-based Excel row number for each data row, captured BEFORE any
        # filtering below — lets callers target a specific spreadsheet row
        # reliably even if blank rows get dropped afterwards.
        df["_ExcelRow"] = range(hdr_idx + 2, hdr_idx + 2 + len(df))

        item_col = next(
            (c for c in [self.ROW_KEY, "Item No.", "Item No"] if c in df.columns), None
        )
        if item_col:
            df = df[
                df[item_col].notna()
                & (df[item_col] != "")
                & (df[item_col] != 0)
            ]

        self.df    = df.reset_index(drop=True)
        self.sheet = target
        wb.close()

        missing = [c for c in SOURCE_COLUMNS if c not in self.df.columns]
        if missing:
            log.warning("[GTM-FW] Expected source column(s) absent: %s", missing)

        log.info("[GTM-FW] %d SKU rows loaded from '%s'", len(self.df), self.sheet)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "N/A", "#N/A", "#VALUE!") else s


def _norm_ws(v) -> str:
    """Collapse newlines/tabs in a multi-line cell to single spaces."""
    return re.sub(r"[\r\n\t]+", " ", _s(v)).strip()


def _fmt_date_ddmmyyyy(v) -> str:
    """AT_IncomingMonth — DD-MM-YYYY (repo-wide convention). Accepts a
    datetime/date cell or 'MM/DD/YYYY' / 'YYYY-MM-DD' strings."""
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
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return ""


def _item_prefix_profile(item_number: str, gender_col: str) -> tuple[dict, str]:
    """
    Item Number 1st letter → gender/age profile (BM rows 37/38/95/96/132).
    Fallback chain for an unrecognised prefix: the Adult Gender/ Kids
    Closure Type column (Mens/Womens/Unisex), then Unisex/Adult. Returns
    (profile, note) — note is "" when the prefix matched directly.
    """
    first = (item_number or "")[:1].upper()
    if first in ITEM_PREFIX_PROFILE:
        return ITEM_PREFIX_PROFILE[first], ""

    gc = (gender_col or "").strip().upper()
    if gc.startswith("WOMEN"):
        return ITEM_PREFIX_PROFILE["W"], f"unknown prefix '{first}' → Womens column"
    if gc.startswith("MEN") or gc.startswith("BOY"):
        return ITEM_PREFIX_PROFILE["M"], f"unknown prefix '{first}' → Mens column"
    return ITEM_PREFIX_PROFILE["U"], f"unknown prefix '{first}' → Unisex default"


def _line_plan_to_sports_cat(line_plan: str) -> str:
    """BM row 141 mapping. Unmapped → "" (omitted, warned — never guessed)."""
    return LOV_LINE_PLAN_TO_SPORTS_CAT.get(_s(line_plan).upper().strip(), "")


def _sports_cat_to_lov_id(sports_cat: str) -> str:
    """Sports Category display name → MDD LOV ID (zero-padded)."""
    return LOV_SPORTS_CATEGORY_ID.get(sports_cat.strip(), "")


def _factory_to_country_origin(factory: str) -> tuple[str, str]:
    """
    "PYV - VNM - P00" → middle 3-letter code (VNM) → 2-letter Country Origin
    LOV (VN/Vietnam). Returns ("", "") when unresolvable.
    """
    tokens = [t for t in re.split(r"[\s\-–—|,;]+", _s(factory)) if t]
    a3 = ""
    for t in tokens:
        if len(t) == 3 and t.isalpha() and t.upper() in ISO_A3_TO_A2:
            a3 = t.upper()
            break
    if not a3:
        return "", ""
    a2 = ISO_A3_TO_A2[a3]
    return a2, LOV_COUNTRY_ORIGIN.get(a2, a2)


def _material_upper_close(catalog_material: str) -> tuple[str, str]:
    """
    Catalog Material Description → closest MDD "Material -Upper LOV" value
    (BM row 103: "system to generate close value"). Tries the first token,
    then falls through the remaining tokens. Returns ("", "") when no token
    resolves — the attribute is then omitted (never a raw-text guess).
    """
    tokens = [t.strip().upper() for t in re.split(r"[/,]", _norm_ws(catalog_material)) if t.strip()]
    for tok in tokens:
        # Fold multi-word tokens ("PIG SUEDE", "SPLIT SUEDE", …)
        for candidate in (tok, " ".join(tok.split())):
            if candidate in MATERIAL_UPPER_CLOSE_MAP:
                return MATERIAL_UPPER_CLOSE_MAP[candidate]
        # Exact multi-word miss → try the first word alone (e.g. "PIGSKIN/
        # leather" splits oddly, "Hairy Suede/Mesh" first token "HAIRY SUEDE")
        first_word = tok.split()[0] if tok.split() else ""
        if first_word in MATERIAL_UPPER_CLOSE_MAP:
            return MATERIAL_UPPER_CLOSE_MAP[first_word]
    return "", ""


def _closure_to_fastening(closure_col: str, sap_age_id: str) -> tuple[str, str]:
    """
    BM row 104:
      • closure value in the column → mapped Fastening (any SAP age —
        "For certain item with SAP Age 'Adult' will be follow the mapping")
      • no closure + SAP Age Adults  → default Lace Up
      • no closure + children        → ("", "") — omitted
    """
    raw = _s(closure_col).upper().strip()
    if raw in CLOSURE_TO_FASTENING:
        return CLOSURE_TO_FASTENING[raw]
    if raw in GENDER_TOKENS or not raw:
        if sap_age_id == "AD":
            return ("LU", "Lace Up")
        return ("", "")
    return ("", "")


def _colorway_name_clean(colorway: str, item_number: str) -> str:
    """
    BM row 16 note: "for certain items include the 'Item Number' in Colorway
    Name, so it is required to be removed" — e.g.
    "M10802FR - BLACK / SLATE GREY" → "BLACK / SLATE GREY".
    """
    val = _s(colorway)
    item = _s(item_number)
    if val and item:
        pattern = r"^\s*" + re.escape(item) + r"\s*[-–—:]\s*"
        val = re.sub(pattern, "", val, flags=re.IGNORECASE).strip()
    return val


def _build_generic_code(brand_code: str, item_number: str) -> str:
    """
    KEY_InboundArticle / AT_InboundGenericCode — repo convention (every NB
    ETL keys the generic article 1:1 with a source row): brand(3) + the
    row-unique Item Number. Keying on Product Number would collapse the
    2784 rows into 470 articles (see DEVIATION 1 in the module docstring).
    """
    brand = re.sub(r"[^A-Z0-9]", "", brand_code.upper())[:3]
    base  = re.sub(r"[^A-Z0-9]", "", item_number.upper())
    return f"{brand}{base}"


def map_sku(row: pd.Series, brand_code: str = "NEW") -> dict:
    item_number    = _s(row.get("Item Number"))
    product_number = _s(row.get("Product Number"))
    short_desc     = _s(row.get("Short Product Description"))
    colorway       = _s(row.get("Colorway Name"))
    gender_raw     = _s(row.get("Adult Gender/ Kids Closure Type", ""))
    size_raw       = _norm_ws(row.get("SIZE"))
    line_plan_biz  = _s(row.get("Line Plan Business", ""))
    category       = _s(row.get("CATEGORY", ""))
    factory        = _s(row.get("Production Factory Name", ""))
    catalog_mat    = _s(row.get("Catalog Material Description", ""))
    series_no      = _s(row.get("Series#", ""))
    intro_date     = row.get("Intro Date - APAC", None)
    excel_row      = row.get("_ExcelRow")

    profile, prefix_note = _item_prefix_profile(item_number, gender_raw)
    sports_cat = _line_plan_to_sports_cat(line_plan_biz)
    fastening_id, fastening_val = _closure_to_fastening(gender_raw, profile["age_id"])
    mu_id, mu_val = _material_upper_close(catalog_mat)
    coo_id, coo_val = _factory_to_country_origin(factory)

    return {
        # ── source values ───────────────────────────────────────
        "item_number":     item_number,
        "product_number":  product_number,
        "short_desc":      short_desc,
        "colorway_raw":    colorway,
        "colorway":        _colorway_name_clean(colorway, item_number),
        "gender_raw":      gender_raw,
        "size_raw":        size_raw,
        "line_plan_biz":   line_plan_biz,
        "category":        category,
        "factory":         factory,
        "catalog_mat":     catalog_mat,
        "series_no":       series_no,
        "intro_date_raw":  intro_date,
        "excel_row":       int(excel_row) if excel_row is not None else None,
        # ── derived (Item Number 1st letter) ─────────────────────
        "prefix_note":     prefix_note,
        "gender_id":       profile["gender_id"],
        "gender_val":      profile["gender_val"],
        "age_id":          profile["age_id"],
        "age_val":         profile["age_val"],
        "by_gender":       profile["by_gender"],
        "by_age":          profile["by_age"],
        "by_age_id":       LOV_BY_AGE.get(profile["age_id"], profile["by_age"].upper()),
        "ecom_age":        profile["ecom_age"],
        "ecom_age_id":     LOV_ECOM_AGES.get(profile["ecom_age"], ""),
        # ── derived (mapping tables) ─────────────────────────────
        "sports_cat":      sports_cat,
        "sports_cat_id":   _sports_cat_to_lov_id(sports_cat),
        "fastening_id":    fastening_id,
        "fastening_val":   fastening_val,
        "material_upper_id": mu_id,
        "material_upper_val": mu_val,
        "coo_id":          coo_id,
        "coo_val":         coo_val,
        "incoming_month":  _fmt_date_ddmmyyyy(intro_date),
        # ── codes / context ──────────────────────────────────────
        "brand_code":      brand_code,
        "generic_code":    _build_generic_code(brand_code, item_number),
        "division":        DIVISION,
        "channel_type":    "Inline",
        "article_type":    "Inline",
        "art_category":    "1",
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """
    Validates only the fields this ETL actually emits. Attributes dropped
    for lack of a source column are reported once per run by
    _log_dropped_attributes(), not once per SKU.
    """
    warns = []
    sku   = mapped["item_number"] or "<no item number>"

    field_to_at = {
        "item_number":    "AT_PrincipalStyleCode",   # row key (Mandatory in MDD)
        "gender_id":      "AT_Gender",
        "age_id":         "AT_SAPAge",
        "brand_code":     "AT_Brand",
        "sports_cat":     "AT_SportsCategoryEN",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {}) if mdd else {}
        card = (meta.get("cardinality") or "").lower() if meta else ""
        val  = mapped.get(field, "")
        if "mandatory" in card and (not val or str(val).strip() in ("", "None", "nan")):
            warns.append(f"[{sku}] MISSING mandatory: {at_id}")

    if not mapped.get("product_number"):
        warns.append(f"[{sku}] MISSING source value: Product Number (AT_PrincipalStyleCode omitted)")
    if not mapped.get("line_plan_biz"):
        warns.append(f"[{sku}] MISSING source value: Line Plan Business")
    if not mapped.get("sports_cat") and mapped.get("line_plan_biz"):
        warns.append(f"[{sku}] UNMAPPED Line Plan Business '{mapped['line_plan_biz']}' — "
                     "AT_SportsCategoryEN omitted")
    if mapped.get("prefix_note"):
        warns.append(f"[{sku}] {mapped['prefix_note']}")
    if not mapped.get("incoming_month") and _s(mapped.get("intro_date_raw")):
        warns.append(f"[{sku}] UNPARSEABLE Intro Date - APAC '{mapped['intro_date_raw']}' — "
                     "AT_IncomingMonth omitted")
    if not mapped.get("coo_id") and mapped.get("factory"):
        warns.append(f"[{sku}] Country of origin unresolved from factory '{mapped['factory']}' — "
                     "AT_CountryOrigin omitted")
    if not mapped.get("material_upper_id") and mapped.get("catalog_mat"):
        warns.append(f"[{sku}] Material -Upper close value unresolved from "
                     "'{mapped['catalog_mat']}' — AT_MaterialUpper omitted")

    return warns


def _log_dropped_attributes(mdd: MDDLoader) -> list[str]:
    """Emit a one-time, run-level report of every attribute not carried over."""
    notes: list[str] = []
    notes.append(
        "Attributes the spec pins whose source column is absent from the GTM sheet "
        "(deliberately NOT emitted):"
    )
    dropped = {
        "AT_LaunchingDate":  "BM row 166 wants 'SEA Intro Date'; GTM sheet only has "
                             "'Intro Date - APAC' (→ AT_IncomingMonth)",
        "AT_FOB":            "BM row 60: Footwear = N/A",
        "AT_FOBCurrency":    "no FOB for footwear — currency not emitted either",
        "AT_BCI":            "BM row 97: Footwear = Manual Input",
        "AT_Silhouette":     "BM row 101: Footwear = Manual Input",
        "AT_TechnologyUsed": "BM row 158: FW = Manual MD input",
        "AT_PrincipalColorCode": "BM row 15: Footwear = N/A",
        "AT_PrincipalGenderCode": "BM row 17: Not Available",
        "AT_PrincipalAgeCode":    "BM row 19: Not Available",
        "AT_PrincipalAgeDescription": "BM row 20: Not Available",
        "AT_PrincipalMerchandiseHierarchyL4": "BM row 81: Footwear = N/A",
        "AT_PrincipalMerchandiseHierarchyL5": "BM row 82: Footwear = N/A",
        "AT_SAPStyleCode / AT_Generic / AT_GenericDescription / AT_Variant / "
        "AT_VariantDescription":
            "BM rows 22-26 'Formula in System' — computed downstream from the BY "
            "SAP colour/size response; never guessed here",
    }
    for at_id, why in dropped.items():
        card = (mdd.attributes.get(at_id, {}).get("cardinality") or "n/a").strip() if mdd else "n/a"
        notes.append(f"  {at_id:44} cardinality={card:12} {why}")
        if "mandatory" in card.lower():
            log.warning("MANDATORY attribute %s not emitted — %s", at_id, why)

    notes.append("")
    notes.append("Source columns present in the sheet but mapped to no attribute row:")
    for col, why in UNMAPPED_COLUMNS.items():
        notes.append(f"  {col:32} {why}")
    for line in notes:
        log.info(line)
    return notes


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER HELPERS
# ══════════════════════════════════════════════════════════════════

def _clean_xml_value(raw) -> str:
    text = "" if raw is None else str(raw).strip()
    return "" if text in ("", "None", "nan", "NaT", "#N/A", "#VALUE!") else text


def _val(parent, attr_id, value="", id_val=""):
    normalized_id    = _clean_xml_value(id_val)
    normalized_value = _clean_xml_value(value)
    if not normalized_id and not normalized_value:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if normalized_id:
        el.set("ID", normalized_id)
        if normalized_value:
            el.text = normalized_value
        return el
    el.text = normalized_value
    return el


def _multival(parent, attr_id, id_val, label=""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)
    return mv


def _lov(code, lookup, default=""):
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def _get_division_code(division: str) -> str:
    div_upper = (division or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in div_upper:
            return code
    return "F"


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC VALUE WRITER
# ══════════════════════════════════════════════════════════════════

# Attribute IDs deliberately exempt from the BrandMapping self-check:
# the sheet row exists but its ID cell is #N/A in this revision — the
# attribute itself is declared in the MDD (Mandatory, LOV "BY Article
# Type") and in the v6 sheet (R105 "Article Type", FW default Inline).
SHEET_ID_GAPS: set[str] = {"AT_BYArticleType"}

# Every AT_ id this module writes (KEY_InboundArticle / <Name> are not
# attributes). Used for the BrandMapping self-check in run().
EMITTED_ATTRIBUTE_IDS: list[str] = [
    "AT_Country", "AT_CompanyCode", "AT_SBU", "AT_Brand", "AT_BrandGroup",
    "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
    "AT_PrincipalColorName", "AT_PrincipalGenderDescription", "AT_PrincipalSize",
    "AT_Gender", "AT_SAPAge", "AT_BYGender", "AT_BYAge", "AT_EComAgesCategory",
    "AT_CountryOrigin", "AT_CountrySize",
    "AT_Season", "AT_SeasonYear",
    "AT_PrincipalMerchandiseHierarchyL1", "AT_PrincipalMerchandiseHierarchyL2",
    "AT_PrincipalMerchandiseHierarchyL3",
    "AT_SportsCategoryEN", "AT_Collection1", "AT_Collection2",
    "AT_MaterialUpper", "AT_Fastening", "AT_IncomingMonth",
    "AT_MaterialType", "AT_SAPArticleCategory", "AT_BYArticleType",
    "AT_SAPProductFlag", "AT_UOM", "AT_ArticleStatus",
]


def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None, bm=None):
    """
    Writes exactly the attributes the NEW BALANCE  LINELIST mapping sheet
    defines as populated for the Footwear GTM line-list flow (see
    EMITTED_ATTRIBUTE_IDS). Every other attribute row in that sheet is
    N/A / Manual input / Formula-in-System and is deliberately NOT emitted
    (see _log_dropped_attributes).
    """
    # ── Context (from filename metadata, not a source column) ─────
    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand",      b_label, id_val=b_code)
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)

    # ── Principal style (BM rows 13/14) ───────────────────────────
    _val(vals_el, "AT_PrincipalStyleCode", art["product_number"])
    # v6 R027-style fallback: blank description → Item Number
    _val(vals_el, "AT_PrincipalStyleDescription",
         art["short_desc"] or art["item_number"])

    # ── Colorway Name minus the Item Number prefix (BM row 16) ────
    _val(vals_el, "AT_PrincipalColorName", art["colorway"])

    # ── Adult Gender/ Kids Closure Type (BM row 18) — plain text ──
    gender_max = MAX_GENDER_DESC_CHARS
    if mdd is not None:
        gender_max = mdd.max_chars("AT_PrincipalGenderDescription", MAX_GENDER_DESC_CHARS)
    _val(vals_el, "AT_PrincipalGenderDescription", art["gender_raw"][:gender_max])

    # ── SIZE → Principal Size (BM row 21) ─────────────────────────
    _val(vals_el, "AT_PrincipalSize", art["size_raw"])

    # ── Item Number 1st letter → SAP/BY gender & age (BM rows 37/38/95/96)
    _val(vals_el, "AT_Gender",   art["gender_val"], id_val=art["gender_id"])
    _val(vals_el, "AT_SAPAge",   art["age_val"],    id_val=art["age_id"])
    _val(vals_el, "AT_BYGender", art["by_gender"],  id_val=art["by_gender"].upper())
    _val(vals_el, "AT_BYAge",    art["by_age"],     id_val=art["by_age_id"])
    _val(vals_el, "AT_EComAgesCategory", art["ecom_age"], id_val=art["ecom_age_id"])

    # ── Production Factory Name → Country of Origin (BM row 41) ───
    if art["coo_id"]:
        _val(vals_el, "AT_CountryOrigin", art["coo_val"], id_val=art["coo_id"])

    # ── Season (from filename metadata) ───────────────────────────
    sea_raw = art.get("season", "") or ""
    sea_m = re.match(r'^([A-Z]{2})(\d{2,4})$', sea_raw.strip().upper())
    if sea_m:
        sea_prefix = sea_m.group(1)
        sea_digits = sea_m.group(2)
        sea_year_4 = sea_digits if len(sea_digits) == 4 else f"20{sea_digits}"
    else:
        sea_prefix = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
        sea_year_4 = ""

    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_prefix, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_prefix)
    _val(vals_el, "AT_SeasonYear", sea_year_4)

    # ── Merchandise hierarchy (BM rows 78/79/80) ──────────────────
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", "Footwear")
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art.get("category", ""))

    # ── Line Plan Business → Sports Category (BM row 121, LOV ID) ─
    if art.get("sports_cat_id"):
        _val(vals_el, "AT_SportsCategoryEN", art["sports_cat"],
             id_val=art["sports_cat_id"])

    # ── Collections (BM rows 122/123) ─────────────────────────────
    _val(vals_el, "AT_Collection1", art.get("catalog_mat", ""))
    _val(vals_el, "AT_Collection2", art.get("series_no", ""))

    # ── Material - Upper (BM row 103 — close value, LOV only) ─────
    if art.get("material_upper_id"):
        _val(vals_el, "AT_MaterialUpper", art["material_upper_val"],
             id_val=art["material_upper_id"])

    # ── Fastening (BM row 104) ────────────────────────────────────
    if art.get("fastening_id"):
        _val(vals_el, "AT_Fastening", art["fastening_val"], id_val=art["fastening_id"])

    # ── Intro Date - APAC → Incoming month (BM row 233) ───────────
    _val(vals_el, "AT_IncomingMonth", art.get("incoming_month", ""))

    # ── Spec defaults (BM rows 83/84/93/77/90/130/232) ────────────
    _val(vals_el, "AT_MaterialType", id_val="ZINA")            # Intercompany

    cat_code  = art["art_category"]                             # "1"
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, "Generic")
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)

    at_norm  = art["article_type"]                              # "Inline"
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)

    _val(vals_el, "AT_SAPProductFlag", "Intercompany", id_val="A")
    _val(vals_el, "AT_UOM", "Each", id_val="EA")

    # MDD "Country Size LOV": US → US (BM row 130: All US Size default).
    country_size_id = "US"
    if mdd is not None:
        cs_lov = mdd.lovs.get("Country Size", {})
        if cs_lov and "US" not in cs_lov.values():
            country_size_id = ""   # MDD revision removed US — omit
    if country_size_id:
        _val(vals_el, "AT_CountrySize", "US", id_val="US")

    _val(vals_el, "AT_ArticleStatus", "Active")

    # ── Derived pipeline key (repo convention) ────────────────────
    _val(vals_el, "AT_InboundGenericCode", art["generic_code"])

    country_val = art.get("country_code", "")
    if country_val:
        _val(vals_el, "AT_Country", country_val, id_val=country_val)


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS
# ══════════════════════════════════════════════════════════════════

def build_classifications(brand, brand_code, season_code):
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent   = f"CLH_{brand.replace(' ', '')}Batches"
    season_label_map = {
        # MDD "Season LOV" names
        "SP": "Spring",        "SM": "Summer",       "FL": "Fall",
        "WN": "Winter",        "CO": "Core",         "SS": "Spring-Summer",
        "FW": "Fall-Winter",   "AL": "All Season",   "AW": "Autumn-Winter",
        "HO": "Holiday",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

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


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — PRODUCT XML BUILDER
# ══════════════════════════════════════════════════════════════════

def build_product_xml(art, brand, brand_code, comp_code, sbu, season_id, mdd=None, bm=None):
    item_number = art["item_number"]
    if not item_number:
        return ""

    div_letter  = _get_division_code(art.get("division", DIVISION))
    parent_id   = f"PPH_{div_letter}-TempSubCat"
    key_generic = art["generic_code"]

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    # <Name> = Principal Style Description with the same v6 fallback chain
    # (blank description → Product Number → Item Number).
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["short_desc"] or art["product_number"] or item_number
    )

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd, bm=bm)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_sku(task):
    row, brand_code, mdd = task
    mapped = map_sku(row, brand_code=brand_code)
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Entry point called by new_balance/lambda_function.py (or CLI).
    args must have: brand, brand_code, comp_code, sbu, season, seq
    """
    def first(d: Path):
        files = [f for pat in ("*.xlsx", "*.xlsm") for f in d.glob(pat)]
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ll_files = [f for pat in ("*.xlsx", "*.xlsm") for f in LINELIST_DIR.glob(pat)]

    for label, val in [
        ("MDD",           mdd_f),
        ("Attributes",    attr_f),
        ("GTM Line List", ll_files),
    ]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    mdd = MDDLoader(mdd_f)
    bm  = BrandMappingLinelistLoader(attr_f)

    # Self-check: every ID this module emits must be declared in the live
    # "NEW BALANCE  LINELIST" sheet (see BrandMappingLinelistLoader).
    # SHEET_ID_GAPS lists rows whose ID cell is #N/A in the current sheet
    # revision but whose attribute is confirmed by the MDD / v6 sheet.
    if bm.declared_ids:
        undeclared = [
            a for a in EMITTED_ATTRIBUTE_IDS
            if a not in bm.declared_ids and a not in SHEET_ID_GAPS
        ]
        if undeclared:
            log.warning(
                "[BrandMapping SELF-CHECK] %d emitted attribute ID(s) not declared "
                "in sheet '%s': %s — mapping sheet may have been revised.",
                len(undeclared), bm.sheet, undeclared,
            )
        else:
            log.info("[BrandMapping SELF-CHECK] all %d emitted attribute IDs are "
                     "declared in '%s'.", len(EMITTED_ATTRIBUTE_IDS), bm.sheet)

    dropped_notes = _log_dropped_attributes(mdd)

    all_warnings: list[str] = []

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    for ll_path in ll_files:
        log.info("─── Processing GTM Footwear Line List: %s ───", ll_path.name)

        ll = GTMFootwearLineListLoader(ll_path)
        if ll.df.empty:
            log.warning("[GTM-FW] Empty dataframe — skipping.")
            continue

        item_col = next(
            (c for c in ["Item Number", "Item No.", "Item No"] if c in ll.df.columns), None
        )
        rows = []
        for _, row in ll.df.iterrows():
            item = str(row.get(item_col or "Item Number", "")).strip()
            if item and item not in ("None", "nan", ""):
                rows.append(row)

        total_rows = len(rows)
        log.info("[GTM-FW] %d valid SKU rows to process", total_rows)

        out_name = f"{ll_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        log.info("Pass 1/2 — parallel map+validate (%d rows) …", total_rows)
        num_workers = min(8, max(1, total_rows))
        task_args   = [(row, args.brand_code, mdd) for row in rows]
        ordered: list[tuple[int, dict, list]] = []

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_sku, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx, (mapped, warns) = futures[fut], fut.result()
                mapped["season"]       = args.season
                mapped["country_code"] = getattr(args, "country_code", "")
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_skus = [m for _, m, _ in ordered]
        del ordered

        # TEST LIMIT: uncomment to process only specific input-Excel rows.
        # Filters on the true 1-based Excel row number (tracked via
        # "_ExcelRow" in GTMFootwearLineListLoader / "excel_row" in map_sku),
        # NOT list position — safe even if blank rows exist elsewhere in the
        # sheet, since positional slicing would shift after row-dropping.
        # TEST_ROWS = {833, 844}
        # mapped_skus = [m for m in mapped_skus if m.get("excel_row") in TEST_ROWS]
        # log.info(
        #     "[GTM-FW][TEST LIMIT] Restricted to Excel rows %s → %d SKU(s) matched",
        #     sorted(TEST_ROWS), len(mapped_skus),
        # )

        # Guard the one deviation: KEY_InboundArticle must stay 1:1 with rows.
        keys = [m["generic_code"] for m in mapped_skus if m.get("item_number")]
        dupes = len(keys) - len(set(keys))
        if dupes:
            log.warning(
                "[GTM-FW] %d duplicate KEY_InboundArticle value(s) — "
                "articles will be merged in STEP.", dupes,
            )
            all_warnings.append(
                f"[FILE {ll_path.name}] {dupes} duplicate KEY_InboundArticle values"
            )

        log.info("Pass 2/2 — streaming XML to %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cls_el  = build_classifications(args.brand, args.brand_code, args.season)
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
            for art in mapped_skus:
                if not art.get("item_number"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, args.brand_code,
                    args.comp_code, args.sbu, season_id, mdd=mdd, bm=bm,
                )
                product_xml = _XMLNS_RE.sub("", product_xml)
                f.write(f"    {product_xml}\n")
                del product_xml
                written_count += 1
            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        del mapped_skus

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%dKB)", out_path, file_kb)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

        print("═══ SKU SUMMARY ════════════════════════════════════", flush=True)
        print(f"  Line List rows      : {total_rows}",    flush=True)
        print(f"  Mapped SKUs         : {written_count}", flush=True)
        print(f"  XML file size       : {file_kb}KB",     flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    rpt_path = LOG_DIR / f"validation_nb_footwear_gtm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
        f.write("\n\n")
        f.write("\n".join(dropped_notes))
        f.write("\n\nEmitted attribute IDs:\n")
        f.write("\n".join(f"  {a}" for a in EMITTED_ATTRIBUTE_IDS))

    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All SKUs passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 10 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — New Balance Footwear GTM Line Sheet v2.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="FW")
    p.add_argument("--season",     default="SS27")
    p.add_argument("--seq",        default=1, type=int)
    p.add_argument("--country-code", default="")
    run(p.parse_args())


if __name__ == "__main__":
    main()
