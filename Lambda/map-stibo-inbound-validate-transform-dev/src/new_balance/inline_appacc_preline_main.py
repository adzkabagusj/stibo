"""
╔══════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — NEW BALANCE (App/Acc Preline) v2.0 ║
║  Line List (App Acc Preline - Updated) → Stibo STEP XML          ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : S2'27 APP ACC Preline Linelist - SEA
               ("S227 APP ACC Preline Linelist - SEA.xlsm")
Primary tab  : FIRST VISIBLE sheet — "APPACC"
               (marker rows "ACC"/"APP" at index 1-2, header row index 3,
                data from index 4; the header row is located dynamically by
                scanning for "Item Number" — same strategy as every NB ETL)

MAPPING AUTHORITY (v2.0 rewrite — every column cross-checked):
  1. "NEW - Brand mapping files Template.xlsx"  → sheet "NEW BALANCE  LINELIST"
     (note: the sheet name carries TWO spaces) — the operative spec for this
     file type: the router fetches THIS file as the "attributes" input for
     linelist_appacc_preline (see lambda_function._NB_NEW_MAPPING_TRIGGER_TYPES).
  2. "Attributes List_complete_v6_25032026.xlsx" → sheet "NEW BALANCE - LL"
     (v6, 25-03-2026) — cross-check; agrees with (1) on every column below.
  3. "Master Data Dictionary (MAA).xlsx"         → Core Attributes + LOV sheets
     (cardinality / LOV IDs / max chars used at runtime for self-checks).

────────────────────────────────────────────────────────────────────
COLUMN → ATTRIBUTE MAP (v2.0 — every column of the S2'27 Preline sheet)
────────────────────────────────────────────────────────────────────
  DIRECT (BM sheet rows 13/14/15/16/18/21/79/80/81/82/60/97/101/106/107/
  109/110/111/123/158/166/233 — "Direct from Principal" / "Mapping from
  Principal", APP/ACC branch):
    Product Number            → AT_PrincipalStyleCode           [BM row 13]
    Product Display Name      → AT_PrincipalStyleDescription    [BM row 14]
                               → <Name> (blank → Item Number fallback)
    Item Number               → AT_PrincipalColorCode           [BM row 15:
                                 "Item number (digits after '_')" —
                                 MB6180TA_AAX → "AAX"]
                               → KEY_InboundArticle / AT_InboundGenericCode
                                 (repo convention: brand(3) + Item Number)
    Color Code - Non-Footwear → AT_PrincipalColorName           [BM row 16]
                                 (e.g. "BLACK (BK)")
    Gender Duty - Non-Footwear→ AT_PrincipalGenderDescription   [BM row 18]
                               → AT_Gender / AT_BYGender         [BM rows 38/96]
                                 (Mens→Male/M, Womens→Female/F, Unisex→U)
    Global Size Offering      → AT_PrincipalSize                [BM row 21]
    Line Plan Business        → AT_PrincipalMerchandiseHierarchyL2 [BM row 79]
                               → AT_SportsCategoryEN (LOV ID)    [BM row 121]
                               → AT_SAPAge / AT_BYAge            [BM rows 37/95:
                                 "Mapping from column 'Line Plan Bussiness'"
                                 — Kids rows → Children/Kids, else Adults]
    Product Line              → AT_PrincipalMerchandiseHierarchyL1 [BM row 78]
    Silhouette                → AT_PrincipalMerchandiseHierarchyL3 [BM row 80]
                               → AT_Silhouette (LOV — ID resolved from the
                                 MDD "Silhouette LOV" at runtime) [BM row 101]
    GBU                       → AT_PrincipalMerchandiseHierarchyL4 [BM row 81]
    Detailed Silhouette       → AT_PrincipalMerchandiseHierarchyL5 [BM row 82]
                               → AT_Style (LOV — matched against the MDD
                                 "Style LOV" values at runtime)    [BM row 110]
    Price Segment             → AT_BCI                          [BM row 97]
    Product Fit               → AT_Fit (LOV close value)        [BM row 109]
    Target Principal Material → AT_Content / AT_Fabric / AT_Material (LOV
                                 close values)                   [BM rows 106/107/111]
    Country Of Origin         → AT_CountryOrigin                [BM row 41]
                                 ("VNM - Viet Nam" → VNM → VN/Vietnam)
    Technologies              → AT_TechnologyUsed               [BM row 158]
                                 ("N/A"/"0" rows skipped)
    Product Collection        → AT_Collection2                  [BM row 123]
    SEA Intro Date            → AT_LaunchingDate (DD-Mon-YYYY)  [BM row 166]
                               → AT_IncomingMonth (DD-MM-YYYY)   [BM row 233]
    Estimated GEP             → AT_FOB                          [BM row 60]
                               → AT_FOBCurrency = USD            [BM row 61]

  DERIVED (BM row 122):
    AT_Collection1 = "Product Line - Silhouette - Detailed Silhouette"
    (the three columns combined with a dash separator)

  CONTEXT (filename/portal metadata — not a source column):
    AT_Country, AT_CompanyCode, AT_SBU, AT_Brand, AT_BrandGroup,
    AT_Season, AT_SeasonYear

  FILE-TYPE / SPEC DEFAULTS (BM rows 83/84/93/77/90/232):
    AT_MaterialType       = ZINA (Intercompany Articles) [BM row 83]
    AT_SAPArticleCategory = 1 / Generic                  [BM row 84]
    AT_BYArticleType      = Inline (this is an *Inline* preline linelist —
                            same constant every NB Inline ETL emits) [BM row 93]
    AT_SAPProductFlag     = A / Intercompany             [BM row 77]
    AT_UOM                = EA / Each                    [BM row 90]
    AT_ArticleStatus      = Active                       [BM row 232]

  PIPELINE KEY (repo convention, every NB ETL):
    KEY_InboundArticle / AT_InboundGenericCode = brand(3) + Item Number
    (row-unique: 1627 distinct Item Numbers in the S2'27 sheet).

────────────────────────────────────────────────────────────────────
COLUMNS READ BUT *NOT* MAPPED TO STIBO
────────────────────────────────────────────────────────────────────
  None of the columns below appear as a "Field Name in the Brand File"
  for any attribute row of the two mapping sheets (APP/ACC branch):

    Thumbnail                     (BM row 94 "Thumbnail Image" exists but its
                                  Stibo Attribute ID is #N/A — no STEP
                                  attribute is configured yet; images flow
                                  through the Brand Dashboard / BY pipeline)
    Item - CarryOver/New          (not referenced by any attribute row)
    Date Canceled                 (not referenced)
    APAC DA/DROP, SMS, EMERGING   (not referenced)
    Category, LPA Category        (not referenced)
    Game Plan                     (not referenced)
    Global Brand Event, Mandatory Global Brand Event (not referenced)
    Global Item Intro Date        (Launching Date uses SEA Intro Date)
    Intro Date - APAC             (Incoming month uses SEA Intro Date per
                                   BM row 233 APP/ACC branch)
    Global Item Phaseout Date, Phaseout Date - APAC (not referenced)
    Global Wholesale Price        (not referenced)
    Region Sales Sample Set - APAC / Region Sales Sample Set (not referenced)
    S127 Price List               (not referenced)
    SGD w/o Tax / SGD / MYRR / THB / PHP / IDR / VND (local retail pricing —
                                  Suggested/Proposed Retail Price arrive via
                                  the BY response, BM rows 234/235)
    "What 3 words would best describe this product to a consumer?" (not
                                  referenced)
    Construction                  (not referenced)
    Item Development Status, Dev Site, Pack Name (not referenced)
    In-Plan DTC Global, Global DTC Door Format, VLP VDA Comments (not ref.)
    APAC DTC DA Scale             (not referenced)
    Local Key Look Capsule, Local Key Look, MYSG, Drop SKU, Exclusive,
    Comments, Key Story, DA Look, DA, x  (planning / merch commentary —
                                  not referenced)

  ATTRIBUTES THE SPEC PINS BUT WHOSE SOURCE IS ABSENT / MANUAL for the
  APP/ACC branch (deliberately NOT emitted):
    AT_PrincipalColorCode is emitted (from Item Number) — see above.
    AT_PrincipalGenderCode / AT_PrincipalAgeCode / AT_PrincipalAgeDescription
                      — BM rows 17/19/20: Not Available.
    AT_EComAgesCategory — BM row 132: APP/ACC = Manual input by MD.
    AT_CountrySize      — BM row 130: APP/ACC = Manual input from MD.
    AT_MaterialUpper    — BM row 103: APP ACC = N/A.
    AT_Fastening        — BM row 104: APP/ACC = N/A.
    AT_SAPStyleCode / AT_Generic / AT_GenericDescription / AT_Variant /
    AT_VariantDescription — BM rows 22-26 "Formula in System" ("Type C"),
                      computed downstream from the BY SAP colour/size
                      response; never guessed here.

────────────────────────────────────────────────────────────────────
DEVIATIONS (documented, deliberate)
────────────────────────────────────────────────────────────────────
 1. AT_PrincipalStyleCode keeps coming from Product Number (BM row 13 —
    v1.0 already did). KEY_InboundArticle stays keyed on the row-unique
    Item Number — the key must remain 1:1 with a source row (1627 rows
    would collapse to 606 articles if keyed on Product Number, and the
    colour component the original formula used is carried by the
    "_AAX" suffix, not a separate column).
 2. SAP Age / BY Age come from Line Plan Business (BM rows 37/95: "Mapping
    from column 'Line Plan Bussiness'"): rows whose LPB starts with "KIDS"
    → Children / Kids; every other value → Adults / Adult. The v6 sheet
    (row 049) states the same rule.
 3. Sports Category LOV IDs are zero-padded 2-digit strings ("07" for
    Running) — the ID format the live Stibo "Sports Category" LOV accepts
    (same rule as asics/footwear_main.py, anta/linelist_main.py, reebok).
 4. LOV attributes (Silhouette / Style / Content / Fabric / Material / Fit)
    are only emitted when the derived value resolves to a real LOV ID;
    a miss is skipped with a warning — writing raw text to a LOV attribute
    throws "Illegal LOV" on STEP import (lesson recorded in
    reebok/article_master_*). Merchandise-hierarchy text attributes still
    receive the raw column values.
 5. AT_BCI carries the Price Segment text (BM row 97). The MDD declares
    LOV "BCI" but ships no BCI LOV sheet, so the ID resolves to "" — the
    text is emitted exactly like lotto/recap_main.py does (display value,
    best-effort ID).
 6. FOB (Estimated GEP) is emitted for non-zero values only; the 27 zero /
    blank rows are omitted and reported (a FOB of 0 would price the
    article wrongly).
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
LINELIST_DIR = INPUT_DIR / "linelist_appacc_preline"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_appacc_preline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# ("APPACC"). Name-based fallbacks cover renamed exports.
# ══════════════════════════════════════════════════════════════════
SHEET_MODE = "first_visible"          # operator instruction (default)
SHEET_NAME_FALLBACKS = ["APPACC", "APP ACC", "Preline", "Linelist", "Sheet1"]


# ══════════════════════════════════════════════════════════════════
# SOURCE COLUMN CONTRACT — every column this ETL reads, and the
# attributes each one feeds (BM = "NEW BALANCE  LINELIST" sheet row).
# ══════════════════════════════════════════════════════════════════
SOURCE_COLUMNS: dict[str, list[str]] = {
    "Product Number":              ["AT_PrincipalStyleCode"],
    "Product Display Name":        ["AT_PrincipalStyleDescription", "<Name>"],
    "Item Number":                 ["AT_PrincipalColorCode", "AT_InboundGenericCode"],
    "Color Code - Non-Footwear":   ["AT_PrincipalColorName"],
    "Gender Duty - Non-Footwear":  ["AT_PrincipalGenderDescription",
                                    "AT_Gender", "AT_BYGender"],
    "Global Size Offering":        ["AT_PrincipalSize"],
    "Line Plan Business":          ["AT_SAPAge", "AT_BYAge",
                                    "AT_SportsCategoryEN",
                                    "AT_PrincipalMerchandiseHierarchyL2"],
    "Product Line":                ["AT_PrincipalMerchandiseHierarchyL1",
                                    "AT_Collection1 (component)"],
    "Silhouette":                  ["AT_PrincipalMerchandiseHierarchyL3",
                                    "AT_Silhouette", "AT_Collection1 (component)"],
    "GBU":                         ["AT_PrincipalMerchandiseHierarchyL4"],
    "Detailed Silhouette":         ["AT_PrincipalMerchandiseHierarchyL5",
                                    "AT_Style", "AT_Collection1 (component)"],
    "Price Segment":               ["AT_BCI"],
    "Product Fit":                 ["AT_Fit"],
    "Target Principal Material":   ["AT_Content", "AT_Fabric", "AT_Material"],
    "Country Of Origin":           ["AT_CountryOrigin"],
    "Technologies":                ["AT_TechnologyUsed"],
    "Product Collection":          ["AT_Collection2"],
    "SEA Intro Date":              ["AT_LaunchingDate", "AT_IncomingMonth"],
    "Estimated GEP":               ["AT_FOB"],
}

# Columns present in the S2'27 Preline sheet that no attribute row of the
# two mapping sheets references (see module docstring for the full audit).
UNMAPPED_COLUMNS: dict[str, str] = {
    "Thumbnail":                    "BM row 94 has no Stibo Attribute ID (#N/A)",
    "Item - CarryOver/New":         "not referenced by any attribute row",
    "Date Canceled":                "not referenced by any attribute row",
    "APAC DA/DROP":                 "not referenced by any attribute row",
    "SMS":                          "not referenced by any attribute row",
    "EMERGING":                     "not referenced by any attribute row",
    "Category":                     "not referenced by any attribute row",
    "LPA Category":                 "not referenced by any attribute row",
    "Game Plan":                    "not referenced by any attribute row",
    "Global Brand Event":           "not referenced by any attribute row",
    "Mandatory Global Brand Event": "not referenced by any attribute row",
    "Global Item Intro Date":       "Launching Date uses SEA Intro Date",
    "Intro Date - APAC":            "Incoming month uses SEA Intro Date (BM row 233)",
    "Global Item Phaseout Date":    "not referenced by any attribute row",
    "Phaseout Date - APAC":         "not referenced by any attribute row",
    "Global Wholesale Price":       "not referenced by any attribute row",
    "Region Sales Sample Set - APAC": "not referenced by any attribute row",
    "Region Sales Sample Set":      "not referenced by any attribute row",
    "S127 Price List":              "not referenced by any attribute row",
    "SGD w/o Tax":                  "local retail price — BY response",
    "SGD":                          "local retail price — BY response",
    "MYRR":                         "local retail price — BY response",
    "THB":                          "local retail price — BY response",
    "PHP":                          "local retail price — BY response",
    "IDR":                          "local retail price — BY response",
    "VND":                          "local retail price — BY response",
    "What 3 words would best describe this product to a consumer?":
                                    "not referenced by any attribute row",
    "Construction":                 "not referenced by any attribute row",
    "Item Development Status":      "not referenced by any attribute row",
    "Dev Site":                     "not referenced by any attribute row",
    "Pack Name":                    "not referenced by any attribute row",
    "In-Plan DTC Global":           "not referenced by any attribute row",
    "Global DTC Door Format":       "not referenced by any attribute row",
    "VLP VDA Comments":             "not referenced by any attribute row",
    "APAC DTC DA Scale":            "not referenced by any attribute row",
    "Local Key Look Capsule":       "not referenced by any attribute row",
    "Local Key Look":               "not referenced by any attribute row",
    "MYSG":                         "not referenced by any attribute row",
    "Drop SKU":                     "not referenced by any attribute row",
    "Exclusive":                    "not referenced by any attribute row",
    "Comments":                     "not referenced by any attribute row",
    "Key Story":                    "not referenced by any attribute row",
    "DA Look":                      "not referenced by any attribute row",
    "DA":                           "not referenced by any attribute row",
    "x":                            "junk helper column",
}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES  (IDs resolved against the MDD at runtime; the hard-coded
# fallbacks below match the MDD LOV sheets — same convention as every
# NB ETL in this repository)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}

# MDD "Gender LOV" — SAP Gender (value → ID)
LOV_GENDER = {"Male": "M", "Female": "F", "Unisex": "U"}

# Gender Duty - Non-Footwear → SAP Gender (BM row 38, verbatim)
GENDER_DUTY_TO_GENDER: dict[str, tuple] = {
    "MENS":   ("M", "Male"),
    "WOMENS": ("F", "Female"),
    "UNISEX": ("U", "Unisex"),
}

# MDD "Season LOV" — code → season name
LOV_SEASON = {
    # MDD "Season LOV" — code → season name
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
    "AP": "Apparel", "AC": "Accessories",
}

# MDD "Article Category LOV" — 1 = Generic (BM row 84 default)
LOV_SAP_ARTICLE_CATEGORY = {
    "0": "Single", "1": "Generic", "2": "Variant",
    "10": "Sell set (Hampers)", "11": "Prepack (Musical Box)",
}

# MDD "BY Article Type" — code → name
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
    "Licensed": "License", "INLINE": "Inline", "TEAM": "Inline",
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

# ── Sports Category → MDD "Sports Category LOV" ID (zero-padded) ────────────
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

# ── Country Of Origin: "VNM - Viet Nam" → 3-letter token → 2-letter LOV ─────
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

# MDD "Country Origin LOV" (ID → name) — subset covering every country in
# the Preline sheet (resolved against the full MDD LOV at runtime).
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

# ── Product Fit → MDD "Fit LOV" close values (BM row 109) ────────────────────
# Fit LOV: F=Fitted, R=Regular, L=Loose, C=Compression Fit.
# "0" / "N/A" rows are skipped (source noise).
FIT_CLOSE_MAP: dict[str, tuple] = {
    "FITTED":    ("F", "Fitted"),
    "SLIM":      ("F", "Fitted"),
    "STANDARD":  ("R", "Regular"),
    "RELAXED":   ("L", "Loose"),
    "OVERSIZED": ("L", "Loose"),
}

# ── Target Principal Material → Material / Content / Fabric LOV close values ─
# Three attributes share one source column (BM rows 106/107/111); each is
# emitted only when its own LOV resolves (see DEVIATION 4).
MATERIAL_LOV_CLOSE: dict[str, tuple] = {          # MDD "Material LOV"
    "COTTON":        ("COT", "Cotton"),
    "COTTON JERSEY": ("COT", "Cotton"),
    "COTTON FLEECE": ("COT", "Cotton"),
    "COTTON TWILL":  ("COT", "Cotton"),
    "POLYESTER":     ("PES", "Polyester"),
    "POLY KNIT":     ("PES", "Polyester"),
    "POLYWOVEN":     ("PES", "Polyester"),
    "NYLON":         ("NYL", "Nylon"),
    "NYLON WOVEN":   ("NYL", "Nylon"),
    "WOOL":          ("WOL", "Wool"),
    "MESH":          ("MES", "Mesh"),
    "SYNTHETIC":     ("SYN", "Synthetic"),
    "ACRYLIC":       ("ACR", "Acrylic"),
    "POLYURETHANE":  ("PU",  "PU"),
}
CONTENT_LOV_CLOSE: dict[str, tuple] = {           # MDD "Content LOV"
    "COTTON":        ("COTTON", "Cotton"),
    "COTTON JERSEY": ("COTTON", "Cotton"),
    "COTTON FLEECE": ("COTTON", "Cotton"),
    "COTTON TWILL":  ("COTTON", "Cotton"),
    "POLYESTER":     ("POLYESTER", "Polyester"),
    "POLY KNIT":     ("POLYESTER", "Polyester"),
    "POLYWOVEN":     ("POLYESTER", "Polyester"),
    "NYLON":         ("NYLON", "Nylon"),
    "NYLON WOVEN":   ("NYLON", "Nylon"),
}
FABRIC_LOV_CLOSE: dict[str, tuple] = {            # MDD "Fabric LOV"
    "COTTON JERSEY": ("JE", "Jersey"),
    "COTTON FLEECE": ("FL", "Fleece"),
    "COTTON TWILL":  ("TW", "Twill"),
}

# Product Line → Division letter (for PPH parent and division code)
DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E", "FOOTWEAR": "F", "APPAREL": "A",
    "EQUIPMENT":   "Q", "TOYS":     "T",
}

# MDD AT_PrincipalGenderDescription max chars (MDD: 10)
MAX_GENDER_DESC_CHARS = 10

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

# Serialise STEP-namespace elements with the DEFAULT namespace — same
# convention as the newest NB ETL modules; guarantees identical output
# standalone AND inside the unified lambda.
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')

_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


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
        self._load_single_col_lovs(wb)
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

    def _load_single_col_lovs(self, wb):
        """
        Single-column LOV sheets ("Style LOV", "Franchise LOV") list values
        only — the value IS the ID. Loaded as {value: value} so runtime
        lookups keep working through the same lovs{} table.
        """
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            if self.lovs.get(display):
                continue          # already loaded from a two-column layout
            rows = list(wb[sn].iter_rows(values_only=True))
            for row in rows[1:]:
                if not row or row[0] is None:
                    continue
                val = str(row[0]).strip()
                if val and (len(row) < 2 or row[1] is None):
                    self.lovs.setdefault(display, {})[val] = val

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

    def lov_value_exists(self, lov_display: str, value: str) -> bool:
        """Case-insensitive membership test against an MDD LOV."""
        table = self.lovs.get(lov_display, {})
        if not value:
            return False
        v = value.strip()
        return v in table or v.title() in table or v.upper() in table


class BrandMappingLinelistLoader:
    """
    Parses the "NEW BALANCE  LINELIST" sheet (note: TWO spaces in the sheet
    name — matched fuzzily on collapsed whitespace) of the NEW Brand Mapping
    file that the router downloads into input/attributes/.

    The ETL's attribute mapping is hard-coded from that sheet (repo
    convention — same as every NB ETL), but the loader lets the module
    SELF-CHECK at runtime that every attribute ID it emits is declared in
    the live mapping sheet (same contract as the sample ETLs).
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


class AppAccPrelineLoader:
    """
    Loads the New Balance APP/ACC Preline Line List.

    Sheet selection (SHEET_MODE):
      * "first_visible"  → the first visible sheet ("APPACC")
      * name fallback    → "APPACC" / "APP ACC" / "Preline" / "Linelist"

    The sheet carries two short marker rows ("ACC", "APP") above the
    header, so the header row is located dynamically by scanning for
    "Item Number" — same strategy as the price-list loader.
    """

    ROW_KEY = "Item Number"

    def __init__(self, path: Path, sheet_mode: str = SHEET_MODE):
        self.path  = path
        self.mode  = sheet_mode
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[AppAccPreline] Loading: %s", self.path.name)
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
        log.info("[AppAccPreline] Using sheet: '%s' (mode=%s)", target, self.mode)

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
            log.error("[AppAccPreline] Cannot find header row in '%s'", target)
            wb.close()
            return

        log.info("[AppAccPreline] Header at row index %d (row %d)", hdr_idx, hdr_idx + 1)
        header = [
            str(h).replace("\n", " ").strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
        # True 1-based Excel row number (captured before any filtering) so
        # test limits can target exact spreadsheet rows.
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
            log.warning("[AppAccPreline] Expected source column(s) absent: %s", missing)

        log.info("[AppAccPreline] %d SKU rows loaded from '%s'", len(self.df), self.sheet)


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


def _fmt_date_ddmonyyyy(v) -> str:
    """AT_LaunchingDate — DD-Mon-YYYY (repo-wide convention)."""
    if v is None:
        return ""
    if hasattr(v, "day") and hasattr(v, "month") and hasattr(v, "year"):
        return f"{v.day:02d}-{_MONTH_ABBR.get(v.month, str(v.month))}-{v.year}"
    s = _s(v)
    if not s:
        return ""
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            d = datetime.strptime(s, fmt)
            return f"{d.day:02d}-{_MONTH_ABBR.get(d.month, str(d.month))}-{d.year}"
        except ValueError:
            continue
    return ""


def _fmt_date_ddmmyyyy(v) -> str:
    """AT_IncomingMonth — DD-MM-YYYY (repo-wide convention)."""
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


def _gender_duty_to_gender(gender_duty: str) -> tuple[str, str, str]:
    """
    Gender Duty - Non-Footwear → SAP Gender (BM row 38: Mens→Male,
    Womens→Female, Unisex→Unisex). Returns (sap_id, sap_value, note);
    note is "" on a direct hit. Unknown values default to Unisex (warned).
    """
    raw = _s(gender_duty).upper().strip()
    if raw in GENDER_DUTY_TO_GENDER:
        gid, gval = GENDER_DUTY_TO_GENDER[raw]
        return gid, gval, ""
    return "U", "Unisex", f"unmapped Gender Duty '{gender_duty}' → Unisex default"


def _line_plan_to_age(line_plan: str) -> tuple[str, str, str, str, str]:
    """
    BM rows 37/95: SAP Age / BY Age "Mapping from column 'Line Plan
    Bussiness'" — Kids rows → Children / Kids, everything else →
    Adults / Adult. Returns (sap_id, sap_val, by_age, by_age_id, note).
    """
    raw = _s(line_plan).upper().strip()
    if not raw:
        return "AD", "Adults", "Adult", "ADULT", ""
    if raw.startswith("KIDS"):
        return "CH", "Children", "Kids", "KIDS", ""
    return "AD", "Adults", "Adult", "ADULT", ""


def _line_plan_to_sports_cat(line_plan: str) -> str:
    """BM row 121 mapping. Unmapped → "" (omitted, warned — never guessed)."""
    return LOV_LINE_PLAN_TO_SPORTS_CAT.get(_s(line_plan).upper().strip(), "")


def _sports_cat_to_lov_id(sports_cat: str) -> str:
    """Sports Category display name → MDD LOV ID (zero-padded)."""
    return LOV_SPORTS_CATEGORY_ID.get(sports_cat.strip(), "")


def _country_origin_tokens(coo_raw: str) -> tuple[str, str]:
    """
    "VNM - Viet Nam" → the leading 3-letter token (VNM) → 2-letter Country
    Origin LOV (VN/Vietnam). Returns ("", "") when unresolvable.
    """
    tokens = [t for t in re.split(r"[\s\-–—|,;()]+", _s(coo_raw)) if t]
    for t in tokens:
        if len(t) == 3 and t.isalpha() and t.upper() in ISO_A3_TO_A2:
            a2 = ISO_A3_TO_A2[t.upper()]
            return a2, LOV_COUNTRY_ORIGIN.get(a2, a2)
    # Also accept a bare 2-letter code straight from the sheet
    if tokens and len(tokens[0]) == 2 and tokens[0].upper() in LOV_COUNTRY_ORIGIN:
        a2 = tokens[0].upper()
        return a2, LOV_COUNTRY_ORIGIN[a2]
    return "", ""


def _item_color_code(item_number: str) -> str:
    """BM row 15: "Item number (digits after '_')" → MB6180TA_AAX → AAX."""
    item = _s(item_number)
    if "_" in item:
        suffix = item.rsplit("_", 1)[-1].strip()
        return suffix if suffix else ""
    return ""


def _fit_close(product_fit: str) -> tuple[str, str]:
    """Product Fit → MDD 'Fit LOV' close value; '' when unmappable.
    '0' / 'N/A' / blank carry no fit information and return ('', '')
    WITHOUT a warning (source noise, not a mapping gap)."""
    raw = _s(product_fit).upper().strip()
    if raw in ("", "0", "N/A", "NA", "NONE"):
        return ("", "")
    return FIT_CLOSE_MAP.get(raw, ("", ""))


def _material_triples(target_material: str) -> dict:
    """
    Target Principal Material → (id, value) per LOV family for
    AT_Material / AT_Content / AT_Fabric. Unresolvable families get "".
    """
    raw = _s(target_material).upper().strip()
    mat  = MATERIAL_LOV_CLOSE.get(raw, ("", ""))
    cont = CONTENT_LOV_CLOSE.get(raw, ("", ""))
    fab  = FABRIC_LOV_CLOSE.get(raw, ("", ""))
    return {"material": mat, "content": cont, "fabric": fab}


def _technology_clean(technologies: str) -> str:
    """Technologies → AT_TechnologyUsed; 'N/A' / '0' rows are dropped."""
    raw = _s(technologies)
    if raw.upper() in ("N/A", "0", "NA", "NONE"):
        return ""
    return raw


def _build_generic_code(brand_code: str, item_number: str) -> str:
    """
    KEY_InboundArticle / AT_InboundGenericCode — repo convention (every NB
    ETL keys the generic article 1:1 with a source row): brand(3) + the
    row-unique Item Number. Keying on Product Number would collapse the
    1627 rows into 606 articles (see DEVIATION 1 in the module docstring).
    """
    brand = re.sub(r"[^A-Z0-9]", "", brand_code.upper())[:3]
    base  = re.sub(r"[^A-Z0-9_]", "", item_number.upper()).replace("_", "")
    return f"{brand}{base}"


def map_sku(row: pd.Series, brand_code: str = "NEW") -> dict:
    item_number    = _s(row.get("Item Number"))
    product_number = _s(row.get("Product Number"))
    display_name   = _s(row.get("Product Display Name"))
    color_name     = _s(row.get("Color Code - Non-Footwear"))
    gender_duty    = _s(row.get("Gender Duty - Non-Footwear"))
    size_offering  = _norm_ws(row.get("Global Size Offering"))
    line_plan_biz  = _s(row.get("Line Plan Business"))
    product_line   = _s(row.get("Product Line"))
    silhouette     = _s(row.get("Silhouette"))
    gbu            = _s(row.get("GBU"))
    detail_sil     = _s(row.get("Detailed Silhouette"))
    price_segment  = _s(row.get("Price Segment"))
    product_fit    = _s(row.get("Product Fit"))
    target_mat     = _s(row.get("Target Principal Material"))
    coo_raw        = _s(row.get("Country Of Origin"))
    technologies   = _s(row.get("Technologies"))
    collection     = _s(row.get("Product Collection"))
    sea_intro      = row.get("SEA Intro Date", None)
    gep_raw        = row.get("Estimated GEP", None)
    excel_row      = row.get("_ExcelRow")

    gender_id, gender_val, gender_note = _gender_duty_to_gender(gender_duty)
    age_id, age_val, by_age, by_age_id, age_note = _line_plan_to_age(line_plan_biz)
    sports_cat  = _line_plan_to_sports_cat(line_plan_biz)
    coo_id, coo_val = _country_origin_tokens(coo_raw)
    fit_id, fit_val = _fit_close(product_fit)
    mat_trip = _material_triples(target_mat)

    # FOB — Estimated GEP (BM row 60): non-zero floats only
    fob_value = ""
    gep_s = _s(gep_raw)
    if gep_s:
        try:
            f = float(gep_s)
            if f > 0:
                fob_value = f"{round(f, 2):g}"
        except (TypeError, ValueError):
            fob_value = ""

    # BM row 122: "Product Line - Silhouette - Detailed Silhouette"
    collection1 = " - ".join(
        p for p in (product_line, silhouette, detail_sil) if p
    ) if (product_line or silhouette or detail_sil) else ""

    return {
        # ── source values ───────────────────────────────────────
        "item_number":     item_number,
        "product_number":  product_number,
        "display_name":    display_name,
        "color_name":      color_name,
        "gender_duty":     gender_duty,
        "size_offering":   size_offering,
        "line_plan_biz":   line_plan_biz,
        "product_line":    product_line,
        "silhouette":      silhouette,
        "gbu":             gbu,
        "detailed_silhouette": detail_sil,
        "price_segment":   price_segment,
        "product_fit":     product_fit,
        "target_material": target_mat,
        "coo_raw":         coo_raw,
        "technologies":    _technology_clean(technologies),
        "collection":      collection,
        "sea_intro_raw":   sea_intro,
        "gep_raw":         gep_raw,
        "excel_row":       int(excel_row) if excel_row is not None else None,
        # ── derived ──────────────────────────────────────────────
        "color_code":      _item_color_code(item_number),
        "gender_id":       gender_id,
        "gender_val":      gender_val,
        "gender_note":     gender_note,
        "age_id":          age_id,
        "age_val":         age_val,
        "by_age":          by_age,
        "by_age_id":       by_age_id,
        "age_note":        age_note,
        "sports_cat":      sports_cat,
        "sports_cat_id":   _sports_cat_to_lov_id(sports_cat),
        "coo_id":          coo_id,
        "coo_val":         coo_val,
        "fit_id":          fit_id,
        "fit_val":         fit_val,
        "material_id":     mat_trip["material"][0],
        "material_val":    mat_trip["material"][1],
        "content_id":      mat_trip["content"][0],
        "content_val":     mat_trip["content"][1],
        "fabric_id":       mat_trip["fabric"][0],
        "fabric_val":      mat_trip["fabric"][1],
        "collection1":     collection1,
        "launching_date":  _fmt_date_ddmonyyyy(sea_intro),
        "incoming_month":  _fmt_date_ddmmyyyy(sea_intro),
        "fob":             fob_value,
        # ── codes / context ──────────────────────────────────────
        "brand_code":      brand_code,
        "division":        product_line,
        "generic_code":    _build_generic_code(brand_code, item_number),
        # File-type constant: this is an NB *Inline* preline linelist.
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
        "product_number": "AT_PrincipalStyleCode",
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
    if mapped.get("gender_note"):
        warns.append(f"[{sku}] {mapped['gender_note']}")
    if not mapped.get("coo_id") and mapped.get("coo_raw"):
        warns.append(f"[{sku}] Country of origin unresolved from '{mapped['coo_raw']}' — "
                     "AT_CountryOrigin omitted")
    if mapped.get("target_material") and not mapped.get("material_id"):
        warns.append(f"[{sku}] Material LOV close value unresolved from "
                     "'{mapped['target_material']}' — AT_Material omitted")
    if mapped.get("product_fit") and not mapped.get("fit_id"):
        pf = _s(mapped.get("product_fit")).upper().strip()
        if pf not in ("0", "N/A", "NA", "NONE"):
            warns.append(f"[{sku}] Fit close value unresolved from '{mapped['product_fit']}' — "
                         "AT_Fit omitted")
    if mapped.get("sea_intro_raw") is not None and not mapped.get("incoming_month"):
        warns.append(f"[{sku}] UNPARSEABLE SEA Intro Date '{mapped['sea_intro_raw']}' — "
                     "AT_LaunchingDate/AT_IncomingMonth omitted")
    if mapped.get("gep_raw") is not None and _s(mapped.get("gep_raw")) and not mapped.get("fob"):
        warns.append(f"[{sku}] FOB (Estimated GEP) zero or unparseable "
                     f"'{mapped['gep_raw']}' — AT_FOB omitted")

    return warns


def _log_dropped_attributes(mdd: MDDLoader) -> list[str]:
    """Emit a one-time, run-level report of every attribute not carried over."""
    notes: list[str] = []
    notes.append(
        "Attributes the spec pins as Manual Input / N/A / Formula-in-System for the "
        "APP/ACC branch (deliberately NOT emitted):"
    )
    dropped = {
        "AT_PrincipalGenderCode":     "BM row 17: Not Available",
        "AT_PrincipalAgeCode":        "BM row 19: Not Available",
        "AT_PrincipalAgeDescription": "BM row 20: Not Available",
        "AT_EComAgesCategory":        "BM row 132: APP/ACC = Manual input by MD",
        "AT_CountrySize":             "BM row 130: APP/ACC = Manual input from MD",
        "AT_MaterialUpper":           "BM row 103: APP ACC = N/A",
        "AT_Fastening":               "BM row 104: APP/ACC = N/A",
        "AT_SAPStyleCode / AT_Generic / AT_GenericDescription / AT_Variant / "
        "AT_VariantDescription":
            "BM rows 22-26 'Formula in System' (Type C) — computed downstream "
            "from the BY SAP colour/size response; never guessed here",
    }
    for at_id, why in dropped.items():
        card = (mdd.attributes.get(at_id, {}).get("cardinality") or "n/a").strip() if mdd else "n/a"
        notes.append(f"  {at_id:44} cardinality={card:12} {why}")
        if "mandatory" in card.lower():
            log.warning("MANDATORY attribute %s not emitted — %s", at_id, why)

    notes.append("")
    notes.append("Source columns present in the sheet but mapped to no attribute row:")
    for col, why in UNMAPPED_COLUMNS.items():
        notes.append(f"  {col[:44]:44} {why}")
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


def _get_division_code(product_line: str) -> str:
    """Map Product Line to single-letter division code for PPH parent ID."""
    pl_upper = (product_line or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in pl_upper:
            return code
    return "X"


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC VALUE WRITER
# ══════════════════════════════════════════════════════════════════

# Attribute IDs deliberately exempt from the BrandMapping self-check:
# the sheet row exists but its ID cell is #N/A in this revision — the
# attribute itself is declared in the MDD (Mandatory, LOV "BY Article
# Type") and in the v6 sheet (R105 "Article Type", APP/ACC Manual Input —
# this module pins the Inline constant, same as every NB Inline ETL).
SHEET_ID_GAPS: set[str] = {"AT_BYArticleType"}

# Every AT_ id this module writes (KEY_InboundArticle / <Name> are not
# attributes). Used for the BrandMapping self-check in run().
EMITTED_ATTRIBUTE_IDS: list[str] = [
    "AT_Country", "AT_CompanyCode", "AT_SBU", "AT_Brand", "AT_BrandGroup",
    "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
    "AT_PrincipalColorCode", "AT_PrincipalColorName",
    "AT_PrincipalGenderDescription", "AT_PrincipalSize",
    "AT_Gender", "AT_SAPAge", "AT_BYGender", "AT_BYAge",
    "AT_CountryOrigin", "AT_FOB", "AT_FOBCurrency",
    "AT_Season", "AT_SeasonYear",
    "AT_PrincipalMerchandiseHierarchyL1", "AT_PrincipalMerchandiseHierarchyL2",
    "AT_PrincipalMerchandiseHierarchyL3", "AT_PrincipalMerchandiseHierarchyL4",
    "AT_PrincipalMerchandiseHierarchyL5",
    "AT_SportsCategoryEN", "AT_Silhouette", "AT_BCI",
    "AT_Content", "AT_Fabric", "AT_Fit", "AT_Style", "AT_Material",
    "AT_Collection1", "AT_Collection2", "AT_TechnologyUsed",
    "AT_LaunchingDate", "AT_IncomingMonth",
    "AT_MaterialType", "AT_SAPArticleCategory", "AT_BYArticleType",
    "AT_SAPProductFlag", "AT_UOM", "AT_ArticleStatus",
]


def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None, bm=None):
    """
    Writes exactly the attributes the NEW BALANCE  LINELIST mapping sheet
    defines as populated for the App/Acc Preline line-list flow (see
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

    # ── Principal style / colour (BM rows 13/14/15/16) ────────────
    _val(vals_el, "AT_PrincipalStyleCode",        art["product_number"])
    _val(vals_el, "AT_PrincipalStyleDescription",
         art["display_name"] or art["item_number"])
    _val(vals_el, "AT_PrincipalColorCode",        art["color_code"])
    _val(vals_el, "AT_PrincipalColorName",        art["color_name"])

    # ── Gender Duty - Non-Footwear (BM rows 18/38/96) ─────────────
    gender_max = MAX_GENDER_DESC_CHARS
    if mdd is not None:
        gender_max = mdd.max_chars("AT_PrincipalGenderDescription", MAX_GENDER_DESC_CHARS)
    _val(vals_el, "AT_PrincipalGenderDescription", art["gender_duty"][:gender_max])
    _val(vals_el, "AT_Gender",   art["gender_val"], id_val=art["gender_id"])
    _val(vals_el, "AT_BYGender", art["gender_val"], id_val=art["gender_val"].upper())

    # ── Line Plan Business → SAP Age / BY Age (BM rows 37/95) ─────
    _val(vals_el, "AT_SAPAge", art["age_val"], id_val=art["age_id"])
    _val(vals_el, "AT_BYAge",  art["by_age"],  id_val=art["by_age_id"])

    # ── Global Size Offering → Principal Size (BM row 21) ─────────
    _val(vals_el, "AT_PrincipalSize", art["size_offering"])

    # ── Country Of Origin (BM row 41) ─────────────────────────────
    if art["coo_id"]:
        _val(vals_el, "AT_CountryOrigin", art["coo_val"], id_val=art["coo_id"])

    # ── Season (from filename metadata) ───────────────────────────
    sea_raw   = art.get("season", "") or ""
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_code)

    year_raw = sea_raw[2:] if len(sea_raw) > 2 else ""
    year_val = f"20{year_raw}" if len(year_raw) == 2 else year_raw
    _val(vals_el, "AT_SeasonYear", year_val)

    # ── Merchandise hierarchy (BM rows 78/79/80/81/82) ────────────
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art.get("product_line", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art.get("silhouette", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL4", art.get("gbu", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL5", art.get("detailed_silhouette", ""))

    # ── Line Plan Business → Sports Category (BM row 121, LOV ID) ─
    if art.get("sports_cat_id"):
        _val(vals_el, "AT_SportsCategoryEN", art["sports_cat"],
             id_val=art["sports_cat_id"])

    # ── Silhouette → AT_Silhouette (LOV ID from the MDD at runtime) ─
    if art.get("silhouette"):
        sil_id = mdd.lov_id("Silhouette", art["silhouette"]) if mdd else ""
        if sil_id:
            _val(vals_el, "AT_Silhouette", art["silhouette"], id_val=sil_id)

    # ── Detailed Silhouette → AT_Style (LOV value match, runtime) ──
    if art.get("detailed_silhouette") and mdd is not None:
        if mdd.lov_value_exists("Style", art["detailed_silhouette"]):
            _val(vals_el, "AT_Style", art["detailed_silhouette"],
                 id_val=art["detailed_silhouette"])

    # ── Price Segment → BCI (BM row 97; ID best-effort from MDD) ──
    if art.get("price_segment"):
        bci_id = mdd.lov_id("BCI", art["price_segment"]) if mdd else ""
        _val(vals_el, "AT_BCI", art["price_segment"], id_val=bci_id)

    # ── Target Principal Material → Content / Fabric / Material ────
    if art.get("content_id"):
        _val(vals_el, "AT_Content", art["content_val"], id_val=art["content_id"])
    if art.get("fabric_id"):
        _val(vals_el, "AT_Fabric",  art["fabric_val"],  id_val=art["fabric_id"])
    if art.get("material_id"):
        _val(vals_el, "AT_Material", art["material_val"], id_val=art["material_id"])

    # ── Product Fit → AT_Fit (close value) ────────────────────────
    if art.get("fit_id"):
        _val(vals_el, "AT_Fit", art["fit_val"], id_val=art["fit_id"])

    # ── Collections (BM rows 122/123) ─────────────────────────────
    _val(vals_el, "AT_Collection1", art.get("collection1", ""))
    _val(vals_el, "AT_Collection2", art.get("collection", ""))

    # ── Technologies (BM row 158) ─────────────────────────────────
    _val(vals_el, "AT_TechnologyUsed", art.get("technologies", ""))

    # ── SEA Intro Date → Launching Date / Incoming month ──────────
    _val(vals_el, "AT_LaunchingDate", art.get("launching_date", ""))
    _val(vals_el, "AT_IncomingMonth", art.get("incoming_month", ""))

    # ── FOB / FOB Currency (BM rows 60/61) ────────────────────────
    _val(vals_el, "AT_FOB", art.get("fob", ""))
    _val(vals_el, "AT_FOBCurrency", "United States Dollar", id_val="USD")

    # ── Spec defaults (BM rows 83/84/93/77/90/232) ────────────────
    _val(vals_el, "AT_MaterialType", id_val="ZINA")            # Intercompany

    cat_code  = art["art_category"]                             # "1"
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, "Generic")
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)

    at_norm  = art["article_type"]                              # "Inline"
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)

    _val(vals_el, "AT_SAPProductFlag", "Intercompany", id_val="A")
    _val(vals_el, "AT_UOM", "Each", id_val="EA")
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

    div_letter  = _get_division_code(art.get("division", ""))
    parent_id   = f"PPH_{div_letter}-TempSubCat"
    key_generic = art["generic_code"]

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    # v6 R027: Product Display Name, blank → Item Number
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["display_name"] or item_number

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
    # Preline arrives as .xlsm (macro-enabled) — accept both extensions.
    ll_files = [f for pat in ("*.xlsx", "*.xlsm") for f in LINELIST_DIR.glob(pat)]

    for label, val in [
        ("MDD",               mdd_f),
        ("Attributes",        attr_f),
        ("Preline Line List", ll_files),
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
        log.info("─── Processing App/Acc Preline Line List: %s ───", ll_path.name)

        ll = AppAccPrelineLoader(ll_path)
        if ll.df.empty:
            log.warning("[AppAccPreline] Empty dataframe — skipping.")
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
        log.info("[AppAccPreline] %d valid SKU rows to process", total_rows)

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
        # "_ExcelRow" in AppAccPrelineLoader / "excel_row" in map_sku).
        # TEST_ROWS = {5, 6}
        # mapped_skus = [m for m in mapped_skus if m.get("excel_row") in TEST_ROWS]
        # log.info(
        #     "[AppAccPreline][TEST LIMIT] Restricted to Excel rows %s → %d SKU(s) matched",
        #     sorted(TEST_ROWS), len(mapped_skus),
        # )

        # Guard the one deviation: KEY_InboundArticle must stay 1:1 with rows.
        keys  = [m["generic_code"] for m in mapped_skus if m.get("item_number")]
        dupes = len(keys) - len(set(keys))
        if dupes:
            log.warning(
                "[AppAccPreline] %d duplicate KEY_InboundArticle value(s) — "
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

    rpt_path = LOG_DIR / f"validation_nb_appacc_preline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
        description="Stibo Inbound XML Generator — New Balance App/Acc Preline v2.0"
    )
    p.add_argument("--brand",        default="New Balance")
    p.add_argument("--brand-code",   default="NEW")
    p.add_argument("--comp-code",    default="0888")
    p.add_argument("--sbu",          default="AP")
    p.add_argument("--season",       default="SS27")
    p.add_argument("--seq",          default=1, type=int)
    p.add_argument("--country-code", default="")
    run(p.parse_args())


if __name__ == "__main__":
    main()
