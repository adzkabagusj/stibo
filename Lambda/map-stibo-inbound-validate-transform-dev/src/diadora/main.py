"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Diadora Inline  v1.0           ║
║   Pricelist (Analisi LineList tab) → Stibo STEP XML            ║
╚══════════════════════════════════════════════════════════════════╝

Diadora Inline-specific details:

Source file
  • Single input file : "FW26_US___UK_SIZE_Pricelist_-_Indonesia.xlsx"
  • Active sheet      : "Analisi LineList"
  • Header row        : index 0 (0-based)  → row 1 in Excel
  • Data starts       : index 1 (0-based)  → row 2 in Excel
  • No filter needed  — all rows are inline articles

Column mapping (0-indexed confirmed from file inspection):
    col[0]  = CAT             → article status / category tier (Active, Spw T1_T2, Spw T3)
    col[1]  = Collection      → Principal Hierarchy L2 / SAP Product Group / BY hierarchy
    col[2]  = Tier            → commercial tier (T1/T2/T3)
    col[3]  = Gender          → Principal Gender Description → drives SAP Gender + Age
    col[4]  = Material ID     → Principal Style Code  (e.g. 101.181495)
    col[5]  = Material        → Principal Style Description
    col[6]  = Color ID        → Principal Color Code  (e.g. C0513)
    col[7]  = Color           → Principal Color Description
    col[8]  = Theme ID        → Theme / franchise grouping
    col[9]  = Theme           → Theme description
    col[10] = Material Group  → SAP Product Division (Footwear → F, Apparel → A, Accessories → E)
    col[11] = Category        → SAP Product Category / Principal Hierarchy L3
    col[12] = Size USA MAIN   → Principal Size Code (range string e.g. "7- 15" / "XXS XXL")
    col[13] = UK Size         → alternate size reference
    col[14] = Made In         → Country of Origin (full name → ISO)
    col[15] = CloseOut        → YES/NO flag
    col[16] = cutoff          → delivery/cutoff window
    col[17] = BEU RRP         → Retail RRP (EUR)
    col[18] = EX$             → FOB price (USD)
    col[19] = LIC             → LIC price (used for FOB)
    col[20] = 0.25            → computed discount (LIC * 0.75) — not mapped to attribute

Article structure
  • Article Type    : "Inline"  (all rows are inline articles)
  • Article Category: Generic (1) — has size variants
  • Each unique (Material ID + Color ID) = one Generic
  • Each row is already one Generic (Material ID + Color ID is the key)
  • Variants        : derived by expanding the size-range string in col[12]
                      e.g. "7- 15" means US sizes 7 to 15 in 0.5 increments (footwear)
                      e.g. "XXS XXL" means apparel sizes XXS through XXL

Generic code  : DIA + last 6 chars of Material ID (dots removed) + full Color ID
                e.g. DIA181495C0513
Variant code  : Generic + SAP Color token (3-char) + SAP Size code (3-char)
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
import os

BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR     = BASE_DIR / "input"
PRICELIST_DIR = INPUT_DIR / "pricelist"
MDD_DIR       = INPUT_DIR / "mdd"
ATTR_DIR      = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PRICELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
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

def _parse_metadata_from_input_filename(stem: str) -> dict:
    parts = re.split(r"\s*-\s*", stem)
    season = ""
    season_idx = None
    for i, p in enumerate(parts):
        tok = p.strip()
        if re.match(r"^[A-Z]{2}\d{2,4}$", tok, re.IGNORECASE):
            season = tok.upper()
            season_idx = i
            break

    file_type_token = parts[3].strip() if len(parts) > 3 else ""
    article_type = ""
    for at in ("Inline", "License", "SSE"):
        if at.lower() in file_type_token.lower():
            article_type = at
            break

    country = ""
    if season_idx is not None:
        trailing = parts[season_idx + 1:]
        for p in trailing:
            tok = p.strip()
            if re.match(r"^[A-Z]{2,3}$", tok, re.IGNORECASE) and not tok.isdigit():
                country = tok.upper()
                break

        # fallback: if country token appears before season, capture first country-like token
        if not country:
            for i, p in enumerate(parts):
                if i in (0, 1, 2, 3, 4, season_idx):
                    continue
                tok = p.strip()
                if re.match(r"^[A-Z]{2,3}$", tok, re.IGNORECASE) and not tok.isdigit():
                    country = tok.upper()
                    break

    return {"season": season, "article_type": article_type, "country": country}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "DIA": "DIADORA",
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "LOT": "LOTTO",  "ELL": "ELLESSE",
}

LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}

LOV_AGE = {
    "AD": "Adults",   "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}

LOV_BY_AGE_ID = {
    "AD": "ADULT", "CH": "KIDS", "IN": "INFANT",
    "AA": "ALL AGES", "JR": "KIDS",
}

LOV_UOM = {"EA": "Each", "PR": "Pair", "PAA": "Pair", "SET": "Set"}

LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter",
    "AL": "All Season",    "HO": "Holiday",
    "AW": "Autumn-Winter",
}

LOV_COUNTRY_ORIGIN = {
    "CN": "China",      "VN": "Vietnam",    "ID": "Indonesia",
    "KH": "Cambodia",   "BD": "Bangladesh", "IN": "India",
    "MY": "Malaysia",   "TH": "Thailand",   "PK": "Pakistan",
    "LK": "Sri Lanka",  "IT": "Italy",      "TR": "Turkey",
    "SG": "Singapore",  "GB": "United Kingdom",
}

LOV_SBU = {
    "SP": "MAA Sport",  "FQ": "Foot Locker",
    "FF": "MAA Fashion Footwear", "CH": "MAA Children",
}

LOV_SAP_ARTICLE_CATEGORY = {
    "0": "Single", "1": "Generic", "2": "Variant",
}

# ── Diadora-specific lookup tables ────────────────────────────────

# COO full name (from pricelist "Made In") → ISO 2-char
DIA_COO_TO_ISO: dict[str, str] = {
    "CHINA":          "CN",
    "VIETNAM":        "VN",
    "INDONESIA":      "ID",
    "CAMBODIA":       "KH",
    "BANGLADESH":     "BD",
    "INDIA":          "IN",
    "MALAYSIA":       "MY",
    "THAILAND":       "TH",
    "PAKISTAN":       "PK",
    "SRI LANKA":      "LK",
    "ITALY":          "IT",
    "TURKEY":         "TR",
    "SINGAPORE":      "SG",
    "UNITED KINGDOM": "GB",
}

# Pricelist Gender field → SAP Gender code (M/F/U)
DIA_GENDER_TO_SAP: dict[str, str] = {
    "ADULT MAN":       "M",
    "ADULT WOMAN":     "F",
    "ADULT UNISEX":    "U",
    "JUNIOR UNISEX":   "U",
    "JUNIOR BOY":      "M",
    "JUNIOR GIRL":     "F",
    "YOUTH UNISEX":    "U",
    "INFANT UNISEX":   "U",
    "UNISEX UOMO-DONNA": "U",
}

# Pricelist Gender field → SAP Age code
DIA_GENDER_TO_AGE: dict[str, str] = {
    "ADULT MAN":       "AD",
    "ADULT WOMAN":     "AD",
    "ADULT UNISEX":    "AD",
    "JUNIOR UNISEX":   "CH",
    "JUNIOR BOY":      "CH",
    "JUNIOR GIRL":     "CH",
    "YOUTH UNISEX":    "CH",
    "INFANT UNISEX":   "IN",
    "UNISEX UOMO-DONNA": "AD",
}

# Material Group (col 10) → SAP Product Division letter
DIA_MATGROUP_TO_DIVISION: dict[str, str] = {
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "ACCESSORIES": "E",
}

# Collection (col 1) → SAP Product Group (2-char)
DIA_COLLECTION_TO_PRODGROUP: dict[str, str] = {
    "ACT RUNNING":          "RU",
    "ACT TENNIS":           "TE",
    "ACT SOCCER":           "SO",
    "LIFESTYLE SPORTSWEAR": "LS",
}

# Category (col 11) → SAP Product Category (2-char)
DIA_CATEGORY_TO_PRODCAT: dict[str, str] = {
    "RUNNING HIGH PERF.":              "RH",
    "RUNNING PERFORMANCE":             "RH",
    "JOGGING":                         "JG",
    "LIFESTYLE":                       "LS",
    "TENNIS HIGH PERFORM.":            "TH",
    "TENNIS PERFORMANCE":              "TP",
    "TENNIS COURT":                    "TC",
    "SHORT SLEEVE T-SHIRT":            "SS",
    "LONG SLEEVE T-SHIRT":             "LS",
    "SHORT SLEEVE POLO":               "PL",
    "POLO PIQUET JERSEY":              "PL",
    "PANTS":                           "PT",
    "SHORTS/UNDERSHORTS":              "SH",
    "BERMUDA":                         "BE",
    "TIGHTS":                          "TG",
    "3/4 TIGHTS":                      "TG",
    "HOODIE SWEATS (HOODIE+FULL ZIP HOODIE)": "HD",
    "SWEATS (CREWNECK+HALF ZIP)":      "SW",
    "LIGHT JACKET (WIND+K-WAY)":       "JK",
    "TRACKJACKET":                     "TJ",
    "PADDED VEST":                     "PV",
    "VEST":                            "VT",
    "ACTIVE BRA":                      "BR",
    "TANK/CROP TOP":                   "TK",
    "SKIRT":                           "SK",
    "COTTON SUITS":                    "CS",
    "POLYESTER SUIT":                  "PS",
    "TUTE TRIACETATO":                 "TT",
    "TRACK & FIELD":                   "TF",
    "SOCKS":                           "SK",
    "CAPS/VISORS":                     "CV",
    "HAT":                             "HT",
    "HEADBANDS/WRISTS/SCARFS":         "HW",
    "GLOVES":                          "GL",
    "MD PU (FIXED)":                   "MD",
    "MD RUBBER (FIXED)":               "MD",
    "PEBAX (FIXED)":                   "PB",
    "MPH (MISTA)":                     "MP",
    "TURF":                            "TU",
    "INDOOR":                          "IN",
    "SANDALS":                         "SA",
    "DOLCE VITA":                      "DV",
    "RUNNING OFF-ROAD":                "OR",
}

# Collection → Sports Category EN
DIA_COLLECTION_TO_SPORTS_CAT: dict[str, str] = {
    "ACT RUNNING":          "Running",
    "ACT TENNIS":           "Tennis / Padel",
    "ACT SOCCER":           "Soccer",
    "LIFESTYLE SPORTSWEAR": "Lifestyle / Casual",
}


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
        wb   = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
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
                "cardinality":  _v("Cardinality \n(Conditional / Mandatory / Optional)") or _v("Cardinality"),
                "validation":   _v("Validation Base Type"),
                "multi_valued": _v("Multi Valued"),
                "lov_name":     _v("Name of LOV\n(Only if Validation Base Type = LOV)") or _v("Name of LOV"),
                "max_chars":    _v("Max Characters\n<Only if Validation Type = Text / Numeric Text / URL>") or _v("Max Characters"),
                "group":        _v("PIM Attribute Group\n(As Configured in STEP)") or _v("PIM Attribute Group"),
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
                raw_id = str(val_id).strip() if val_id else str(val_name).strip()
                if raw_id.isdigit():
                    raw_id = raw_id.zfill(3)
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = raw_id

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper() and sn not in (
                "Brand LOV", "SBU LOV", "Company Code LOV",
                "Gender LOV", "Age LOV", "Season LOV",
                "Country Origin LOV", "Nature of Article LOV",
                "Sports Category LOV", "Sports Categor LOV", "BY Article Type",
                "Retail Price Currency LOV", "Article Category LOV",
            ):
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                
                # Special handling for Sports Category/Categor LOV: columns are reversed
                # Column A = Label/Name, Column B = ID/Code
                if sn in ("Sports Category LOV", "Sports Categor LOV"):
                    name, code = row[0], row[1]  # Reversed!
                else:
                    code, name = row[0], row[1]  # Standard order
                
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )


class AttributesListLoader:
    """Loads the DIADORA tab from the Attributes List workbook."""

    def __init__(self, path: Path, brand: str = "DIADORA"):
        self.path      = path
        self.brand     = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames
             if s.upper() == self.brand and "V4" not in s.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames if self.brand in s.upper()),
            wb.sheetnames[0],
        )

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"),
            None,
        )
        if hdr_idx is None:
            log.warning("[AttrList] Cannot find header in sheet '%s'", sheet_name)
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
        log.info(
            "[AttrList] %d attributes loaded from sheet '%s'",
            len(self.attr_map), sheet_name,
        )

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class DiadoraPricelistLoader:
    """
    Loads the Diadora inline pricelist.

    Active sheet  : "Analisi LineList"
    Header row    : index 0 (0-based)
    Data rows     : index 1+ (0-based)
    All rows are inline — no channel-type filter needed.
    Drops rows where Material ID is blank.
    """

    SHEET_NAME = "Analisi LineList"

    def __init__(self, path: Path):
        self.path  = path
        self.rows: list[dict] = []
        self._load()

    def _load(self):
        log.info("[Pricelist-DIA] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in wb.sheetnames if s.upper() == self.SHEET_NAME.upper()),
            None,
        ) or wb.sheetnames[0]

        log.info("[Pricelist-DIA] Using sheet: '%s'", target)
        ws      = wb[target]
        all_rows = list(ws.iter_rows(values_only=True))

        if not all_rows:
            log.error("[Pricelist-DIA] Empty sheet '%s'", target)
            wb.close()
            return

        # Row 0 is the header
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(all_rows[0])
        ]

        valid = 0
        for raw in all_rows[1:]:
            if len(raw) < 5:
                continue
            material_id = raw[4]
            if material_id is None:
                continue
            sid = str(material_id).strip()
            if not sid or sid in ("None", "nan", ""):
                continue
            row_dict = {header[i]: raw[i] for i in range(len(header))}
            self.rows.append(row_dict)
            valid += 1

        wb.close()
        log.info("[Pricelist-DIA] %d valid data rows loaded", valid)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — HELPERS
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "N/A", "n/a") else s


def _fmt_date(v) -> str:
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def _price(v) -> str:
    if v is None:
        return ""
    try:
        f = float(v)
        return "" if f == 0.0 else str(round(f, 2))
    except (TypeError, ValueError):
        m = re.search(r"[\d.]+", str(v))
        return m.group() if m else str(v)


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — MAPPER
# ══════════════════════════════════════════════════════════════════

def map_article_diadora(row: dict, brand_code: str = "DIA") -> dict:
    """
    Map one Diadora pricelist row → unified article dict.

    Column names from "Analisi LineList" header row:
        CAT, Collection, Tier, Gender, Material ID, Material,
        Color ID, Color, Theme ID, Theme, Material Group, Category,
        Size USA MAIN, UK Size, Made In, CloseOut, cutoff,
        BEU RRP, EX$, LIC, 0.25
    """
    cat           = _s(row.get("CAT"))
    collection    = _s(row.get("Collection"))
    tier          = _s(row.get("Tier ") or row.get("Tier"))
    gender_raw    = _s(row.get("Gender"))
    material_id   = _s(row.get("Material ID"))
    material      = _s(row.get("Material"))
    color_id      = _s(row.get("Color ID"))
    color_desc    = _s(row.get("Color"))
    theme_id      = _s(row.get("Theme ID"))
    theme         = _s(row.get("Theme"))
    material_group = _s(row.get("Material Group"))
    category      = _s(row.get("Category"))
    # Handle multiline column name
    size_usa_main = _s(
        row.get("Size USA \nMAIN ") or
        row.get("Size USA MAIN ") or
        row.get("Size USA \nMAIN") or
        row.get("Size USA MAIN")
    )
    uk_size       = _s(row.get("UK Size"))
    made_in       = _s(row.get("Made In"))
    closeout      = _s(row.get("CloseOut"))
    cutoff        = _s(row.get("cutoff"))
    beu_rrp       = row.get("BEU RRP")
    ex_dollar     = row.get("EX$")
    pct_25        = row.get("0.25") or row.get("25%")

    # ── Derived fields ──────────────────────────────────────────

    gender_upper  = gender_raw.upper()
    sap_gender    = DIA_GENDER_TO_SAP.get(gender_upper, "U")
    sap_age       = DIA_GENDER_TO_AGE.get(gender_upper, "AD")

    # SAP Product Division from Material Group
    div_letter    = DIA_MATGROUP_TO_DIVISION.get(material_group.upper(), "F")

    # SAP Product Group from Collection
    prod_group    = DIA_COLLECTION_TO_PRODGROUP.get(collection.upper(), "OT")

    # SAP Product Category from Category
    prod_category = DIA_CATEGORY_TO_PRODCAT.get(category.upper(), "OT")

    # COO → ISO
    coo_iso       = DIA_COO_TO_ISO.get(made_in.upper(), "CN")

    # Sports Category
    sports_cat    = DIA_COLLECTION_TO_SPORTS_CAT.get(collection.upper(), "Other")

    # Nature of Article — CloseOut = YES → C/O, else REG
    noa_code      = "C/O" if closeout.upper() == "YES" else "REG"

    # Prices
    rrp           = _price(beu_rrp)    # BEU RRP in EUR
    fob           = _price(pct_25)     # 25% column as FOB

    # SAP Style Code: Material ID with dots removed, up to 9 chars
    sap_style_code = re.sub(r"\.", "", material_id)[:9]

    # Generic code: brand_code (3) + last 6 chars of style + full color_id.
    generic_code  = f"{brand_code}{sap_style_code[-6:]}{color_id}"

    # Article type = always Inline
    article_type  = "Inline"

    # UOM: Footwear → PAA (Pair), others → EA
    uom_code      = "PAA" if material_group.upper() == "FOOTWEAR" else "EA"

    return {
        # Identifiers
        "material_id":     material_id,       # e.g. 101.181495
        "sap_style_code":  sap_style_code,    # e.g. 101181495
        "generic_code":    generic_code,      # e.g. DIA181495513
        "color_id":        color_id,          # e.g. C0513

        # Descriptions
        "material_name":   material,          # style description
        "color_desc":      color_desc,        # color description
        "gender_raw":      gender_raw,        # original gender field

        # Taxonomy
        "cat":             cat,
        "collection":      collection,
        "tier":            tier,
        "material_group":  material_group,
        "category":        category,
        "theme_id":        theme_id,
        "theme":           theme,

        # SAP codes
        "sap_gender":      sap_gender,
        "sap_age":         sap_age,
        "div_letter":      div_letter,
        "prod_group":      prod_group,
        "prod_category":   prod_category,
        "sports_cat":      sports_cat,

        # Commercial
        "article_type":    article_type,
        "art_category":    "1",              # Generic
        "noa_code":        noa_code,
        "uom_code":        uom_code,

        # Origin
        "coo":             coo_iso,
        "made_in":         made_in,

        # Prices
        "rrp":             rrp,             # BEU RRP EUR
        "fob":             fob,             # LIC price USD

        # Misc
        "brand_code":      brand_code,
        "closeout":        closeout,
        "cutoff":          cutoff,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Validate mandatory fields against MDD."""
    warns = []
    art   = mapped["material_id"]
    checks = {
        "sap_style_code": "AT_PrincipalStyleCode",
        "sap_gender":     "AT_Gender",
        "sap_age":        "AT_SAPAge",
        "brand_code":     "AT_Brand",
        "material_group": "AT_SAPProductDivision",
    }
    for field, at_id in checks.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id} (field: {field})")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(
    parent:  ET.Element,
    attr_id: str,
    value:   str = "",
    id_val:  str = "",
) -> ET.Element | None:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip()  if id_val  else ""
    clean_val = str(value).strip()   if value   else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v  = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _add_generic_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd:        MDDLoader | None = None,
) -> None:
    """Write all Generic-level <Value> elements for a Diadora inline article."""
    written: set[str] = set()

    def _w(attr_id, value="", id_val=""):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val)
            written.add(attr_id)

    def _mw(attr_id, id_val):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    _mw("AT_SBU",         sbu)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    b_code  = art["brand_code"]
    b_label = LOV_BRAND.get(b_code, b_code)
    _w("AT_Brand",      id_val=b_code)
    _w("AT_BrandGroup", id_val=b_label)

    # ── Principal identifiers ────────────────────────────────────
    # AT_PrincipalStyleCode: Stibo max 9 chars — use dot-free SAP style code
    _w("AT_PrincipalStyleCode",        art["sap_style_code"])
    _w("AT_PrincipalStyleDescription", art["material_name"])
    _w("AT_PrincipalColorCode",        art["color_id"])
    # AT_PrincipalColorDescription not configured in Stibo — omitted
    _w("AT_PrincipalGenderDescription", art["gender_raw"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
    _w("AT_InboundGenericCode",                   art["generic_code"])

    # ── Color ────────────────────────────────────────────────────
    # Look up color ID from MDD Color Code LOV only — do not fall back
    # to derived token since unmapped values are rejected as illegal LOV
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id = (
            color_lov.get(art["color_desc"], "") or
            color_lov.get(art["color_desc"].upper(), "") or
            ""
        )
    # ── Color Description (Pricelist: Color → AT_PrincipalColorName) ──────────
    # AT_Color was WRONG — correct AttributeID is AT_PrincipalColorName (MDD confirmed)
    _w("AT_PrincipalColorName", art["color_desc"])

    # ── Gender ───────────────────────────────────────────────────
    g_code = art["sap_gender"]
    _w("AT_Gender",   id_val=g_code)
    _w("AT_BYGender", id_val=g_code)

    # ── Age ──────────────────────────────────────────────────────
    age_code  = art["sap_age"]
    by_age_id = LOV_BY_AGE_ID.get(age_code, "ADULT")
    _w("AT_SAPAge",          id_val=age_code)
    _w("AT_BYAge",           id_val=by_age_id)
    _w("AT_PrincipalAgeDescription", LOV_AGE.get(age_code, age_code))

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season_from_filename") or art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    _w("AT_Season", id_val=sea_code)
    year_match = re.search(r"20\d{2}", sea_raw)
    if year_match:
        sea_year_4 = year_match.group()
    else:
        year_2_match = re.search(r"([A-Z]{2})(\d{2})$", sea_raw.upper())
        sea_year_4 = f"20{year_2_match.group(2)}" if year_2_match else ""
    _w("AT_SeasonYear", sea_year_4)

    # ── Country of Origin ────────────────────────────────────────
    coo_code = art["coo"]

    # ── SAP Article Category — Generic (1) ──────────────────────
    _w("AT_SAPArticleCategory", id_val="1")

    # ── BY / SAP Article Type ────────────────────────────────────
    _w("AT_BYArticleType", id_val=art.get("article_type_from_filename") or art.get("article_type") or "Inline")
    country_code = (art.get("country_from_filename", "") or "").strip().upper()
    country_label = LOV_COUNTRY_ORIGIN.get(country_code, country_code)
    country_lov = mdd.lovs.get("Country", {}) if mdd else {}
    country_id = country_lov.get(country_label) or country_lov.get(country_code) or country_code
    _w("AT_Country", id_val=country_id)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   id_val="Y")
    _w("AT_SAPIndicator",  id_val="N")
    # _w("AT_EcomIndicator", "")
    _w("AT_SAPProductFlag", id_val="A")

    # ── UOM ──────────────────────────────────────────────────────
    _w("AT_UOM", id_val=art["uom_code"])

    # ── Pricing ──────────────────────────────────────────────────
    # AT_OriginalPrice / AT_CurrentPrice TEMPORARILY COMMENTED OUT:
    # BEU RRP has no documented mapping row in the DIADORA Attribute List.
    # Needs MD team confirmation before re-enabling.
    # _w("AT_OriginalPrice",       art["rrp"])   # BEU RRP
    # _w("AT_CurrentPrice",        art["rrp"])
    # RRP is in EUR for Diadora
    currency_lov = mdd.lovs.get("Retail Price Currency", {}) if mdd else {}
    eur_id = currency_lov.get("European Euro", "EUR")
    _w("AT_RetailPriceCurrency", id_val=eur_id)
    _w("AT_FOB",                 art["fob"])   # 25% column
    usd_id = currency_lov.get("United States Dollar", "USD")
    _w("AT_FOBCurrency",         id_val=usd_id)

    # ── Nature of Article ────────────────────────────────────────
    # AT_NatureOfArticle TEMPORARILY COMMENTED OUT:
    # NOA (Row 105) mapping logic is BLANK in the DIADORA Attribute List.
    # CloseOut → NOA link is undocumented. Needs MD team confirmation.
    # _w("AT_NatureOfArticle", id_val=art["noa_code"])

    # ── Product Hierarchy / Classification (TEMPORARILY COMMENTED OUT AS THEY DO NOT EXIST IN STIBO) ──
    # _w("AT_SAPProductDivision",  art["div_letter"])
    # _w("AT_SAPProductGroup",     art["prod_group"])
    # _w("AT_SAPProductCategory",  art["prod_category"])
    
    
    _w("AT_PrincipalMerchandiseHierarchyL1", art["material_group"])   # Material Group
    _w("AT_PrincipalMerchandiseHierarchyL2", art["collection"])        # Collection
    _w("AT_PrincipalMerchandiseHierarchyL3", art["category"])          # Category
    # _w("AT_PrincipalMerchandiseHierarchyL4", "")                       # N/A for Inline
    # _w("AT_PrincipalMerchandiseHierarchyL5", "")                       # N/A for Inline

    # ── Sports Category ──────────────────────────────────────────
    # Map Collection → Sports Category LOV ID
    sports_cat_label = art["sports_cat"]  # Already mapped from Collection via DIA_COLLECTION_TO_SPORTS_CAT
    sports_cat_id = ""
    if mdd and sports_cat_label and sports_cat_label != "Other":
        sports_lov = mdd.lovs.get("Sports Category", {})
        # Try exact match first
        sports_cat_id = sports_lov.get(sports_cat_label, "")
        
        # If not found, try case-insensitive lookup
        if not sports_cat_id:
            for key, val in sports_lov.items():
                if key.lower() == sports_cat_label.lower():
                    sports_cat_id = val
                    break
        
        # Zero-pad 1-digit IDs to 2 digits (e.g. "5" → "05")
        if sports_cat_id and sports_cat_id.isdigit() and len(sports_cat_id) == 1:
            sports_cat_id = f"0{sports_cat_id}"
    
    if sports_cat_id:
        _w("AT_SportsCategoryEN", id_val=sports_cat_id)

    # ── Country Size — REMOVED: manual input from MD team ───────
    # _w("AT_CountrySize", "US")

    # ── Material Type — default ZINA for Diadora Inline ─────────
    _w("AT_MaterialType", id_val="ZINA")



# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    full_season    = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_season}"
    batches_parent = f"CLH_{brand.title().replace(' ', '')}Batches"

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
    ET.SubElement(season_cls,    f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed,    f"{{{STIBO_NS}}}Name").text = (
        f"{season_short} Confirmed Articles"
    )

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed,    f"{{{STIBO_NS}}}Name").text = (
        f"{season_short} Unconfirmed Articles"
    )

    return cls_root


def build_product_xml(
    art:        dict,
    brand:      str,
    brand_code: str,
    comp_code:  str,
    sbu:        str,
    season_id:  str,
    mdd:        MDDLoader | None = None,
) -> str:
    """
    Build a <Product> XML fragment for one Diadora inline article (Generic only).

    Structure:
      Product (Generic, UserTypeID = PRD_GenericArticle)

    ParentID for Generic  : PPH_{div_letter}-TempSubCat
    KeyID for Generic     : KEY_InboundArticle  → generic_code

    ClassificationReferences on Generic:
        CPL_Merchandiser         → CLH_{Brand}Articles
        CPL_UnConfirmedForSeason → {season_id}UA
    """
    material_id = art["material_id"]
    if not material_id:
        return ""

    div_letter = art.get("div_letter", "F")
    parent_id  = f"PPH_{div_letter}-TempSubCat"
    key_generic = art["generic_code"]

    # ── Generic Product element ──────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["material_name"] or material_id
    )

    # ── Classification references ────────────────────────────────
    brand_title = brand.title().replace(" ", "")
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand_title}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Generic-level Values ─────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one Diadora row."""
    row, brand_code, mdd, season = row_tuple
    mapped = map_article_diadora(row, brand_code=brand_code)
    mapped["season"] = season
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

    args must have:
        brand       str   e.g. "Diadora"
        brand_code  str   e.g. "DIA"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SP"
        season      str   e.g. "FW26"
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f     = first(MDD_DIR)
    attr_f    = first(ATTR_DIR)
    pl_files  = list(PRICELIST_DIR.glob("*.xlsx"))

    # ── Mandatory file checks ────────────────────────────────────
    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Pricelist", pl_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise FileNotFoundError(f"No {label} file found in expected input directory.")

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand="DIADORA")

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Season ID ────────────────────────────────────────────────
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for pl_path in pl_files:
        log.info("─── Processing Pricelist: %s ───", pl_path.name)
        filename_meta = _parse_metadata_from_input_filename(pl_path.stem)

        pl = DiadoraPricelistLoader(pl_path)
        if not pl.rows:
            log.warning("[Pricelist-DIA] No rows loaded — skipping.")
            continue

        rows = pl.rows   # TEST MODE: limit to first 5 rows
        total_rows = len(rows)
        log.info("Total inline rows (limited to 5 for testing): %d", total_rows)

        # ── Output filename ──────────────────────────────────────
        _stem  = pl_path.stem
        _parts = re.split(r"\s*-\s*", _stem)
        file_type_from_input  = _parts[3] if len(_parts) >= 7 else "Pricelist"
        multi_mono_from_input = _parts[4] if len(_parts) >= 7 else "Multi"

        out_name = f"{pl_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (row, args.brand_code, mdd, args.season) for row in rows
        ]
        ordered: list[tuple[int, dict, list]] = []

        log.info(
            "Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
            total_rows, num_workers,
        )

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {
                pool.submit(_process_article, t): i
                for i, t in enumerate(task_args)
            }
            for fut in as_completed(futures):
                idx           = futures[fut]
                mapped, warns = fut.result()
                all_warnings.extend(warns)
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered]
        for _m in mapped_articles:
            _m["article_type_from_filename"] = filename_meta.get("article_type", "")
            _m["season_from_filename"] = filename_meta.get("season", "")
            # Diadora requirement: AT_Country ID/text must come from input filename token.
            _m["country_from_filename"] = (filename_meta.get("country", "") or "").strip().upper()
        del ordered

        log.info("Pass 1 done — mapped=%d articles", len(mapped_articles))

        # ── Pass 2: stream XML to file ───────────────────────────
        log.info("Pass 2/2 — streaming XML → %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cls_el  = build_classifications(
            args.brand, args.brand_code, args.comp_code, args.sbu, args.season,
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
                if not art.get("material_id"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, args.brand_code,
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

        print("═══ ARTICLE SUMMARY (DIADORA INLINE) ══════════════", flush=True)
        print(f"  Pricelist rows   : {total_rows}",    flush=True)
        print(f"  Generics written : {written_count}",  flush=True)
        print(f"  XML file size    : {file_kb}KB",      flush=True)
        print("════════════════════════════════════════════════════", flush=True)

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
# SECTION 9 — CLI (local testing)
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — Diadora Inline v1.0"
    )
    p.add_argument("--brand",      default="Diadora")
    p.add_argument("--brand-code", default="DIA")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="FW26")
    p.add_argument("--country-code", default="")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
