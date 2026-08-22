"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — Smiggle  v1.0               ║
║  Ecommerce File (Master Asset / Online Content) → STEP XML     ║
╚══════════════════════════════════════════════════════════════════╝

Smiggle Ecommerce File specifics:
  • Source file   : Ecommerce / Master Asset file (e.g. 27MARCH_W26.xlsx)
  • Active sheet  : "Sheet1"
  • Header row    : row index 0 (0-based)
  • Rows with LINE NUMBER = "#N/A" or blank are skipped.

  Column → STEP Attribute mapping:
  ─────────────────────────────────────────────────────────────────
  MUST BE UPDATED        → composite key (LINE#+COLOUR) — internal only
  LINE NUMBER            → AT_PrincipalStyleCode / article_no
  DEPARTMENT             → SAP Product Division  (AT_SAPProductDivision)
  SUBCAT                 → SAP Product Category / AT_Interest / AT_Material
  LINE DESCRIPTION       → AT_PrincipalStyleDescription (principal code-style desc)
  COLOUR                 → AT_Color / AT_PrincipalColorCode
  PRODUCT COLLECTION NAME→ AT_Collection1
  ONLINE PRODUCT NAME    → AT_EComProductNameEN  (English ecom name)
  PRODUCT DESCRIPTION    → AT_ShortDescriptionEN (short desc EN — used when
                           ONLINE PRODUCT NAME is blank)
  ONLINE PRODUCT DESCRIPTION → AT_LongDescriptionEN  (long desc EN)
  FEATURES & BENEFITS    → AT_LongDescriptionEN  (appended when ONLINE PRODUCT
                           DESCRIPTION is blank; otherwise standalone features)
  DIMENSIONS             → AT_ProductLengthWidthHeight
  CARE FOR ME/SAFETY     → AT_CareInstructionEN
  FABRIC COMPOSITION     → AT_Fabric / AT_Material (supplementary)

  Translation fields (ID, PH, MY, VN, TH, KH) are emitted as empty
  placeholders — AI translation happens downstream in Stibo.

  Article Category = always Single (00) — Smiggle accessories, no sizing.
  SAP Size         = always "000" / No Size.
  COO              = default CN (China) — no port column in ecom file.
  Season           = derived from pipeline args (filename-parsed by lambda).
  Gender           = derived from colour keyword heuristic (default U).
  Age              = default CH (Children) — Smiggle kids brand.

  One Generic product per (LINE NUMBER / COLOUR) row.

S3 Layout:
    raw/metadata/smiggle/{ecommerce_file}.xlsx  ← Smiggle uploads here
    raw/metadata/mdd_file.xlsx                  ← shared global MDD
    raw/metadata/attributes_list.xlsx           ← shared global attr list

Output:
    s3://map-stibo-inbound-processed-dev/
        processed/stepxml/smiggle/{filename}.xml
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

INPUT_DIR    = BASE_DIR / "input"
ECOMMERCE_DIR = INPUT_DIR / "ecommerce"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [ECOMMERCE_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_ecommerce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}
LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior",
}
LOV_BY_AGE = {
    "ADULT": "Adult", "ADULTS": "Adult", "AD": "Adult",
    "JUNIOR": "Junior", "CH": "Children",
    "CHILDREN": "Children", "CHILD": "Child",
    "ALL AGES": "All Ages", "AA": "All Ages",
    "CORE": "Children", "TEENY TINY": "Children",
}
LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set", "PK": "Pack"}
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

# ── Smiggle-specific lookup tables ───────────────────────────────

# Department → SAP Product Division letter
SMIGGLE_DEPT_TO_DIVISION: dict[str, str] = {
    "PACK & CARRY": "E",
    "STATIONERY":   "E",
    "ACCESSORIES":  "E",
}

# Category / SUBCAT → SAP Product Group code (2-char)
SMIGGLE_CATEGORY_TO_PROD_GROUP: dict[str, str] = {
    "BAGS":               "BG",
    "BACKPACKS":          "BG",
    "PRESCHOOL BACKPACKS":"BG",
    "TROLLEY BAGS":       "BG",
    "LUNCHBAGS":          "FD",
    "DRINK":              "DK",
    "PLASTIC BOTTLES":    "DK",
    "STEEL BOTTLES":      "DK",
    "FOOD":               "FD",
    "PENCIL CASE":        "PC",
    "PENCIL CASE SOFT":   "PC",
    "PENCIL CASE HARDTOP":"PC",
    "PENCIL CASE POP OUT":"PC",
    "WRITE SINGLES":      "WS",
    "WRITE PACKS":        "WP",
    "DESKTOP":            "DS",
    "STATIONERY":         "ST",
    "ACCESSORIES":        "AC",
    "FASHION ACCESSORIES":"FA",
    "ACTIVITY":           "AV",
    "ACTIVITY BOOKS":     "AV",
    "KITS":               "KT",
    "KITS ART & CRAFT":   "KT",
    "KITS ESSENTIALS":    "KT",
    "KITS LICENSED":      "KT",
    "KITS MIDI":          "KT",
    "KITS OTHER":         "KT",
    "KITS POPOUT":        "KT",
    "GAMES":              "GM",
    "BOOKS":              "BK",
    "CONSUMABLES":        "CO",
    "LIGHTS":             "LT",
    "TIME":               "TM",
    "WATCHES":            "TM",
    "CLOCKS":             "TM",
    "WALLETS AND PURSES": "AC",
    "HEADPHONES":         "AC",
    "JEWELLERY":          "FA",
    "WRISTBANDS":         "FA",
    "LANYARDS":           "AC",
    "NOTEBOOKS":          "ST",
    "NOVELTY KEYRINGS":   "AC",
    "ALPHA KEYRINGS":     "AC",
    "MINI COLLECT KEYRINGS":"AC",
    "PEN PALS":           "WS",
    "PEN SINGLES":        "WS",
    "PEN PACKS":          "WP",
    "PEN NOVELTY":        "WS",
    "PENCIL SINGLES":     "WS",
    "PENCIL PACKS":       "WP",
    "HIGHLIGHTER PACKS":  "WP",
    "MARKER SINGLES":     "WS",
    "MARKER PACKS":       "WP",
    "ERASER SINGLES":     "WS",
    "ERASER PACKS":       "WP",
    "STICKERS":           "AC",
    "STORAGE":            "AC",
    "TRAVEL ACCESSORIES": "AC",
    "REUSE ME BAGS":      "BG",
    "MYSTERY BAGS":       "BG",
    "BENTO":              "FD",
    "CONTAINERS":         "FD",
    "BATH AND BEAUTY":    "AC",
    "GIGGLE":             "AC",
    "SPARE PARTS":        "AC",
    "TECH OTHER":         "AC",
    "ARTS & CRAFT PACK":  "AV",
    "ADVENT":             "KT",
    "EBOOKS":             "BK",
    "FOOD OTHER":         "FD",
    "DRINK OTHER":        "DK",
    "BAGS OTHER":         "BG",
}


def _colour_to_gender(colour: str) -> str:
    """Heuristic: derive SAP Gender code from Smiggle colour string."""
    c = (colour or "").upper()
    if any(k in c for k in ("PINK", "LILAC", "ROSE", "PURPLE")):
        return "F"
    return "U"


def _parse_dimensions(dim_str: str) -> dict[str, str]:
    """
    Parse a Smiggle dimensions string into L / W / H components.

    Input examples:
        "H 15cm x W 24.5cm x D 15.5cm"
        "W 22cm x H 12cm x D 5cm"
        "30cm x 20cm x 10cm"
        "L30 x W20 x H10"

    Returns dict with keys "length", "width", "height" (strings, cm values).
    Falls back to the raw string in "raw" if not parseable.
    """
    if not dim_str:
        return {}
    raw = dim_str.strip()
    result: dict[str, str] = {}

    # Pattern: optional label (H/W/D/L) followed by number and optional cm
    pattern = re.compile(
        r"""
        (?:
            ([HhWwDdLl])\s*       # optional dimension label
        )?
        ([\d]+(?:\.\d+)?)         # number
        \s*(?:cm)?                # optional cm unit
        """,
        re.VERBOSE,
    )
    matches = pattern.findall(raw)
    # Filter out empty matches
    matches = [(lbl.upper(), val) for lbl, val in matches if val]

    # Map labels to L/W/H
    label_map = {"L": "length", "W": "width", "H": "height", "D": "depth"}
    unlabelled = []
    for lbl, val in matches:
        if lbl in label_map:
            key = label_map[lbl]
            if key == "depth":
                key = "height"   # treat Depth as Height for most bags
            result[key] = val
        else:
            unlabelled.append(val)

    # If no labels, fall back to positional L / W / H
    if not result and unlabelled:
        keys = ["length", "width", "height"]
        for i, v in enumerate(unlabelled[:3]):
            result[keys[i]] = v

    result["raw"] = raw
    return result


def _clean_multiline(text: str) -> str:
    """
    Normalise multi-line cell content for XML.
    Converts \\n to space-separated sentences; collapses whitespace.
    """
    if not text:
        return ""
    lines = [ln.strip() for ln in str(text).splitlines() if ln.strip()]
    return " | ".join(lines)


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
                "cardinality":  _v("Cardinality"),
                "validation":   _v("Validation Base Type"),
                "lov_name":     _v("Name of LOV"),
                "group":        _v("PIM Attribute Group"),
            }

        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        self._load_age_lov(wb) 
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_age_lov(self, wb):
        """
        Load Age LOV: col A = display name, col C = SAP Age Code, col F = BY Age value.
        Stores:
        self.lovs["SAPAge"] = { "Children": "K", "Adults": "A", "All Ages": "O" }
        self.lovs["BYAge"]  = { "Children": "Kids", "Adults": "Adult", ... }
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None
        )
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
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
        log.info("[MDD] Age LOV loaded — SAPAge: %d, BYAge: %d",
                len(sap_age_map), len(by_age_map))


    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = str(val_id).strip() if val_id else str(val_name).strip()

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
    """Loads the SMIGGLE tab from the Attributes List workbook."""

    def __init__(self, path: Path, brand: str = "SMIGGLE"):
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
        log.info("[AttrList] %d attributes loaded from '%s'", len(self.attr_map), sheet_name)

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class SmiggleEcommerceLoader:
    """
    Loads the Smiggle ecommerce / master asset workbook.

    Active sheet  : "Sheet1" (preferred) — fallback to largest sheet.
    Header row    : row 0 (0-based). No blank rows above data.

    Columns (0-indexed):
        0   MUST BE UPDATED    — composite LINE#+COLOUR key (internal use)
        1   LINE NUMBER        — article line number (maps to article_no)
        2   DEPARTMENT         — SAP Product Division source
        3   SUBCAT             — SAP Product Category / sub-category
        4   LINE DESCRIPTION   — principal style description
        5   COLOUR             — colour (free-text)
        6   PRODUCT COLLECTION NAME — Collection 1
        7   ONLINE PRODUCT NAME     — AT_EComProductNameEN
        8   PRODUCT DESCRIPTION     — AT_ShortDescriptionEN (fallback name)
        9   ONLINE PRODUCT DESCRIPTION → AT_LongDescriptionEN
        10  FEATURES & BENEFITS     — appended to long desc / standalone
        11  DIMENSIONS              — AT_ProductLengthWidthHeight
        12  CARE FOR ME/SAFETY      — AT_CareInstructionEN
        13  FABRIC COMPOSITION      — AT_Fabric / supplementary material

    Rows where LINE NUMBER is "#N/A", blank, or "None" are skipped.
    """

    PREFERRED_SHEETS = ["Sheet1", "Ecommerce", "Online Content", "Master Asset"]

    # Invalid line number sentinel values
    SKIP_VALUES = {"#N/A", "#REF!", "#VALUE!", "#NAME?", "", "NONE", "NAN", "LINE NUMBER"}

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[Ecommerce-SMI] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in self.PREFERRED_SHEETS if s in wb.sheetnames), None
        )
        if target is None:
            best, best_rows = wb.sheetnames[0], 0
            for sn in wb.sheetnames:
                n = wb[sn].max_row or 0
                if n > best_rows:
                    best, best_rows = sn, n
            target = best

        log.info("[Ecommerce-SMI] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            log.error("[Ecommerce-SMI] Sheet is empty")
            wb.close()
            return

        # Detect header row: first row containing "LINE NUMBER" or "LINE"
        hdr_idx = None
        for i, r in enumerate(rows[:5]):
            if any(isinstance(v, str) and "LINE" in v.upper() for v in r if v):
                hdr_idx = i
                break
        if hdr_idx is None:
            hdr_idx = 0

        log.info("[Ecommerce-SMI] Header at row %d (0-indexed)", hdr_idx)
        raw_header = rows[hdr_idx]
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(raw_header)
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Normalise column name for LINE NUMBER (may have trailing space)
        line_col = next(
            (c for c in df.columns if "LINE" in c.upper() and "NUMBER" in c.upper()),
            None,
        ) or next(
            (c for c in df.columns if c.strip().upper() == "LINE NUMBER"),
            None,
        )

        if line_col:
            # Skip invalid / formula-error rows
            df = df[
                df[line_col].notna()
                & (~df[line_col].astype(str).str.strip().str.upper().isin(self.SKIP_VALUES))
            ]
            df = df.rename(columns={line_col: "LINE NUMBER"})

        # Normalise other column names (strip whitespace)
        df.columns = [c.strip() for c in df.columns]

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info(
            "[Ecommerce-SMI] %d rows loaded (invalid LINE# skipped)", len(self.df)
        )


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT", "#N/A", "#REF!", "#VALUE!") else s


def _normalise_header(v) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(v or "").upper())


def _row_value(row, *headers) -> str:
    """Fetch a row value by exact or normalised header alias."""
    keys = list(row.keys()) if hasattr(row, "keys") else []
    blank = ""

    for header in headers:
        if header in keys:
            val = row.get(header)
            if _s(val):
                return val
            blank = val

    wanted = {_normalise_header(h) for h in headers}
    for key in keys:
        if _normalise_header(key) in wanted:
            val = row.get(key)
            if _s(val):
                return val
            blank = val

    return blank


def _fmt_date(v) -> str:
    """Format a date to dd-Mon-YYYY lowercase (e.g. 15-may-2026)."""
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


def map_article_ecommerce(ec_row: dict, brand_code: str = "SMI") -> dict:
    """
    Map one Smiggle ecommerce row → unified article dict.

    Key decisions vs linelist / packing list mappers:
      - article_no        = LINE NUMBER
      - model_name        = LINE DESCRIPTION (principal style description)
      - colour            = COLOUR (free-text)
      - dept              = DEPARTMENT  → division letter
      - subcat            = SUBCAT      → product group / category
      - collection        = PRODUCT COLLECTION NAME → AT_Collection1
      - ecom_name_en      = "Smiggle " + ONLINE PRODUCT NAME + " - " + Generic Article
                            (if ONLINE PRODUCT NAME is blank, use PRODUCT DESCRIPTION)
      - short_desc_en     = PRODUCT DESCRIPTION (or ONLINE PRODUCT NAME if no desc)
      - long_desc_en      = ONLINE PRODUCT DESCRIPTION
                            (if blank, falls back to FEATURES & BENEFITS)
      - features          = FEATURES & BENEFITS (supplementary)
      - dimensions        = DIMENSIONS  → AT_ProductLengthWidthHeight
      - care_en           = CARE FOR ME/SAFETY → AT_CareInstructionEN
      - fabric            = FABRIC COMPOSITION → AT_Fabric
      - gender_code       = derived from COLOUR heuristic (default U)
      - age_code          = default CH (Children — Smiggle kids brand)
      - coo               = default CN (no port column in ecom file)
      - art_category      = always "0" (Single)
    """
    line_no   = _s(ec_row.get("LINE NUMBER"))
    if line_no.endswith(".0"):
        line_no = line_no[:-2]

    key_val   = _s(ec_row.get("MUST BE UPDATED"))
    line_desc = _s(ec_row.get("LINE DESCRIPTION"))
    colour    = _s(ec_row.get("COLOUR"))
    dept      = _s(ec_row.get("DEPARTMENT"))
    subcat    = _s(ec_row.get("SUBCAT"))
    collection = _s(ec_row.get("PRODUCT COLLECTION NAME"))

    online_name    = _s(_row_value(ec_row, "ONLINE PRODUCT NAME"))
    prod_desc      = _s(_row_value(
        ec_row,
        "PRODUCT DESCRIPTION (PRINCIPAL)",
        "PRODUCT DESCRIPTION PRINCIPAL",
        "PRODUCT DESCRIPTION",
    ))
    online_desc    = _s(_row_value(ec_row, "ONLINE PRODUCT DESCRIPTION"))
    features_raw   = ec_row.get("FEATURES & BENEFITS")
    dimensions_raw = ec_row.get("DIMENSIONS")
    care_raw       = ec_row.get("CARE FOR ME/SAFETY")
    fabric_raw     = ec_row.get("FABRIC COMPOSITION")

    # ── Content field processing ─────────────────────────────────
    # Features: clean multiline
    features = _clean_multiline(_s(features_raw)) if features_raw else ""

    # Dimensions: store raw for the composite field
    dims_raw  = _s(dimensions_raw) if dimensions_raw else ""
    dims      = dims_raw.replace("\n", " ").strip()
    dims_parsed = _parse_dimensions(dims)

    # Care instruction: clean multiline
    care_en = _clean_multiline(_s(care_raw)) if care_raw else ""

    # Fabric: keep newlines collapsed
    fabric  = _clean_multiline(_s(fabric_raw)) if fabric_raw else ""

    # Generic / variant codes
    generic_code  = f"{brand_code}{line_no}"
    colour_token  = re.sub(r"[^A-Z0-9]", "", colour.upper())[:3].ljust(3, "0")
    variant_code  = f"{generic_code}{colour_token}000"

    # ── Ecom product name EN ─────────────────────────────────────
    # Rule: Smiggle + selected source name + Generic Article.
    # Source priority: ONLINE PRODUCT NAME, then PRODUCT DESCRIPTION (PRINCIPAL).
    source_name = online_name or prod_desc
    if source_name:
        base_name = (
            source_name
            if source_name.upper().startswith("SMIGGLE ")
            else f"Smiggle {source_name}"
        )
        ecom_name_en = (
            base_name
            if not generic_code or base_name.upper().endswith(f"- {generic_code}".upper())
            else f"{base_name} - {generic_code}"
        )
    else:
        ecom_name_en = ""

    # ── Short description EN → PRODUCT DESCRIPTION ───────────────
    short_desc_en = prod_desc or online_name or ""

    # ── Long description EN ──────────────────────────────────────
    # Priority: ONLINE PRODUCT DESCRIPTION → FEATURES & BENEFITS → empty
    if online_desc:
        long_desc_en = online_desc
    elif features:
        long_desc_en = features
    else:
        long_desc_en = ""

    # ── Derived fields ───────────────────────────────────────────
    gender_code = _colour_to_gender(colour)
    age_code    = "CH"   # always Children for Smiggle

    div_letter  = SMIGGLE_DEPT_TO_DIVISION.get(
        (dept or "").upper().strip(), "E"
    )
    cat_code    = SMIGGLE_CATEGORY_TO_PROD_GROUP.get(
        (subcat or "").upper().strip(),
        SMIGGLE_CATEGORY_TO_PROD_GROUP.get(
            (dept or "").upper().strip(), "AC"
        ),
    )

    return {
        # Core identifiers
        "article_no":       line_no,
        "key_val":          key_val,         # composite LINE#+COLOUR — internal
        "model_name":       line_desc,       # LINE DESCRIPTION (principal style desc)
        "brand_code":       brand_code,

        # Colour
        "colour":           colour,
        "colour_token":     colour_token,

        # Gender / Age
        "gender_code":      gender_code,
        "gender_raw":       colour,
        "age_code":         age_code,
        "age_raw":          "CORE",          # default source label for Smiggle

        # Classification
        "division":         dept,
        "div_letter":       div_letter,
        "cat_code":         cat_code,
        "subcat":           subcat,

        # Online / ecommerce content ─────────────────────────────
        "collection1":      collection,
        "collection2":      subcat,          # SUBCAT doubles as Collection 2

        "ecom_name_en":     ecom_name_en,
        "short_desc_en":    short_desc_en,
        "long_desc_en":     long_desc_en,
        "features":         features,
        "dimensions":       dims,            # raw cleaned dimensions string
        "dimensions_parsed":dims_parsed,     # dict: length/width/height
        "care_en":          care_en,
        "fabric":           fabric,

        # Commercial
        "art_category":     "0",            # always Single
        "article_type":     "Inline",       # default; no Design col in ecom file
        "franchise":        collection,      # collection name as franchise proxy

        # Origin — default CN (no port in ecom file)
        "coo":              "CN",

        # Pricing — not in ecom file; left empty for downstream enrichment
        "rrp":              "",
        "currency":         "SGD",
        "fob":              "",

        # Codes
        "generic_code":     generic_code,
        "variant_code":     variant_code,
        "sap_style_code":   line_no,
        "_tdd_sizes":       [],
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Validate mandatory MDD attributes against the mapped article dict."""
    warns = []
    art   = mapped["article_no"]

    field_checks = {
        "article_no":  "AT_PrincipalStyleCode",
        "gender_code": "AT_Gender",
        "age_code":    "AT_SAPAge",
        "brand_code":  "AT_Brand",
    }
    for field, at_id in field_checks.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id}")

    # Ecom-specific: warn if all content fields are empty
    content_fields = ["ecom_name_en", "short_desc_en", "long_desc_en"]
    if all(not mapped.get(f) for f in content_fields):
        warns.append(
            f"[{art}/{mapped.get('colour','')}] All ecom content fields are empty "
            f"(ONLINE PRODUCT NAME, PRODUCT DESCRIPTION, ONLINE PRODUCT DESCRIPTION)"
        )

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


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


def resolve_brand_lov_id(mdd: "MDDLoader", brand_code: str) -> tuple[str, str]:
    """
    Resolve brand display name + canonical MDD Brand LOV ID once per run.
    Returns (brand_display_name, brand_lov_id).
    """
    b_code, b_label = _lov(brand_code, LOV_BRAND, brand_code)
    brand_display = b_label.upper()
    brand_lov_id = b_code
    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})
        for display_name, lov_id in brand_lov.items():
            if display_name.upper() == brand_display:
                brand_lov_id = lov_id
                break
    return brand_display, brand_lov_id


# ── Translation language suffixes — emitted as empty placeholders ─
_LOCALES = ["ID", "PH", "MY", "VN", "TH", "KH"]

# ── Full ecommerce attribute placeholder list ─────────────────────
ECOMMERCE_PLACEHOLDERS = [
    # BY / offline fields not sourced from ecom file
    "AT_SPUGrouping",           "AT_SPUGroupingName",
    # "AT_ECommPriority",
    "AT_MainEANIndicator",
    "AT_HSCode",
    "AT_EstimatedLandedCost",
    "AT_FOBCurrency",
    "AT_Royalty",
    "AT_FreightCost",
    "AT_MarketingFee",
    "AT_SGS",
    "AT_HaddadOfficeCharge",
    "AT_HaddadOfficeChargeDeduction",
    "AT_Others",
    "AT_Commission",
    "AT_TechnologyUsed",
    "AT_CertificateNumber",
    "AT_ProductWeight",
    "AT_PackagingLength",
    "AT_PackagingWidth",
    "AT_PackagingHeight",
    "AT_PackagingWeight",
    "AT_ImagesSource",
    "AT_EcomGenderDescriptionEN",
    "AT_EComAgesCategory",
    "AT_CountrySize",
    # "AT_EComSize",
    "AT_Createdon",
    "AT_Updatedon",
    "AT_Updatedby",
    "AT_LaunchingDate",
    # Image slots — populated from Google Drive by downstream process
    # *[f"AT_Image{i}" for i in range(1, 11)],
]


def _add_ecommerce_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,
    brand_lov_id: str = "",
) -> None:
    """
    Write all Generic-level <Value> elements for a Smiggle ecommerce article.

    Ecommerce-specific attributes written (beyond standard BY/SAP fields):
      AT_EComProductNameEN + locale translations (placeholder)
      AT_ShortDescriptionEN + locale translations (placeholder)
      AT_LongDescriptionEN + locale translations (placeholder)
      AT_CareInstructionEN + locale translations (placeholder)
      AT_ProductLengthWidthHeight  (raw dimensions string)
      AT_ProductLength / AT_ProductWidth / AT_ProductHeight (parsed)
      AT_Collection1 / AT_Collection2
      AT_Interest / AT_Material
      AT_EcomIndicator = Yes
    """
    written: set[str] = set()

    def _w(attr_id: str, value: str = "", id_val: str = "", derived: bool = False):
        has_value = str(value).strip() != ""
        has_id = str(id_val).strip() != ""
        if attr_id not in written and (has_value or has_id):
            _val(vals_el, attr_id, value=value, id_val=id_val, derived=derived)
            written.add(attr_id)

    def _mw(attr_id: str, id_val: str):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    # ── System / organisational ──────────────────────────────────
    sbu_code, _ = _lov(sbu, LOV_SBU, sbu)
    _mw("AT_SBU", sbu_code)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    brand_lov_id = art["brand_code"]   # IGL — correct for AT_Brand

    _w("AT_Brand",      id_val=brand_lov_id)          # ID="IGL"  ✅
    _w("AT_BrandGroup", id_val=brand_name.upper())

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["article_no"])
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["colour"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
   

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
    else:
        # Not found in MDD — send uppercase colour name as plain value, no ID
        _w("AT_Color", value=art["colour"].upper())

    # AT_Generic = BrandCode + PrincipalStyleCode + SAPColorCode
    sap_color_code = colour_id if colour_id else art["colour"].upper()
    at_generic_val = f"{brand_lov_id}{art['article_no']}{sap_color_code}"

    # Store on art so build_product_xml can use it for KEY_Article
    art["at_generic_val"] = at_generic_val

    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",          g_label, id_val=g_code)
    _w("AT_BYGender",        g_label, id_val=g_code)



    # ── Age ──────────────────────────────────────────────────────
    age_raw_upper = (art["age_raw"] or "").strip()

    sap_age_map = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map  = mdd.lovs.get("BYAge",  {}) if mdd else {}

    matched_age_display = next(
        (k for k in sap_age_map if k.upper() == age_raw_upper.upper()), None
    )

    sap_age_code  = sap_age_map.get(matched_age_display, art["age_code"]) if matched_age_display else art["age_code"]
    sap_age_label = matched_age_display or LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge", sap_age_label, id_val=sap_age_code)

    by_age_val = by_age_map.get(matched_age_display, "") if matched_age_display else ""
    if not by_age_val:
        by_age_val = next(
            (v for k, v in by_age_map.items() if k.upper() == age_raw_upper.upper()), "KIDS"
        )
    # After — send uppercase, ID = uppercase value (matches LOV_BYAge)
    _w("AT_BYAge", by_age_val.upper(), id_val=by_age_val.upper())

    _w("AT_PrincipalAgeDescription", art["age_raw"] or sap_age_code)   # renamed from AT_PrincipalAge

    # ── Season ───────────────────────────────────────────────────
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    _w("AT_SeasonYear", year_match.group() if year_match else "")

    # ── Country of Origin — default CN ───────────────────────────
    coo_code, coo_label = _lov(art["coo"], LOV_COUNTRY_ORIGIN, art["coo"])
    _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    
    _w("AT_SAPArticleCategory", id_val="0")

    # ── BY Article Type (prefer filename-derived Inline/License) ──
    at_code = art["article_type"]
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   "Yes", id_val="Y")
    _w("AT_SAPIndicator",  "No",  id_val="N")
    # Ecom file → ecom indicator = Yes
    _w("AT_EcomIndicator", "Yes", id_val="Y")

    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)

    # ── UOM / Size ───────────────────────────────────────────────
    _w("AT_UOM",           "Each",    id_val="EA")
    _w("AT_Size",          "No Size", id_val="000")
    _w("AT_PrincipalSize", "One Size")
    _w("AT_CountrySize", "ONE SIZE")
    # _w("AT_EComSize",      "One Size")

    # ── Pricing placeholders (not in ecom file) ───────────────────
    _w("AT_OriginalPrice",       art.get("rrp", ""))
    _w("AT_CurrentPrice",        art.get("rrp", ""))
    _w("AT_RetailPriceCurrency", art.get("currency", "SGD"), id_val=art.get("currency", "SGD"))
    _w("AT_FOB",                 art.get("fob", ""))
    _w("AT_FOBCurrency",         "USD", id_val="USD")

    # ── Nature of Article ────────────────────────────────────────
    _w("AT_NatureOfArticle", "Regular", id_val="REG")

    # ── Vendor ───────────────────────────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Collections / hierarchy ──────────────────────────────────
    _w("AT_Collection1", art["collection1"])
    _w("AT_Collection2", art["collection2"])

    # ── Interest / Material (from SUBCAT) ─────────────────────────
    _w("AT_Interest", art["subcat"])
    _w("AT_Material", art["subcat"])

    # ── Ecom Product Name (EN + locale placeholders) ─────────────
    _w("AT_EComProductNameEN", art["ecom_name_en"])
    for loc in _LOCALES:
        _w(f"AT_EComProductName{loc}", "")

    # ── Short Description (EN + locale placeholders) ─────────────
    _w("AT_ShortDescriptionEN", art["short_desc_en"])
    for loc in _LOCALES:
        _w(f"AT_ShortDescription{loc}", "")

    # ── Long Description (EN + locale placeholders) ───────────────
    _w("AT_LongDescriptionEN", art["long_desc_en"])
    for loc in _LOCALES:
        _w(f"AT_LongDescription{loc}", "")

    # ── Care Instructions (EN + locale placeholders) ─────────────
    _w("AT_CareInstructionEN", art["care_en"])
    for loc in _LOCALES:
        _w(f"AT_CareInstruction{loc}", "")

    # ── Dimensions ───────────────────────────────────────────────
    _w("AT_ProductLengthWidthHeight", art["dimensions"])
    dp = art.get("dimensions_parsed", {})
    # _w("AT_ProductLength", dp.get("length", ""))
    # _w("AT_ProductWidth",  dp.get("width",  ""))
    # _w("AT_ProductHeight", dp.get("height", ""))

    # ── Principal Color Name (EN + locale placeholders) ───────────
    # _w("AT_PrincipalColorName",   art["colour"])
    # _w("AT_PrincipalColorNameEN", art["colour"])
    # for loc in _LOCALES:
    #     _w(f"AT_PrincipalColorName{loc}", "")



# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_code}_{full_season_code}"
    batches_parent   = f"CLH_{brand}Batches"

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
    mdd=None,
    brand_lov_id: str = "",
) -> str:
    """
    Build a <Product> XML fragment for one Smiggle ecommerce article.

    UserTypeID = PRD_SingleArticle  (Single — accessories brand, no size)
    ParentID   = PPH_E-TempSubCat   (Equipment / Accessories)
    KeyID      = KEY_Article, value = {brand_code}{Line#}{colour_token}
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = brand_lov_id or art.get("brand_code", brand_code) or brand_code
    div_letter = art.get("div_letter", "E")
    parent_id  = f"PPH_{div_letter}-TempSubCat"
    cat_code   = art.get("cat_code", "AC")

    # KEY_Article must match AT_Generic — both include the SAP color code
    # at_generic_val is set by _add_generic_values, but we need it here first.
    # Recompute it the same way so KEY_Article and AT_Generic always match.
    colour_id_for_key = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        colour_id_for_key = color_lov.get(art["colour"], "") \
                        or color_lov.get(art["colour"].upper(), "")
    if colour_id_for_key and colour_id_for_key.isdigit():
        colour_id_for_key = colour_id_for_key.zfill(3)

    sap_color_for_key = colour_id_for_key if colour_id_for_key else art["colour"].upper()
    key_article = f"{b_code}{art['article_no']}{sap_color_for_key}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_SingleArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["model_name"] or article_no
    )

    # ── Classification references ────────────────────────────────
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{b_code}_{div_letter}_{cat_code}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{div_letter}")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Values ───────────────────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_ecommerce_values(vals_el, art, brand, comp_code, sbu, mdd=mdd, brand_lov_id=b_code)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one ecommerce row."""
    row, brand_code, mdd, season, country_code, article_type_from_filename = row_tuple
    mapped = map_article_ecommerce(row, brand_code=brand_code)
    mapped["season"] = season
    mapped["country_code"] = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def _extract_file_type_segment(filename_stem: str) -> str:
    """
    Extract the file-type portion from a filename like:
        0888-SP-SMIGGLE-Ecommerce File Inline-Multi-SS28-1
    Returns 'Ecommerce File Inline'. Falls back to 'Ecommerce'.
    """
    import re as _re
    parts = _re.split(r'\s*-\s*', filename_stem)
    # Layout: [CompCode, SBU, Brand, ...FileType..., Multi/Mono, Season, Seq]
    # So file type = everything between index 3 and -3
    if len(parts) >= 7:
        segment = "-".join(parts[3:-3]).strip()
        return segment if segment else "Ecommerce"
    return "Ecommerce"



def run(args, auditor=None):
    """
    Main entry point for ecommerce ETL.

    args must have attributes:
        brand       str   e.g. "Smiggle"
        brand_code  str   e.g. "SMI"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SM"
        season      str   e.g. "SS26"
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ec_files = list(ECOMMERCE_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Ecommerce", ec_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand=args.brand)
    brand_display, brand_lov_id = resolve_brand_lov_id(mdd, args.brand_code)

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # Build season ID: CLH_{BrandLOVID}_{Season}
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for ec_path in ec_files:
        log.info("─── Processing Ecommerce File: %s ───", ec_path.name)

        ec = SmiggleEcommerceLoader(ec_path)
        if ec.df.empty:
            log.warning("[Ecommerce-SMI] Empty dataframe — skipping.")
            continue

        rows       = [row for _, row in ec.df.iterrows()]
        # ── TEST MODE: limit to 1 product ───────────────────────
        # rows = rows[:5]
        # ────────────────────────────────────────────────────────
        total_rows = len(rows)
        log.info("Total valid rows (after invalid LINE# filter): %d", total_rows)

        # ── Output filename ──────────────────────────────────────
        file_type_segment = _extract_file_type_segment(ec_path.stem)
        multi_mono        = getattr(args, "multi_mono", "Multi")

        out_name = ec_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [(row, brand_lov_id, mdd, args.season, getattr(args, "country_code", ""), getattr(args, "article_type_from_filename", "")) for row in rows]
        ordered: list[tuple[int, dict, list]] = []

        log.info(
            "Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
            total_rows, num_workers,
        )

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
            brand_display.title(), brand_lov_id, args.comp_code, args.sbu, args.season,
        )
        cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
        del cls_el

        written_count     = 0
        empty_content_ct  = 0

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
                if not art.get("ecom_name_en") and not art.get("short_desc_en"):
                    empty_content_ct += 1
                product_xml = build_product_xml(
                    art, brand_display.title(), brand_lov_id,
                    args.comp_code, args.sbu, season_id, mdd=mdd, brand_lov_id=brand_lov_id,
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

        print("═══ ARTICLE SUMMARY (SMIGGLE ECOMMERCE) ══════════════", flush=True)
        print(f"  Ecommerce rows (valid)       : {total_rows}",       flush=True)
        print(f"  Articles written             : {written_count}",    flush=True)
        print(f"  Articles with empty content  : {empty_content_ct}", flush=True)
        print(f"  XML file size                : {file_kb}KB",        flush=True)
        print("═══════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_ecommerce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
        description="Stibo Inbound XML Generator — Smiggle Ecommerce v1.0"
    )
    p.add_argument("--brand",      default="Smiggle")
    p.add_argument("--brand-code", default="SMI")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SM")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
