"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — Smiggle  v1.0               ║
║  Linelist (MAPI sheet) → Stibo STEP XML                        ║
╚══════════════════════════════════════════════════════════════════╝

Smiggle-specific differences vs Adidas:
  • Single input file only — no TDD, no Backlog
  • Header is at row index 8 (0-based) in the MAPI sheet
  • Principal Style Code  = "Line #" column (integer or "TBC")
  • Principal Style Desc  = "Description" column
  • Colour                = "Colour" column  (free-text, no numeric code)
  • Age                   = "Age" column → mapped to SAP Age (default CH)
  • SAP Gender            = derived from Colour keywords / default U
  • Article Category      = always Single (00) — accessories brand, no sizing
  • SAP Size              = always "000" / No Size
  • COO                   = hardcoded CN — all Smiggle sourcing is China
  • FOB                   = "FOB USD" column
  • RRP / Currency        = "SG RRP" + "SGD" (Singapore price list)
  • Article Type (BY)     = "Design" column → Licence / Proprietary → maps to License / Inline
  • Collection 1          = "Collection" column
  • Collection 2          = "Sub Category" column
  • Franchise             = "Design" column  (same as article type for Smiggle)
  • Launching Date        = "Embargo Date" column
  • Department            = → SAP Product Division  (PACK & CARRY / STATIONERY / ACCESSORIES)
  • Category              = → SAP Product Group
  • Sub Category          = → SAP Product Category / Interest / Material

  One Generic per (Line # + Colour) row.  TBC Line # rows are skipped.
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
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
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
    # Smiggle-specific raw values from Age column
    "CORE": "Children", "TEENY TINY": "Children",
}
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

# ── Smiggle-specific lookup tables ───────────────────────────────

# NOTE: COO is hardcoded to CN in map_article_smiggle — SMIGGLE_PORT_TO_COO
# is retained here for reference only and is NOT used in the mapping pipeline.
SMIGGLE_PORT_TO_COO: dict[str, str] = {
    "SHENZHEN":    "CN",
    "SHANGHAI":    "CN",
    "NINGBO":      "CN",
    "XIAMEN":      "CN",
    "CHINA":       "CN",
    "GUANGZHOU":   "CN",
    "VIETNAM":     "VN",
    "HO CHI MINH": "VN",
}

# ── Smiggle Material Type LOV ─────────────────────────────────────
SMIGGLE_MATERIAL_TYPE_LOV: dict[str, str] = {
    "ZHAW": "Direct Articles (QVB)",
    "ZINA": "Intercompany",           # XML emits "Intercompany" — not the full "Intercompany Articles"
    "ZUNB": "Nonvaluated Articles (QB)",
    "ZDIN": "Nonstock Articles (NQVB)",
}

# Default Material Type for Smiggle (Intercompany brand)
SMIGGLE_DEFAULT_MATERIAL_TYPE = "ZINA"

# Smiggle Age column value → SAP Age code
SMIGGLE_AGE_MAP: dict[str, str] = {
    "CORE":       "CH",   # Core range — still children
    "JUNIOR":     "CH",   # Junior  → Children
    "TEENY TINY": "CH",   # Infant/toddler — map to Children (Smiggle's youngest tier)
}

# Smiggle Design column → BY Article Type LOV
SMIGGLE_DESIGN_TO_BY_ARTICLE_TYPE: dict[str, str] = {
    "LICENCE":     "License",
    "LICENSE":     "License",
    "PROPRIETARY": "Inline",
}

# Smiggle Department → SAP Product Division letter
SMIGGLE_DEPT_TO_DIVISION: dict[str, str] = {
    "PACK & CARRY": "E",    # Equipment / Accessories
    "STATIONERY":   "E",
    "ACCESSORIES":  "E",
}

# Smiggle Category → SAP Product Group code (2-char)
SMIGGLE_CATEGORY_TO_PROD_GROUP: dict[str, str] = {
    "BAGS":               "BG",
    "DRINK":              "DK",
    "FOOD":               "FD",
    "PENCIL CASE":        "PC",
    "WRITE SINGLES":      "WS",
    "WRITE PACKS":        "WP",
    "DESKTOP":            "DS",
    "STATIONERY":         "ST",
    "ACCESSORIES":        "AC",
    "FASHION ACCESSORIES":"FA",
    "ACTIVITY":           "AV",
    "GAMES":              "GM",
    "BOOKS":              "BK",
    "KITS":               "KT",
    "CONSUMABLES":        "CO",
    "LIGHTS":             "LT",
    "TIME":               "TM",
}

# Smiggle Colour string → SAP Gender
# Smiggle is a kids brand so gender is largely U (Unisex), but
# some colour names signal a specific gender target.
def _colour_to_gender(colour: str) -> str:
    """Heuristic: derive SAP Gender code from Smiggle colour string."""
    c = (colour or "").upper()
    if any(k in c for k in ("PINK", "LILAC", "ROSE", "PURPLE")):
        return "F"
    if any(k in c for k in ("BLUE", "NAVY", "CHARCOAL", "DARK GREY")):
        return "U"   # neutral — don't force Male
    return "U"


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
        """
        Load Age LOV sheet with both SAP Age Code (col C) and BY Age (col F).
        Stores:
        self.lovs["SAPAge"]  = { "Children": "K", "Adults": "A", "All Ages": "O" }
        self.lovs["BYAge"]   = { "Children": "Kids", "Adults": "Adult", "All Ages": "All Ages" }
        """
        sheet_name = next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return

        sap_age_map = {}
        by_age_map  = {}

        for row in rows[1:]:   # skip header
            if not row or not row[0]:
                continue
            age_display  = str(row[0]).strip()           # col A: "Children"
            sap_age_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""   # col C: "K"
            by_age_val   = str(row[5]).strip() if len(row) > 5 and row[5] else ""   # col F: "Kids"

            if age_display and sap_age_code:
                sap_age_map[age_display] = sap_age_code
            if age_display and by_age_val:
                by_age_map[age_display] = by_age_val

        self.lovs["SAPAge"] = sap_age_map
        self.lovs["BYAge"]  = by_age_map
        log.info("[MDD] Age LOV loaded — SAPAge: %d entries, BYAge: %d entries",
                len(sap_age_map), len(by_age_map))

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

        # Prefer exact match, skip v4 variant
        sheet_name = next(
            (s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()),
            None,
        ) or next(
            (s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()),
            wb.sheetnames[0],
        )
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
            # Attribute name is in col index 1
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


class SmiggleLinelistLoader:
    """
    Loads the Smiggle linelist workbook.

    The active sheet is determined by name (defaults to "MAPI").
    Header row is auto-detected by looking for the "Line #" column.

    Each row in the resulting DataFrame represents one
    (Line # / Colour) combination — which maps to one Generic article
    in STEP (Smiggle uses Single Article Category 00, no variants).

    Rows where Line # is "TBC" or blank are skipped.
    """

    # Candidate sheet names, in preference order
    PREFERRED_SHEETS = ["MAPI", "Linelist", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[Linelist-SMI] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in self.PREFERRED_SHEETS if s in wb.sheetnames), None
        )
        if target is None:
            # Fall back to the sheet with most rows
            best, best_rows = wb.sheetnames[0], 0
            for sn in wb.sheetnames:
                n = wb[sn].max_row or 0
                if n > best_rows:
                    best, best_rows = sn, n
            target = best

        log.info("[Linelist-SMI] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Auto-detect header row: first row that contains "Line #" or "Line#"
        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and "LINE" in v.upper() for v in r if v):
                hdr_idx = i
                break

        # Fall back: first row with ≥ 5 non-null string cells
        if hdr_idx is None:
            hdr_idx = next(
                (
                    i
                    for i, r in enumerate(rows)
                    if sum(1 for v in r[:12] if isinstance(v, str) and v.strip()) >= 5
                ),
                None,
            )

        if hdr_idx is None:
            log.error("[Linelist-SMI] Cannot find header row"); wb.close(); return

        log.info("[Linelist-SMI] Header at row %d (0-indexed)", hdr_idx)
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(rows[hdr_idx])
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Identify the Line # column (may be "Line # " with trailing space)
        line_col = next(
            (c for c in df.columns if c.strip().upper().startswith("LINE")), None
        )
        if line_col:
            # Drop rows where Line # is blank, None, or "TBC"
            df = df[
                df[line_col].notna()
                & (~df[line_col].astype(str).str.strip().isin(["", "None", "nan", "TBC"]))
            ]
            # Rename to a canonical name for downstream use
            df = df.rename(columns={line_col: "Line #"})

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info(
            "[Linelist-SMI] %d rows loaded (TBC / blank Line# skipped)", len(self.df)
        )


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely convert any cell value to a clean string; return '' for empties."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT") else s


def _fmt_date(v) -> str:
    """Format a date to dd-Mon-YYYY lowercase (e.g. 15-may-2026), matching Stibo's expected format."""
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y").lower()
    if hasattr(v, "strftime"):          # date object
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


def map_article_smiggle(ll_row: dict, brand_code: str = "SMI") -> dict:
    """
    Map one Smiggle linelist row → unified article dict.

    Key decisions:
      - article_no   = Line # (numeric, cast to string, zero-stripped)
      - model_name   = Description
      - colour       = Colour (free-text, used as-is for Principal Color)
      - gender_code  = derived from Colour keyword heuristic (default U)
      - age_code     = from Age column, default CH
      - art_category = always "0" (Single — accessories brand)
      - division     = from Department column
      - coo          = hardcoded CN — all Smiggle sourcing is China
      - rrp          = SG RRP
      - currency     = SGD (Singapore price list)
      - article_type = Design column (Licence → License / Proprietary → Inline)
      - collection1  = Collection column
      - collection2  = Sub Category column
      - franchise    = Design column value (Licence/Proprietary)
      - embargo_date = Embargo Date column
      - fob          = FOB USD column
      - department   = Department column (for SAP Division)
      - category     = Category column (for SAP Product Group)
      - sub_category = Sub Category column (for SAP Product Category)
    """
    line_no  = _s(ll_row.get("Line #"))
    # Strip any trailing .0 from numeric-looking values (pandas int→float risk)
    if line_no.endswith(".0"):
        line_no = line_no[:-2]

    desc         = _s(ll_row.get("Description"))
    colour       = _s(ll_row.get("Colour"))
    age_raw      = _s(ll_row.get("Age"))
    design       = _s(ll_row.get("Design"))
    dept         = _s(ll_row.get("Department"))
    category     = _s(ll_row.get("Category"))
    sub_cat      = _s(ll_row.get("Sub Category"))
    collection   = _s(ll_row.get("Collection"))
    sg_rrp       = _s(ll_row.get("SG RRP"))
    fob          = _s(ll_row.get("FOB USD"))
    port         = _s(ll_row.get("Port"))
    embargo_raw  = ll_row.get("Embargo Date")
    specification = _s(ll_row.get("Specification"))   # Fashion / Graphic

    # ── Derived fields ──────────────────────────────────────────
    # SAP Gender: heuristic on colour string
    gender_code = _colour_to_gender(colour)   # "U" / "F" / "M"

    # SAP Age: map from Smiggle Age column, default CH (Children)
    age_code = SMIGGLE_AGE_MAP.get((age_raw or "").upper(), "CH")

    # Country of Origin: hardcoded CN — all Smiggle sourcing is China.
    # Port column is retained on the dict for auditing but does NOT affect COO.
    coo = "CN"

    # BY Article Type / Franchise: from Design column
    by_art_type = SMIGGLE_DESIGN_TO_BY_ARTICLE_TYPE.get(
        (design or "").upper().strip(), "Inline"
    )

    # SAP Product Division letter
    div_letter = SMIGGLE_DEPT_TO_DIVISION.get(
        (dept or "").upper().strip(), "E"
    )

    # Embargo / Launching Date
    embargo_date = _fmt_date(embargo_raw)

    # RRP: strip non-numeric chars (e.g. currency prefix)
    rrp_str = ""
    if sg_rrp:
        m = re.search(r"[\d.]+", str(sg_rrp))
        rrp_str = m.group() if m else str(sg_rrp)

    # FOB: same
    fob_str = ""
    if fob:
        m = re.search(r"[\d.]+", str(fob))
        fob_str = m.group() if m else str(fob)

    # Generic code: SMI + Line #  (brand_code 3 chars + style code)
    generic_code = f"{brand_code}{line_no}"

    # Variant key: Generic + colour token (3 chars, padded) + "000" (no size)
    colour_token = re.sub(r"[^A-Z0-9]", "", colour.upper())[:3].ljust(3, "0")
    variant_code = f"{generic_code}{colour_token}000"

    return {
        # Core identifiers
        "article_no":       line_no,
        "model_name":       desc,
        "brand_code":       brand_code,

        # Colour
        "colour":           colour,
        "colour_token":     colour_token,

        # Gender / Age
        "gender_code":      gender_code,
        "gender_raw":       colour,          # raw source for PrincipalGender
        "age_code":         age_code,
        "age_raw":          age_raw,

        # Classification
        "division":         dept,            # Department → division letter via SMIGGLE_DEPT_TO_DIVISION
        "div_letter":       div_letter,
        "category":         category,
        "sub_category":     sub_cat,
        "specification":    specification,

        # Commercial
        "article_type":     by_art_type,     # Inline / License
        "art_category":     "0",             # always Single (00)
        "collection1":      collection,
        "collection2":      sub_cat,
        "franchise":        design,           # "Licence" / "Proprietary" raw
        "embargo_date":     embargo_date,

        # Pricing
        "rrp":              rrp_str,
        "currency":         "SGD",
        "fob":              fob_str,

        # Origin
        "coo":              coo,
        "port":             port,

        # Derived codes
        "generic_code":     generic_code,
        "variant_code":     variant_code,

        # SAP style code = Line # (same as article_no for Smiggle)
        "sap_style_code":   line_no,

        # No TDD / Backlog data for Smiggle
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
        "article_no":   "AT_PrincipalStyleCode",
        "gender_code":  "AT_Gender",
        "age_code":     "AT_SAPAge",
        "brand_code":   "AT_Brand",
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
    """
    Append a <Value AttributeID="..."> element to *parent*.

    Rules:
      • derived=True  → skip (calculated attributes go in Stibo, not inbound XML)
      • id_val given   → emit ID="..." attribute only (LOV reference), no text
      • value given    → emit text content
      • neither        → emit empty element (placeholder)
    """
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
    """
    Resolve brand identifiers once from source brand code.
    Returns:
      - brand_lov_id: the MDD Brand LOV ID (e.g. IGL for Smiggle)
      - brand_label: display label from LOV_BRAND (e.g. SMIGGLE)
    """
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
    """
    Write all Generic-level <Value> elements for a Smiggle article.
    Order mirrors the Stibo attribute group sequence in the MDD.
    """
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
    brand_lov_id = art["brand_code"]   # IGL — correct for AT_Brand

    _w("AT_Brand",      id_val=brand_lov_id)          # ID="IGL"  ✅
    _w("AT_BrandGroup", id_val=brand_name.upper())    # ID="SMIGGLE" ✅    

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
    # _w("AT_PrincipalGender", art["gender_raw"])   # raw colour string as principal gender label

    # ── Age ──────────────────────────────────────────────────────
    # Look up SAP Age Code and BY Age from MDD Age LOV tab
    # Age LOV: col A = display name (e.g. "Children"), col C = SAP code (e.g. "K"), col F = BY Age (e.g. "Kids")
    age_raw_upper = (art["age_raw"] or "").strip()

    sap_age_map = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map  = mdd.lovs.get("BYAge",  {}) if mdd else {}

    # Match age_raw against col A display names (case-insensitive)
    matched_age_display = next(
        (k for k in sap_age_map if k.upper() == age_raw_upper.upper()), None
    )

    sap_age_code  = sap_age_map.get(matched_age_display, art["age_code"]) if matched_age_display else art["age_code"]
    sap_age_label = matched_age_display or LOV_AGE.get(sap_age_code, sap_age_code)
    _w("AT_SAPAge", sap_age_label, id_val=sap_age_code)

    by_age_val = by_age_map.get(matched_age_display, "") if matched_age_display else ""
    if not by_age_val:
        # fallback: try matching directly on age_raw
        by_age_val = next(
            (v for k, v in by_age_map.items() if k.upper() == age_raw_upper.upper()), "KIDS"
        )
    _w("AT_BYAge", by_age_val.upper(), id_val=by_age_val.upper())

    _w("AT_PrincipalAgeDescription", art["age_raw"] or sap_age_code)

    # ── Season ───────────────────────────────────────────────────
    # season comes from args — stored on art via the orchestrator
    sea_raw   = art.get("season", "")
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _w("AT_Season", sea_label, id_val=sea_code)

    year_match = re.search(r"20\d{2}", sea_raw)
    if not year_match:
        # Handle short format: SS28 → 2028, FW26 → 2026
        short_match = re.search(r"\d{2}$", sea_raw)
        season_year = f"20{short_match.group()}" if short_match else ""
    else:
        season_year = year_match.group()
    _w("AT_SeasonYear", season_year)

    # ── Country of Origin ────────────────────────────────────────
    coo_code, coo_label = _lov(art["coo"], LOV_COUNTRY_ORIGIN, art["coo"])
    _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    # AT_SAPArticleCategory is intentionally hardcoded to id_val="0" (Single / 00)
    # for ALL Smiggle articles. Smiggle is an accessories brand — there are no size
    # variants and no Generic/Variant split. This value will never vary.
    _w("AT_SAPArticleCategory", id_val="0")

    # ── BY Article Type ──────────────────────────────────────────
    at_code  = art["article_type"]          # "License" or "Inline"
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   "Yes",          id_val="Y")
    _w("AT_SAPIndicator",  "No",           id_val="N")
    _w("AT_EcomIndicator", "")
    _w("AT_SAPProductFlag", "A",  id_val="A")

    # ── UOM — always EA for Smiggle ──────────────────────────────
    _w("AT_UOM", "Each", id_val="EA")

    # ── SAP Size — always No Size (000) for Smiggle ─────────────
    _w("AT_Size",          "No Size", id_val="000")
    _w("AT_PrincipalSize", "One Size")

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice",        art["rrp"])
    _w("AT_CurrentPrice",         art["rrp"])
    _w("AT_RetailPriceCurrency",  art["currency"],  id_val=art["currency"])
    _w("AT_FOB",                  art["fob"])
    _w("AT_FOBCurrency",          "USD",            id_val="USD")

    # ── Nature of Article — default Regular ──────────────────────
    _w("AT_NatureOfArticle", "Regular", id_val="REG")

    # ── Vendor identification ─────────────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Collections / hierarchy ──────────────────────────────────
    _w("AT_Collection1", art["collection1"])
    _w("AT_Collection2", art["collection2"])
    _w("AT_Franchise", (art["franchise"] or "").upper())

    # ── Ecom Product Name EN ─────────────────────────────────────
    # Rule: "Smiggle" + " " + Description  (per SMIGGLE attributes list)
    ecom_name_en = f"Smiggle {art['model_name']}".strip() if art["model_name"] else ""
    _w("AT_EComProductNameEN", ecom_name_en)

    # ── Launching / Embargo date ─────────────────────────────────
    _w("AT_LaunchingDate", art["embargo_date"])

    # ── Principal Merchandise Hierarchy (L1–L5) ──────────────────
    _w("AT_PrincipalMerchandiseHierarchyL1", art.get("category", ""))
    _w("AT_PrincipalMerchandiseHierarchyL2", art.get("sub_category", ""))
    # L3, L4, L5 — no mapping defined for Smiggle in the attributes list

    # ── Sports Category EN — default "Other" for Smiggle ─────────
    _w("AT_SportsCategoryEN", "Other")

    # ── Interest / Material (from Sub Category) ───────────────────
    _w("AT_Interest",  art["sub_category"])

    mat_type_code  = art.get("material_type", SMIGGLE_DEFAULT_MATERIAL_TYPE)
    mat_type_label = SMIGGLE_MATERIAL_TYPE_LOV.get(mat_type_code, mat_type_code)
    _w("AT_MaterialType", mat_type_label, id_val=mat_type_code)

    # ── Country Size / E-com Size — default ONE SIZE ─────────────
    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)

    _w("AT_CountrySize", "ONE SIZE")


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_code: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """
    Build the <Classifications> block.
    Season ID: CLH_{brand_code}_{sea_prefix}{sea_year}  e.g. CLH_SMI_SS2026
    """
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
    """
    Build a <Product> XML fragment for one Smiggle article.

    Smiggle specifics:
      • UserTypeID = PRD_SingleArticle  (always — Article Category = Single / 00)
      • ParentID   = PPH_E-TempSubCat   (all Smiggle = Equipment/Accessories)
      • KeyID      = KEY_InboundArticle, value = {brand_code}{Line#}{SAPColorCode}
        (colour is part of the key because the same Line# can have multiple colours
         and each is a separate Generic in STEP)
      • No Variant sub-products (Single article, no size run)
      • ClassificationReferences: Merchandiser, BYHierarchy, SAPHierarchy, Unconfirmed
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code      = art.get("brand_code", brand_code) or brand_code
    div_letter  = art.get("div_letter", "E")
    parent_id   = f"PPH_{div_letter}-TempSubCat"

    # Include colour token in KEY so that same Line# + different colour = different STEP article
    # KEY_Article must match AT_Generic — both include the SAP color code.
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

    # SAP product group / category codes for BY Hierarchy reference
    cat_code = SMIGGLE_CATEGORY_TO_PROD_GROUP.get(
        (art.get("category") or "").upper().strip(), "AC"
    )
    sub_cat_raw = re.sub(r"[^A-Z0-9]", "", (art.get("sub_category") or "").upper())[:4]

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
    # 1. Merchandiser
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # 2. BY Hierarchy  (brand_code + division + category)
    # cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{b_code}_{div_letter}_{cat_code}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # 3. SAP Hierarchy  (division letter)
    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{div_letter}")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

    # 4. Unconfirmed season (all inbound articles start unconfirmed)
    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Values ───────────────────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    # NOTE: No Variant sub-products for Smiggle (Single Article Category)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one Smiggle row. Returns (mapped, warns)."""
    row, brand_code, mdd, season, country_code, article_type_from_filename = row_tuple
    mapped = map_article_smiggle(row, brand_code=brand_code)
    mapped["season"] = season
    mapped["country_code"] = country_code
    mapped["article_type_from_filename"] = article_type_from_filename
    warns  = validate(mapped, mdd)
    return mapped, warns


# ══════════════════════════════════════════════════════════════════
# SECTION 7 — ORCHESTRATOR  (called by lambda_function.py as run(args))
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point.

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
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    # ── Mandatory file checks ────────────────────────────────────
    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Linelist", ll_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand=args.brand)   # loaded for logging / future use

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    # ── Season ID (CLH_SMI_SS2026) ───────────────────────────────
    brand_lov_id, brand_lov_label = resolve_brand_lov_id(mdd, args.brand_code)
    log.info(
        "Resolved brand LOV → source_brand_code=%s  brand_label=%s  brand_lov_id=%s",
        args.brand_code, brand_lov_label, brand_lov_id
    )

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{brand_lov_id}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing Linelist: %s ───", ll_path.name)

        ll = SmiggleLinelistLoader(ll_path)
        if ll.df.empty:
            log.warning("[Linelist-SMI] Empty dataframe — skipping."); continue

        rows     = [row for _, row in ll.df.iterrows()]
        # ── TEST MODE: limit to 5 products ──────────────────────
        # rows = rows[:5]
        # ────────────────────────────────────────────────────────
        total_rows = len(rows)
        log.info("Total valid rows (after TBC/blank filter): %d", total_rows)

        # ── Parse file_type and multi_mono from input filename ───────
        # Convention: {CompCode}-{SBU}-{Brand}-{FileType}-{Multi/Mono}-{Season}-{Seq}
        # [0]          [1]       [2]    [3]       [4]        [-2]       [-1]
        # FileType may contain spaces/hyphens ("Linelist Inline") so we
        # anchor from both ends: [3:-3] = FileType, [-3] = Multi/Mono.
        # Additionally use regex on the stem as a robust fallback for Multi/Mono.
        _stem  = ll_path.stem
        _parts = re.split(r"\s*-\s*", _stem)

        if len(_parts) >= 7:
            file_type_from_input = "-".join(_parts[3:-3])   # e.g. "Linelist Inline"
        else:
            file_type_from_input = "Catalogue"

        # Use regex on full stem — same approach as NB Licensed reference file
        _mm_match = re.search(r'\b(Multi|Mono)\b', _stem, re.IGNORECASE)
        multi_mono_from_input = _mm_match.group(1).capitalize() if _mm_match else "Multi"

        log.info(
            "Parsed from filename → file_type='%s'  multi_mono='%s'",
            file_type_from_input, multi_mono_from_input,
        )

        # ── Output filename ──────────────────────────────────────────
        out_name = ll_path.stem + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers   = min(8, max(1, total_rows))
        task_args     = [
            (
                row, brand_lov_id, mdd, args.season,
                getattr(args, "country_code", ""),
                getattr(args, "article_type_from_filename", ""),
            )
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

        print("═══ ARTICLE SUMMARY (SMIGGLE) ══════════════════════", flush=True)
        print(f"  Linelist rows (valid)  : {total_rows}",   flush=True)
        print(f"  Articles written       : {written_count}", flush=True)
        print(f"  XML file size          : {file_kb}KB",     flush=True)
        print("════════════════════════════════════════════════════",  flush=True)

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
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator — Smiggle v1.0")
    p.add_argument("--brand",      default="Smiggle")
    p.add_argument("--brand-code", default="SMI")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SM")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()