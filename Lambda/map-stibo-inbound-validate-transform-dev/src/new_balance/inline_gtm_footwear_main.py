"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — NEW BALANCE  (GTM Footwear) v1.0 ║
║   Line List (Footwear GTM) → Stibo STEP XML                      ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : S2'27 APAC Footwear GTM 2 Line Sheet
               ("Copy of S227 GTM 2 APAC Footwear Line List July 2026.xlsx")
Primary tab  : APAC   (title row index 0, header row index 3, data from index 4)

────────────────────────────────────────────────────────────────────
WHY THIS MODULE EXISTS
────────────────────────────────────────────────────────────────────
The GTM line sheet is a *different* export from the Footwear Price List that
`inline_main_footwear.py` consumes. Only 4 of the 15 source columns that the
price-list ETL reads are present in the GTM sheet, so every attribute whose
source column is absent is DELIBERATELY NOT EMITTED here (rather than emitted
empty or back-filled from a look-alike column).

Column availability audit — Price List columns vs. GTM sheet:

  PRESENT (mapping preserved 1:1)
  ───────────────────────────────
    Item Number          → AT_PrincipalStyleCode
                         → AT_PrincipalStyleDescription
                           (price-list ETL writes `display_name or item_number`;
                            with no Product Name column the Item Number
                            fallback applies — which is also the mapping the
                            MDD / brand attributes sheet declares)
    Product Number       → (read, but emits nothing — AT_SAPStyleCode was
                            withdrawn at MAA's request, see REMOVED_BY_REQUEST)
    Line Plan Business   → AT_SportsCategoryEN (LOV — emits the MDD LOV *ID*)
                         → AT_PrincipalMerchandiseHierarchyL2
    CATEGORY             → (read, but the price-list ETL emits no attribute
                            from it — so nothing is emitted here either)

  ABSENT (attribute dropped — see DROPPED_ATTRIBUTES below)
  ────────────────────────────────────────────────────────
    NRF Color, Size Profile, Product Line, Segment, Sizes - Region,
    Country of Origin, NBIL Price, Retail Price, Factory Name,
    Product Name, Intro Period

Attributes with no source-column dependency (brand/season/company context and
fixed constants) are emitted unchanged, exactly as the price-list ETL does.

────────────────────────────────────────────────────────────────────
DEVIATION (one, deliberate)
────────────────────────────────────────────────────────────────────
AT_InboundGenericCode / KEY_InboundArticle is built from Item Number instead
of Product Number. In the GTM sheet Product Number is style-level, not
colorway-level (2781 rows collapse to 468 Product Numbers), and the colorway
component of the original key (NRF Color) is gone — so keying on Product
Number would merge ~2300 distinct articles into 468. Item Number is present
and unique per row, so it carries the key. See _build_generic_code().
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
# SOURCE COLUMN CONTRACT
# ══════════════════════════════════════════════════════════════════
# Every column this ETL reads, and the attributes each one feeds.
SOURCE_COLUMNS: dict[str, list[str]] = {
    "Item Number":        ["AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
                           "AT_InboundGenericCode"],
    "Product Number":     [],   # AT_SAPStyleCode withdrawn — emits nothing now
    "Line Plan Business": ["AT_SportsCategoryEN",
                           "AT_PrincipalMerchandiseHierarchyL2"],
    "Adult Gender/ Kids Closure Type": ["AT_PrincipalGenderDescription"],
    "CATEGORY":           [],   # read for logging/QA only — emits no attribute
}

# Attributes withdrawn on MAA instruction, NOT because a source column is
# absent. Their source data is present (or they were constants) — MAA asked for
# them to be left out of the payload. Kept separate from DROPPED_ATTRIBUTES so
# the two reasons never get confused.
REMOVED_BY_REQUEST: dict[str, str] = {
    "AT_Gender":   "MAA asked not to send it",
    "AT_BYGender": "MAA asked not to send it",
    "AT_SAPAge":   "MAA asked not to send it",
    "AT_BYAge":    "MAA asked not to send it",
    "AT_SAPStyleCode":             "Product Number was available; MAA asked not to send it",
    "AT_BYIndicator":              "constant 'Y'; MAA asked not to send it",
    "AT_SAPIndicator":             "constant 'N'; MAA asked not to send it",
    "AT_EcomIndicator":            "constant 'N'; MAA asked not to send it",
    "AT_UOM":                      "constant 'EA'; MAA asked not to send it",
    "AT_MainVendorIdentification": "constant '1'; MAA asked not to send it",
    "AT_NatureOfArticle":          "constant 'Retail' (RT1); MAA asked not to send it",
}

# Attributes the Price List ETL (inline_main_footwear.py) emits that are NOT
# emitted here, keyed by the source column that is missing from the GTM sheet.
DROPPED_ATTRIBUTES: dict[str, list[str]] = {
    # AT_PrincipalStyleDescription is NOT listed here: Product Name is only its
    # first choice — it falls back to Item Number, which the GTM sheet has.
    "Product Name":      ["AT_PrincipalMerchandiseHierarchyL3"],
    "NRF Color":         ["AT_PrincipalColorCode", "AT_PrincipalColorName",
                          "AT_Color"],
    # AT_PrincipalGenderDescription moved OFF Size Profile — MAA sources it from
    # "Adult Gender/ Kids Closure Type" and wants it as a plain text value.
    "Size Profile":      ["AT_PrincipalGenderCode",
                          "AT_PrincipalAgeCode", "AT_PrincipalAgeDescription"],
    "Country of Origin": ["AT_CountryOrigin"],
    "Retail Price":      ["AT_OriginalPrice", "AT_CurrentPrice"],
    "NBIL Price":        ["AT_FOB"],
    "Segment":           ["AT_MerchandiseCategory"],
    "Product Line":      ["AT_PrincipalMerchandiseHierarchyL1"],
    "Sizes - Region":    [],   # fed size variants, which are disabled anyway
    "Intro Period":      [],   # mapped but never emitted by the price-list ETL
    "Factory Name":      [],   # mapped but never emitted by the price-list ETL
}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES  (identical to inline_main_footwear.py — only the tables
#              still reachable from the surviving columns are kept)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "NEW": "NEW BALANCE", "ADI": "ADIDAS", "NIK": "NIKE",
    "SMI": "SMIGGLE",     "ALD": "ALDO",   "CRO": "CROCS",
    "LOT": "LOTTO",       "BIR": "BIRKENSTOCK",
}

LOV_SAP_ARTICLE_CATEGORY = {
    "1": "Single Article", "2": "Prepack", "3": "Display", "10": "Structured Article",
}
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "Licensed": "Licensed", "Basic": "Basic",
}
LOV_SEASON = {
    "SS": "Spring Summer", "FW": "Fall Winter",
    "AW": "Autumn Winter", "HO": "Holiday",
}
LOV_COMPANY_CODE = {
    "0888": "MAP Active Indonesia", "0000": "Default",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
    "FW": "Footwear",
}

# ── Footwear: Line Plan Business → Sports Category ───────────────────────────
# Kept identical to the App/Acc ETL's table (new_balance/inline_main_accessories.py)
# per MAA instruction — same source column ("Line Plan Business"), same mapping.
LOV_LINE_PLAN_TO_SPORTS_CAT: dict[str, str] = {
    "RUNNING":               "Running",
    "TRAINING":              "Running",
    "OTHER FOP":             "Other",
    "LIFESTYLE":             "Lifestyle / Casual",
    "BASKETBALL":            "Basketball",
    "KIDS LIFESTYLE":        "Lifestyle / Casual",
    "TENNIS":                "Tennis / Padel",
    "KIDS PERFORMANCE":      "Running",
    "BASKETBALL AND SOFTBALL": "Other",
    "SKATE":                 "Skateboarding",
    "SOCCER":                "Soccer",
    "GOLF":                  "Golf",
    "WALKING":               "Walking",
    "BADMINTON":             "Badminton",
}

# ── Sports Category → MDD LOV ID  (MDD sheet "Sports Category LOV") ──────────
# AT_SportsCategoryEN is an LOV attribute, so the XML carries the ID, not the
# display name. Value IDs are ZERO-PADDED for 1-9 ("01".."09"), matching the
# App/Acc ETL's table (new_balance/inline_appacc_preline_main.py) and the
# live Stibo "Sports Category" LOV.
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

DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E", "FOOTWEAR": "F",
    "APPAREL": "A", "EQUIPMENT": "Q", "TOYS": "T",
    "SHOES": "F",
}

# The GTM sheet is footwear-only (no Product Line column), so the division is
# fixed by the file type rather than read from a column.
DIVISION = "Footwear"

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

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


class AttributesListLoader:
    def __init__(self, path: Path, brand: str = "NEW BALANCE"):
        self.path  = path
        self.brand = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = (
            next((s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()), None)
            or next((s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()), wb.sheetnames[0])
        )
        log.info("[AttrList] Using sheet: '%s'", sheet_name)
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "Attributes"), None
        )
        if hdr_idx is None:
            wb.close()
            return
        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}
        for row in rows[hdr_idx + 1:]:
            attr = row[col.get("Attributes", 1)] if len(row) > col.get("Attributes", 1) else None
            if not attr or str(attr).strip() in ("", "None"):
                continue
            self.attr_map.append({
                "attribute":      str(attr).strip(),
                "cluster":        self._s(row, col, "Cluster"),
                "indicator":      self._s(row, col, "Indicator"),
                "mapping_logic":  self._s(row, col, "Field Name / Mapping Logic"),
                "from_pricelist": bool(self._s(row, col, "Price List")),
            })
        wb.close()
        log.info("[AttrList] %d attributes loaded", len(self.attr_map))

    @staticmethod
    def _s(row, col_map, key):
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        return str(v).strip() if v else None


class GTMFootwearLineListLoader:
    """
    Loads the New Balance APAC Footwear GTM Line Sheet.

    The sheet carries a merged title row ("S2'27 APAC Footwear GTM 2 Line
    Sheet") followed by blank rows, so the header row is located dynamically by
    scanning for "Item Number" — same strategy as FootwearPriceListLoader.
    """
    SHEET_NAMES = ["APAC", "Line Sheet", "Line List", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[GTM-FW] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next((s for s in self.SHEET_NAMES if s in wb.sheetnames), None)
        if target is None:
            target = max(wb.sheetnames, key=lambda s: wb[s].max_row or 0)
        log.info("[GTM-FW] Using sheet: '%s'", target)

        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(str(v).strip() == "Item Number" for v in r if v)),
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
        # (e.g. row 833) reliably even if blank rows get dropped afterwards.
        df["_ExcelRow"] = range(hdr_idx + 2, hdr_idx + 2 + len(df))

        item_col = next(
            (c for c in ["Item Number", "Item No.", "Item No"] if c in df.columns), None
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
    return "" if s in ("None", "nan", "0", "N/A") else s


def _line_plan_to_sports_cat(line_plan: str) -> str:
    return LOV_LINE_PLAN_TO_SPORTS_CAT.get(line_plan.upper().strip(), "Other")


def _sports_cat_to_lov_id(sports_cat: str) -> str:
    """Sports Category display name → MDD LOV ID. Unmapped → "16" (Other)."""
    return LOV_SPORTS_CATEGORY_ID.get(sports_cat.strip(), "16")


def _build_generic_code(brand_code: str, item_number: str) -> str:
    """
    Price-list ETL builds this as brand(3) + Product Number(9) + NRF colour
    suffix. The GTM sheet has no colour column and its Product Number is
    style-level (2781 rows → 468 Product Numbers), so the colourway-unique
    Item Number carries the key instead. Keeps the key 1:1 with a source row.
    """
    brand = re.sub(r"[^A-Z0-9]", "", brand_code.upper())[:3]
    base  = re.sub(r"[^A-Z0-9]", "", item_number.upper())
    return f"{brand}{base}"


def map_sku(row: pd.Series, brand_code: str = "NEW") -> dict:
    item_number    = _s(row.get("Item Number"))
    product_number = _s(row.get("Product Number"))
    line_plan_biz  = _s(row.get("Line Plan Business", ""))
    gender_raw     = _s(row.get("Adult Gender/ Kids Closure Type", ""))
    category       = _s(row.get("CATEGORY", ""))
    excel_row      = row.get("_ExcelRow")

    return {
        "item_number":    item_number,
        "product_number": product_number,
        # No Product Name column in the GTM sheet, so the price-list ETL's
        # `display_name or item_number` fallback always resolves to Item Number.
        "display_name":   "",
        "brand_code":     brand_code,
        "line_plan_biz":  line_plan_biz,
        "gender_raw":     gender_raw,
        "category":       category,
        "sports_cat":     _line_plan_to_sports_cat(line_plan_biz),
        "generic_code":   _build_generic_code(brand_code, item_number),
        "division":       DIVISION,
        "channel_type":   "Inline",
        # True 1-based Excel row number this SKU came from (see
        # GTMFootwearLineListLoader._load "_ExcelRow") — used by the run()
        # test-limit filter to target exact spreadsheet rows.
        "excel_row":      int(excel_row) if excel_row is not None else None,
        "article_type":   "Inline",
        "art_category":   "1",
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """
    Only validates fields this ETL actually emits. Attributes dropped for lack
    of a source column are reported once per run by _log_dropped_attributes(),
    not once per SKU.
    """
    warns = []
    sku   = mapped["item_number"] or "<no item number>"

    field_to_at = {
        "item_number":    "AT_PrincipalStyleCode",
        "brand_code":     "AT_Brand",
        "sports_cat":     "AT_SportsCategoryEN",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {})
        if "mandatory" in (meta.get("cardinality") or "").lower():
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{sku}] MISSING mandatory: {at_id}")

    if not mapped.get("line_plan_biz"):
        warns.append(f"[{sku}] MISSING source value: Line Plan Business")

    return warns


def _log_dropped_attributes(mdd: MDDLoader) -> list[str]:
    """Emit a one-time, run-level report of every attribute not carried over."""
    notes: list[str] = []
    notes.append(
        "Attributes NOT emitted — source column absent from the GTM line sheet:"
    )
    for src_col, attrs in DROPPED_ATTRIBUTES.items():
        if not attrs:
            continue
        for at_id in attrs:
            card = (mdd.attributes.get(at_id, {}).get("cardinality") or "n/a").strip()
            notes.append(f"  {at_id:38} (source: {src_col})  cardinality={card}")
            if "mandatory" in card.lower():
                log.warning(
                    "MANDATORY attribute %s dropped — source column '%s' "
                    "is not in the GTM line sheet.", at_id, src_col,
                )
    notes.append("")
    notes.append("Attributes withdrawn at MAA's request (source data not the issue):")
    for at_id, why in REMOVED_BY_REQUEST.items():
        notes.append(f"  {at_id:38} {why}")
    for line in notes:
        log.info(line)
    return notes


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER HELPERS
# ══════════════════════════════════════════════════════════════════

def _clean_xml_value(raw) -> str:
    text = "" if raw is None else str(raw).strip()
    return "" if text in ("", "None", "nan", "NaT") else text


def _val(parent, attr_id, value="", id_val=""):
    normalized_id = _clean_xml_value(id_val)
    normalized_value = _clean_xml_value(value)
    if not normalized_id and not normalized_value:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if normalized_id:
        el.set("ID", normalized_id)
        return el
    if normalized_value:
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
    pl_upper = (product_line or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in pl_upper:
            return code
    return "F"


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC VALUE WRITER
# ══════════════════════════════════════════════════════════════════

def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None):
    """
    Writes only:
      • attributes whose source column exists in the GTM line sheet, and
      • attributes with no source-column dependency (context + constants).

    Every other attribute the price-list ETL writes is intentionally absent —
    see DROPPED_ATTRIBUTES (source column missing) and REMOVED_BY_REQUEST
    (withdrawn on MAA instruction) at the top of this module.
    """
    # ── Context (from filename metadata, not a source column) ─────
    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand",      b_label, id_val=b_code)
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)

    # ── Item Number ───────────────────────────────────────────────
    _val(vals_el, "AT_PrincipalStyleCode", art["item_number"])
    # Same expression as the price-list ETL. display_name is always "" here
    # (no Product Name column), so this resolves to Item Number.
    _val(vals_el, "AT_PrincipalStyleDescription",
         art.get("display_name") or art["item_number"])

    # ── Adult Gender/ Kids Closure Type → Principal Gender Description
    # Sent as a plain text VALUE (not a LOV ID). Truncated to the MDD
    # max_chars, same as the footwear price-list ETL does.
    _gender_max = 20
    if mdd is not None:
        _meta = mdd.attributes.get("AT_PrincipalGenderDescription", {})
        _gender_max = int(_meta.get("max_chars") or 20)
    _val(vals_el, "AT_PrincipalGenderDescription", art.get("gender_raw", "")[:_gender_max])

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

    # ── Article category & type (constants) ───────────────────────
    cat_code  = art["art_category"]
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, cat_code)
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)
    at_norm  = art["article_type"].capitalize()
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)

    # ── Material Type — ZINA (constant) ───────────────────────────
    _val(vals_el, "AT_MaterialType", id_val="ZINA")

    # ── Line Plan Business ────────────────────────────────────────
    _val(vals_el, "AT_SportsCategoryEN", art["sports_cat"],
         id_val=_sports_cat_to_lov_id(art["sports_cat"]))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))

    # ── Derived key ────────────────────────────────────────────────
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
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
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

def build_product_xml(art, brand, brand_code, comp_code, sbu, season_id, mdd=None):
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

    # No Product Name column — same `display_name or item_number` fallback the
    # price-list ETL applies to <Name> and AT_PrincipalStyleDescription.
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art.get("display_name") or item_number
    )

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand.replace(' ', '')}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

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
    def first(d: Path, ext="*.xlsx"):
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f    = first(MDD_DIR)
    attr_f   = first(ATTR_DIR)
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    for label, val in [
        ("MDD",           mdd_f),
        ("Attributes",    attr_f),
        ("GTM Line List", ll_files),
    ]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    mdd    = MDDLoader(mdd_f)
    _      = AttributesListLoader(attr_f, brand=args.brand)

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
                    args.comp_code, args.sbu, season_id, mdd=mdd,
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

    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All SKUs passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 10 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — New Balance Footwear GTM Line Sheet v1.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0000")
    p.add_argument("--sbu",        default="FW")
    p.add_argument("--season",     default="SS27")
    p.add_argument("--seq",        default=1, type=int)
    p.add_argument("--country-code", default="")
    run(p.parse_args())


if __name__ == "__main__":
    main()
