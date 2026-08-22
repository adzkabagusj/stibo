"""
╔══════════════════════════════════════════════════════════════════╗
║      STIBO INBOUND XML GENERATOR — Smiggle  v1.0               ║
║  Packing List (EAN Source / 2-month before incoming) → STEP XML ║
╚══════════════════════════════════════════════════════════════════╝

Smiggle Packing List specifics vs Linelist:
  • Source file   : "Packing List (2m before incoming)" / EAN Source file
  • Active sheet  : "Order Details"  (summary sheet is ignored for ETL)
  • Header row    : row index 1 (0-based), first row is blank
  • Key columns:
      LINE #                  → Principal Style Code  (article_no)
      DESCRIPTION             → Principal Style Description / model_name
      COLOUR                  → Colour (free-text)
      EAN                     → Principal Barcode  (AT_PrincipalBarcode)
      HS CODE                 → HS Code             (AT_HSCode)
      DEPARTMENT              → SAP Product Division
      CATEGORY                → SAP Product Group
      SUB CAT                 → SAP Product Category / Interest / Material
      COLLECTION              → Collection 1
      AGE                     → SAP Age (default CH)
      IP OWNER                → BY Article Type  (Licence → License / Proprietary → Inline)
      SPEC                    → Specification  (Fashion / Graphic)
      PORT                    → Country of Origin  (CN / VN)
      MRP                     → Retail Price
      MRP CURRENCY            → Retail Price Currency
      SMIGGLE SELL PRICE USD  → FOB USD
      HANDOVER DATE           → Handover Date  (informational)
      EMBARGO DATE            → Launching / Embargo Date
      SEASON                  → Season code  (overrides pipeline args when present)
      BUY TYPE                → Buy Type  (NEW / REBUY / SPLIT)
      CUSTOMER                → Customer / Channel
      SMIGGLE SO #            → Sales Order reference
       ORDER UNITS            → Order Quantity
      OUTER CARTON            → Units per outer carton
      TOTAL OUTER CARTONS     → Total cartons
      QA/TESTING              → QA flag
      ORDER NOTES             → Order notes

  • One Generic product node per unique (LINE # / COLOUR / EAN) combination.
    The EAN is included in the STEP key so that rebuys / splits of the same
    article with the same colour but different EAN are treated separately.
  • TBC / blank LINE # rows are skipped (same behaviour as linelist).
  • SAP Article Category = always Single (00) — accessories brand, no sizing.
  • SAP Size = always "000" / No Size.
  • COO derived from PORT column (all Smiggle ports → CN).

S3 Layout (consumed by lambda_function.py):
    raw/metadata/smiggle/{packing_list_file}.xlsx   ← Smiggle uploads here
    raw/metadata/mdd_file.xlsx                      ← shared global MDD
    raw/metadata/attributes_list.xlsx               ← shared global attr list

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
from collections import defaultdict
import openpyxl
import pandas as pd

import os

BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
PACKING_DIR  = INPUT_DIR / "packing"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [PACKING_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_packing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# LOV TABLES  (identical to linelist_main — shared brand vocabulary)
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

# Port of loading → ISO country code for Country of Origin
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

# Smiggle Age column value → SAP Age code
SMIGGLE_AGE_MAP: dict[str, str] = {
    "CORE":       "CH",
    "JUNIOR":     "CH",
    "TEENY TINY": "CH",
}

# Smiggle IP Owner column → BY Article Type LOV
# (packing list uses "IP OWNER" column rather than "Design")
SMIGGLE_IP_OWNER_TO_BY_ARTICLE_TYPE: dict[str, str] = {
    "LICENCE":     "License",
    "LICENSE":     "License",
    "PROPRIETARY": "Inline",
}

# Smiggle Department → SAP Product Division letter
SMIGGLE_DEPT_TO_DIVISION: dict[str, str] = {
    "PACK & CARRY": "E",
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

# Buy Type labels (NEW / REBUY / SPLIT → STEP LOV IDs)
SMIGGLE_BUY_TYPE_MAP: dict[str, str] = {
    "NEW":   "New",
    "REBUY": "Rebuy",
    "SPLIT": "Split",
}


def _colour_to_gender(colour: str) -> str:
    """Heuristic: derive SAP Gender code from Smiggle colour string."""
    c = (colour or "").upper()
    if any(k in c for k in ("PINK", "LILAC", "ROSE", "PURPLE")):
        return "F"
    return "U"


# ── Packing list season code normalisation ────────────────────────
# Source file uses short codes: "S25" / "S26" / "W25" / "W26"
# Map to standard: "SS25" / "SS26" / "FW25" / "FW26"
def _normalise_season(raw: str) -> str:
    """
    Convert Smiggle packing list season codes to standard pipeline format.
    S25 → SS25,  W25 → FW25,  SS26 → SS26  (passthrough if already full)
    """
    r = (raw or "").strip().upper()
    if re.match(r"^(SS|FW|AW|HO|AL)\d{2,4}$", r):
        return r
    m = re.match(r"^([SW])(\d{2})$", r)
    if m:
        prefix = "SS" if m.group(1) == "S" else "FW"
        return f"{prefix}{m.group(2)}"
    return r


def resolve_brand_lov_id(brand: str, brand_code: str, mdd: "MDDLoader" = None) -> str:
    """
    Resolve the canonical Brand LOV ID once, then reuse it everywhere IDs/keys are built.
    Falls back to the incoming brand_code if no MDD match is found.
    """
    input_code = (brand_code or "").strip().upper()
    brand_display = (brand or LOV_BRAND.get(input_code, input_code)).strip().upper()
    if not mdd:
        return input_code
    brand_lov = mdd.lovs.get("Brand", {})
    for display_name, lov_id in brand_lov.items():
        if str(display_name).strip().upper() == brand_display and lov_id:
            return str(lov_id).strip().upper()
    return input_code


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

    def _load_age_lov(self, wb):
        sheet_name = next((s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()), None)
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if len(rows) < 2:
            return

        sap_age_map = {}
        by_age_map  = {}

        for row in rows[1:]:
            if not row or not row[0]:
                continue
            age_display  = str(row[0]).strip()
            sap_age_code = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            by_age_val   = str(row[5]).strip() if len(row) > 5 and row[5] else ""

            if age_display and sap_age_code:
                sap_age_map[age_display] = sap_age_code
            if age_display and by_age_val:
                by_age_map[age_display] = by_age_val

        self.lovs["SAPAge"] = sap_age_map
        self.lovs["BYAge"]  = by_age_map
        log.info("[MDD] Age LOV loaded — SAPAge: %d entries, BYAge: %d entries",
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


class SmigglePackingListLoader:
    """
    Loads the Smiggle packing list workbook (EAN Source / 2-month before incoming).

    Active sheet: "Order Details"
    Header:       row index 1 (0-based); row 0 is blank.

    Each row represents one (LINE # / COLOUR / EAN) combination.
    Rows where LINE # is blank, None, or "TBC" are skipped.

    Columns used:
        FINANCIAL HANDOVER MONTH  FINANCIAL HANDOVER WEEK  SEASON
        BUY TYPE  CUSTOMER  SMIGGLE SO #  HANDOVER DATE  EMBARGO DATE
        COLLECTION  AGE  IP OWNER  SPEC  DEPARTMENT  CATEGORY  SUB CAT
        LINE #  DESCRIPTION  COLOUR   ORDER UNITS   OUTER CARTON
        TOTAL OUTER CARTONS  EAN  HS CODE  QA/TESTING  PORT
        MRP  MRP CURRENCY  SMIGGLE SELL PRICE USD  SALES VALUE IN USD
        ORDER NOTES
    """

    PREFERRED_SHEETS = ["Order Details", "OrderDetails", "Sheet1"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[PackingList-SMI] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target = next(
            (s for s in self.PREFERRED_SHEETS if s in wb.sheetnames), None
        )
        if target is None:
            # Fall back to sheet with most rows
            best, best_rows = wb.sheetnames[0], 0
            for sn in wb.sheetnames:
                n = wb[sn].max_row or 0
                if n > best_rows:
                    best, best_rows = sn, n
            target = best

        log.info("[PackingList-SMI] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Auto-detect header row: first row that has "LINE" and "EAN" in it
        hdr_idx = None
        for i, r in enumerate(rows):
            if any(isinstance(v, str) and "LINE" in v.upper() for v in r if v) and \
               any(isinstance(v, str) and "EAN" in v.upper() for v in r if v):
                hdr_idx = i
                break

        # Fallback: first row with ≥ 10 non-null string cells
        if hdr_idx is None:
            hdr_idx = next(
                (
                    i for i, r in enumerate(rows)
                    if sum(1 for v in r[:20] if isinstance(v, str) and v.strip()) >= 10
                ),
                None,
            )

        if hdr_idx is None:
            log.error("[PackingList-SMI] Cannot find header row")
            wb.close()
            return

        log.info("[PackingList-SMI] Header at row %d (0-indexed)", hdr_idx)

        # Build clean header list
        raw_header = rows[hdr_idx]
        header = [
            str(h).strip() if h else f"col_{i}"
            for i, h in enumerate(raw_header)
        ]
        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Normalise the LINE # column name (may have trailing space: "LINE # ")
        line_col = next(
            (c for c in df.columns if c.strip().upper().startswith("LINE")), None
        )
        if line_col:
            df = df[
                df[line_col].notna()
                & (~df[line_col].astype(str).str.strip().isin(["", "None", "nan", "TBC"]))
            ]
            df = df.rename(columns={line_col: "LINE #"})

        # Normalise ORDER UNITS column (may have leading/trailing spaces)
        units_col = next(
            (c for c in df.columns if "ORDER" in c.upper() and "UNIT" in c.upper()), None
        )
        if units_col and units_col != "ORDER UNITS":
            df = df.rename(columns={units_col: "ORDER UNITS"})

        wb.close()
        self.df    = df.reset_index(drop=True)
        self.sheet = target
        log.info(
            "[PackingList-SMI] %d rows loaded (TBC / blank LINE# skipped)", len(self.df)
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
    if hasattr(v, "strftime"):
        return v.strftime("%d-%b-%Y").lower()
    raw = _s(v)
    if not raw or raw in ("NO EMBARGO", "N/A"):
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d-%b-%Y").lower()
        except (ValueError, TypeError):
            pass
    return raw


def map_article_packing(pl_row: dict, brand_code: str = "SMI", brand_input_code: str = "SMI") -> dict:
    """
    Map one Smiggle packing list row → unified article dict.

    Key decisions vs linelist mapper:
      - article_no    = LINE # (integer cast to string, .0 stripped)
      - ean           = EAN column  → AT_PrincipalBarcode
      - hs_code       = HS CODE column → AT_HSCode
      - ip_owner      = IP OWNER  → BY Article Type (Licence/Proprietary)
      - specification = SPEC column  (Fashion / Graphic)
      - buy_type      = BUY TYPE column  (NEW / REBUY / SPLIT)
      - so_number     = SMIGGLE SO # column
      - order_units   = ORDER UNITS column
      - outer_carton  = OUTER CARTON column
      - total_cartons = TOTAL OUTER CARTONS column
      - sales_value   = SALES VALUE IN USD column
      - handover_date = HANDOVER DATE column
      - embargo_date  = EMBARGO DATE column  → AT_LaunchingDate
      - mrp           = MRP column  → Retail Price
      - mrp_currency  = MRP CURRENCY  → Retail Price Currency
      - fob_usd       = SMIGGLE SELL PRICE USD  → AT_FOB
      - season_raw    = SEASON column  (normalised to SS/FW prefix)
    """
    line_no = _s(pl_row.get("LINE #"))
    if line_no.endswith(".0"):
        line_no = line_no[:-2]

    desc         = _s(pl_row.get("DESCRIPTION"))
    colour       = _s(pl_row.get("COLOUR"))
    age_raw      = _s(pl_row.get("AGE"))
    ip_owner     = _s(pl_row.get("IP OWNER"))
    spec         = _s(pl_row.get("SPEC"))
    dept         = _s(pl_row.get("DEPARTMENT"))
    category     = _s(pl_row.get("CATEGORY"))
    sub_cat      = _s(pl_row.get("SUB CAT"))
    collection   = _s(pl_row.get("COLLECTION"))
    port         = _s(pl_row.get("PORT"))
    buy_type     = _s(pl_row.get("BUY TYPE"))
    customer     = _s(pl_row.get("CUSTOMER"))
    so_number    = _s(pl_row.get("SMIGGLE SO #"))
    order_units  = _s(pl_row.get("ORDER UNITS"))
    outer_carton = _s(pl_row.get("OUTER CARTON"))
    total_cartons = _s(pl_row.get("TOTAL OUTER CARTONS"))
    ean_raw      = pl_row.get("EAN")
    hs_code_raw  = pl_row.get("HS CODE")
    mrp_raw      = pl_row.get("MRP")
    mrp_currency = _s(pl_row.get("MRP CURRENCY"))
    fob_usd_raw  = pl_row.get("SMIGGLE SELL PRICE USD")
    sales_val_raw = pl_row.get("SALES VALUE IN USD")
    handover_raw = pl_row.get("HANDOVER DATE")
    embargo_raw  = pl_row.get("EMBARGO DATE")
    season_raw   = _s(pl_row.get("SEASON"))
    ho_month     = _s(pl_row.get("FINANCIAL HANDOVER MONTH"))
    ho_week      = _s(pl_row.get("FINANCIAL HANDOVER WEEK"))
    qa_testing   = _s(pl_row.get("QA/TESTING"))
    order_notes  = _s(pl_row.get("ORDER NOTES"))

    # ── EAN: ensure it's a clean string (may be a large integer) ─
    ean = ""
    if ean_raw is not None:
        ean_str = str(ean_raw).strip()
        if ean_str.endswith(".0"):
            ean_str = ean_str[:-2]
        ean = ean_str if ean_str not in ("", "None", "nan") else ""

    # ── HS Code ──────────────────────────────────────────────────
    hs_code = ""
    if hs_code_raw is not None:
        hs_str = str(hs_code_raw).strip()
        if hs_str.endswith(".0"):
            hs_str = hs_str[:-2]
        hs_code = hs_str if hs_str not in ("", "None", "nan") else ""

    # ── Derived fields ──────────────────────────────────────────
    gender_code = _colour_to_gender(colour)
    age_code    = SMIGGLE_AGE_MAP.get((age_raw or "").upper(), "CH")
    coo         = SMIGGLE_PORT_TO_COO.get((port or "").upper().strip(), "CN")

    # BY Article Type from IP OWNER column
    by_art_type = SMIGGLE_IP_OWNER_TO_BY_ARTICLE_TYPE.get(
        (ip_owner or "").upper().strip(), "Inline"
    )

    div_letter  = SMIGGLE_DEPT_TO_DIVISION.get(
        (dept or "").upper().strip(), "E"
    )

    handover_date = _fmt_date(handover_raw)
    embargo_date  = _fmt_date(embargo_raw)

    # Normalise season from packing list format (S25 → SS25)
    season_normalised = _normalise_season(season_raw)

    # MRP (retail price)
    mrp_str = ""
    if mrp_raw is not None:
        m = re.search(r"[\d.]+", str(mrp_raw))
        mrp_str = m.group() if m else str(mrp_raw)

    # FOB = SMIGGLE SELL PRICE USD
    fob_str = ""
    if fob_usd_raw is not None:
        m = re.search(r"[\d.]+", str(fob_usd_raw))
        fob_str = m.group() if m else str(fob_usd_raw)

    # Sales value (informational)
    sales_value = ""
    if sales_val_raw is not None:
        m = re.search(r"[\d.]+", str(sales_val_raw))
        sales_value = m.group() if m else str(sales_val_raw)

    # Generic code: Brand LOV ID + Line #
    generic_code = f"{brand_code}{line_no}"

    # Colour token (3 chars, alphanumeric, padded)
    colour_token = re.sub(r"[^A-Z0-9]", "", colour.upper())[:3].ljust(3, "0")

    # Variant key: Generic + colour token + "000" (no size)
    variant_code = f"{generic_code}{colour_token}000"

    # EAN token for STEP key uniqueness
    ean_token = re.sub(r"[^0-9]", "", ean)[-4:] if ean else "0000"

    return {
        # Core identifiers
        "article_no":       line_no,
        "model_name":       desc,
        "brand_code":       brand_code,
        "brand_input_code": brand_input_code,

        # Colour
        "colour":           colour,
        "colour_token":     colour_token,

        # EAN / barcode
        "ean":              ean,

        # HS Code
        "hs_code":          hs_code,

        # Gender / Age
        "gender_code":      gender_code,
        "gender_raw":       colour,
        "age_code":         age_code,
        "age_raw":          age_raw,

        # Classification
        "division":         dept,
        "div_letter":       div_letter,
        "category":         category,
        "sub_category":     sub_cat,
        "specification":    spec,

        # Commercial
        "article_type":     by_art_type,
        "art_category":     "0",            # always Single
        "collection1":      collection,
        "collection2":      sub_cat,
        "franchise":        ip_owner,       # IP OWNER used as franchise label
        "embargo_date":     embargo_date,
        "handover_date":    handover_date,

        # Packing-specific
        "buy_type":         SMIGGLE_BUY_TYPE_MAP.get((buy_type or "").upper(), buy_type or ""),
        "customer":         customer,
        "so_number":        so_number,
        "order_units":      order_units,
        "outer_carton":     outer_carton,
        "total_cartons":    total_cartons,
        "ho_month":         ho_month,
        "ho_week":          ho_week,
        "qa_testing":       qa_testing,
        "order_notes":      order_notes,
        "sales_value":      sales_value,

        # Pricing
        "rrp":              mrp_str,
        "currency":         mrp_currency or "SGD",
        "fob":              fob_str,

        # Origin
        "coo":              coo,
        "port":             port,

        # Season (normalised)
        "season_source":    season_normalised,

        # Derived codes
        "generic_code":     generic_code,
        "variant_code":     variant_code,
        "sap_style_code":   line_no,

        # No TDD / Backlog data for Smiggle packing list
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

    # Packing-list specific: warn if EAN missing
    if not mapped.get("ean"):
        warns.append(f"[{art}] MISSING EAN — AT_PrincipalBarcode will be empty")

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
    clean_id = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip() if value else ""
    if not clean_id and (not clean_val or clean_val in ("None", "nan")):
        return None
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
        return el
    el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = "") -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _lov(code: str, lookup: dict, default: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default or code)
    return code, label


# ── Generic-level attribute placeholder list (packing list superset)
# Extends the linelist placeholder list with packing-specific fields.
SMIGGLE_PACKING_PLACEHOLDERS = [
    "AT_SPUGrouping",           "AT_SPUGroupingName",
    "AT_Collection1",           "AT_Collection2",
    "AT_EComProductNameEN",     "AT_EComProductNameID",
    "AT_EComProductNamePH",     "AT_EComProductNameMY",
    "AT_EComProductNameVN",     "AT_EComProductNameTH",
    "AT_EComProductNameKH",
    "AT_ShortDescriptionEN",    "AT_ShortDescriptionID",
    "AT_ShortDescriptionPH",    "AT_ShortDescriptionMY",
    "AT_ShortDescriptionVN",    "AT_ShortDescriptionTH",
    "AT_ShortDescriptionKH",
    "AT_LongDescriptionEN",     "AT_LongDescriptionID",
    "AT_LongDescriptionPH",     "AT_LongDescriptionMY",
    "AT_LongDescriptionVN",     "AT_LongDescriptionTH",
    "AT_LongDescriptionKH",
    "AT_CareInstructionEN",     "AT_CareInstructionID",
    "AT_CareInstructionPH",     "AT_CareInstructionMY",
    "AT_CareInstructionVN",     "AT_CareInstructionTH",
    "AT_CareInstructionKH",
    "AT_PrincipalColorName",
    "AT_EComAgesCategory",
    "AT_Interest",
    "AT_Material",
    "AT_CountrySize",
    "AT_TechnologyUsed",
    "AT_CertificateNumber",
    "AT_ProductWeight",
    "AT_ProductLengthWidthHeight",
    "AT_PackagingLength",       "AT_PackagingWidth",
    "AT_PackagingHeight",       "AT_PackagingWeight",
    "AT_ImagesSource",
    "AT_EstimatedLandedCost",
    "AT_ExchangeRate",
    "AT_FOBCurrency",
    "AT_MainEANIndicator",
    "AT_Royalty",
    "AT_FreightCost",
    "AT_MarketingFee",
    "AT_SGS",
    "AT_HaddadOfficeCharge",
    "AT_HaddadOfficeChargeDeduction",
    "AT_Others",
    "AT_Commission",
    "AT_LaunchingDate",
    "AT_Createdon",
    "AT_Updatedon",
    "AT_Updatedby",
]


def _build_data_containers(art: dict) -> ET.Element | None:
    """
    Build <DataContainers> block for a Smiggle packing list article.
    Creates one DC_Barcode DataContainer per EAN on the product.
    
    Structure:
    <DataContainers>
      <MultiDataContainer Type="DC_Barcode">
        <DataContainer ID="{ean}">
          <Values>
            <Value AttributeID="AT_Barcode">{ean}</Value>
            <Value AttributeID="AT_BarcodeType" ID="P"/>
            <Value AttributeID="AT_MainEANIndicator" ID="Y"/>
          </Values>
        </DataContainer>
        ... (one per EAN)
      </MultiDataContainer>
    </DataContainers>
    """
    ean = art.get("ean", "")
    if not ean:
        return None

    # Collect all EANs — currently one per row, but future-proofed
    # for multi-EAN products by checking art.get("eans", [ean])
    eans = art.get("eans", [ean])
    eans = [e for e in eans if e]  # filter empties
    if not eans:
        return None

    dc_root = ET.Element(f"{{{STIBO_NS}}}DataContainers")
    mdc = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")

    for ean_val in eans:
        dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")

        vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

        barcode_val = ET.SubElement(vals, f"{{{STIBO_NS}}}Value")
        barcode_val.set("AttributeID", "AT_Barcode")
        barcode_val.text = ean_val

        barcode_type = ET.SubElement(vals, f"{{{STIBO_NS}}}Value")
        barcode_type.set("AttributeID", "AT_BarcodeType")
        barcode_type.set("ID", "P")

        main_ean_indicator = ET.SubElement(vals, f"{{{STIBO_NS}}}Value")
        main_ean_indicator.set("AttributeID", "AT_MainEANIndicator")
        main_ean_indicator.set("ID", "Y")

    return dc_root

    

def _add_packing_values(
    vals_el:    ET.Element,
    art:        dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd=None,  
) -> None:
    """
    Write all Generic-level <Value> elements for a Smiggle packing list article.

    Extends linelist attribute writing with packing-specific fields:
      AT_PrincipalBarcode  ← EAN
      AT_HSCode            ← HS CODE
      AT_BuyType           ← BUY TYPE
      AT_SONumber          ← SMIGGLE SO #
      AT_OrderQuantity     ← ORDER UNITS
      AT_OuterCarton       ← OUTER CARTON
      AT_TotalCartons      ← TOTAL OUTER CARTONS
      AT_HandoverDate      ← HANDOVER DATE (informational)
      AT_SalesValueUSD     ← SALES VALUE IN USD
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
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    b_code, b_label = _lov(art["brand_input_code"], LOV_BRAND, art["brand_input_code"])
    brand_display = b_label.upper()
    brand_lov_id  = b_code
    if mdd:
        brand_lov = mdd.lovs.get("Brand", {})
        for display_name, lov_id in brand_lov.items():
            if display_name.upper() == brand_display:
                brand_lov_id = lov_id
                break
    _w("AT_Brand",      id_val=brand_lov_id)
    _w("AT_BrandGroup", id_val=b_label)

    # ── Principal identifiers ────────────────────────────────────
    _w("AT_PrincipalStyleCode",        art["article_no"])
    _w("AT_PrincipalStyleDescription", art["model_name"])
    _w("AT_PrincipalColorCode",        art["colour"])
    _w("AT_SAPStyleCode",              art["sap_style_code"])
   

    # ── Barcode (EAN) — key packing list field ───────────────────
    _w("AT_PrincipalBarcode", art["ean"])

    # ── HS Code ──────────────────────────────────────────────────
    _w("AT_HSCode", art["hs_code"])


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
    at_generic_val = f"{art['brand_code']}{art['article_no']}{sap_color_code}"

    # Store on art so build_product_xml can use it for KEY_Article
    art["at_generic_val"] = at_generic_val

    _w("AT_InboundGenericCode", at_generic_val)

    # ── Gender ───────────────────────────────────────────────────
    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _w("AT_Gender",          g_label, id_val=g_code)
    _w("AT_BYGender",        g_label, id_val=g_code)
    _w("AT_PrincipalGenderDescription", art["gender_raw"])

    # ── Age ──────────────────────────────────────────────────────
    age_raw_upper = (art["age_raw"] or "").strip()
    sap_age_map   = mdd.lovs.get("SAPAge", {}) if mdd else {}
    by_age_map    = mdd.lovs.get("BYAge",  {}) if mdd else {}

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
    _w("AT_BYAge", by_age_val.upper(), id_val=by_age_val.upper())
    _w("AT_PrincipalAgeDescription", art["age_raw"] or sap_age_code)

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
    coo_code, coo_label = _lov(art["coo"], LOV_COUNTRY_ORIGIN, art["coo"])
    _w("AT_CountryOrigin", coo_label, id_val=coo_code)

    # ── SAP Article Category — always Single (00) ────────────────
    _w("AT_SAPArticleCategory", "Single", id_val="0")

    # ── BY Article Type (prefer filename-derived Inline/License) ──
    at_code = art["article_type"]
    at_from_filename = (art.get("article_type_from_filename") or "").strip()
    if at_from_filename:
        at_code = "License" if at_from_filename.lower().startswith("lic") else "Inline"
    _w("AT_BYArticleType", at_code, id_val=at_code)

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   "Yes", id_val="Y")
    _w("AT_SAPIndicator",  "No",  id_val="N")
    _w("AT_EcomIndicator", "")

    country_val = (art.get("country_code") or "").strip().upper()
    if country_val:
        _w("AT_Country", country_val, id_val=country_val)

    # ── UOM / Size — fixed for Smiggle ───────────────────────────
    _w("AT_UOM",           "Each",    id_val="EA")
    _w("AT_Size",          "No Size", id_val="000")
    _w("AT_PrincipalSize", "One Size")

    # ── Pricing ──────────────────────────────────────────────────
    _w("AT_OriginalPrice",       art["rrp"])
    _w("AT_CurrentPrice",        art["rrp"])
    _w("AT_RetailPriceCurrency", art["currency"], id_val=art["currency"])
    _w("AT_FOB",                 art["fob"])
    _w("AT_FOBCurrency",         "USD", id_val="USD")

    # ── Nature of Article ────────────────────────────────────────
    _w("AT_NatureOfArticle", "Regular", id_val="REG")

    # ── Vendor ───────────────────────────────────────────────────
    _w("AT_MainVendorIdentification", "1")

    # ── Collections / hierarchy ──────────────────────────────────
    _w("AT_Collection1", art["collection1"])
    _w("AT_Collection2", art["collection2"])

    # ── EcomProductName EN ───────────────────────────────────────
    ecom_name_en = f"Smiggle {art['model_name']}".strip() if art["model_name"] else ""
    _w("AT_EComProductNameEN", ecom_name_en)

    # ── Launching / Embargo date ─────────────────────────────────
    _w("AT_LaunchingDate", art["embargo_date"])

    # ── Interest / Material (Sub Category) ───────────────────────
    _w("AT_Interest", art["sub_category"])
    _w("AT_Material", art["sub_category"])

    # ── Country Size ─────────────────────────────────────────────
    _w("AT_CountrySize", "ONE SIZE")

    # ── Packing-list specific attributes ─────────────────────────
    # _w("AT_BuyType",         art["buy_type"])
    # _w("AT_SONumber",        art["so_number"])
    # _w("AT_OrderQuantity",   art["order_units"])
    # _w("AT_OuterCarton",     art["outer_carton"])
    # _w("AT_TotalCartons",    art["total_cartons"])
    # _w("AT_HandoverDate",    art["handover_date"])
    # _w("AT_SalesValueUSD",   art["sales_value"])

    # ── Remaining placeholders ───────────────────────────────────
    for attr_id in SMIGGLE_PACKING_PLACEHOLDERS:
        if attr_id not in written:
            created = _val(vals_el, attr_id)
            if created is not None:
                written.add(attr_id)


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — CLASSIFICATIONS + PRODUCT XML BUILDERS
# ══════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str, brand_lov_id: str, comp_code: str, sbu: str, season_code: str
) -> ET.Element:
    """Build the <Classifications> block."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short

    full_season_code = f"{sea_prefix}{sea_year}"
    season_id        = f"CLH_{brand_lov_id}_{full_season_code}"
    batches_parent   = f"CLH_{brand_lov_id}Batches"

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
    Build a <Product> XML fragment for one Smiggle packing list article.

    UserTypeID = PRD_SingleArticle  (Single article category, no size run)
    ParentID   = PPH_E-TempSubCat   (Equipment / Accessories)
    KeyID      = KEY_Article, value = {brand_code}{Line#}{colour_token}{ean_suffix}
    """
    article_no = art["article_no"]
    if not article_no:
        return ""

    b_code     = art.get("brand_code", brand_code) or brand_code
    div_letter = art.get("div_letter", "E")
    parent_id  = f"PPH_{div_letter}-TempSubCat"

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
    key_article = f"{art['brand_code']}{art['article_no']}{sap_color_for_key}"

    cat_code = SMIGGLE_CATEGORY_TO_PROD_GROUP.get(
        (art.get("category") or "").upper().strip(), "AC"
    )

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
    cr_merch.set("ClassificationID", f"CLH_{b_code}Articles")
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
    _add_packing_values(vals_el, art, brand, comp_code, sbu,mdd=mdd)


    # ── DataContainers (Barcode / EAN) ───────────────────────────
    dc_el = _build_data_containers(art)
    if dc_el is not None:
        g_el.append(dc_el)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — THREAD WORKER
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    """Thread worker: map + validate one packing list row."""
    row, brand_code, brand_input_code, mdd, season, country_code, article_type_from_filename = row_tuple
    mapped = map_article_packing(row, brand_code=brand_code, brand_input_code=brand_input_code)
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
    Main entry point for packing list ETL.

    args must have attributes:
        brand       str   e.g. "Smiggle"
        brand_code  str   e.g. "SMI"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SM"
        season      str   e.g. "SS25"  (already normalised by lambda_function.py)
        seq         int   e.g. 1
    """

    def first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f     = first(MDD_DIR)
    attr_f    = first(ATTR_DIR)
    pl_files  = list(PACKING_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Packing List", pl_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd    = MDDLoader(mdd_f)
    _al    = AttributesListLoader(attr_f, brand=args.brand)

    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
        args.brand, args.brand_code, args.comp_code, args.sbu, args.season, args.seq,
    )

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    resolved_brand_lov_id = resolve_brand_lov_id(args.brand, args.brand_code, mdd)
    season_id      = f"CLH_{resolved_brand_lov_id}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for pl_path in pl_files:
        log.info("─── Processing Packing List: %s ───", pl_path.name)

        pl = SmigglePackingListLoader(pl_path)
        if pl.df.empty:
            log.warning("[PackingList-SMI] Empty dataframe — skipping.")
            continue

        rows = [row for _, row in pl.df.iterrows()]
        # ── TEST MODE: limit to 1 product ───────────────────────
        # rows = rows[:10]
        # ────────────────────────────────────────────────────────


        # ── Pre-scan: group all EANs per generic_code + colour ───
        ean_map: dict[str, list[str]] = defaultdict(list)
        for row in rows:
            line_no = str(row.get("LINE #", "") or "").strip()
            if line_no.endswith(".0"):
                line_no = line_no[:-2]
            colour = _s(row.get("COLOUR")).upper()
            ean_raw = row.get("EAN")
            if ean_raw:
                ean_str = str(ean_raw).strip()
                if ean_str.endswith(".0"):
                    ean_str = ean_str[:-2]
                key = f"{resolved_brand_lov_id}{line_no}|{colour}"
                if ean_str and ean_str not in ean_map[key]:
                    ean_map[key].append(ean_str)
        # ─────────────────────────────────────────────────────────


        total_rows = len(rows)
        log.info("Total valid rows (after TBC/blank filter): %d", total_rows)

        season_to_use = args.season

        # ── Output filename ──────────────────────────────────────
        # ── Parse file_type and multi_mono from input filename ───────
        _stem  = pl_path.stem
        _parts = re.split(r"\s*-\s*", _stem)

        # Structure: {CompCode}-{SBU}-{Brand}-{FileType (may have spaces)}-{Multi/Mono}-{Season}-{Seq}
        # Anchored:  [0]        [1]   [2]     [3:-3]                        [-3]         [-2]     [-1]
        if len(_parts) >= 7:
            file_type_from_input = "-".join(_parts[3:-3])
        else:
            file_type_from_input = "Packing List"

        # Multi/Mono: regex on full stem is more reliable than positional
        _mm_match = re.search(r'\b(Multi|Mono)\b', _stem, re.IGNORECASE)
        multi_mono_from_input = _mm_match.group(1).capitalize() if _mm_match else "Multi"

        log.info(
            "Parsed from filename → file_type='%s'  multi_mono='%s'",
            file_type_from_input, multi_mono_from_input,
        )

        # ── Output filename ──────────────────────────────────────────
        out_name = pl_path.name + ".xml"
        out_path = XML_OUT_DIR / out_name

        # ── Pass 1: parallel map + validate ─────────────────────
        num_workers = min(8, max(1, total_rows))
        task_args   = [
            (
                row,
                resolved_brand_lov_id,
                args.brand_code,
                mdd,
                season_to_use,
                getattr(args, "country_code", ""),
                getattr(args, "article_type_from_filename", ""),
            )
            for row in rows
        ]
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
                # Inject EAN list scoped to this specific article+colour
                generic_key = mapped.get("generic_code", "")
                colour_key = (mapped.get("colour") or "").upper()
                ean_key = f"{generic_key}|{colour_key}"
                mapped["eans"] = ean_map.get(ean_key, [mapped.get("ean", "")])
                ordered.append((idx, mapped, warns))

        ordered.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered]
        del ordered

        log.info("Pass 1 done — mapped=%d articles", len(mapped_articles))

        # ── Pass 2: stream XML to file ───────────────────────────
        log.info("Pass 2/2 — streaming XML to %s …", out_name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cls_el  = build_classifications(
            args.brand, resolved_brand_lov_id, args.comp_code, args.sbu, season_to_use,
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
                    art, args.brand, resolved_brand_lov_id,
                    args.comp_code, args.sbu, season_id,mdd=mdd
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

        print("═══ ARTICLE SUMMARY (SMIGGLE PACKING LIST) ════════════", flush=True)
        print(f"  Packing list rows (valid)  : {total_rows}",    flush=True)
        print(f"  Articles written           : {written_count}", flush=True)
        print(f"  XML file size              : {file_kb}KB",     flush=True)
        print("═══════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_packing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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
        description="Stibo Inbound XML Generator — Smiggle Packing List v1.0"
    )
    p.add_argument("--brand",      default="Smiggle")
    p.add_argument("--brand-code", default="SMI")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SM")
    p.add_argument("--season",     default="SS25")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
