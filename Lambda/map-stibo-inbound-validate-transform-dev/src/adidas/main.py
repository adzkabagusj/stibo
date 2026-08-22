"""
╔══════════════════════════════════════════════════════════════════╗
║          STIBO INBOUND XML GENERATOR  — v3.2                   ║
║  Linelist (Multi) + Backlog + TDD → Stibo STEP XML             ║
║                                                                  ║
║  v3.2 changes vs v3.1:                                          ║
║  • LinelistLoader: strips (M)/(A) suffixes from column names    ║
║  • LinelistLoader: skips duplicate header row                   ║
║  • map_article: maps new SS27 column names                      ║
║    (Colorway Name, Country of Origin (A-Fact), IDN-Retail Price ║
║     SEA SEGMENTATION, IND SEGMENTATION, CONCEPT, SUB CONCEPT   ║
║     Retail Intro Date, Retail Exit Date, Hard Launch etc.)      ║
║  • LOV_COUNTRY_ORIGIN: full country name lookup added           ║
║  • _add_generic_values: new AT_ attributes for new columns      ║
║                                                                  ║
║  v3.2.1 fix:                                                     ║
║  • _s() and TDDLoader.g() now guard against pandas.NaT          ║
║    (NaT passes isinstance(v, datetime) but crashes on strftime) ║
╚══════════════════════════════════════════════════════════════════╝
"""

import re
import sys
import logging
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom

import openpyxl
import pandas as pd

import os
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
BACKLOG_DIR  = INPUT_DIR / "backlog"
TDD_DIR      = INPUT_DIR / "tdd"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, BACKLOG_DIR, TDD_DIR, MDD_DIR,
          ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
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

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}
LOV_GENDER   = {"M": "Male (M)", "F": "Female (W)", "W": "Female (W)", "U": "Unisex (U)"}
LOV_BY_GENDER = {"M": "Male (M)", "M": "Men (M)", "F": "Female (W)", "F": "Women (W)", "U": "Unisex (U)"}
LOV_AGE = {
    "AD": "Adult", "ADULTS": "Adult", "CH": "Children", "IN": "Children",
    "AA": "All Ages", "JR": "Children",
}
# BY Age ID Mapping: maps input age values to BY Age LOV IDs
BY_AGE_ID_MAP = {
    "ADULT":        "ADULT",
    "ADULT":       "ADULT",
    "AD":           "ADULT",
    "INFANT":       "INFANT",
    "JUNIOR":       "KIDS",
    "YOUTH":        "KIDS",
    "KIDS":         "KIDS",
    "KID":          "KIDS",
    "CHILDREN":     "KIDS",
    "CHILD":        "KIDS",
    "CH":           "KIDS",
    "ALL AGES":     "ALL AGES",
    "ALL":          "ALL AGES",
    "AA":           "ALL AGES",
    "GRADE SCHOOL": "GRADE SCHOOL",
    "PRESCHOOL":    "PRESCHOOL",
}
LOV_UOM = {
    "EA": "Each", "PR": "Pair", "SET": "Set", "SLV": "Sleeve",
    "PK": "Pack", "DZ": "Dozen",
}
LOV_BY_INDICATOR   = {"Yes": "Y", "No": "N", "Y": "Y", "N": "N"}
LOV_SAP_INDICATOR  = {"Yes": "Y", "No": "N", "Y": "Y", "N": "N"}
LOV_ECOM_INDICATOR = {"Yes": "Y", "No": "N", "Y": "Y", "N": "N"}
LOV_SAP_ARTICLE_CATEGORY = {"1": "Generic", "0": "Single", "10": "Sell set (Hampers)"}
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
    "Licensed": "License", "INLINE": "Inline",
}
LOV_SEASON = {
    "SS": "Spring-Summer", "FW": "Fall-Winter", "AL": "All Season",
    "SS26": "Spring-Summer", "SS25": "Spring-Summer", "SS27": "Spring-Summer",
    "25F": "Fall-Winter",   "26S": "Spring-Summer",
}

# ── v3.2: full country name lookup added (linelist uses full names) ──────────
LOV_COUNTRY_ORIGIN = {
    # 2-letter ISO codes
    "ID": "Indonesia",   "CN": "China",       "KH": "Cambodia",
    "VN": "Vietnam",     "BD": "Bangladesh",  "IN": "India",
    "MY": "Malaysia",    "TH": "Thailand",    "PH": "Philippines",
    "PK": "Pakistan",    "TR": "Turkey",
    # Full country names → normalise to same label, store ISO code separately
    "INDONESIA":   "Indonesia",   "CHINA":       "China",
    "CAMBODIA":    "Cambodia",    "VIETNAM":     "Vietnam",
    "BANGLADESH":  "Bangladesh",  "INDIA":       "India",
    "MALAYSIA":    "Malaysia",    "THAILAND":    "Thailand",
    "PHILIPPINES": "Philippines", "PAKISTAN":    "Pakistan",
    "TURKEY":      "Turkey",
}
# Full name → ISO code (for ID attribute on <Value>)
_COUNTRY_NAME_TO_CODE = {
    "Indonesia": "ID", "China": "CN", "Cambodia": "KH", "Vietnam": "VN",
    "Bangladesh": "BD", "India": "IN", "Malaysia": "MY", "Thailand": "TH",
    "Philippines": "PH", "Pakistan": "PK", "Turkey": "TR",
}

LOV_COMPANY_CODE = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}
LOV_SBU = {
    "SP": "Sports", "FQ": "Footlocker", "FL": "Fashion Footwear",
}
LOV_SPORTS_CATEGORY_EN = {
    "MAA SPORTS CATEGORY": "MAA Sports Category",
    "FOOTBALL/SOCCER":     "Soccer",
    "BASKETBALL":          "Basketball",
    "RUNNING":             "Running",
    "TRAINING":            "Fitness / Training",
    "ORIGINALS":           "Lifestyle / Casual",
    "TENNIS":              "Tennis / Padel",
    "MOTORSPORT":          "Other",
    "HIKING":              "Outdoor / Trail / Hiking",
    "TRAIL RUNNING":       "Running",
    "NOT SPORTS SPECIFIC": "Other",
    "SKATEBOARDING":       "Skateboarding",
    "SWIM":                "Swimming",
    "OUTDOOR":             "Lifestyle / Casual",
    "PADEL":               "Tennis / Padel",
    "SPORTSWEAR":          "Lifestyle / Casual",
    "INDOOR":              "Tennis / Padel",
}

# Set to True once your Stibo admin creates these attributes in the system
INCLUDE_SS27_ATTRS = False


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
        wb   = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws   = wb["Core Attributes"]
        rows = list(ws.iter_rows(values_only=True))
        hdr  = rows[1]
        col  = {}
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
        self._load_gender_lov(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_gender_lov(self, wb):
        """
        Load Gender LOV: Col A = display value (e.g. 'Male'),
                         Col C = LOV ID (e.g. 'M').
        Also registers common aliases (W, M, U, MEN, WOMEN, FEMALE, etc.).
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "GENDER" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Gender LOV sheet not found")
            return
        gender_lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or not row[0]:
                continue
            display_value = str(row[0]).strip()           # Col A = display value
            lov_id        = str(row[2]).strip() if len(row) > 2 and row[2] else ""  # Col C = LOV ID
            if display_value and lov_id:
                gender_lov[display_value.upper()] = lov_id
        # Register common aliases so short codes / alternate spellings resolve too
        for alias, canonical in [
            ("W",     "FEMALE"),
            ("F",     "FEMALE"),
            ("M",     "MALE"),
            ("MEN",   "MALE"),
            ("WOMEN", "FEMALE"),
            ("U",     "UNISEX"),
        ]:
            if canonical in gender_lov:
                gender_lov.setdefault(alias, gender_lov[canonical])
        self.lovs["GenderLOV"] = gender_lov
        log.info("[MDD] Gender LOV loaded: %d entries", len(gender_lov))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[str(val_name).strip()] = \
                    str(val_id).strip() if val_id else str(val_name).strip()

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
                    self.lovs.setdefault(display, {})[str(name).strip()] = \
                        str(code).strip() if code else str(name).strip()


class AttributesListLoader:
    def __init__(self, path: Path, brand: str = "ADIDAS"):
        self.path  = path
        self.brand = brand.upper()
        self.attr_map: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrList] Loading brand '%s' from: %s", self.brand, self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = next(
            (s for s in wb.sheetnames if s.upper() == self.brand and "V4" not in s.upper()), None
        ) or next(
            (s for s in wb.sheetnames if self.brand in s.upper() and "V4" not in s.upper()),
            wb.sheetnames[0]
        )
        ws      = wb[sheet_name]
        rows    = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows) if len(r) > 0 and r[0] == "Attributes"), None
        ) or next(
            (i for i, r in enumerate(rows) if any(v == "Attributes" for v in r[:3] if v)), None
        )
        if hdr_idx is None:
            wb.close(); return
        hdr = rows[hdr_idx]
        col = {str(h).strip(): i for i, h in enumerate(hdr) if h}
        for row in rows[hdr_idx + 1:]:
            attr = row[col.get("Attributes", 0)] if col.get("Attributes", 0) < len(row) else None
            if not attr:
                continue
            self.attr_map.append({
                "attribute":     str(attr).strip(),
                "cluster":       self._s(row, col, "Cluster"),
                "indicator":     self._s(row, col, "Indicator"),
                "mapping_logic": self._s(row, col, "Field Name / Mapping Logic"),
                "from_linelist": bool(self._s(row, col, "Linelist")),
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
def _norm_col(c) -> str:
    """
    Strip trailing (M) / (A) / (M-Fact) / (A-Fact) suffixes and whitespace.
    Examples:
        "Model Name (M)"              → "Model Name"
        "Country of Origin (A-Fact)"  → "Country of Origin"
        "Age Group (A)"               → "Age Group"
        "Article Number"              → "Article Number"   (unchanged)
    """
    return re.sub(r'\s*\([^)]*\)\s*$', '', str(c or "")).strip()

class LinelistLoader:
    """
    v3.2 changes:
    • Column headers are normalised with _norm_col() to strip (M)/(A) suffixes.
    • If two consecutive rows both have "Article Number" in col-0, the second
      is treated as a duplicate header and skipped — data starts on row after.
    • PREFERRED sheet list includes generic fallbacks for arbitrary filenames.
    """
    PREFERRED = ["SS27 V2 Linelist", "SS26 V2 Linelist", "SS25 V2 Linelist",
                 "Linelist", "Sheet1", "Sheet2"]

    def __init__(self, path: Path):
        self.path  = path
        self.df    = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[Linelist] Loading: %s", self.path.name)
        wb     = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        target = next((s for s in self.PREFERRED if s in wb.sheetnames), None)
        if target is None:
            best, best_rows = wb.sheetnames[0], 0
            for sn in wb.sheetnames:
                n = wb[sn].max_row or 0
                if n > best_rows:
                    best, best_rows = sn, n
            target = best
        log.info("[Linelist] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # ── Find header row: first row containing "Article Number" ────────────
        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(str(v or "").strip() == "Article Number" for v in r)), None
        )
        if hdr_idx is None:
            # Fallback: first row with 5+ non-empty string values in cols 0-9
            hdr_idx = next(
                (i for i, r in enumerate(rows)
                 if sum(1 for v in r[:10] if isinstance(v, str) and v.strip()) >= 5), None
            )
        if hdr_idx is None:
            log.error("[Linelist] Cannot find header row")
            wb.close()
            return

        # ── v3.2 FIX: skip duplicate header row ───────────────────────────────
        # Some linelists repeat the header on the very next row.
        next_row = rows[hdr_idx + 1] if hdr_idx + 1 < len(rows) else None
        if next_row and str(next_row[0] or "").strip() == "Article Number":
            log.info("[Linelist] Duplicate header detected at row %d — skipping", hdr_idx + 1)
            hdr_idx += 1   # use the second copy; data starts after it

        log.info("[Linelist] Header at row %d (0-indexed)", hdr_idx)

        # ── v3.2 FIX: normalise column names — strip (M) / (A) etc. ──────────
        raw_header = rows[hdr_idx]
        header     = []
        seen: dict[str, int] = {}
        for h in raw_header:
            normed = _norm_col(h) if h else ""
            if not normed:
                normed = f"col_{len(header)}"
            # Deduplicate: "Retail Intro Date" appears twice → keep first as-is,
            # second becomes "Retail Intro Date_1"
            if normed in seen:
                seen[normed] += 1
                normed = f"{normed}_{seen[normed]}"
            else:
                seen[normed] = 0
            header.append(normed)

        df      = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
        art_col = next((c for c in ["Article Number", "Article No.", "Article No"]
                        if c in df.columns), None)
        if art_col:
            df = df[df[art_col].notna() & (df[art_col] != 0) & (df[art_col] != "")]

        self.df    = df
        self.sheet = target
        wb.close()
        log.info("[Linelist] %d articles in '%s'", len(self.df), self.sheet)


class BacklogLoader:
    def __init__(self, path: Path):
        self.path = path
        self.backlog_by_material:      dict[str, dict] = {}
        self.availability_by_material: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[Backlog] Loading 'Data' sheet using pandas: %s", self.path.name)
        try:
            df = pd.read_excel(self.path, sheet_name="Data", header=None)
        except Exception as e:
            log.error("[Backlog] Could not load Data sheet: %s", e)
            return

        hdr_idx = None
        for i, row in df.head(20).iterrows():
            if any(str(v).strip().upper() == "MATERIAL" for v in row.values if pd.notnull(v)):
                hdr_idx = i
                break

        if hdr_idx is None:
            log.error("[Backlog] Cannot find header row in Data tab containing 'Material'")
            return

        df.columns = df.iloc[hdr_idx]
        df = df.iloc[hdr_idx + 1:].reset_index(drop=True)

        mat_col          = next((c for c in df.columns if pd.notnull(c) and str(c).strip().upper() == "MATERIAL"), None)
        div_col          = next((c for c in df.columns if pd.notnull(c) and "PRODUCT DIVISION" in str(c).strip().upper()), None)
        soldto_col       = next((c for c in df.columns if pd.notnull(c) and "SOLD-TO PARTY" in str(c).strip().upper()), None)
        qty_col          = next((c for c in df.columns if pd.notnull(c) and str(c).strip().upper() == "QUANTITY"), None)
        ns_col           = next((c for c in df.columns if pd.notnull(c) and "NS VALUE" in str(c).strip().upper()), None)
        season_jun_col   = next((c for c in df.columns if pd.notnull(c) and "SEASON" in str(c).strip().upper() and "JUN" in str(c).strip().upper()), None)
        seasonal_ind_col = next((c for c in df.columns if pd.notnull(c) and "SEASONAL INDICATOR" in str(c).strip().upper()), None)

        if not mat_col:
            log.error("[Backlog] Column 'Material' not found.")
            return

        df[mat_col] = df[mat_col].astype(str).str.strip()
        if qty_col:
            df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
        if ns_col:
            df[ns_col] = pd.to_numeric(df[ns_col], errors='coerce').fillna(0.0)

        grouped = df.groupby(mat_col)

        for mat, group in grouped:
            if mat in ("nan", "None", ""):
                continue

            total_qty = group[qty_col].sum() if qty_col else 0
            total_ns  = group[ns_col].sum()  if ns_col  else 0.0

            div      = group[div_col].dropna().iloc[0]          if (div_col          and len(group[div_col].dropna()))          else ""
            soldto   = group[soldto_col].dropna().iloc[0]        if (soldto_col       and len(group[soldto_col].dropna()))       else ""
            seasonal_ind = group[seasonal_ind_col].dropna().iloc[0] if (seasonal_ind_col and len(group[seasonal_ind_col].dropna())) else ""

            tdd_windows = {}
            if season_jun_col:
                for _, row in group.iterrows():
                    val = str(row[season_jun_col]).strip()
                    if val.upper().startswith("TDD"):
                        tdd_windows[val] = tdd_windows.get(val, 0) + (row[qty_col] if (qty_col and pd.notnull(row[qty_col])) else 0)

            self.backlog_by_material[mat] = {
                "product_division": str(div),
                "sold_to":          str(soldto),
                "qty":              total_qty,
                "ns_value":         total_ns,
                "tdd_windows":      tdd_windows,
            }
            self.availability_by_material[mat] = {
                "product_division":   str(div),
                "seasonal_indicator": str(seasonal_ind),
                "qty":                total_qty,
            }

        log.info("[Backlog] Processed %d materials directly from Data tab.", len(self.backlog_by_material))

    def get(self, material: str) -> dict:
        b = self.backlog_by_material.get(material, {})
        a = self.availability_by_material.get(material, {})
        return {
            "backlog_qty":        b.get("qty", 0),
            "ns_value":           b.get("ns_value", 0.0),
            "sold_to":            b.get("sold_to", ""),
            "tdd_windows":        b.get("tdd_windows", {}),
            "available_qty":      a.get("qty", 0),
            "seasonal_indicator": a.get("seasonal_indicator", ""),
        }


class TDDLoader:
    HEADER_IDX = 5
    DATA_IDX   = 6

    def __init__(self, path: Path):
        self.path          = path
        self.by_article:   dict[str, list[dict]] = {}
        self.article_meta: dict[str, dict]        = {}
        self._load()

    def _load(self):
        log.info("[TDD] Loading: %s", self.path.name)
        wb   = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws   = wb["Sheet1"]
        rows = list(ws.iter_rows(values_only=True))
        hdr_row = rows[self.HEADER_IDX]
        col     = {str(h).strip(): i for i, h in enumerate(hdr_row) if h}
        rrp_col = col.get("RRP FINAL") or next(
            (i for i, h in enumerate(hdr_row)
             if h and "RRP" in str(h).upper() and "FINAL" in str(h).upper()), None
        )
        log.info("[TDD] RRP column index: %s", rrp_col)

        def g(row, name):
            idx = col.get(name)
            if idx is None or idx >= len(row):
                return ""
            v = row[idx]
            if v is None:
                return ""
            # ── v3.2.1 FIX: guard against pandas NaT ──────────────────────
            try:
                if pd.isnull(v):
                    return ""
            except (TypeError, ValueError):
                pass
            # ──────────────────────────────────────────────────────────────
            if isinstance(v, datetime):
                return v.strftime("%d.%m.%Y")
            return str(v).strip()

        def g_rrp(row):
            if rrp_col is not None and rrp_col < len(row) and row[rrp_col] is not None:
                v = row[rrp_col]
                # ── v3.2.1 FIX: guard against NaT in RRP column ──────────
                try:
                    if pd.isnull(v):
                        return ""
                except (TypeError, ValueError):
                    pass
                return str(v).strip()
            return ""

        for row in rows[self.DATA_IDX:]:
            article_no = g(row, "Article No")
            if not article_no: continue
            if article_no not in self.article_meta:
                self.article_meta[article_no] = {
                    "article_no":        article_no,
                    "material_desc":     g(row, "IAM Material description"),
                    "color_description": g(row, "Color/ Description 1"),
                    "product_division":  g(row, "Product Division"),
                    "product_group":     g(row, "Product Group"),
                    "business_segment":  g(row, "Business Segment"),
                    "product_type":      g(row, "Product Type"),
                    "age_group":         g(row, "Age Group"),
                    "gender":            g(row, "Gender"),
                    "rrp":               g_rrp(row),
                    "currency":          g(row, "Currency"),
                    "season_indicator":  g(row, "Season Indicator"),
                    "country_of_origin": g(row, "Country of origin"),
                    "frs_whs":           g(row, "FRS / WHS"),
                    "channel":           g(row, "Channel - Jun"),
                    "remarks_season":    g(row, "Remarks - Season ( Jun )"),
                }
            self.by_article.setdefault(article_no, []).append({
                "tech_size":     g(row, "Tech.Size"),
                "size":          g(row, "Size"),
                "ean":           g(row, "EAN/UPC"),
                "qty_ordered":   g(row, "Quantity Ordered"),
                "allocated_qty": g(row, "Allocated Qty - Batch"),
                "rdd":           g(row, "RDD"),
                "order_no":      g(row, "Order Number"),
                "sold_to":       g(row, "Sold-to Name"),
                "ship_to":       g(row, "Ship-to Name"),
            })

        wb.close()
        log.info("[TDD] %d articles | %d size rows",
                 len(self.by_article), sum(len(v) for v in self.by_article.values()))

    def get_meta(self, a): return self.article_meta.get(a, {})
    def get_sizes(self, a): return self.by_article.get(a, [])


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

GENDER_MAP = {
    "M": "M", "W": "W", "F": "W", "U": "U",
    "MALE": "M", "MEN": "M",
    "FEMALE": "W", "WOMEN": "W",
    "UNISEX": "U"
}
AGE_MAP = {
    "ADULT": "AD", "ADULTS": "AD", "AD": "AD",
    "JUNIOR": "CH", "YOUTH": "CH", "KIDS": "CH", "KID": "CH",
    "CHILDREN": "CH", "CHILD": "CH", "INFANT": "CH", "IN": "CH",
    "ALL AGES": "AA", "ALL": "AA",
}


def _s(v) -> str:
    """
    Safely convert any value to a clean string.

    v3.2.1 fix: guards against pandas.NaT, which passes isinstance(v, datetime)
    but raises 'NaTType does not support strftime' when .strftime() is called.
    Also catches numpy.nan and any other pandas null sentinel via pd.isnull().
    """
    if v is None:
        return ""
    # Guard: catch NaT, NaN, and any other pandas/numpy null sentinel
    try:
        if pd.isnull(v):
            return ""
    except (TypeError, ValueError):
        # pd.isnull raises TypeError for some non-scalar types — safe to ignore
        pass
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "0") else s


def _format_date_dd_mm_yyyy(v) -> str:
    """
    Format date to DD-MM-YYYY format for AT_IncomingMonth.
    Handles datetime objects and common date string formats.
    """
    if v is None:
        return ""
    try:
        if pd.isnull(v):
            return ""
    except (TypeError, ValueError):
        pass
    
    if isinstance(v, datetime):
        return v.strftime("%d-%m-%Y")
    
    raw = str(v).strip()
    if not raw or raw in ("None", "nan", "NaT", "0"):
        return ""
    
    # Try parsing common date formats
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            dt = datetime.strptime(raw, fmt)
            return dt.strftime("%d-%m-%Y")
        except (ValueError, TypeError):
            continue
    
    return raw


def _pad(val: str, n: int) -> str:
    v = re.sub(r"[^A-Z0-9]", "", val.upper())[:n]
    return v.ljust(n, "0")


def _resolve_coo(raw: str) -> tuple[str, str]:
    """
    Return (iso_code, label) for a country of origin value.
    Handles both ISO codes ("VN") and full names ("VIETNAM").
    """
    raw = (raw or "").strip()
    upper = raw.upper()
    # Direct ISO code hit
    if upper in LOV_COUNTRY_ORIGIN and len(upper) == 2:
        label = LOV_COUNTRY_ORIGIN[upper]
        return upper, label
    # Full name hit
    label = LOV_COUNTRY_ORIGIN.get(upper, raw)
    code  = _COUNTRY_NAME_TO_CODE.get(label, raw[:2].upper())
    return code, label


def map_article(ll_row, tdd: TDDLoader, backlog: BacklogLoader, brand_code="ADI") -> dict:
    """
    v3.2: reads both old column names and new SS27 column names.
    New columns added: Colorway Name, Country of Origin (A-Fact),
    IDN-Retail Price, SEA SEGMENTATION, IND SEGMENTATION, CONCEPT,
    SUB CONCEPT, Retail Intro Date, Retail Exit Date, Retail Intro Month,
    Hard Launch, Carry Over F26, ADI_EM_SEA-PLC (PLC code).
    """
    # ── Core identity ─────────────────────────────────────────────────────────
    article_no   = _s(ll_row.get("Article Number"))
    model_name   = _s(ll_row.get("Model Name"))
    model_no     = _s(ll_row.get("Model Number"))

    # ── Demographics ─────────────────────────────────────────────────────────
    gender_raw   = _s(ll_row.get("Gender") or ll_row.get("SEA GENDER"))
    age_raw      = _s(ll_row.get("Age Group"))

    # ── Hierarchy ─────────────────────────────────────────────────────────────
    division_raw = _s(ll_row.get("Product Division"))
    biz_seg      = _s(ll_row.get("Business Segment"))
    prod_type    = _s(ll_row.get("Product Type"))
    prod_group   = _s(ll_row.get("Product Group"))
    sports_cat   = _s(ll_row.get("Sports Category"))
    key_cat_cls  = _s(ll_row.get("Key Category Cluster"))
    key_cat      = _s(ll_row.get("Key Category"))
    sales_line   = _s(ll_row.get("Sales Line"))
    franchise    = _s(ll_row.get("Franchise") or ll_row.get("Monobrand Franchise"))

    # ── Colour / COO ─────────────────────────────────────────────────────────
    # v3.2: "Colorway Name" is the new column; "Colour" is the old one
    colour = _s(ll_row.get("Colorway Name") or ll_row.get("Colour"))
    # v3.2: full country name in "Country of Origin (A-Fact)"
    coo_raw = _s(
        ll_row.get("Country of Origin (A-Fact)")
        or ll_row.get("Country Of Origin")
        or ll_row.get("Country of Origin")
    )

    # ── Pricing ───────────────────────────────────────────────────────────────
    # v3.2: "IDN-Retail Price" replaces / supplements "IND RRP Local"
    rrp_ll      = _s(ll_row.get("IDN-Retail Price") or ll_row.get("IND RRP Local"))
    currency_ll = _s(ll_row.get("AREA RRP Currency") or "IDR")

    # ── Article classification ────────────────────────────────────────────────
    art_type = _s(ll_row.get("ARTICLE TYPE"))
    quarter  = _s(ll_row.get("Quarter"))

    # ── v3.2 new fields ───────────────────────────────────────────────────────
    sea_seg      = _s(ll_row.get("SEA SEGMENTATION"))
    ind_seg      = _s(ll_row.get("IND SEGMENTATION"))
    concept      = _s(ll_row.get("CONCEPT"))
    sub_concept  = _s(ll_row.get("SUB CONCEPT"))
    intro_date   = _s(ll_row.get("Retail Intro Date"))
    exit_date    = _s(ll_row.get("Retail Exit Date"))
    intro_month  = _format_date_dd_mm_yyyy(ll_row.get("Retail Intro Date"))
    hard_launch  = _s(ll_row.get("Hard Launch"))
    carry_over   = _s(ll_row.get("Carry\nOver\nF26") or ll_row.get("Carry Over F26"))
    plc_code     = _s(ll_row.get("ADI_EM_SEA-PLC"))
    order_cutoff = _s(ll_row.get("SEA Country Order Cutoff"))
    allocation   = _s(ll_row.get("ALLOCATION"))
    comms        = _s(ll_row.get("COMMS SUPPORT"))
    priority     = _s(ll_row.get("PRIORITY LEVEL"))
    camp_name    = _s(ll_row.get("CAMP. NAME / PACK NAME"))
    local_camp   = _s(ll_row.get("LOCAL CAMPAIGN NAME"))
    launch_cls   = _s(ll_row.get("Launch Classification"))
    early_access = _s(ll_row.get("Early Access Classification"))
    early_date   = _s(ll_row.get("Early Access Date"))
    exclusive    = _s(ll_row.get("EXCLUSIVE"))

    # ── TDD enrichment ────────────────────────────────────────────────────────
    tdd_meta  = tdd.get_meta(article_no)
    tdd_sizes = tdd.get_sizes(article_no)

    gender_fin   = _s(tdd_meta.get("gender"))            or gender_raw
    age_fin      = _s(tdd_meta.get("age_group"))         or age_raw
    coo_fin_raw  = _s(tdd_meta.get("country_of_origin")) or coo_raw
    rrp_fin      = _s(tdd_meta.get("rrp"))               or rrp_ll
    currency_fin = _s(tdd_meta.get("currency"))          or currency_ll
    colour_fin   = _s(tdd_meta.get("color_description")) or colour
    season_fin   = _s(tdd_meta.get("season_indicator"))  or quarter
    div_fin      = _s(tdd_meta.get("product_division"))  or division_raw
    biz_fin      = _s(tdd_meta.get("business_segment"))  or biz_seg
    ptype_fin    = _s(tdd_meta.get("product_type"))      or prod_type

    # ── Backlog enrichment ────────────────────────────────────────────────────
    bl           = backlog.get(article_no)
    seasonal_ind = bl["seasonal_indicator"] or season_fin

    # ── Code resolution ───────────────────────────────────────────────────────
    sap_gender   = GENDER_MAP.get(gender_fin.upper(), gender_fin)
    sap_age      = AGE_MAP.get(age_fin.upper(), "AD")
    art_category = "0" if (art_type or "").lower() == "single" else "1"
    coo_code, coo_label = _resolve_coo(coo_fin_raw)

    return {
        # ── core
        "article_no":      article_no,
        "model_name":      model_name,
        "model_no":        model_no,
        "brand_code":      brand_code,
        # ── demographics
        "gender_code":     sap_gender,
        "gender_raw":      gender_fin,
        "age_code":        sap_age,
        "age_raw":         age_fin,
        # ── colour / coo
        "colour":          colour_fin,
        "colour_code":     colour_fin,
        "coo_code":        coo_code,
        "coo_label":       coo_label,
        # ── hierarchy
        "division":        div_fin,
        "biz_segment":     biz_fin,
        "key_category":    key_cat,
        "key_cat_cluster": key_cat_cls,
        "product_type":    ptype_fin,
        "product_group":   prod_group,
        "sports_category": sports_cat,
        "sports_cat_en":   LOV_SPORTS_CATEGORY_EN.get((sports_cat or "").upper(), ""),
        "sales_line":      sales_line,
        "franchise":       franchise,
        # ── classification
        "article_type":    art_type or _s(tdd_meta.get("remarks_season")),
        "art_category":    art_category,
        # ── season / timing
        "season":          season_fin,
        "seasonal_ind":    seasonal_ind,
        "intro_date":      intro_date,
        "exit_date":       exit_date,
        "intro_month":     intro_month,
        # ── pricing
        "rrp":             rrp_fin,
        "currency":        currency_fin,
        # ── channel / logistics
        "channel":         _s(tdd_meta.get("channel")),
        "frs_whs":         _s(tdd_meta.get("frs_whs")),
        "order_cutoff":    order_cutoff,
        "allocation":      allocation,
        # ── campaign / marketing
        "sea_seg":         sea_seg,
        "ind_seg":         ind_seg,
        "concept":         concept,
        "sub_concept":     sub_concept,
        "comms":           comms,
        "priority":        priority,
        "camp_name":       camp_name,
        "local_camp":      local_camp,
        "launch_cls":      launch_cls,
        "early_access":    early_access,
        "early_date":      early_date,
        "exclusive":       exclusive,
        "plc_code":        plc_code,
        "hard_launch":     hard_launch,
        "carry_over":      carry_over,
        # ── backlog
        "backlog_qty":     bl["backlog_qty"],
        "available_qty":   bl["available_qty"],
        "ns_value":        bl["ns_value"],
        "tdd_windows":     bl["tdd_windows"],
        "_tdd_sizes":      tdd_sizes,
        # ── filename-level fields (injected by run() after mapping)
        "country_code":              "",   # filled by run() from args.country_code
        "article_type_from_filename": "",  # filled by run() from args.article_type_from_filename
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    warns = []
    art   = mapped["article_no"]
    field_to_at = {
        "article_no":  "AT_PrincipalStyleCode",
        "gender_code": "AT_Gender",
        "age_code":    "AT_SAPAge",
        "brand_code":  "AT_Brand",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {})
        if (meta.get("cardinality") or "").lower() == "mandatory":
            val = mapped.get(field, "")
            if not val or str(val).strip() in ("", "None", "nan"):
                warns.append(f"[{art}] MISSING mandatory: {at_id}")
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER
# ══════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

import re as _re
_XMLNS_RE = _re.compile(r'\s+xmlns="[^"]*"')

ET.register_namespace('', STIBO_NS)

DIVISION_PARENT_MAP = {
    "ACCESSORIES": "E",
    "FOOTWEAR":    "F",
    "APPAREL":     "A",
    "EQUIPMENT":   "Q",
    "HARDWARE":    "Q",
    "TOYS":        "T",
}


def _val(parent: ET.Element, attr_id: str, value: str = "",
         id_val: str = "", derived: bool = False, empty_ok: bool = True):
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


def _multival(parent: ET.Element, attr_id: str, id_val: str, label: str = ""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))


def _lov(code: str, lookup: dict, default_label: str = "") -> tuple[str, str]:
    code  = (code or "").strip()
    label = lookup.get(code, default_label or code)
    return code, label


def _resolve_currency_from_country(country_code: str) -> str:
    """
    Resolve Retail Price Currency LOV ID from country code.
    Based on the Retail Price Currency LOV mapping table.
    """
    country_to_currency = {
        "ID": "IDR",  # Indonesia → Indonesian Rupiah
        "PH": "PHP",  # Philippines → Philippine Peso
        "TH": "THB",  # Thailand → Thailand Baht
        "SG": "SGD",  # Singapore → Singapore Dollar
        "MY": "MYR",  # Malaysia → Malaysian Ringgit
        "VN": "VND",  # Vietnam → Vietnamese Dong
        "KH": "USD",  # Cambodia → United States Dollar
    }
    return country_to_currency.get((country_code or "").strip().upper(), "")


GENERIC_ATTR_PLACEHOLDERS = [
    "AT_SPUGrouping", "AT_EComProductNameTH",
    "AT_EComProductNameKH", "AT_SPUGroupingName", "AT_EComProductNameMY",
    "AT_Collection1", "AT_EComProductNameVN",
    "AT_Collection2", "AT_EComProductNameEN", "AT_Interest",
    "AT_EComProductNamePH", "AT_PrincipalColorName", "AT_HeelHeight",
    "AT_EComAgesCategory", "AT_GolfClubLength", "AT_Occasion",
    "AT_EComProductNameID", "AT_HeelType", "AT_CountrySize",
    "AT_PrincipalStyleDescription", "AT_ShortDescriptionKH",
    "AT_CareInstructionID", "AT_EstimatedLandedCost",
    "AT_CareInstructionEN", "AT_LongDescriptionEN", "AT_Updatedby",
    "AT_ShortDescriptionVN", "AT_FOBCurrency", "AT_ShortDescriptionTH",
    "AT_ProductWeight", "AT_MainEANIndicator", "AT_Updatedon",
    "AT_ShortDescriptionPH", "AT_ImagesSource", "AT_ShortDescriptionMY",
    "AT_ExchangeRate", "AT_ShortDescriptionEN", "AT_PrincipalSize",
    "AT_ShortDescriptionID", "AT_FOB", "AT_CertificateNumber",
    "AT_PackagingLength", "AT_SAPStyleCode", "AT_PackagingWidth",
    "AT_MainVendorIdentification", "AT_LongDescriptionTH",
    "AT_CareInstructionKH", "AT_LongDescriptionKH", "AT_TechnologyUsed", "AT_Royalty", "AT_LongDescriptionMY",
    "AT_CareInstructionVN", "AT_Size", "AT_LongDescriptionVN",
    "AT_CareInstructionTH", "AT_LongDescriptionID",
    "AT_CareInstructionPH", "AT_FreightCost",
    "AT_HaddadOfficeChargeDeduction", "AT_LongDescriptionPH",
    "AT_CareInstructionMY", "AT_Createdon", "AT_HaddadOfficeCharge",
    "AT_MarketingFee", "AT_Others", "AT_ECommEstimatedLaunchDate",
    "AT_SGS", "AT_PackagingHeight", "AT_PackagingWeight", "AT_Commission",
]

VARIANT_ATTR_PLACEHOLDERS = [
    "AT_EComProductNameTH", "AT_SPUGrouping", "AT_EComProductNameKH",
    "AT_SPUGroupingName", "AT_EComProductNameMY",
    "AT_Collection1", "AT_ProductLengthWidthHeight", "AT_LabelCost",
    "AT_EComProductNameVN", "AT_Collection2", "AT_FGBarcode",
    "AT_EComProductNameEN", "AT_PrincipalBarcode", "AT_EComProductNamePH",
    "AT_InternalBarcode", "AT_SafeguardDutyIdOnly", "AT_EComProductNameID",
    "AT_LongDescriptionEN", "AT_HSCode", "AT_Updatedby",
    "AT_ShortDescriptionVN", "AT_ShortDescriptionTH", "AT_ProductWeight",
    "AT_MainEANIndicator", "AT_Updatedon", "AT_ShortDescriptionPH",
    "AT_ShortDescriptionMY", "AT_ExchangeRate", "AT_SAPProductFlag",
    "AT_PrincipalSize", "AT_ShortDescriptionID", "AT_FOB",
    "AT_CertificateNumber", "AT_CurrentPrice", "AT_PackagingLength",
    "AT_ColorDescriptionVN", "AT_SAPStyleCode", "AT_PackagingWidth",
    "AT_ColorDescriptionTH", "AT_MainVendorIdentification",
    "AT_LongDescriptionTH", "AT_CareInstructionKH", "AT_ColorDescriptionPH",
    "AT_OriginalPrice", "AT_LongDescriptionKH", "AT_TechnologyUsed",
    "AT_ColorDescriptionMY", "AT_Royalty", "AT_LongDescriptionMY",
    "AT_CareInstructionVN", "AT_LongDescriptionVN", "AT_CareInstructionTH",
    "AT_ColorDescriptionID", "AT_LongDescriptionID", "AT_CareInstructionPH",
    "AT_FreightCost", "AT_HaddadOfficeChargeDeduction", "AT_LongDescriptionPH",
    "AT_CareInstructionMY", "AT_Createdon", "AT_EANCategory",
    "AT_HaddadOfficeCharge", "AT_MarketingFee", "AT_PrincipalStyleCode",
    "AT_Others", "AT_ECommEstimatedLaunchDate", "AT_SGS",
    "AT_PackagingHeight", "AT_ColorDescriptionKH", "AT_PackagingWeight",
    "AT_Commission",
]



def _parse_season(season_raw: str) -> tuple[str, str, str]:
    """
    Returns (sea_prefix, sea_year_4digit, full_season_code)
    Handles both SP26 (2-digit) and SP2029 (4-digit) formats.
    """
    m = re.match(r'^([A-Z]{2})(\d{2,4})$', (season_raw or "").strip().upper())
    if not m:
        return season_raw[:2].upper(), "", season_raw
    prefix   = m.group(1)
    year_raw = m.group(2)
    year_4   = year_raw if len(year_raw) == 4 else f"20{year_raw}"
    return prefix, year_4, f"{prefix}{year_4}"

    

def _add_generic_values(vals_el: ET.Element, art: dict, brand_name: str,
                        comp_code: str, sbu: str,mdd=None):
    written = set()

    # ── Brand / Company ───────────────────────────────────────────────────────
    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand", b_label, id_val=b_code);      written.add("AT_Brand")
    brand_lov_id = LOV_BRAND.get(b_code, b_code)  # "ADI" → "ADIDAS"
    _val(vals_el, "AT_BrandGroup", brand_lov_id, id_val=brand_lov_id); written.add("AT_BrandGroup")
    


    # ── Article identity ──────────────────────────────────────────────────────
    _val(vals_el, "AT_PrincipalStyleCode", art["article_no"]);    written.add("AT_PrincipalStyleCode")
    _val(vals_el, "AT_SAPStyleCode", art["article_no"]);          written.add("AT_SAPStyleCode")
    _val(vals_el, "AT_PrincipalStyleDescription", art["model_name"]); written.add("AT_PrincipalStyleDescription")
    # _val(vals_el, "AT_PrincipalColorCode", art["colour"])

    # if art["colour"]:
    #     colour_upper = art["colour"].strip().upper()
    #     color_lov    = (mdd.lovs.get("Color Code") or {}) if mdd else {}
    #     # color_lov keys are descriptions (e.g. "BLACK"), values are codes (e.g. "005")
    #     # Build reverse map: description → code
    #     color_id = color_lov.get(colour_upper, "")
    #     if not color_id:
    #         # Try matching by code directly (in case colour is already a code)
    #         rev = {v.upper(): v for v in color_lov.values()}
    #         color_id = rev.get(colour_upper, art["colour"])
    #     _val(vals_el, "AT_Color", art["colour"], id_val=color_id)
    #     written.add("AT_Color")

    # ── Gender ── (dynamic lookup from MDD Gender LOV: Col A → Col C) ──────────
    gender_lov    = (mdd.lovs.get("GenderLOV") or {}) if mdd else {}
    gender_raw_up = (art.get("gender_raw") or "").strip().upper()
    g_lov_id      = (
        gender_lov.get(gender_raw_up)
        or gender_lov.get((art.get("gender_code") or "").upper())
        or art["gender_code"]
    )
    _val(vals_el, "AT_Gender",   id_val=g_lov_id); written.add("AT_Gender")
    _val(vals_el, "AT_BYGender", id_val=g_lov_id); written.add("AT_BYGender")
    _val(vals_el, "AT_PrincipalGenderDescription", art["gender_raw"] or g_lov_id); written.add("AT_PrincipalGenderDescription")

    # ── Age ── (dynamic lookup from MDD Age LOV: Col A → Col B) ──────────────
    age_lov       = (mdd.lovs.get("AgeLOV") or {}) if mdd else {}
    age_raw_up    = (art.get("age_raw") or "").strip().upper()
    sap_age_id    = (
        age_lov.get(age_raw_up)
        or age_lov.get((art.get("age_code") or "").upper())
        or art["age_code"]
    )
    _val(vals_el, "AT_SAPAge", id_val=sap_age_id); written.add("AT_SAPAge")
    by_age_raw   = art["age_raw"].upper() if art["age_raw"] else sap_age_id
    by_age_id    = BY_AGE_ID_MAP.get(by_age_raw, by_age_raw)
    _val(vals_el, "AT_BYAge", id_val=by_age_id); written.add("AT_BYAge")
    _val(vals_el, "AT_PrincipalAgeDescription", art["age_raw"] or sap_age_id); written.add("AT_PrincipalAgeDescription")

    # ── Season ────────────────────────────────────────────────────────────────
    sea_raw                          = art.get("season") or ""
    sea_prefix, sea_year_4, _        = _parse_season(sea_raw)
    sea_label                        = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_prefix, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_prefix); written.add("AT_Season")
    #always emit 4-digit year (handles SP2029 correctly)
    _val(vals_el, "AT_SeasonYear", sea_year_4); written.add("AT_SeasonYear")

    # ── Country of origin (v3.2: uses pre-resolved code+label) ───────────────
    country_code = art.get("country_code", "")
    if country_code:
        country_label = LOV_COUNTRY_ORIGIN.get(country_code, country_code)
        _val(vals_el, "AT_Country", country_label, id_val=country_code)
        _val(vals_el, "AT_CountryOrigin", country_label, id_val=country_code)
        written.add("AT_Country")
        written.add("AT_CountryOrigin")

    # ── Article category / type ───────────────────────────────────────────────
    cat_code  = art["art_category"]
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, cat_code)
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code); written.add("AT_SAPArticleCategory")

    # Requirement 1: prefer article_type parsed from filename (Inline/License/SSE),
    # fall back to linelist column value
    at_raw   = art.get("article_type_from_filename") or art.get("article_type") or ""
    at_norm  = at_raw.capitalize()
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm) or at_norm
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm); written.add("AT_BYArticleType")

    # ── Material Type — ZHAW ─────────────────────────────────────
    _val(vals_el, "AT_MaterialType", id_val="ZHAW")
    written |= {"AT_MaterialType"}

    # ── SAP Product Flag — 5 ─────────────────────────────────────
    _val(vals_el, "AT_SAPProductFlag", id_val="5")
    written |= {"AT_SAPProductFlag"}

    # ── Indicators ────────────────────────────────────────────────────────────
    _val(vals_el, "AT_UOM", "Each", id_val="EA");        written.add("AT_UOM")

    # ── Pricing ───────────────────────────────────────────────────────────────
    _val(vals_el, "AT_OriginalPrice", art["rrp"]); written |= {"AT_OriginalPrice", "AT_CurrentPrice"}
    curr_code = art["currency"] or ""
    _val(vals_el, "AT_RetailPriceCurrency", curr_code, id_val=curr_code); written.add("AT_RetailPriceCurrency")
    
    # Resolve currency from country_code for AT_FOBCurrency
    currency_lov_id = _resolve_currency_from_country(art.get("country_code", ""))
    if currency_lov_id:
        _val(vals_el, "AT_FOBCurrency", id_val=currency_lov_id)
    else:
        _val(vals_el, "AT_FOBCurrency", curr_code, id_val=curr_code)
    written.add("AT_FOBCurrency")

    # ── Channel ───────────────────────────────────────────────────────────────
    # frs = art.get("frs_whs", "")
    # if "WHS" in frs.upper():
    #     _val(vals_el, "AT_NatureOfArticle", "Wholesale", id_val="WH1")
    # else:
    #     _val(vals_el, "AT_NatureOfArticle")
    # written.add("AT_NatureOfArticle")

    _val(vals_el, "AT_MainVendorIdentification", "1"); written.add("AT_MainVendorIdentification")
    _val(vals_el, "AT_PricingDistributionChannel", id_val="01"); written.add("AT_PricingDistributionChannel")
    _val(vals_el, "AT_IncomingMonth", art.get("intro_month", "")); written.add("AT_IncomingMonth")

    # ── Generic key ───────────────────────────────────────────────────────────
    generic_code = f"{art['brand_code']}{art['article_no']}"
    _val(vals_el, "AT_InboundGenericCode", generic_code); written.add("AT_InboundGenericCode")

    #  Country (from filename country token e.g. MY, ID) ─────
    country_val = art.get("country_code", "")
    if country_val and "AT_Country" not in written:
        _val(vals_el, "AT_Country", country_val, id_val=country_val)
        written.add("AT_Country")


    # ── v3.2 new attributes ───────────────────────────────────────────────────
    # FIXED — wrap the entire block:
    if INCLUDE_SS27_ATTRS:
        _val(vals_el, "AT_SEASegmentation",      art.get("sea_seg", ""))
        _val(vals_el, "AT_INDSegmentation",      art.get("ind_seg", ""))
        _val(vals_el, "AT_Concept",              art.get("concept", ""))
        _val(vals_el, "AT_SubConcept",           art.get("sub_concept", ""))
        _val(vals_el, "AT_RetailIntroDate",      art.get("intro_date", ""))
        _val(vals_el, "AT_RetailExitDate",       art.get("exit_date", ""))
        _val(vals_el, "AT_RetailIntroMonth",     art.get("intro_month", ""))
        hl = art.get("hard_launch", "")
        _val(vals_el, "AT_HardLaunch",           hl, id_val=hl)
        co = art.get("carry_over", "")
        _val(vals_el, "AT_CarryOver",            co, id_val=co)
        _val(vals_el, "AT_PLCCode",              art.get("plc_code", ""))
        _val(vals_el, "AT_OrderCutoff",          art.get("order_cutoff", ""))
        _val(vals_el, "AT_Allocation",           art.get("allocation", ""))
        _val(vals_el, "AT_CommsSupport",         art.get("comms", ""))
        _val(vals_el, "AT_PriorityLevel",        art.get("priority", ""))
        _val(vals_el, "AT_CampaignName",         art.get("camp_name", ""))
        _val(vals_el, "AT_LocalCampaignName",    art.get("local_camp", ""))
        _val(vals_el, "AT_LaunchClassification", art.get("launch_cls", ""))
        _val(vals_el, "AT_EarlyAccessClass",     art.get("early_access", ""))
        _val(vals_el, "AT_EarlyAccessDate",      art.get("early_date", ""))
        _val(vals_el, "AT_Exclusive",            art.get("exclusive", ""))
        _val(vals_el, "AT_KeyCategoryCluster",   art.get("key_cat_cluster", ""))
        _val(vals_el, "AT_SalesLine",            art.get("sales_line", ""))
        _val(vals_el, "AT_ProductGroup",         art.get("product_group", ""))

        

    # # Fill remaining placeholders
    # for a in GENERIC_ATTR_PLACEHOLDERS:
    #     if a not in written:
    #         _val(vals_el, a)
    #         written.add(a)

    # ── Principal Merchandise Hierarchy (L1–L5) ───────────────────────────
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art.get("division", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("biz_segment", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", art.get("product_group", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL4", art.get("product_type", ""))
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL5", art.get("sports_category", ""))
    _val(vals_el, "AT_SportsCategoryEN", art.get("sports_cat_en", ""))


def _add_variant_values(vals_el: ET.Element, art: dict, size_rec: dict,
                        comp_code: str, sbu: str):
    written = set()

    curr_code = art["currency"] or ""
    _val(vals_el, "AT_RetailPriceCurrency", curr_code, id_val=curr_code); written.add("AT_RetailPriceCurrency")
    
    # Resolve currency from country_code for AT_FOBCurrency
    currency_lov_id = _resolve_currency_from_country(art.get("country_code", ""))
    if currency_lov_id:
        _val(vals_el, "AT_FOBCurrency", id_val=currency_lov_id)
    else:
        _val(vals_el, "AT_FOBCurrency", curr_code, id_val=curr_code)
    written.add("AT_FOBCurrency")

    g_code, g_label = _lov(art["gender_code"], LOV_GENDER, art["gender_code"])
    _val(vals_el, "AT_Gender", g_label, id_val=g_code); written.add("AT_Gender")

    sap_age_label = LOV_AGE.get(art["age_code"], art["age_code"])
    _val(vals_el, "AT_SAPAge", sap_age_label, id_val=art["age_code"]); written.add("AT_SAPAge")

    size_val = size_rec.get("size", "")
    size_id  = re.sub(r"[^A-Z0-9/]", "", size_val.upper()) if size_val else ""
    _val(vals_el, "AT_Size",          size_val, id_val=size_id); written.add("AT_Size")
    _val(vals_el, "AT_PrincipalSize", size_val);                 written.add("AT_PrincipalSize")

    ean = size_rec.get("ean", "")
    _val(vals_el, "AT_PrincipalBarcode", ean); written.add("AT_PrincipalBarcode")
    _val(vals_el, "AT_EANCategory");           written.add("AT_EANCategory")

    # v3.2: use pre-resolved coo_code / coo_label
    country_code = art.get("country_code", "")
    if country_code:
        country_label = LOV_COUNTRY_ORIGIN.get(country_code, country_code)
        _val(vals_el, "AT_Country", country_label, id_val=country_code)
        _val(vals_el, "AT_CountryOrigin", country_label, id_val=country_code)
        written.add("AT_Country")
        written.add("AT_CountryOrigin")

    _val(vals_el, "AT_OriginalPrice",           art["rrp"])
    _val(vals_el, "AT_CurrentPrice",            art["rrp"])
    _val(vals_el, "AT_MainVendorIdentification","1")
    _val(vals_el, "AT_PricingDistributionChannel", id_val="01")
    _val(vals_el, "AT_PrincipalStyleCode",      art["article_no"])
    _val(vals_el, "AT_SAPStyleCode",            art["article_no"])
    _val(vals_el, "AT_PrincipalStyleDescription", art["model_name"])
    written |= {"AT_OriginalPrice", "AT_CurrentPrice", "AT_MainVendorIdentification",
                "AT_PricingDistributionChannel", "AT_PrincipalStyleCode","AT_SAPStyleCode","AT_PrincipalStyleDescription"}

    generic_code = f"{art['brand_code']}{art['article_no']}"
    _val(vals_el, "AT_InboundGenericCode", generic_code); written.add("AT_InboundGenericCode")

    # for a in VARIANT_ATTR_PLACEHOLDERS:
    #     if a not in written:
    #         _val(vals_el, a)
    #         written.add(a)


def _get_division_code(division: str) -> str:
    div_upper = (division or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in div_upper:
            return code
    return "X"


def build_classifications(articles: list[dict], brand: str,
                          brand_code: str, comp_code: str, sbu: str,
                          season_code: str) -> ET.Element:
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")



    sea_prefix, sea_year, full_season_code = _parse_season(season_code)
    season_id      = f"CLH_{brand_code}_{full_season_code}"
    batches_parent = f"CLH_{brand.title()}Batches"

    season_label_map = {"SS": "Spring Summer", "FW": "Fall Winter",
                        "AW": "Autumn Winter", "HO": "Holiday"}



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


def build_product_xml(art: dict, brand: str, brand_code: str,
                      comp_code: str, sbu: str,
                      season_id: str,
                      source_type: str = "linelist",
                      mdd=None) -> str:
    article_no = art["article_no"]
    if not article_no:
        return ""

    tdd_sizes = art.get("_tdd_sizes", [])
    is_single = (art["art_category"] == "0")
    user_type = "PRD_SingleArticle" if is_single else "PRD_GenericArticle"

    div_letter  = _get_division_code(art.get("division", ""))
    parent_id   = f"PPH_{div_letter}-TempSubCat"
    b_code      = art.get("brand_code", brand_code) or brand_code
    key_article = f"{b_code}{article_no}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", user_type)
    g_el.set("ParentID",   parent_id)
    g_el.set("update",     "true")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["model_name"] or article_no

    div_id = re.sub(r"[^A-Z0-9]", "", (art["division"] or "").upper())[:2]
    seg_id = re.sub(r"[^A-Z0-9]", "", (art["biz_segment"] or "").upper())[:4]

    for cls_id, cpl_type in [
        (f"CLH_{brand.title()}Articles",          "CPL_Merchandiser"),
        # (f"MA_{brand_code}_{div_id}_{seg_id}", "CPL_BYHierarchy"),
        # (f"CLH_{div_id}",                 "CPL_SAPHierarchy"),
        (f"{season_id}UA",                "CPL_UnConfirmedForSeason"),
    ]:
        cr = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
        cr.set("ClassificationID", cls_id)
        cr.set("Type", cpl_type)

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, brand, comp_code, sbu, mdd=mdd)

    if source_type in ("tdd", "backlog") and not is_single:
        if tdd_sizes:
            seen_eans   = set()
            var_counter = 0
            for size_rec in tdd_sizes:
                ean = size_rec.get("ean", "").strip()
                if not ean or ean in seen_eans:
                    continue
                seen_eans.add(ean)
                var_counter += 1

                size_val = size_rec.get("size", "")
                size_id  = re.sub(r"[^A-Z0-9/]", "", size_val.upper()) if size_val else ""

                v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
                v_el.set("UserTypeID", "PRD_VariantArticle")
                v_el.set("update",     "true")

                v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
                v_kv.set("KeyID", "KEY_Variant")
                v_kv.text = f"{b_code}{article_no}{size_id}"

                ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = f"Size {var_counter}"

                v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
                _add_variant_values(v_vals, art, size_rec, comp_code, sbu)
        else:
            v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
            v_el.set("UserTypeID", "PRD_VariantArticle")
            v_el.set("update",     "true")

            v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
            v_kv.set("KeyID", "KEY_Variant")
            v_kv.text = f"{b_code}{article_no}"

            ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = "Size 1"

            v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
            _add_variant_values(v_vals, art, {}, comp_code, sbu)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def _process_article(row_tuple):
    row, tdd, backlog, brand_code, mdd = row_tuple
    mapped = map_article(row, tdd, backlog, brand_code=brand_code)
    warns  = validate(mapped, mdd)
    return mapped, warns


def run(args, auditor=None):
    def first(d: Path, ext="*.xlsx"):
        f = list(d.glob(ext))
        return f[0] if f else None

    mdd_f     = first(MDD_DIR)
    attr_f    = first(ATTR_DIR)
    ll_files  = list(LINELIST_DIR.glob("*.xlsx"))
    bl_files  = list(BACKLOG_DIR.glob("*.xlsx"))
    tdd_files = list(TDD_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("Attributes List", attr_f), ("Linelist", ll_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            sys.exit(1)

    mdd     = MDDLoader(mdd_f)
    _       = AttributesListLoader(attr_f, brand=args.brand)
    backlog = BacklogLoader(bl_files[0]) if bl_files else None
    tdd     = TDDLoader(tdd_files[0])   if tdd_files else None

    if not backlog:
        log.warning("No Backlog file — availability data will be empty.")
    if not tdd:
        log.warning("No TDD file — EAN / size data will be empty.")

    if not backlog:
        class _EB:
            def get(self, _):
                return {"backlog_qty": 0, "available_qty": 0, "ns_value": 0,
                        "tdd_windows": {}, "seasonal_indicator": "", "sold_to": ""}
        backlog = _EB()

    if not tdd:
        class _ET:
            def get_meta(self, _): return {}
            def get_sizes(self, _): return []
        tdd = _ET()

    all_warnings = []

    _, _, full_season_code = _parse_season(args.season)
    season_id = f"CLH_{args.brand_code}_{full_season_code}"

    for ll_path in ll_files:
        log.info("─── Processing Linelist: %s ───", ll_path.name)
        ll = LinelistLoader(ll_path)
        if ll.df.empty:
            log.warning("Empty dataframe, skipping.")
            continue

        rows = []
        for _, row in ll.df.iterrows():
            art = str(
                row.get("Article Number")
                or row.get("Article No.")
                or row.get("Article No")
            ).strip()
            if art and art not in ("None", "nan", ""):
                rows.append(row)
                

        total_rows = len(rows)

        # Use input filename stem + .xml extension
        input_filename = ll_path.stem  # e.g. "0888-SP-ADIDAS-Line List MAA Sport Inline-Multi-SS2026-MY-1"
        out_name = f"{input_filename}.xml"
        out_path = XML_OUT_DIR / out_name

        log.info("Pass 1/2 — parallel map+validate (%d rows, %d workers) …",
                 total_rows, min(8, total_rows))

        num_workers    = min(8, total_rows)
        task_args      = [(row, tdd, backlog, args.brand_code, mdd) for row in rows]
        ordered_results: list[tuple[int, dict, list]] = []

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            futures = {pool.submit(_process_article, t): i for i, t in enumerate(task_args)}
            for fut in as_completed(futures):
                idx           = futures[fut]
                mapped, warns = fut.result()
                all_warnings.extend(warns)
                ordered_results.append((idx, mapped, warns))

        ordered_results.sort(key=lambda x: x[0])
        mapped_articles = [m for _, m, _ in ordered_results]
        del ordered_results

        # Inject filename-level fields that apply uniformly to all articles
        _country    = getattr(args, "country_code", "")
        _art_type   = getattr(args, "article_type_from_filename", "")
        for art in mapped_articles:
            if not art.get("country_code"):
                art["country_code"] = _country
            if not art.get("article_type_from_filename"):
                art["article_type_from_filename"] = _art_type
            # ↓ ADD THIS LINE:
            if not art.get("season"):
                art["season"] = args.season

        total_variants = sum(len(a.get("_tdd_sizes", [])) for a in mapped_articles)
        log.info("Pass 1 done — mapped=%d  variant_rows=%d", len(mapped_articles), total_variants)

        log.info("Pass 2/2 — streaming XML to %s …", out_name)

        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cls_el      = build_classifications(
            mapped_articles, args.brand, args.brand_code,
            args.comp_code, args.sbu, args.season,
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
                f' UseContextLocale="false"'
                f' update="true">\n\n'
            )
            f.write(f"  {cls_str}\n\n")
            del cls_str

            f.write('  <Products>\n')
            # ── TEST MODE: limit to 5 articles ──────────────────────────
            # mapped_articles = mapped_articles[:150]
            for art in mapped_articles:  # ONLY ONE PRODUCT AS REQUESTED BY USER
                if not art.get("article_no"):
                    continue
                product_xml = build_product_xml(
                    art, args.brand, args.brand_code,
                    args.comp_code, args.sbu,
                    season_id,
                    source_type="linelist",
                    mdd=mdd
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

        print("═══ ARTICLE SUMMARY ════════════════════════════════", flush=True)
        print(f"  Linelist rows      : {total_rows}",                                        flush=True)
        print(f"  Mapped articles    : {written_count}",                                     flush=True)
        print(f"  Variant rows (EANs): {total_variants}",                                    flush=True)
        print(f"  Total XML products : {written_count}  (Generic only — linelist source)",   flush=True)
        print(f"  XML file size      : {file_kb}KB",                                         flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    rpt_path = LOG_DIR / f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(description="Stibo Inbound XML Generator v3.2")
    p.add_argument("--brand",      default="Adidas")
    p.add_argument("--brand-code", default="ADI")
    p.add_argument("--comp-code",  default="0888")
    p.add_argument("--sbu",        default="SP")
    p.add_argument("--season",     default="SS27")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()
