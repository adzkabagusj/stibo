"""
╔══════════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — NEW BALANCE (Sample / Accessories) v1.0║
║  Licensed Accessories Sample Breakout → Stibo STEP XML               ║
╚══════════════════════════════════════════════════════════════════════╝

Source file  : "S127 Licensed Accessories Sample Breakout file _ 03202026.xlsx"
Primary tab  : FIRST VISIBLE sheet — "Sheet1"
               (title row 1, blank row 2, header row 3, data from row 4;
                the header row is located dynamically by scanning for "SKU")
Input folder : /tmp/stibo_workdir/input/sample_accessories/

MAPPING AUTHORITY (two sheets, cross-checked column-by-column):
  1. "Attributes List_complete_v6_25032026.xlsx" → sheet "NEW BALANCE - SAMPLE"
  2. "NEW - Brand mapping files Template.xlsx"   → sheet "NEW BALANCE SAMPLE"
  3. "Master Data Dictionary (MAA).xlsx"         → Core Attributes + LOV sheets

────────────────────────────────────────────────────────────────────────
THE "SAMPLE" TYPE — KEY BUSINESS RULE (per user + mapping sheets)
────────────────────────────────────────────────────────────────────────
New Balance "Sample" articles are inline-category templates whose generic
article codes carry the literal prefix "S" when ingested into Stibo:

    AT_PrincipalStyleCode   = "S" + <SKU>                      (BM row 14:
                              Mapping Logic = "S" + Principal Style Code,
                              Acc column = "SKU")
    AT_SAPStyleCode         = "S" + clean(SKU)[:7] + <running letter>
                              (BM row 23 / v6 R035 App/Acc: "Prefix S +
                              Up to 7 Digits Principal Style Code + 1 Digit
                              Running Alphabeth")
    AT_Generic              = "NEW" + "S" + clean(SKU)[:7] + <running letter>
                              (BM row 24 / v6 R036 — 3+1+7+1 = 12 chars,
                              exactly the MDD AT_Generic max of 12)
    AT_Variant              = AT_Generic + SAP color(3) + SAP size(3)
                              (BM row 26 / v6 R038 — 12+3+3 = 18 chars,
                              exactly the MDD AT_Variant max of 18)
    KEY_InboundArticle /
    AT_InboundGenericCode   = "NEW" + "S" + clean(SKU)          (row-unique —
                              repo convention, keeps KEY 1:1 with a source row)

THE RUNNING ALPHABET (BM rows 23/24/26 "1 Digit Running Alphabeth"):
    The breakout sheet is SKU-level (one row per style+colourway), but the
    SAP-style formula only uses the first 7 chars of the SKU, so multiple
    rows share the same 7-char base. The disambiguating letter is assigned
    sequentially per 7-char base in file order: 1st row of a base → "A",
    2nd → "B", … (>26 rows per base would wrap — a warning is logged; this
    live file peaks at 18 rows per base). Both distinct 8-char styles that
    collide on a 7-char base (e.g. LAU13009 / LAU13001) and the multiple
    colours of one style are disambiguated by this counter.

────────────────────────────────────────────────────────────────────────
COLUMN → ATTRIBUTE MAP (accessories sheet — every mapped column)
────────────────────────────────────────────────────────────────────────
  SKU                  → AT_PrincipalStyleCode   ("S"+SKU, raw)  [v6 R026]
                       → AT_SAPStyleCode / AT_Generic / AT_Variant base
                       → KEY_InboundArticle / AT_InboundGenericCode
  Product Display Name → AT_PrincipalStyleDescription               [R027]
                       → <Name>  (blank → SKU fallback)
  Color/Width          → AT_PrincipalColorCode  (max 7 chars)       [R028]
  Color Family         → AT_PrincipalColorName                      [R029]
                       → AT_Color  (SAP color — Color Code LOV)     [R040]
  Size                 → AT_PrincipalSize  (value "SS")             [R034]
  COO                  → AT_CountryOrigin (ISO-3 → 2-letter LOV)   [R054]
                       CHN→CN, LKA→LK, IDN→ID, JPN→JP, VNM→VN, THA→TH
  Product Line         → AT_PrincipalMerchandiseHierarchyL1          [R091]
                       → division parent (Accessories → PPH_E-…)
  Line Plan Business   → AT_PrincipalMerchandiseHierarchyL2          [R092]

  Accessories defaults (no source column — v6 "NEW BALANCE - SAMPLE"):
    SAP Gender  → Unisex / U (v6 R051: "Acc default: Unisex" — the Acc SKUs
                  start with A/L, which carry no W/M/Y gender signal) [R051]
    SAP Age     → Adults / AD (v6 R050 App/Acc default)             [R050]
    SAP Size    → AT_Size = 0SS / SS (v6 R041 default — "will be
                  automatically generated SAP Code SS (only one
                  variant)", confirmed by the all-"SS" Size column)
    SAP Color   → Color Family via the MDD "Color Code LOV" (the
                  103-colour SAP palette), v6 R040 note: "Mapping only
                  for RED = RED . or system to choose the close value
                  from LOV"

────────────────────────────────────────────────────────────────────────
COLUMNS READ BUT *NOT* MAPPED TO STIBO (per the two SAMPLE mapping sheets)
────────────────────────────────────────────────────────────────────────
  GBU, Channel Type, Category, UPC, Style/Part No., Construction,
  Primary Fabric Content, Supplier, Vendor Code, Factory Address Code,
  and ALL 39 country sample-quantity columns (AUSTRALIA … HONG KONG -
  REGIONAL). None of these appear as a "Field Name in the Brand File" for
  any attribute row in the two SAMPLE mapping sheets, so they are
  deliberately not emitted.

  Style/Part No. is intentionally NOT a source on its own: the SKU embeds it
  ("Style/Part No." + "_" + "Color/Width" == SKU for every row — verified)
  and both SAP code formulas reference the SKU column, not this column.

────────────────────────────────────────────────────────────────────────
DEVIATIONS (documented, deliberate)
────────────────────────────────────────────────────────────────────────
 1. SHEET SCOPE — the Brand Mapping sheet pins App/Acc to "Sheet : Sheet 1";
    the first visible sheet of this workbook IS "Sheet1", so the operator
    instruction (ingest the FIRST VISIBLE sheet) and the mapping spec agree.
    Default SHEET_MODE = "first_visible"; "all_data_sheets" also supported.
 2. AT_BYArticleType is emitted as the constant ID="Inline" for every row
    (v6 R106 / BM row 94: "Article Type — Default : Inline"), even though a
    "Channel Type" column exists and is mixed in this file (212 "Licensed" /
    55 "Inline" — including in the file title "Licensed Accessories").
    The SAMPLE spec pins the default and does not read Channel Type: the
    user's brief also classes Sample under the "inline" category.
 3. SAP Style/Generic/Variant codes ARE emitted with formula-computed
    values (v6 R035-R039 mandate them for the sample flow). See the
    [SAP-COMPUTED] flags in _add_generic_values() if MAA later asks for
    Stibo-internal handling instead.
 4. "SPL + Season" note on v6 R027 (Principal Style Description) is not
    applied — the sheet's own primary rule (Column "Product Display Name",
    blank → SKU fallback) is implemented instead, exactly like every other
    NB inline ETL in this repository.
 5. AT_PrincipalGenderDescription is NOT emitted — v6 R031 / BM row 19 map
    it to "Size Profile" for Footwear only and "N/A" for App/Acc.
 6. Prices (Original/Current/FOB) are NOT emitted — the SAMPLE mapping says
    "Manual Input via Smartsheet (Special for Sample)", i.e. no source
    column in this workbook.
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
SAMPLE_DIR   = INPUT_DIR / "sample_accessories"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [SAMPLE_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_sample_accessories_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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
# SHEET SCOPE — see DEVIATION 1 in the module docstring
# ══════════════════════════════════════════════════════════════════
SHEET_MODE = "first_visible"        # operator instruction + BM "Sheet 1"
# SHEET_MODE = "all_data_sheets"    # process every visible sheet
SUMMARY_SHEET_KEYWORDS = ("summary",)   # never ingested in any mode


# ══════════════════════════════════════════════════════════════════
# SOURCE COLUMN CONTRACT — every column this ETL reads
# ══════════════════════════════════════════════════════════════════
SOURCE_COLUMNS: dict[str, list[str]] = {
    "SKU":                  ["AT_PrincipalStyleCode", "AT_Gender",
                             "AT_SAPStyleCode", "AT_Generic", "AT_Variant",
                             "AT_InboundGenericCode"],
    "Product Display Name": ["AT_PrincipalStyleDescription", "<Name>"],
    "Color/Width":          ["AT_PrincipalColorCode"],
    "Color Family":         ["AT_PrincipalColorName", "AT_Color"],
    "Size":                 ["AT_PrincipalSize"],
    "COO":                  ["AT_CountryOrigin"],
    "Product Line":         ["AT_PrincipalMerchandiseHierarchyL1"],
    "Line Plan Business":   ["AT_PrincipalMerchandiseHierarchyL2"],
    "GBU":                  [],   # read for logging/QA only — no attribute
    "Channel Type":         [],   # read for logging/QA only — see DEVIATION 2
    "Category":             [],   # read for logging/QA only — no attribute
}


# ══════════════════════════════════════════════════════════════════
# LOV TABLES  (IDs resolved against the MDD; hard-coded fallbacks match
#              the MDD LOV sheets — same convention as every NB ETL)
# ══════════════════════════════════════════════════════════════════

LOV_BRAND = {
    "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
    "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
    "LOT": "LOTTO",   "BIR": "BIRKENSTOCK",
}

# MDD "Gender LOV" / master "SAP Gender" — value → ID
LOV_GENDER = {"Male": "M", "Female": "F", "Unisex": "U"}

# MDD "Age LOV" / master "SAP Age" — value → ID
LOV_AGE = {"Adults": "AD", "Children": "CH", "All Ages": "AA"}

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

# MDD "Article Category LOV" — code → name (1 = Generic, per v6 R097 default)
LOV_SAP_ARTICLE_CATEGORY = {
    "1": "Generic", "0": "Single", "2": "Variant",
    "10": "Sell set (Hampers)", "11": "Prepack (Musical Box)",
}

# MDD "BY Article Type" — code → name
LOV_BY_ARTICLE_TYPE = {
    "Inline": "Inline", "License": "License", "SSE": "SSE",
}

# MDD "Nature of Article LOV" — SMP = Sample (v6 R105 default)
LOV_NATURE_OF_ARTICLE = {
    "SMP": "Sample",
}

# MDD "Material LOV" / master "Material Type" — ZINA = Intercompany Articles
# (v6 R096 "Material Type — Intercompany")
MATERIAL_TYPE_ID = "ZINA"

# MDD "SAP Product Flag LOV" — Intercompany = A (v6 R090 default)
SAP_PRODUCT_FLAG_INTERCOMPANY = ("A", "Intercompany")

# MDD "UOM LOV" — EA = Each (v6 R103 default)
UOM_EACH = ("EA", "Each")

# ── v6 R051: SKU first letter → SAP Gender code (App rule) ─────────────────
#   W : Female   M : Male   Y : Unisex — used by the App/Acc sample pair;
#   this Acc module ignores it ("Acc default: Unisex"); the constant stays
#   so both sibling modules share one vocabulary.
SKU_FIRST_CHAR_TO_GENDER: dict[str, str] = {
    "W": "F", "M": "M", "Y": "U",
}

# Gender code → MDD Gender LOV display name (label used in descriptions)
GENDER_CODE_TO_LABEL = {"M": "MALE", "F": "FEMALE", "U": "UNISEX"}

# ── Color Family → SAP Color Code (MDD "Color Code LOV", the 103-colour ────
#    SAP palette). Fallback table verified against the MDD sheet; the live
#    LOV is preferred at runtime (see _resolve_sap_color_id). v6 R040 note:
#    "Mapping only for RED = RED . or system to choose the close value from
#    LOV" — hence the fuzzy handling of the LOV's "RED ." spelling.
COLOR_FAMILY_TO_SAP_COLOR: dict[str, str] = {
    "BLACK":    "005",   "WHITE":    "W",     "GREY":    "GRE",
    "BLUE":     "12W",   "GREEN":    "G",     "RED":     "R",
    "BEIGE":    "18",    "BROWN":    "700",   "YELLOW":  "Y",
    "ORANGE":   "O",     "PINK":     "PK",    "PURPLE":  "P",
    "NO COLOR": "000",
}

# ── COO (ISO-3 code) → MDD "Country Origin LOV" (2-letter) ─────────────────
#    v6 R054: CHN = China, LKA = Sri Lanka, IDN = Indonesia, JPN = Japan,
#             VNM = Vietnam, THA = Thailand. (USA appears in live data and
#             resolves to US/USA in the MDD LOV.)
COO_ISO3_TO_ORIGIN: dict[str, tuple[str, str]] = {
    "CHN": ("CN", "China"),
    "LKA": ("LK", "Sri Lanka"),
    "IDN": ("ID", "Indonesia"),
    "JPN": ("JP", "Japan"),
    "VNM": ("VN", "Vietnam"),
    "THA": ("TH", "Thailand"),
    "USA": ("US", "USA"),
}

# MDD AT_Generic / AT_Variant / description length limits
MAX_GENERIC_LEN           = 12   # 3 brand + 1 "S" + 7 style + 1 letter
MAX_VARIANT_LEN           = 18   # generic + 3 color + 3 size
MAX_GENERIC_DESC_LEN      = 40
MAX_VARIANT_DESC_LEN      = 40
MAX_PRINCIPAL_COLOR_CODE  = 7    # MDD AT_PrincipalColorCode

# ── The sample running alphabet (BM rows 23/24/26) ──────────────────────────
RUNNING_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
STYLE_BASE_LEN    = 7            # "Up to 7 Digits Principal Style Code"

DIVISION = "Accessories"   # fixed by file type → PPH parent letter "E"

DIVISION_PARENT_MAP: dict[str, str] = {
    "ACCESSORIES": "E", "FOOTWEAR": "F", "APPAREL": "A",
    "EQUIPMENT":   "Q", "TOYS":     "T",
}

# SAP Size default (v6 R041 — one variant per sample row → SS;
# MDD Size Code LOV: description "SS" ↔ code "0SS")
SAP_SIZE_SS          = ("0SS", "SS")

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

# Serialise STEP-namespace elements with the DEFAULT namespace (clean,
# unprefixed tags after the _XMLNS_RE strip below) — same convention as
# every other NB ETL module. Without this, output depends on whether a
# sibling module happened to register the namespace first.
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDLoader:
    """Master Data Dictionary loader — same contract as every NB ETL."""

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
                "cardinality":  _v("Mandatory"),
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


class BrandMappingSampleLoader:
    """
    Parses the "NEW BALANCE SAMPLE" sheet of the NEW Brand Mapping file.

    The ETL's attribute mapping is hard-coded from that sheet (repo
    convention — same as GTM / Preline / the footwear sample ETL), but the
    loader lets the module SELF-CHECK at runtime that every attribute ID it
    emits is declared in the live mapping sheet. Undeclared IDs are logged
    as warnings so a mapping-sheet revision surfaces immediately instead of
    silently diverging.
    """

    SHEET_NAME = "NEW BALANCE SAMPLE"

    def __init__(self, path: Path):
        self.path  = path
        self.sheet = self.SHEET_NAME
        self.declared_ids: set[str] = set()
        self.rows: list[dict] = []
        self._load()

    def _load(self):
        log.info("[BrandMapping] Loading: %s → sheet '%s'", self.path.name, self.SHEET_NAME)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        if self.SHEET_NAME not in wb.sheetnames:
            log.warning(
                "[BrandMapping] Sheet '%s' not found in %s — self-check disabled "
                "(available: %s)", self.SHEET_NAME, self.path.name, wb.sheetnames[:8],
            )
            wb.close()
            return
        ws   = wb[self.SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(
            (i for i, r in enumerate(rows[:5])
             if r and r[1] is not None and str(r[1]).strip() == "Stibo Attribute ID"),
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
                "attribute": str(row[0]).strip() if row[0] else "",
                "id":        attr_id,
                "validation": str(row[2]).strip() if len(row) > 2 and row[2] else "",
            })
        wb.close()
        log.info("[BrandMapping] %d attribute IDs declared in '%s'", len(self.declared_ids), self.SHEET_NAME)


class AccessoriesBreakoutLoader:
    """
    Loads the New Balance Licensed Accessories Sample Breakout workbook.

    Sheet selection (see SHEET_MODE / DEVIATION 1):
      * "first_visible"  → the first visible sheet ("Sheet1")
      * "all_data_sheets"→ every visible sheet whose name is not a Summary tab

    The header row is located dynamically by scanning for a row whose first
    populated cell is exactly "SKU" — the sheet carries a title row and a
    blank row above it (header on physical row 3).

    Running-alphabet letters (BM rows 23/24/26) are assigned per 7-char
    SKU base in file order here, deterministically, BEFORE any parallel
    mapping — see the module docstring for the full rationale.
    """

    ROW_KEY = "SKU"

    def __init__(self, path: Path, sheet_mode: str = SHEET_MODE):
        self.path   = path
        self.mode   = sheet_mode
        self.frames: list[pd.DataFrame] = []
        self.sheets: list[str] = []
        self.running_letter_wraps = 0
        self._load()

    def _load(self):
        log.info("[Sample-ACC] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        visible = [sn for sn in wb.sheetnames if wb[sn].sheet_state == "visible"]
        if self.mode == "all_data_sheets":
            targets = [
                sn for sn in visible
                if not any(k in sn.lower() for k in SUMMARY_SHEET_KEYWORDS)
            ]
        else:
            targets = visible[:1]     # first visible sheet only

        for target in targets:
            ws   = wb[target]
            rows = list(ws.iter_rows(values_only=True))

            hdr_idx = next(
                (i for i, r in enumerate(rows)
                 if any(str(v).strip() == self.ROW_KEY for v in r if v is not None)),
                None,
            )
            if hdr_idx is None:
                log.warning("[Sample-ACC] No '%s' header row in sheet '%s' — skipping.",
                            self.ROW_KEY, target)
                continue

            log.info("[Sample-ACC] Sheet '%s': header at row index %d (row %d)",
                     target, hdr_idx, hdr_idx + 1)
            header = [
                str(h).replace("\n", " ").strip() if h is not None else f"col_{i}"
                for i, h in enumerate(rows[hdr_idx])
            ]
            df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
            df["_ExcelRow"] = range(hdr_idx + 2, hdr_idx + 2 + len(df))
            df["_Sheet"]    = target

            key_col = next(
                (c for c in (self.ROW_KEY, "Sku", "SKU Code") if c in df.columns), None
            )
            if key_col:
                df = df[
                    df[key_col].notna()
                    & (df[key_col] != "")
                    & (df[key_col] != 0)
                    & (~df[key_col].astype(str).str.strip().isin(["", "None", "nan"]))
                ]

            df = df.reset_index(drop=True)
            df["_RunningLetter"] = _assign_running_letters(
                df[self.ROW_KEY], counter=self,
            ) if self.ROW_KEY in df.columns else ""

            self.frames.append(df)
            self.sheets.append(target)

        wb.close()

        missing = [c for c in SOURCE_COLUMNS if c not in (self.frames[0].columns if self.frames else [])]
        if missing:
            log.warning("[Sample-ACC] Expected source column(s) absent: %s", missing)

        total = sum(len(f) for f in self.frames)
        log.info("[Sample-ACC] %d SKU row(s) loaded from sheet(s): %s", total, self.sheets)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "N/A") else s


def _clean(code: str) -> str:
    """Alphanumeric-only, upper-cased (used for SAP-side codes)."""
    return re.sub(r"[^A-Z0-9]", "", (code or "").upper())


def _sku_base(sku: str) -> str:
    """
    The 7-char style base of a SKU (BM row 23: "Up to 7 Digits Principal
    Style Code"). The SKU embeds Style/Part No. + "_" + Color/Width, so the
    first 7 alphanumeric chars of the SKU are the style's first 7 chars.
    """
    return _clean(sku)[:STYLE_BASE_LEN]


def _assign_running_letters(sku_series, counter=None) -> list[str]:
    """
    Assigns the "1 Digit Running Alphabeth" (BM rows 23/24/26).

    One letter per row, sequential per 7-char SKU base in file order:
    the first row whose base is X gets "A", the second "B", …  This keeps
    AT_SAPStyleCode / AT_Generic unique even when distinct styles collide
    on a 7-char base (live data: MT71W0NL vs MT71W0NC) or when one style
    carries many colours (live max: 11 rows per base).

    Bases with more than 26 rows wrap back to "A" — a collision the caller
    flags as a validation warning (no such base exists in the live file).
    """
    counts: dict[str, int] = {}
    letters: list[str] = []
    for raw in sku_series:
        base = _sku_base(_s(raw))
        n = counts.get(base, 0)
        counts[base] = n + 1
        if n >= len(RUNNING_ALPHABET) and counter is not None:
            counter.running_letter_wraps += 1
        letters.append(RUNNING_ALPHABET[n % len(RUNNING_ALPHABET)])
    return letters


def _resolve_sap_color_id(color_family: str, mdd=None) -> str:
    """
    Color Family → SAP Color Code (AT_Color LOV — MDD "Color Code LOV",
    the 103-colour SAP palette).

    Resolution order (mirrors the App/Acc price-list ETL):
      1. exact match on the LOV description        (e.g. "BLACK"  → "005")
      2. fuzzy match, LOV desc trailing " ."        (e.g. "RED ." → "R";
         this is exactly what v6 R040 means by "Mapping only for RED = RED .")
      3. reverse lookup if the family is already a code (e.g. "EE0" → "EE0")
      4. hard-coded fallback table (verified against the MDD)
      5. empty string + caller warning — never the raw string, so Stibo
         cannot receive an invalid LOV ID
    """
    fam = (color_family or "").strip().upper()
    if not fam:
        return ""

    lov = (mdd.lovs.get("Color Code") or {}) if mdd else {}
    cid = lov.get(fam, "")
    if not cid:
        for desc, code in lov.items():
            if str(desc).rstrip(". ").strip().upper() == fam:
                cid = code
                break
    if not cid and lov:
        rev = {str(v).upper(): str(v) for v in lov.values()}
        cid = rev.get(fam, "")
    if not cid:
        cid = COLOR_FAMILY_TO_SAP_COLOR.get(fam, "")
    return str(cid).strip()


def _coo_to_origin(coo: str) -> tuple[str, str]:
    """
    COO (ISO-3, e.g. "CHN") → (MDD Country Origin LOV 2-letter ID, name),
    per v6 R054. Unknown codes fall back to (raw, raw) and are flagged by
    the validator so a bad country never silently maps.
    """
    code = (coo or "").strip().upper()
    if not code:
        return "", ""
    if code in COO_ISO3_TO_ORIGIN:
        return COO_ISO3_TO_ORIGIN[code]
    return code[:2], code


def _sku_first_char_to_gender_code(sku: str) -> str:
    """
    v6 R051 Acc rule: "Acc default: Unisex" — the Acc SKUs start with A/L
    (no W/M/Y gender signal), so the default Unisex is returned for every
    row. Kept as a named function so the rule's provenance stays visible.
    """
    return "U"


# ── THE SAMPLE "S" PREFIX ────────────────────────────────────────────────────
def _sample_prefix(sku: str) -> str:
    """BM row 14 Mapping Logic: '"S" + Principal Style Code' (App col: SKU)."""
    return f"S{(sku or '').strip()}"


def _build_sap_style_code(sku: str, running_letter: str) -> str:
    """
    BM row 23 / v6 R035 (App/Acc): Prefix S + up to 7 chars of the
    Principal Style Code + 1 running-alphabet letter  → 9 chars.
    """
    return f"S{_sku_base(sku)}{running_letter}"


def _build_generic_code(brand_code: str, sku: str, running_letter: str) -> str:
    """
    BM row 24 / v6 R036 (App/Acc): 3-digit brand code + prefix "S" +
    up to 7 chars of the Principal Style Code + 1 running-alphabet letter
    (3+1+7+1 = 12 chars — exactly the MDD AT_Generic maximum).
    """
    brand = _clean(brand_code)[:3]
    return f"{brand}S{_sku_base(sku)}{running_letter}"[:MAX_GENERIC_LEN]


def _build_variant_code(generic_code: str, sap_color_id: str, sap_size_id: str) -> str:
    """
    BM row 26 / v6 R038 (App/Acc): Generic + 3-digit SAP color code +
    3-digit SAP size code (12+3+3 = 18 chars — exactly the MDD maximum).
    """
    return f"{generic_code}{(sap_color_id or '')}{(sap_size_id or '')}"[:MAX_VARIANT_LEN]


def _build_generic_description(
    brand_code: str, principal_style_code: str,
    gender_code: str, sap_color_name: str,
) -> str:
    """
    BM row 25 / v6 R037: MAA Generic Description Mapping —
    max 40 chars: 3-digit brand code + Principal style + (Age/Gender) + Color.
    Same formula astec/recap_main.py and the FW sample ETL implement.
    """
    parts = (
        _clean(brand_code)[:3],
        (principal_style_code or "").strip().upper(),
        GENDER_CODE_TO_LABEL.get(gender_code, ""),
        (sap_color_name or "").strip().upper(),
    )
    return " ".join(p for p in parts if p)[:MAX_GENERIC_DESC_LEN]


def _build_variant_description(generic_desc: str, sap_size_name: str) -> str:
    """BM row 27 / v6 R039: generic description + Size, max 40 chars."""
    return f"{generic_desc} {(sap_size_name or '').strip().upper()}".strip()[:MAX_VARIANT_DESC_LEN]


def _build_inbound_key(brand_code: str, sku: str) -> str:
    """
    KEY_InboundArticle / AT_InboundGenericCode — repo convention
    (every NB ETL keys the generic article 1:1 with a source row), carrying
    the sample "S" prefix:  brand(3) + "S" + <row-unique SKU>.
    """
    brand = _clean(brand_code)[:3]
    return f"{brand}S{_clean(sku)}"


def map_sku(row: pd.Series, brand_code: str = "NEW", mdd=None) -> dict:
    """Map one Licensed Accessories Sample Breakout row to a normalized sample-article dict."""
    sku             = _s(row.get("SKU"))
    display_name    = _s(row.get("Product Display Name"))
    color_code      = _s(row.get("Color/Width"))
    color_family    = _s(row.get("Color Family"))
    size            = _s(row.get("Size"))
    coo             = _s(row.get("COO"))
    product_line    = _s(row.get("Product Line"))
    line_plan_biz   = _s(row.get("Line Plan Business"))
    gbu             = _s(row.get("GBU"))
    channel_type    = _s(row.get("Channel Type"))
    category        = _s(row.get("Category"))
    excel_row       = row.get("_ExcelRow")
    sheet_name      = row.get("_Sheet", "")
    running_letter  = _s(row.get("_RunningLetter")) or "A"

    # v6 R051 Acc: default Unisex; v6 R050: Adults default
    gender_code = _sku_first_char_to_gender_code(sku)
    age_code    = "AD"                       # v6 R050 App/Acc default: Adults

    # SAP-side values
    sap_color_id   = _resolve_sap_color_id(color_family, mdd)
    sap_color_nm   = color_family            # family text ≈ LOV description
    sap_size_id, sap_size_nm = SAP_SIZE_SS   # v6 R041 default SS (one variant)
    coo_id, coo_name = _coo_to_origin(coo)

    principal_style_code = _sample_prefix(sku)          # "S" + SKU (raw)
    sap_style_code       = _build_sap_style_code(sku, running_letter)
    generic_code         = _build_generic_code(brand_code, sku, running_letter)
    variant_code         = _build_variant_code(generic_code, sap_color_id, sap_size_id)
    generic_desc         = _build_generic_description(
        brand_code, principal_style_code, gender_code, sap_color_nm,
    )
    variant_desc         = _build_variant_description(generic_desc, sap_size_nm)
    inbound_key          = _build_inbound_key(brand_code, sku)

    return {
        # ── source values ───────────────────────────────────────
        "sku":               sku,
        "display_name":      display_name,
        "color_code":        color_code,
        "color_family":      color_family,
        "size":              size,
        "coo":               coo,
        "product_line":      product_line,
        "line_plan_biz":     line_plan_biz,
        "gbu":               gbu,
        "channel_type":      channel_type,
        "category":          category,
        "running_letter":    running_letter,
        "excel_row":         int(excel_row) if excel_row is not None else None,
        "sheet_name":        sheet_name,
        # ── derived / codes ─────────────────────────────────────
        "brand_code":          brand_code,
        "principal_style_code": principal_style_code,   # "S"+SKU
        "sap_style_code":      sap_style_code,          # "S"+base7+letter
        "generic_code":        generic_code,            # NEW+S+base7+letter
        "variant_code":        variant_code,            # generic+color+size
        "generic_desc":        generic_desc,
        "variant_desc":        variant_desc,
        "inbound_key":         inbound_key,             # NEW+S+SKU
        # ── gender / age ────────────────────────────────────────
        "gender_code":         gender_code,
        "gender_label":        {v: k for k, v in LOV_GENDER.items()}.get(gender_code, "Unisex"),
        "age_code":            age_code,
        "age_label":           {v: k for k, v in LOV_AGE.items()}.get(age_code, "Adults"),
        # ── SAP color / size / COO ──────────────────────────────
        "sap_color_id":        sap_color_id,
        "sap_color_name":      sap_color_nm,
        "sap_size_id":         sap_size_id,
        "sap_size_name":       sap_size_nm,
        "coo_id":              coo_id,
        "coo_name":            coo_name,
        # ── constants ───────────────────────────────────────────
        "division":            product_line or DIVISION,
        "art_category":        "1",                    # Generic
        "article_type":        "Inline",               # v6 R106 default
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VALIDATOR
# ══════════════════════════════════════════════════════════════════

def validate(mapped: dict, mdd: MDDLoader) -> list[str]:
    """Validates only the fields this ETL actually emits."""
    warns = []
    sku   = mapped["sku"] or "<no SKU>"

    # Mandatory (per MDD cardinality) emitted attributes
    field_to_at = {
        "sku":            "AT_PrincipalStyleCode",
        "generic_code":   "AT_Generic",
        "variant_code":   "AT_Variant",
        "gender_code":    "AT_Gender",
        "age_code":       "AT_SAPAge",
        "brand_code":     "AT_Brand",
        "sap_style_code": "AT_SAPStyleCode",
    }
    for field, at_id in field_to_at.items():
        meta = mdd.attributes.get(at_id, {}) if mdd else {}
        card = (meta.get("cardinality") or "").lower() if meta else ""
        val  = mapped.get(field, "")
        if (not card or "mandatory" in card or "conditional" in card) and (
            not val or str(val).strip() in ("", "None", "nan")
        ):
            warns.append(f"[{sku}] MISSING mandatory: {at_id}")

    # Source-quality warnings (v6 notes)
    if not mapped.get("display_name"):
        warns.append(f"[{sku}] BLANK Product Display Name — SKU fallback applied")
    if not mapped.get("color_family"):
        warns.append(f"[{sku}] BLANK Color Family — AT_PrincipalColorName/AT_Color omitted")
    elif not mapped.get("sap_color_id"):
        warns.append(
            f"[{sku}] Color Family '{mapped['color_family']}' not found in Color Code LOV — "
            "AT_Color ID empty (v6 R040: system to choose the close value)"
        )
    if not mapped.get("coo"):
        warns.append(f"[{sku}] BLANK COO — AT_CountryOrigin omitted")
    elif mapped.get("coo", "").upper() not in COO_ISO3_TO_ORIGIN:
        warns.append(f"[{sku}] COO '{mapped['coo']}' outside the v6 R054 list — raw code sent")
    if not mapped.get("line_plan_biz"):
        warns.append(f"[{sku}] MISSING source value: Line Plan Business")
    if not mapped.get("product_line"):
        warns.append(f"[{sku}] MISSING source value: Product Line (division defaulted)")
    if not mapped.get("size"):
        warns.append(f"[{sku}] MISSING source value: Size (AT_PrincipalSize omitted)")
    if len(mapped.get("color_code", "")) > MAX_PRINCIPAL_COLOR_CODE:
        warns.append(
            f"[{sku}] Color/Width '{mapped['color_code']}' exceeds MDD max "
            f"{MAX_PRINCIPAL_COLOR_CODE} chars — truncated"
        )
    return warns


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — XML BUILDER HELPERS
# ══════════════════════════════════════════════════════════════════

def _clean_xml_value(raw) -> str:
    text = "" if raw is None else str(raw).strip()
    return "" if text in ("", "None", "nan", "NaT") else text


def _val(parent, attr_id, value="", id_val=""):
    normalized_id    = _clean_xml_value(id_val)
    normalized_value = _clean_xml_value(value)
    if not normalized_id and not normalized_value:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if normalized_id:
        el.set("ID", normalized_id)
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
    return "E"   # this module's division default: Accessories


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — GENERIC VALUE WRITER
# ══════════════════════════════════════════════════════════════════

EMITTED_ATTRIBUTE_IDS: list[str] = [
    "AT_Country", "AT_CompanyCode", "AT_SBU", "AT_Brand", "AT_BrandGroup",
    "AT_PrincipalStyleCode", "AT_PrincipalStyleDescription",
    "AT_PrincipalColorCode", "AT_PrincipalColorName", "AT_PrincipalSize",
    "AT_SAPStyleCode", "AT_Generic", "AT_GenericDescription",
    "AT_Variant", "AT_VariantDescription",
    "AT_Color", "AT_Size", "AT_Gender", "AT_SAPAge",
    "AT_CountryOrigin", "AT_Season", "AT_SeasonYear",
    "AT_PrincipalMerchandiseHierarchyL1", "AT_PrincipalMerchandiseHierarchyL2",
    "AT_SAPArticleCategory", "AT_BYArticleType", "AT_NatureOfArticle",
    "AT_MaterialType", "AT_SAPProductFlag", "AT_UOM",
    "AT_EcomIndicator", "AT_MainVendorIdentification", "AT_ArticleStatus",
    "AT_InboundGenericCode",
]


def _add_generic_values(vals_el, art, brand_name, comp_code, sbu, mdd=None, bm=None):
    """
    Writes exactly the attributes the two SAMPLE mapping sheets define as
    populated for the accessories sample flow (see EMITTED_ATTRIBUTE_IDS).
    Every other attribute row in those sheets is N/A / Manual input /
    Smartsheet-only and is deliberately NOT emitted.
    """
    # ── Context (portal / filename metadata, not a source column) ──
    sbu_code, sbu_label = _lov(sbu, LOV_SBU, sbu)
    _multival(vals_el, "AT_SBU", sbu_code, sbu_label)

    cc_label = LOV_COMPANY_CODE.get(comp_code, comp_code)
    _multival(vals_el, "AT_CompanyCode", comp_code, cc_label)

    b_code, b_label = _lov(art["brand_code"], LOV_BRAND, art["brand_code"])
    _val(vals_el, "AT_Brand",      b_label, id_val=b_code)
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)

    # ── Principal style — THE SAMPLE "S" PREFIX (BM row 14) ────────
    _val(vals_el, "AT_PrincipalStyleCode", art["principal_style_code"])
    # v6 R027: Product Display Name, blank → SKU fallback
    _val(vals_el, "AT_PrincipalStyleDescription",
         art["display_name"] or art["sku"])

    # ── Principal colour (App/Acc — v6 R028 / R029) ────────────────
    _val(vals_el, "AT_PrincipalColorCode", art["color_code"][:MAX_PRINCIPAL_COLOR_CODE])
    _val(vals_el, "AT_PrincipalColorName", art["color_family"])

    # ── Principal size (App/Acc — v6 R034) ─────────────────────────
    _val(vals_el, "AT_PrincipalSize", art["size"])

    # ── Gender / age (Acc — Unisex default; Adults default) ──────
    _val(vals_el, "AT_Gender", art["gender_label"], id_val=art["gender_code"])
    _val(vals_el, "AT_SAPAge",  art["age_label"],  id_val=art["age_code"])

    # ── SAP-computed codes [SAP-COMPUTED] (v6 R035-R039) ───────────
    _val(vals_el, "AT_SAPStyleCode", art["sap_style_code"])
    _val(vals_el, "AT_Generic",      art["generic_code"])
    _val(vals_el, "AT_Variant",      art["variant_code"])
    _val(vals_el, "AT_GenericDescription", art["generic_desc"])
    _val(vals_el, "AT_VariantDescription",  art["variant_desc"])

    # ── SAP color (Color Family → Color Code LOV) / SAP size ───────
    if art["sap_color_id"]:
        _val(vals_el, "AT_Color", art["sap_color_name"], id_val=art["sap_color_id"])
    _val(vals_el, "AT_Size", art["sap_size_name"], id_val=art["sap_size_id"])

    # ── Country of origin (Acc — COO column, v6 R054) ─────────────
    if art["coo_id"]:
        _val(vals_el, "AT_CountryOrigin", art["coo_name"], id_val=art["coo_id"])

    # ── Season (from filename/portal metadata) ─────────────────────
    sea_raw   = art.get("season", "") or ""
    sea_code  = sea_raw[:2].upper() if len(sea_raw) >= 2 else sea_raw
    sea_label = LOV_SEASON.get(sea_raw, LOV_SEASON.get(sea_code, sea_raw))
    _val(vals_el, "AT_Season", sea_label, id_val=sea_code)

    year_raw = sea_raw[2:] if len(sea_raw) > 2 else ""
    year_val = f"20{year_raw}" if len(year_raw) == 2 else year_raw
    _val(vals_el, "AT_SeasonYear", year_val)

    # ── Merchandise hierarchy (Acc — Product Line / Line Plan) ─────
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", art["product_line"])
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", art.get("line_plan_biz", ""))

    # ── Sample-flow constants (v6 "NEW BALANCE - SAMPLE") ──────────
    cat_code  = art["art_category"]                       # "1" = Generic
    cat_label = LOV_SAP_ARTICLE_CATEGORY.get(cat_code, "Generic")
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)

    at_norm  = art["article_type"]                        # "Inline"
    at_label = LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm)
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)

    noa_code  = "SMP"                                      # Sample
    noa_label = LOV_NATURE_OF_ARTICLE.get(noa_code, "Sample")
    _val(vals_el, "AT_NatureOfArticle", noa_label, id_val=noa_code)

    _val(vals_el, "AT_MaterialType", id_val=MATERIAL_TYPE_ID)          # ZINA

    flag_id, flag_label = SAP_PRODUCT_FLAG_INTERCOMPANY               # A
    _val(vals_el, "AT_SAPProductFlag", flag_label, id_val=flag_id)

    uom_id, uom_label = UOM_EACH                                       # EA
    _val(vals_el, "AT_UOM", uom_label, id_val=uom_id)

    # v6 R025: Ecom Indicator — Manual input / Default : No
    _val(vals_el, "AT_EcomIndicator", "No", id_val="N")

    # v6 R056: Main Vendor Identification — Default "1"
    _val(vals_el, "AT_MainVendorIdentification", "1")

    # v6 R245: Article Status — Default when created: Active
    # (no LOV sheet defines an ID for it — value text only)
    _val(vals_el, "AT_ArticleStatus", "Active")

    # ── Pipeline key (repo convention) ─────────────────────────────
    _val(vals_el, "AT_InboundGenericCode", art["inbound_key"])

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
    season_label_map = {   # MDD "Season LOV" names
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
    sku = art["sku"]
    if not sku:
        return ""

    div_letter  = _get_division_code(art.get("division", DIVISION))
    parent_id   = f"PPH_{div_letter}-TempSubCat"
    key_generic = art["inbound_key"]

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   parent_id)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_generic

    # v6 R027: Product Display Name, blank → SKU fallback
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = (
        art["display_name"] or sku
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
    mapped = map_sku(row, brand_code=brand_code, mdd=mdd)
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

    mdd_f  = first(MDD_DIR)
    attr_f = first(ATTR_DIR)
    sm_f   = first(SAMPLE_DIR)

    for label, val in [
        ("MDD",                 mdd_f),
        ("Attributes",          attr_f),
        ("Accessories Sample",  sm_f),
    ]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    mdd = MDDLoader(mdd_f)
    bm  = BrandMappingSampleLoader(attr_f)

    # Self-check: every ID this module emits must be declared in the
    # live "NEW BALANCE SAMPLE" sheet (see BrandMappingSampleLoader).
    if bm.declared_ids:
        undeclared = [a for a in EMITTED_ATTRIBUTE_IDS if a not in bm.declared_ids]
        if undeclared:
            log.warning(
                "[BrandMapping SELF-CHECK] %d emitted attribute ID(s) not declared "
                "in sheet '%s': %s — mapping sheet may have been revised.",
                len(undeclared), bm.SHEET_NAME, undeclared,
            )
        else:
            log.info("[BrandMapping SELF-CHECK] all %d emitted attribute IDs are "
                     "declared in '%s'.", len(EMITTED_ATTRIBUTE_IDS), bm.SHEET_NAME)

    all_warnings: list[str] = []

    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    for sm_path in [sm_f]:
        log.info("─── Processing Licensed Accessories Sample Breakout file: %s ───", sm_path.name)

        loader = AccessoriesBreakoutLoader(sm_path, sheet_mode=SHEET_MODE)
        if not loader.frames:
            log.warning("[Sample-ACC] No data frames loaded — skipping file.")
            continue

        total_rows = sum(len(f) for f in loader.frames)
        log.info("[Sample-ACC] %d valid SKU row(s) to process", total_rows)

        if loader.running_letter_wraps:
            warn_msg = (
                f"[FILE {sm_path.name}] {loader.running_letter_wraps} running-letter "
                "wrap(s): a 7-char SKU base carries more than 26 rows — SAP style / "
                "generic codes will collide for those rows"
            )
            log.warning("[Sample-ACC] %s", warn_msg)
            all_warnings.append(warn_msg)

        out_name = f"{sm_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        log.info("Pass 1/2 — parallel map+validate (%d rows) …", total_rows)
        mapped_skus: list[dict] = []
        for df in loader.frames:
            rows = [
                row for _, row in df.iterrows()
                if str(row.get("SKU", "")).strip() not in ("", "None", "nan")
            ]
            num_workers = min(8, max(1, len(rows)))
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
            mapped_skus.extend(m for _, m, _ in ordered)
            del ordered

        # TEST LIMIT: restrict to specific 1-based Excel rows, e.g. {833, 844}
        # TEST_ROWS = {833, 844}
        # mapped_skus = [m for m in mapped_skus if m.get("excel_row") in TEST_ROWS]

        # KEY_InboundArticle must stay 1:1 with source rows.
        keys  = [m["inbound_key"] for m in mapped_skus if m.get("sku")]
        dupes = len(keys) - len(set(keys))
        if dupes:
            log.warning(
                "[Sample-ACC] %d duplicate KEY_InboundArticle value(s) — "
                "articles will be merged in STEP.", dupes,
            )
            all_warnings.append(
                f"[FILE {sm_path.name}] {dupes} duplicate KEY_InboundArticle values"
            )

        # SAP style / generic codes must be unique (running alphabet contract).
        sap_keys  = [m["sap_style_code"] for m in mapped_skus if m.get("sku")]
        sap_dupes = len(sap_keys) - len(set(sap_keys))
        if sap_dupes:
            log.warning(
                "[Sample-ACC] %d duplicate AT_SAPStyleCode value(s) — "
                "check running-alphabet assignment.", sap_dupes,
            )
            all_warnings.append(
                f"[FILE {sm_path.name}] {sap_dupes} duplicate AT_SAPStyleCode values"
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
                if not art.get("sku"):
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
        print(f"  Sample ACC rows     : {total_rows}",    flush=True)
        print(f"  Mapped SKUs         : {written_count}", flush=True)
        print(f"  XML file size       : {file_kb}KB",     flush=True)
        print("════════════════════════════════════════════════════", flush=True)

    rpt_path = LOG_DIR / f"validation_nb_sample_accessories_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")
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
        description="Stibo Inbound XML Generator — New Balance Sample (Accessories) v1.0"
    )
    p.add_argument("--brand",        default="New Balance")
    p.add_argument("--brand-code",   default="NEW")
    p.add_argument("--comp-code",    default="0888")
    p.add_argument("--sbu",          default="AC")
    p.add_argument("--season",       default="SS27")
    p.add_argument("--seq",          default=1, type=int)
    p.add_argument("--country-code", default="")
    run(p.parse_args())


if __name__ == "__main__":
    main()
