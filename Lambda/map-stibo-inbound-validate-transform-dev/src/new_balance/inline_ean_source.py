"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — NEW BALANCE  v2.0               ║
║   EAN Source (UPC / Barcode) → Stibo STEP XML                   ║
╚══════════════════════════════════════════════════════════════════╝

Source file  : Master_UPC_List  (EAN Source — Footwear / Apparel)
Primary tab  : UPC-Report-D365 (or first sheet)

Variant code formula  (MAA Article Code — max 18 chars):
    Generic Code (max 12) + 3-digit MAA Color Code + 3-digit MAA Size Code

Column → AT_ attribute mappings (MDD confirmed)
─────────────────────────────────────────────────────────────────
  ITEM_NUMBER     → AT_PrincipalStyleCode  / generic parent key
  SKU             → full variant identifier
  SKU Color Code  → AT_ColorCode + AT_Color + AT_MAAColor
  SKU Size        → AT_SizeCode  + AT_Size  + AT_MAASize
  UPC#            → AT_Barcode   (DataContainer → DC_Barcode)
  UK Size         → AT_UKSize    (when present)
  Country         → AT_Country   (mandatory per MDD)
─────────────────────────────────────────────────────────────────

v2.0 Changes:
  - Correct MDD attributes: AT_ColorCode, AT_Color, AT_SizeCode, AT_Size
  - Grouping logic: same item_number → one Generic, multiple Variants
  - Same variant_code (same size+color) → EAN updated in existing variant
  - New variant_code → new Variant element created inside Generic
─────────────────────────────────────────────────────────────────
"""

import re
import os
import sys
import logging
import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR   = BASE_DIR / "input"
EAN_DIR     = INPUT_DIR / "ean_source"
MDD_DIR     = INPUT_DIR / "mdd"
OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for _d in [EAN_DIR, MDD_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"run_nb_ean_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — MDD LOADER
# ══════════════════════════════════════════════════════════════════

class MDDLoader:

    def __init__(self, path: Path):
        self.company_codes: set[str] = set()
        self.sbu_codes: set[str] = set()
        self.path = path
        self.color_desc_to_code: dict[str, str] = {}
        self.color_code_set: set[str] = set()
        self.size_to_code: dict[str, str] = {}
        self.country_desc_to_code: dict[str, str] = {}
        self._load()

    # ── private loaders ──────────────────────────────────────────

    def _load_company_lov(self, wb):
        if "Company Code LOV" not in wb.sheetnames:
            return
        ws = wb["Company Code LOV"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if row[0]:
                self.company_codes.add(str(row[0]).strip())

    def _load_sbu_lov(self, wb):
        if "SBU LOV" not in wb.sheetnames:
            return
        ws = wb["SBU LOV"]
        for row in list(ws.iter_rows(values_only=True))[1:]:
            if row[0]:
                self.sbu_codes.add(str(row[0]).strip())

    def _load_color_lov(self, wb):
        if "Color Code LOV" not in wb.sheetnames:
            log.warning("[MDD] 'Color Code LOV' sheet not found — color mapping disabled")
            return
        ws   = wb["Color Code LOV"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[1:]:
            if not row or row[0] is None:
                continue
            code = str(row[0]).strip()
            desc = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ""
            if code:
                self.color_code_set.add(code)
                if desc:
                    self.color_desc_to_code[desc] = code
        log.info("[MDD] Color LOV: %d entries", len(self.color_desc_to_code))

    def _load_size_lov(self, wb):
        if "Size Code LOV" not in wb.sheetnames:
            log.warning("[MDD] 'Size Code LOV' sheet not found — size mapping disabled")
            return
        ws   = wb["Size Code LOV"]
        rows = list(ws.iter_rows(values_only=True))

        # Pass 1 — col 5/6 (MAA Size Code / Description)
        for row in rows[1:]:
            if len(row) > 6 and row[5] is not None:
                code = str(row[5]).strip()
                desc = str(row[6]).strip() if row[6] else ""
                if code:
                    self.size_to_code[code.upper()] = _pad_size_code(code)
                    if desc:
                        self.size_to_code[desc.upper()] = _pad_size_code(code)

        # Pass 2 — col 2/3 (SAP Size Code / Description)
        for row in rows[1:]:
            if len(row) > 3 and row[2] is not None:
                sap_code = str(row[2]).strip()
                desc     = str(row[3]).strip() if row[3] else ""
                key_sap  = sap_code.upper()
                key_desc = desc.upper()
                if key_sap not in self.size_to_code and sap_code:
                    self.size_to_code[key_sap] = _pad_size_code(sap_code)
                if key_desc and key_desc not in self.size_to_code:
                    self.size_to_code[key_desc] = _pad_size_code(sap_code)

        log.info("[MDD] Size LOV: %d entries", len(self.size_to_code))

    def _load_country_lov(self, wb):
        if "Country Origin LOV" not in wb.sheetnames:
            log.warning("[MDD] 'Country Origin LOV' sheet not found — country mapping disabled")
            return
        ws   = wb["Country Origin LOV"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[1:]:
            if not row or row[0] is None:
                continue
            code = str(row[0]).strip()
            desc = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ""
            if code:
                if desc:
                    self.country_desc_to_code[desc] = code
                self.country_desc_to_code[code.upper()] = code
        log.info("[MDD] Country LOV: %d entries", len(self.country_desc_to_code))

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        self._load_color_lov(wb)
        self._load_size_lov(wb)
        self._load_country_lov(wb)
        self._load_company_lov(wb)
        self._load_sbu_lov(wb)
        wb.close()

    # ── public resolvers ─────────────────────────────────────────

    def resolve_color_code(self, raw: str) -> str:
        raw = (raw or "").strip().upper()
        if not raw:
            return "000"
        if raw in {c.upper() for c in self.color_code_set}:
            return _pad_color_code(raw)
        if raw in self.color_desc_to_code:
            return _pad_color_code(self.color_desc_to_code[raw])
        for desc, code in self.color_desc_to_code.items():
            if desc.startswith(raw):
                return _pad_color_code(code)
        log.debug("[MDD] Color not found: %r — using padded raw", raw)
        return _pad_color_code(raw)

    def resolve_size_code(self, raw) -> str:
        raw_str = str(raw).strip().upper() if raw is not None else ""
        if not raw_str:
            return "000"
        if raw_str in self.size_to_code:
            return self.size_to_code[raw_str]
        if "." in raw_str:
            trimmed = raw_str.rstrip("0").rstrip(".")
            if trimmed in self.size_to_code:
                return self.size_to_code[trimmed]
        log.debug("[MDD] Size not found: %r — using padded raw", raw_str)
        return _pad_size_code(raw_str)

    def resolve_country_code(self, raw: str) -> str:
        raw = (raw or "").strip().upper()
        if not raw:
            return ""
        if raw in self.country_desc_to_code:
            return self.country_desc_to_code[raw]
        log.debug("[MDD] Country not found: %r — using raw", raw)
        return raw


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — PAD HELPERS
# ══════════════════════════════════════════════════════════════════

def _pad_color_code(code: str) -> str:
    c = re.sub(r"[^A-Z0-9/]", "", code.upper())[:3]
    return c.ljust(3, "0")


def _pad_size_code(code: str) -> str:
    c = re.sub(r"[^A-Z0-9.\-]", "", str(code).upper())[:3]
    return c.rjust(3, "0")


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — VARIANT CODE BUILDER
# ══════════════════════════════════════════════════════════════════

def _build_generic_code(brand_code: str, item_number: str, color_raw: str, is_footwear: bool = False) -> str:
    """
    Generic Code = brand_code(3) + item_number[:9] + raw color code (no padding)
    e.g. "NEW" + "AC0048U" + "BK" → "NEWAC0048UBK"
    For footwear, Generic Code = brand_code + item_number[:9] (no color)
    """
    clean       = re.sub(r"[^A-Z0-9]", "", item_number.upper())[:9]   
    if is_footwear:
        return f"{brand_code}{clean}"
    color_clean = re.sub(r"[^A-Z0-9]", "", color_raw.upper())          
    return f"{brand_code}{clean}{color_clean}"


def build_variant_code(generic_code: str, maa_size_3: str) -> str:
    return f"{generic_code}{maa_size_3}"[:18]


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — EAN SOURCE LOADER
# ══════════════════════════════════════════════════════════════════

class EANSourceLoader:
    EXPECTED_COLS = {
        "ITEM_NUMBER", "SKU", "UPC#", "SKU Size", "SKU Color Code",
    }

    def __init__(self, path: Path):
        self.path = path
        self.df   = pd.DataFrame()
        self._load()

    def _load(self):
        log.info("[EAN] Loading: %s", self.path.name)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        target_sheet = None
        for sname in wb.sheetnames:
            ws   = wb[sname]
            row1 = next(ws.iter_rows(max_row=1, values_only=True), ())
            if any(str(v).strip().upper() == "ITEM_NUMBER" for v in row1 if v):
                target_sheet = sname
                break
        if target_sheet is None:
            target_sheet = wb.sheetnames[0]

        log.info("[EAN] Using sheet: '%s'", target_sheet)
        ws   = wb[target_sheet]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.error("[EAN] Sheet is empty")
            return

        header  = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[0])]
        df      = pd.DataFrame(rows[1:], columns=header)
        df      = df[df["ITEM_NUMBER"].notna() & (df["ITEM_NUMBER"] != "")]
        self.df = df.reset_index(drop=True)
        log.info("[EAN] %d rows loaded", len(self.df))


# ══════════════════════════════════════════════════════════════════
# SECTION 5 — ROW MAPPER
# ══════════════════════════════════════════════════════════════════

def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "N/A") else s


def map_ean_row(row: pd.Series, mdd: MDDLoader, brand_code: str = "NEW") -> dict:
    item_number = _s(row.get("ITEM_NUMBER"))
    sku         = _s(row.get("SKU"))
    color_raw   = _s(row.get("SKU Color Code"))
    size_raw    = row.get("SKU Size")
    upc_raw     = row.get("UPC#")
    uk_size     = _s(row.get("UK Size"))
    country_raw = _s(
        row.get("Country") or row.get("Country of Origin") or row.get("COO") or ""
    )
    is_footwear = "Width" in row

    maa_color_3  = mdd.resolve_color_code(color_raw)
    maa_size_3   = mdd.resolve_size_code(size_raw)
    generic_code = _build_generic_code(brand_code, item_number, color_raw, is_footwear)
    variant_code = build_variant_code(generic_code, maa_size_3)
    country_code = mdd.resolve_country_code(country_raw)

    upc_str = ""
    if upc_raw is not None:
        try:
            upc_str = str(int(float(str(upc_raw)))).strip()
        except (ValueError, TypeError):
            upc_str = str(upc_raw).strip()

    return {
        "item_number":    item_number,
        "sku":            sku,
        "color_raw":      color_raw,
        "size_raw":       str(size_raw).strip() if size_raw is not None else "",
        "maa_color_3":    maa_color_3,
        "maa_size_3":     maa_size_3,
        "generic_code":   generic_code,   # pre-computed — same value used for KEY_InboundArticle
        "variant_code":   variant_code,
        "upc":            upc_str,
        "uk_size":        uk_size,
        "country_code":   country_code,
        "is_footwear":    is_footwear,
    }


# ══════════════════════════════════════════════════════════════════
# SECTION 6 — XML HELPERS
# ══════════════════════════════════════════════════════════════════

def _val(parent, attr_id: str, value: str = "", id_val: str = ""):
    """Add a <Value> element — uses ID= for LOV, text for plain value."""
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


def _add_barcode(v_el: ET.Element, upc: str):
    """
    Append one DC_Barcode DataContainer to a variant element.
    Called both when creating a new variant AND when adding extra
    EANs to an existing variant (same size, different barcode).
    """
    # Find or create DataContainers → MultiDataContainer
    dcs = v_el.find(f"{{{STIBO_NS}}}DataContainers")
    if dcs is None:
        dcs = ET.SubElement(v_el, f"{{{STIBO_NS}}}DataContainers")

    mdc = dcs.find(f"{{{STIBO_NS}}}MultiDataContainer[@Type='DC_Barcode']")
    if mdc is None:
        mdc = ET.SubElement(dcs, f"{{{STIBO_NS}}}MultiDataContainer")
        mdc.set("Type", "DC_Barcode")

    # Check — same EAN already added? skip
    for dc in mdc.findall(f"{{{STIBO_NS}}}DataContainer"):
        existing_vals = dc.find(f"{{{STIBO_NS}}}Values")
        if existing_vals is not None:
            for v in existing_vals.findall(f"{{{STIBO_NS}}}Value[@AttributeID='AT_Barcode']"):
                if v.text == upc:
                    log.debug("[XML] Duplicate EAN skipped: %s", upc)
                    return

    dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    # Type attribute should not be passed here, only Analyzer
    dc.set("Analyzer", "true")
    dc_vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")
    _val(dc_vals, "AT_Barcode",          upc)
    _val(dc_vals, "AT_BarcodeType",      id_val="P")
    _val(dc_vals, "AT_MainEANIndicator", id_val="Y")



# ══════════════════════════════════════════════════════════════════
# SECTION 7 — XML BUILDER
# ══════════════════════════════════════════════════════════════════

def _build_variant_element(mapped: dict) -> ET.Element | None:
    """
    Build one PRD_VariantArticle element.

    
      AT_Variant   — MAA Article Code 18 chars        [text, value=]
      AT_Color     — SAP Color Code   LOV_ColorCode   [ID=color_raw e.g. BK0]
      AT_Size      — SAP Size Code    LOV_SizeCode    [ID=size_raw  e.g. L, XL]
      AT_Country   — Country of Origin LOV            [ID=country_code]
      AT_Barcode   — EAN barcode in DC_Barcode container

    REMOVED (not on PRD_VariantArticle): AT_ColorCode, AT_MAAColor,
      AT_SizeCode, AT_MAASize, AT_UKSize
    """
    item_number  = mapped["item_number"]
    variant_code = mapped["variant_code"]

    if not item_number or not variant_code:
        return None

    # if sizes is empty then don't create a variant at all 
    if not mapped["size_raw"]:
        log.warning("[XML] Skipping variant %s — size_raw is empty", variant_code)
        return None

    v_el = ET.Element(f"{{{STIBO_NS}}}Product")
    v_el.set("UserTypeID", "PRD_VariantArticle")

    # Key
    var_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
    var_kv.set("KeyID", "KEY_InboundVariant")
    var_kv.text = variant_code

    # Name — size is the variant name- not needed in EAN
    # ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = mapped["size_raw"]

    # ── Values ────────────────────────────────────────────────────
    vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")

    # Variant code — MAA 18-char code (plain text)
    # _val(vals, "AT_Variant", variant_code)

    
    # AT_Size 
    if mapped["maa_size_3"]:
        pass # _val(vals, "AT_Size", id_val=mapped["maa_size_3"])

    # Country of Origin — LOV
    if mapped.get("country_code"):
        _val(vals, "AT_Country", id_val=mapped["country_code"])

    # Barcode — DataContainer
    if mapped["upc"]:
        _add_barcode(v_el, mapped["upc"])

    return v_el


def build_generic_xml(
    variants: list[dict],
    brand: str,
    brand_code: str,
    comp_code: str,
    season_id: str,
    by_article_type: str,
    season_code: str,
    sea_year=""
) -> str:
    """
    Build one PRD_GenericArticle XML with ALL its PRD_VariantArticle
    children nested inside.

    Grouping rules (v2.0):
      • Same variant_code (same color+size) → EAN added to existing
        variant's DC_Barcode (update, not duplicate)
      • New variant_code → new Variant element created
    """
    if not variants:
        return ""

    # Use first row for generic-level data
    first        = variants[0]
    item_number  = first["item_number"]
    # build_generic_xml — fallback
    generic_code = first.get("generic_code") or _build_generic_code(
        brand_code,
        first["item_number"],
        first.get("color_raw", ""),
        first.get("is_footwear", False),
    )

    if not item_number or not generic_code:
        return ""

    # ── Generic element ───────────────────────────────────────────
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_X-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic_code   # same value as AT_InboundGenericCode below

    # ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = item_number

    # Classification references
    # cr1 = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr1.set("ClassificationID", f"CLH_{brand_code}Articles")
    # cr1.set("Type", "CPL_Merchandiser")

    # cr2 = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr2.set("ClassificationID", "CLH_F_TempProductHierarchy")
    # cr2.set("Type", "CPL_SAPHierarchy")

    # cr3 = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr3.set("ClassificationID", f"{season_id}UA")
    # cr3.set("Type", "CPL_UnConfirmedForSeason")

    # Generic-level values
    vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    # _val(vals, "AT_PrincipalStyleCode",  item_number)
    # _val(vals, "AT_InboundGenericCode",  generic_code)

    if by_article_type:
        pass # _val(vals, "AT_BYArticleType", by_article_type, id_val=by_article_type)

    # Replace with:
    LOV_SEASON_LABEL = {
        "SS": "Spring-Summer", "FW": "Fall-Winter",
        "AL": "All Season",    "HO": "Holiday",   "AW": "Autumn-Winter",
    }
    if season_code:
        sea_label = LOV_SEASON_LABEL.get(season_code, season_code)
        # _val(vals, "AT_Season", sea_label, id_val=season_code)
    if sea_year:
        pass # _val(vals, "AT_SeasonYear", sea_year)

    # AT_Color — Color on Generic (Parent), from first variant row
    
    if first.get("color_raw"):
        pass # _val(vals, "AT_Color", id_val=first["color_raw"])

    # Country on Generic from first variant (all variants same item)
    if first.get("country_code"):
        _val(vals, "AT_Country", id_val=first["country_code"])

    # ── Variants — grouped by variant_code ───────────────────────
    # key = variant_code, value = ET.Element (already appended to g_el)
    seen: dict[str, ET.Element] = {}

    for mapped in variants:
        vc = mapped["variant_code"]
        if not vc:
            continue

        if vc in seen:
            # Same size+color already exists — just add barcode if new EAN
            if mapped["upc"]:
                _add_barcode(seen[vc], mapped["upc"])
                log.debug("[XML] EAN updated on existing variant %s → %s", vc, mapped["upc"])
        else:
            # New size/color → create variant element
            v_el = _build_variant_element(mapped)
            if v_el is not None:
                seen[vc] = v_el
                g_el.append(v_el)
                log.debug("[XML] New variant created: %s", vc)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════
# SECTION 8 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None):

    def first(d: Path, ext="*.xlsx"):
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f     = first(MDD_DIR)
    ean_files = list(EAN_DIR.glob("*.xlsx"))

    for label, val in [("MDD", mdd_f), ("EAN Source", ean_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise RuntimeError(f"Required input not found: {label}")

    mdd = MDDLoader(mdd_f)

    
    sea_prefix = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year   = (
        f"20{args.season[2:]}"
        if len(args.season) > 2 and len(args.season[2:]) == 2
        else args.season[2:]
    )
    season_id = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    
    ean_filename = ean_files[0].stem.lower() if ean_files else ""
    by_article_type = "Licensed" if "licensed" in ean_filename else "Inline"

    # AT_Season — SS or FW
    season_code = sea_prefix  # "SS" or "FW"

    log.info(
        "BY Article Type: %s  |  Season Code: %s  |  Season ID: %s",
        by_article_type, season_code, season_id,
    )

    for ean_path in ean_files:
        log.info("─── Processing EAN Source: %s ───", ean_path.name)

        loader = EANSourceLoader(ean_path)
        if loader.df.empty:
            log.warning("[EAN] Empty dataframe — skipping.")
            continue

        df = loader.df

        # TestMode
        # df = df.head(10)

        total_rows = len(df)
        log.info("[EAN] %d rows to process", total_rows)

        # ── Step 1: All Rows  ──────────────────────────
        all_mapped: list[dict] = []
        map_errors = 0
        for _, row in df.iterrows():
            try:
                mapped = map_ean_row(row, mdd, args.brand_code) 
                if mapped["item_number"] and mapped["upc"]:
                    all_mapped.append(mapped)
                else:
                    log.warning(         
                        "[SKIP] item=%r  upc=%r  size=%r",
                        mapped["item_number"], mapped["upc"], mapped["size_raw"]
                    )
            except Exception as exc:
                map_errors += 1
                log.warning("[EAN] Map error: %s", exc)

        # ── Step 2: item_number  ────────────────────
        
        groups: dict[str, list[dict]] = defaultdict(list)
        item_order: list[str] = []
        for m in all_mapped:
            key = m["generic_code"]  
            if key not in groups:
                item_order.append(key)
            groups[key].append(m)
            
        # TestMode: Only 2 generic and 2 variant
        # item_order = item_order[:2]
        # for k in item_order:
        #     groups[k] = groups[k][:2]

        log.info(
            "[EAN] %d mapped rows → %d unique generics",
            len(all_mapped), len(item_order),
        )

        # ── Step 3: write an xml ───────────────────────────────
        out_name = f"{ean_path.stem}.xml"
        out_path = XML_OUT_DIR / out_name

        export_time   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        written_generics  = 0
        written_variants  = 0
        updated_eans      = 0
        error_count       = map_errors

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
            f.write("  <Products>\n")

            for item_number in item_order:
                variant_list = groups[item_number]
                try:
                    xml_str = build_generic_xml(
                        variant_list,
                        args.brand, args.brand_code,
                        args.comp_code, season_id,
                        by_article_type, season_code, sea_year,   # ← pass sea_year
                    )
                    if xml_str:
                        xml_str = _XMLNS_RE.sub("", xml_str)
                        f.write(f"    {xml_str}\n")
                        written_generics += 1

                        # Count unique variants vs EAN updates
                        seen_vc: set[str] = set()
                        for m in variant_list:
                            vc = m["variant_code"]
                            if vc in seen_vc:
                                updated_eans += 1
                            else:
                                seen_vc.add(vc)
                                written_variants += 1

                except Exception as exc:
                    error_count += 1
                    log.warning("[XML] Generic error (%s): %s", item_number, exc)

            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ EAN XML written → %s  (%dKB)", out_path, file_kb)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

        print("═══ EAN SUMMARY ════════════════════════════════════", flush=True)
        print(f"  EAN Source rows     : {total_rows}",          flush=True)
        print(f"  Generics written    : {written_generics}",    flush=True)
        print(f"  Variants created    : {written_variants}",    flush=True)
        print(f"  EANs updated        : {updated_eans}",        flush=True)
        print(f"  Rows skipped/error  : {error_count}",         flush=True)
        print(f"  XML file size       : {file_kb}KB",           flush=True)
        print("════════════════════════════════════════════════════", flush=True)


# ══════════════════════════════════════════════════════════════════
# SECTION 9 — CLI
# ══════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — New Balance EAN Source v2.0"
    )
    p.add_argument("--brand",      default="New Balance")
    p.add_argument("--brand-code", default="NEW")
    p.add_argument("--comp-code",  default="0000")
    p.add_argument("--sbu",        default="FW")
    p.add_argument("--season",     default="SS26")
    p.add_argument("--seq",        default=1, type=int)
    run(p.parse_args())


if __name__ == "__main__":
    main()