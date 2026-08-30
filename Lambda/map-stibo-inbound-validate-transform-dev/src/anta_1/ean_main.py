"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Anta EAN Source v1.0             ║
║   EAN Source File → Stibo STEP XML (Generic + Variants)         ║
╚══════════════════════════════════════════════════════════════════╝

Input file pattern:
    0888-SP-ANTA-ANTA_EAN Source (MAA Sport) Inline-Multi-SP2027-ID-1

Sheet: "EAN"

Column layout (Anta EAN Source):
    Item Code  -> Principal Style Code (used in KEY_InboundArticle).
                  The sheet has TWO columns headed "Item Code" — column A
                  (leftmost) is the one used; the second is ignored.
    Sizes      -> Size value (looked up in MDD "Size Code LOV" for KEY_InboundVariant)
    EAN        -> Barcode

Article structure:
  One Generic per unique Item Code
  One Variant per row (size LOV ID from MDD lookup)
  Barcode → Variant DC_Barcode DataContainer

KEY_InboundArticle formula:
    BrandCode + Item Code

KEY_InboundVariant formula:
    KEY_InboundArticle + Size LOV ID (from MDD "Size Code LOV" Col G description
    -> Col F LOV ID lookup, same mechanism as crocs/shipment_confirmation_main.py;
    falls back to a formula when the size isn't found in the LOV).
"""

import re
import sys
import logging
import argparse
import os
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

log = logging.getLogger(__name__)

# ======================================================================
# PATHS
# ======================================================================
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", Path(__file__).resolve().parent))

EAN_DIR     = BASE_DIR / "input" / "ean"
MDD_DIR     = BASE_DIR / "input" / "mdd"
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (EAN_DIR, MDD_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Anta_1"
BRAND_CODE = "ATA"

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


# ======================================================================
# SIZE CODE RESOLUTION  (ported from crocs/shipment_confirmation_main.py)
# ======================================================================

def _maa_size_code_fallback(size_val: str) -> str:
    """Fallback formula when size is not found in MDD Size Code LOV.

    Rules (from brand spec):
      value → lov_id
      10    → 010   (integer numeric: zero-pad to 3 digits)
      7     → 007
      6.5   → 06H   (half-size: int part zero-padded to 2 digits + 'H')
      M     → 00M   (1-char alpha: prefix '00')
      XL    → 0XL   (2-char alpha: prefix '0')
      ABC   → ABC   (3+ char: take first 3 as-is)
    """
    s = size_val.strip().upper()
    if not s:
        return "000"

    # Pure integer numeric  → zero-pad to 3  (10→010, 7→007)
    if re.match(r'^\d+$', s):
        try:
            return str(int(s)).zfill(3)[:3]
        except ValueError:
            pass

    # Half-size numeric  → 2-digit int part + 'H'  (6.5→06H, 10.5→10H)
    m = re.match(r'^(\d+)\.5$', s)
    if m:
        return str(int(m.group(1))).zfill(2)[:2] + "H"

    # Other decimal  → zero-pad int part to 3 digits  (e.g. 6.0→060)
    m2 = re.match(r'^(\d+)\.\d+$', s)
    if m2:
        return str(int(m2.group(1))).zfill(3)[:3]

    # Alpha / alphanumeric  → left-pad with '0' to reach 3 chars
    alnum = re.sub(r"[^A-Z0-9/]", "", s)
    if not alnum:
        return "XXX"

    alnum_no_slash = alnum.replace("/", "")

    if len(alnum_no_slash) == 1:   return "00" + alnum_no_slash       # M   → 00M
    if len(alnum_no_slash) == 2:   return "0"  + alnum_no_slash       # XL  → 0XL
    return alnum_no_slash[:3]                                          # ABC → ABC


def _maa_size_code(size_val: str, size_lov: dict | None = None) -> str:
    """Return a 3-character MAA size code for use in KEY_Variant.

    Priority:
      1. Dynamic lookup in MDD 'Size Code LOV'  (Col G description → Col F LOV ID).
      2. Fallback formula (zero-pad numerics, 'H' for half-sizes, pad alpha).
    """
    s = size_val.strip().upper()

    # ── 1. MDD Size Code LOV lookup (Col G → Col F) ───────────────
    if size_lov:
        if s in size_lov:
            return size_lov[s]
        stripped = s.lstrip("0") or "0"
        if stripped in size_lov:
            return size_lov[stripped]
        # Try without slash
        s_no_slash = s.replace("/", "")
        if s_no_slash in size_lov:
            return size_lov[s_no_slash]

    # ── 2. Fallback formula ───────────────────────────────────────
    return _maa_size_code_fallback(s)


# ======================================================================
# MDD LOADER  (Size Code LOV only)
# ======================================================================

class MDDLoader:
    """Loads the Size Code LOV from the MDD Excel."""

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        self._load_size_lov(wb)
        wb.close()
        log.info("[MDD] %d LOV type(s) loaded", len(self.lovs))

    def _load_size_lov(self, wb):
        """Load Size Code LOV: Col G = description, Col F = Stibo LOV ID.
        Stored as lovs['Size Code'] = { description: lov_id }."""
        sheet = next((s for s in wb.sheetnames if s.strip().upper() == "SIZE CODE LOV"), None)
        if not sheet:
            log.warning("[MDD] 'Size Code LOV' sheet not found — %s", wb.sheetnames)
            return
        rows = list(wb[sheet].iter_rows(values_only=True))
        size_map: dict = {}
        for row in rows[1:]:
            if not row or len(row) < 7:
                continue
            lov_id = row[5]   # Col F
            desc   = row[6]   # Col G
            if lov_id is None or desc is None:
                continue
            lid = str(lov_id).strip()
            key = str(desc).strip().upper()
            if key.endswith('.'): key = key[:-1]   # strip trailing dot
            if key and lid:
                size_map.setdefault(key, lid)       # first occurrence wins
        if size_map:
            self.lovs["Size Code"] = size_map
            log.info("[MDD] Size Code LOV (colF/G): %d entries", len(size_map))


# ======================================================================
# LOADER
# ======================================================================

class AntaEANLoader:
    """Loads Anta EAN source data from the 'EAN' sheet.

    Header auto-detection: scans the first 15 rows for signals like
    'item code', 'sizes', 'ean'.
    """

    HEADER_SIGNALS = {"item code", "sizes", "ean", "size"}
    TARGET_SHEET   = "EAN"

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheet = ""
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Anta-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[Anta-EAN] Available sheets: %s", wb.sheetnames)

        # Find TARGET_SHEET (case-insensitive)
        target = None
        for sn in wb.sheetnames:
            if sn.strip().lower() == self.TARGET_SHEET.lower():
                target = sn
                break
        if not target:
            wb.close()
            raise ValueError(
                f"Sheet '{self.TARGET_SHEET}' not found in {self.path.name}. "
                f"Available: {wb.sheetnames}"
            )
        log.info("[Anta-EAN] Using sheet: '%s'", target)
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Find header row
        hdr_idx = 0
        for i in range(min(15, len(rows))):
            if rows[i]:
                rv = {str(v).strip().lower() for v in rows[i] if v is not None}
                if rv & self.HEADER_SIGNALS:
                    hdr_idx = i
                    break

        log.info("[Anta-EAN] Header at row %d (0-indexed)", hdr_idx)
        header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[hdr_idx])]

        # Handle duplicate column names — first occurrence (leftmost, i.e.
        # column A) keeps the bare name; later duplicates get a "_1", "_2"
        # suffix. The sheet has two "Item Code" columns; this guarantees
        # find_column("Item Code") resolves to column A, as required.
        seen_cols: dict = {}
        for ci, col in enumerate(header):
            if col in seen_cols:
                seen_cols[col] += 1
                header[ci] = f"{col}_{seen_cols[col]}"
            else:
                seen_cols[col] = 0

        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)

        # Filter rows with a valid Item Code
        item_code_col = self._find_column(header, ["Item Code", "ITEM CODE", "item code"])
        if item_code_col:
            df = df[df[item_code_col].notna() & (~df[item_code_col].astype(str).str.strip().isin(["", "None", "nan"]))]

        wb.close()
        self.df = df.reset_index(drop=True)
        self.sheet = target
        self.headers = header
        log.info("[Anta-EAN] %d rows loaded", len(self.df))
        if self.headers:
            log.info("[Anta-EAN] Headers: %s", self.headers[:20])

    @staticmethod
    def _find_column(headers: list, candidates: list) -> str | None:
        for c in candidates:
            if c in headers:
                return c
            for h in headers:
                if h.lower() == c.lower():
                    return h
        return None

    def find_column(self, *candidates) -> str | None:
        return self._find_column(self.headers, list(candidates))


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "NaT") else s


def _fmt_barcode(val) -> str:
    """Format barcode/EAN value."""
    if val is None:
        return ""
    try:
        return str(int(float(str(val).strip())))
    except (ValueError, TypeError):
        s = str(val).strip()
        return "" if s in ("None", "nan") else s


def _clean_item_code(val) -> str:
    """Clean Item Code (style code)."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan") else s


def _clean_size(val) -> str:
    """Clean size value — used as-is in KEY_InboundVariant."""
    if val is None:
        return ""
    s = str(val).strip()
    # Remove trailing .0 from numeric sizes read as float (e.g. "10.0" → "10")
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    return "" if s in ("None", "nan") else s


def _clean_hs_code(val) -> str:
    """Clean H.S.Code value while preserving leading zeros."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan") else s


# ======================================================================
# MAPPER
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str,
                   loader: AntaEANLoader, mdd: "MDDLoader | None" = None) -> dict[str, dict]:
    """
    Groups EAN rows into generics.

    KEY_InboundArticle = BrandCode + Item Code (column A)
    KEY_InboundVariant = KEY_InboundArticle + Size LOV ID (from MDD lookup)

    Each unique Item Code = One Generic
    Each row = One Variant (size LOV ID from 'Sizes', barcode from 'EAN')
    """
    result: dict[str, dict] = {}

    # Locate columns
    item_code_col = loader.find_column("Item Code", "ITEM CODE", "item code")
    size_col      = loader.find_column("Sizes", "SIZES", "sizes", "Size", "SIZE")
    barcode_col   = loader.find_column("EAN", "Ean", "ean")
    hs_code_col   = loader.find_column("H.S.Code", "H.S. Code", "HS Code", "H S Code", "HSCODE")

    if not item_code_col:
        log.error("[Anta-EAN] Cannot find Item Code column")
        return result

    log.info("[Anta-EAN] All available headers: %s", loader.headers)
    log.info("[Anta-EAN] Using columns — Item Code: '%s'  Sizes: '%s'  EAN: '%s'  H.S.Code: '%s'",
             item_code_col, size_col, barcode_col, hs_code_col)

    if not size_col:
        log.warning("[Anta-EAN] Sizes column NOT FOUND — size will default to '000'!")
    if not barcode_col:
        log.warning("[Anta-EAN] EAN column NOT FOUND — barcodes will be empty!")
    if not hs_code_col:
        log.warning("[Anta-EAN] H.S.Code column NOT FOUND — AT_HSCode will be empty on generic!")

    # Get size LOV from MDD
    size_lov = mdd.lovs.get("Size Code") if mdd else None
    if size_lov:
        log.info("[Anta-EAN] Size Code LOV loaded with %d entries", len(size_lov))
    else:
        log.warning("[Anta-EAN] Size Code LOV not found in MDD — using fallback formulas")

    for row in raw_rows:
        item_code_raw = _clean_item_code(row.get(item_code_col))
        if not item_code_raw:
            continue

        size_raw    = _clean_size(row.get(size_col)) if size_col else ""
        sap_size    = _maa_size_code(size_raw, size_lov) if size_raw else "000"
        barcode_str = _fmt_barcode(row.get(barcode_col)) if barcode_col else ""
        hs_code     = _clean_hs_code(row.get(hs_code_col)) if hs_code_col else ""

        # KEY_InboundArticle = BrandCode + Item Code
        generic_code = f"{brand_code}{item_code_raw}"

        if generic_code not in result:
            result[generic_code] = {
                "style_code":   item_code_raw,
                "generic_code": generic_code,
                "brand_code":   brand_code,
                "hs_code":      hs_code,
                "variants":     {},
            }
        elif not result[generic_code].get("hs_code") and hs_code:
            # Keep first non-empty H.S.Code seen for this generic.
            result[generic_code]["hs_code"] = hs_code

        # KEY_InboundVariant = KEY_InboundArticle + Size LOV ID
        variant_code = f"{generic_code}{sap_size}"

        g = result[generic_code]
        # Use size LOV ID as key; keep first occurrence if duplicates exist
        if sap_size not in g["variants"]:
            g["variants"][sap_size] = {
                "size_raw":     size_raw,
                "sap_size":     sap_size,
                "barcode":      barcode_str,
                "variant_code": variant_code,
            }

    log.info("[Mapper] %d rows → %d generics  %d variants",
             len(raw_rows), len(result),
             sum(len(g["variants"]) for g in result.values()))
    return result


# ======================================================================
# XML HELPERS
# ======================================================================

def _add_barcode_datacontainer(parent_el: ET.Element, barcode: str) -> None:
    """Add DC_Barcode DataContainer with barcode attributes."""
    if not barcode:
        return
    dc_root = ET.SubElement(parent_el, f"{{{STIBO_NS}}}DataContainers")
    mdc = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")
    dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc.set("Analyzer", "true")
    dcv = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

    # AT_Barcode — the actual barcode value
    v1 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v1.set("AttributeID", "AT_Barcode")
    v1.text = barcode

    # AT_BarcodeType — P for Principal
    v2 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v2.set("AttributeID", "AT_BarcodeType")
    v2.set("ID", "P")

    # AT_MainEANIndicator — Y for Yes
    v3 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v3.set("AttributeID", "AT_MainEANIndicator")
    v3.set("ID", "Y")


# ======================================================================
# PRODUCT XML BUILDER
# ======================================================================

def build_product_xml(generic: dict) -> str:
    """Build XML for one Generic article with its Variants."""
    if not generic.get("style_code"):
        return ""

    # Generic Article element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")

    # KEY_InboundArticle
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_code"]

    # Values element for generic
    g_values = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # AT_HSCode from input column "H.S.Code"
    if generic.get("hs_code"):
        gv = ET.SubElement(g_values, f"{{{STIBO_NS}}}Value")
        gv.set("AttributeID", "AT_HSCode")
        gv.text = generic["hs_code"]

    # Variant Articles — sorted by size for stable output
    for size_key in sorted(generic["variants"].keys()):
        variant = generic["variants"][size_key]

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")

        # KEY_InboundVariant
        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant["variant_code"]

        # Empty Values element for variant
        ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")

        # Add barcode DataContainer if barcode exists
        if variant.get("barcode"):
            _add_barcode_datacontainer(v_el, variant["barcode"])

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """Main entry point for Anta EAN processing."""
    brand      = getattr(args, "brand",      BRAND_NAME)
    brand_code = getattr(args, "brand_code", BRAND_CODE)
    season     = getattr(args, "season",     "SS27")

    log.info("[Anta-EAN] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    # Load MDD (Size Code LOV)
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None) or next(MDD_DIR.glob("*.xlsm"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Anta-EAN] No MDD file found in %s — using fallback size formulas", MDD_DIR)

    # Load input file
    ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file:
        raise FileNotFoundError(f"No EAN source file found in {EAN_DIR}")

    loader = AntaEANLoader(ean_file)

    if loader.df.empty:
        log.warning("[Anta-EAN] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_ean_rows(raw_rows, brand_code, loader, mdd=mdd)

    if not generics:
        log.warning("[Anta-EAN] No valid generics produced")
        return

    log.info("[Anta-EAN] Total generics to write: %d", len(generics))

    # Generate XML
    xml_filename = f"{ean_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    log.info("[Anta-EAN] Writing XML → %s", xml_filename)
    with open(xml_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            f'<STEP-ProductInformation '
            f'xmlns="{STIBO_NS}" '
            f'xmlns:xsi="{STIBO_XSI}" '
            f'xsi:schemaLocation="{STIBO_SCHEMA}" '
            f'ExportTime="{export_time}" '
            f'ExportContext="Context1" '
            f'WorkspaceID="Main" '
            f'UseContextLocale="false">\n\n'
        )
        f.write("  <Products>\n")

        for generic in generics.values():
            px = build_product_xml(generic)
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])

        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[Anta-EAN] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(
        f"=== ANTA EAN SUMMARY ===\n"
        f"  Input rows : {len(raw_rows)}\n"
        f"  Total generics found : {len(generics)}\n"
        f"  Generics written     : {generic_count}\n"
        f"  Variants written     : {variant_count}\n"
        f"  XML size   : {file_kb} KB",
        flush=True,
    )

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return xml_path


# ======================================================================
# MAIN
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Anta EAN Source → Stibo XML")
    parser.add_argument("--brand",      default=BRAND_NAME)
    parser.add_argument("--brand_code", default=BRAND_CODE)
    parser.add_argument("--season",     default="SS27")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    run(args)
