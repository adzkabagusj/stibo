"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Clarks EAN Source v1.0          ║
║   EAN Source → Stibo STEP XML (Generic + Variants)              ║
╚══════════════════════════════════════════════════════════════════╝

Input file:
    CLARKS_EAN Source.xlsx

Column layout:
    Generic Article - Key         → Principal Style Code (for KEY_InboundArticle)
    Product(SKU) - Key            → Last 3 digits = Size Code (for KEY_InboundVariant)
    Product(SKU) - EAN/UPC (Key)  → Principal Barcode (AT_Barcode / EAN)

Article structure:
  One Generic per unique "Generic Article - Key"
  One Variant per row (size from last 3 digits of Product(SKU) - Key)
  Principal Barcode → Variant DC_Barcode DataContainer

KEY_InboundArticle formula:
    BrandCode + PrincipalStyleCode (from Generic Article - Key)

KEY_InboundVariant formula:
    KEY_InboundArticle + Size Code (last 3 digits from Product(SKU) - Key)

Business rule:
    Rows with NO Generic Article - Key OR NO Product(SKU) - Key are SKIPPED.
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

EAN_DIR     = BASE_DIR / "input" / "ean"          # Lambda downloads EAN source here
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"    # Brand mapping file
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (EAN_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Clarks"
BRAND_CODE = "CLK"          # 3-letter SAP brand code for Clarks

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


# ======================================================================
# MDD LOADER
# ======================================================================

class MDDLoader:
    """Loads Brand LOV from the MDD Excel."""

    def __init__(self, path: Path):
        self.path = path
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        self._load_brand_lov(wb)
        wb.close()
        log.info("[MDD] %d LOVs", len(self.lovs))

    def _load_brand_lov(self, wb):
        """Brand LOV: col A = Value ID of LOV, col B = Values of LOV (display name).
        Lookup key: display name (col B) → LOV ID (col A).
        e.g. row: A='CKS'  B='CLARKS' → stored as {'CLARKS': 'CKS'}
        """
        sheet = next(
            (s for s in wb.sheetnames if "BRAND" in s.upper() and "LOV" in s.upper()
             and "TYPE" not in s.upper() and "GROUP" not in s.upper()),
            None,
        )
        if not sheet:
            log.warning("[MDD] Brand LOV sheet not found")
            return
        log.info("[MDD] Loading Brand LOV from sheet: '%s'", sheet)
        lov: dict[str, str] = {}
        for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
            if not row or len(row) < 2:
                continue
            lov_id  = str(row[0]).strip() if row[0] else ""   # Col A = Value ID of LOV
            display = str(row[1]).strip() if row[1] else ""   # Col B = Values of LOV
            if display and lov_id:
                lov[display.upper()] = lov_id
        self.lovs["BrandLOV"] = lov
        self.lovs["AT_Brand"] = lov
        log.info("[MDD] Brand LOV: %d entries loaded", len(lov))
        log.info("[MDD] Brand LOV sample (first 10): %s",
                 dict(list(lov.items())[:10]))


# ======================================================================
# EAN SOURCE LOADER
# ======================================================================

class ClarksEANLoader:
    """Loads Clarks EAN Source data."""

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheet = ""
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Clarks-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[Clarks-EAN] Available sheets: %s", wb.sheetnames)

        # Use first sheet
        target = wb.sheetnames[0]
        log.info("[Clarks-EAN] Using sheet: '%s'", target)
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Find header row — look for key column names
        HEADER_SIGNALS = {"generic article", "product(sku)", "ean/upc", "key", "article", "sku"}
        hdr_idx = 0
        for i in range(min(20, len(rows))):
            if rows[i]:
                # Check if row has substantive content (not just numbers/dates)
                row_text = " ".join(str(v).strip().lower() for v in rows[i] if v is not None)
                if any(signal in row_text for signal in HEADER_SIGNALS):
                    hdr_idx = i
                    log.info("[Clarks-EAN] Found header signals at row %d: %s", i, rows[i][:5])
                    break

        log.info("[Clarks-EAN] Header at row %d (0-indexed)", hdr_idx)
        header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[hdr_idx])]
        
        # Log first few headers for debugging
        log.info("[Clarks-EAN] First 10 raw headers: %s", header[:10])

        # Handle duplicate column names
        seen_cols: dict = {}
        for ci, col in enumerate(header):
            if col in seen_cols:
                seen_cols[col] += 1
                header[ci] = f"{col}_{seen_cols[col]}"
            else:
                seen_cols[col] = 0

        df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
        wb.close()

        self.df = df.reset_index(drop=True)
        self.sheet = target
        self.headers = header
        log.info("[Clarks-EAN] %d rows loaded", len(self.df))
        if self.headers:
            log.info("[Clarks-EAN] Headers: %s", self.headers[:15])

    def find_column(self, *candidates) -> str | None:
        """Find column name from candidates (case-insensitive)."""
        for c in candidates:
            if c in self.headers:
                return c
            for h in self.headers:
                if h.strip().lower() == c.strip().lower():
                    return h
        return None


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _fmt_barcode(val) -> str:
    """Format barcode/EAN — strip trailing '.0' from float-typed cells."""
    if val is None:
        return ""
    try:
        return str(int(float(str(val).strip())))
    except (ValueError, TypeError):
        s = str(val).strip()
        return "" if s in ("None", "nan") else s


def _extract_size_code(sku_key: str) -> str:
    """Extract last 3 digits/characters from Product(SKU) - Key as size code."""
    s = _s(sku_key)
    if not s:
        return ""
    # Take last 3 characters
    return s[-3:] if len(s) >= 3 else s


# ======================================================================
# MAPPER — groups flat rows into Generic + Variants
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str,
                   loader: ClarksEANLoader) -> dict[str, dict]:
    """
    Groups EAN rows into generics.

    KEY_InboundArticle = BrandCode + PrincipalStyleCode (from Generic Article - Key)
    KEY_InboundVariant = KEY_InboundArticle + Size Code (last 3 from Product(SKU) - Key)

    Business rule: rows with NO Generic Article - Key OR NO Product(SKU) - Key are SKIPPED.
    """
    result: dict[str, dict] = {}

    generic_col = loader.find_column("Generic Article - Key", "Generic Article Key",
                                     "GENERIC ARTICLE - KEY", "Generic Article", "Generic Article-Key",
                                     "GenericArticle-Key", "Article Key", "Article-Key")
    sku_col     = loader.find_column("Product(SKU) - Key", "Product(SKU) Key",
                                     "PRODUCT(SKU) - KEY", "SKU Key", "SKU - Key", "Product(SKU)-Key",
                                     "SKU-Key", "Product Key", "Product-Key")
    barcode_col = loader.find_column("Product(SKU) - EAN/UPC (Key)", "EAN/UPC (Key)",
                                     "PRODUCT(SKU) - EAN/UPC (KEY)", "EAN/UPC", "Barcode",
                                     "Product(SKU)-EAN/UPC(Key)", "EAN", "UPC", "EAN/UPC Key")
    tariff_col  = loader.find_column("Product(SKU) - Tariff Code (Key)", "Tariff Code (Key)",
                                     "PRODUCT(SKU) - TARIFF CODE (KEY)", "Tariff Code", "HS Code",
                                     "Product(SKU)-Tariff Code(Key)", "Tariff", "HSCode")

    if not generic_col:
        log.error("[Clarks-EAN] Cannot find 'Generic Article - Key' column. Available headers: %s",
                  loader.headers)
        return result

    if not sku_col:
        log.error("[Clarks-EAN] Cannot find 'Product(SKU) - Key' column. Available headers: %s",
                  loader.headers)
        return result

    log.info("[Clarks-EAN] Columns: Generic='%s' SKU='%s' Barcode='%s' Tariff='%s'",
             generic_col, sku_col, barcode_col, tariff_col)

    skipped_no_generic = skipped_no_sku = 0

    for row in raw_rows:
        generic_key = _s(row.get(generic_col)) if generic_col else ""
        sku_key     = _s(row.get(sku_col)) if sku_col else ""

        # ── Business rule: skip rows with no Generic Article - Key OR no Product(SKU) - Key ──
        if not generic_key:
            skipped_no_generic += 1
            continue
        if not sku_key:
            skipped_no_sku += 1
            continue

        # Extract size code from last 3 digits of SKU key
        size_code = _extract_size_code(sku_key)
        if not size_code:
            skipped_no_sku += 1
            continue

        barcode     = _fmt_barcode(row.get(barcode_col)) if barcode_col else ""
        tariff_code = _s(row.get(tariff_col)) if tariff_col else ""

        # KEY_InboundArticle = BrandCode + Generic Article - Key
        generic_code = f"{brand_code}{generic_key}"

        if generic_code not in result:
            result[generic_code] = {
                "article_no":   generic_key,
                "style_code":   generic_key,
                "generic_code": generic_code,
                "brand_code":   brand_code,
                "tariff_code":  tariff_code,  # Store at generic level
                "variants":     {},
            }

        # KEY_InboundVariant = generic_code + Size Code
        variant_code = f"{generic_code}{size_code}"
        variant_key = size_code

        g = result[generic_code]
        if variant_key not in g["variants"]:
            g["variants"][variant_key] = {
                "sku_key":      sku_key,
                "size_code":    size_code,
                "barcode":      barcode,
                "variant_code": variant_code,
            }

    log.info("[Clarks-EAN] %d rows -> %d generics  %d variants "
             "(skipped: no-generic=%d, no-sku=%d)",
             len(raw_rows), len(result),
             sum(len(g["variants"]) for g in result.values()),
             skipped_no_generic, skipped_no_sku)
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
    if not generic.get("article_no"):
        return ""

    # Generic Article element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_code"]

    # Values element for generic
    g_vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    
    # AT_HSCode (tariff code) at generic level
    if generic.get("tariff_code"):
        g_hs = ET.SubElement(g_vals, f"{{{STIBO_NS}}}Value")
        g_hs.set("AttributeID", "AT_HSCode")
        g_hs.text = generic["tariff_code"]

    # Variant Articles
    for var_key in sorted(generic["variants"].keys()):
        variant = generic["variants"][var_key]

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant["variant_code"]

        # Empty Values element for variant
        ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")

        # Barcode DataContainer
        if variant.get("barcode"):
            _add_barcode_datacontainer(v_el, variant["barcode"])

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """Main entry point for Clarks EAN Source processing."""
    brand      = getattr(args, "brand",      BRAND_NAME)
    brand_code = getattr(args, "brand_code", BRAND_CODE)

    log.info("[Clarks-EAN] Starting: brand=%s code=%s", brand, brand_code)

    # Load MDD for Brand LOV lookup
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Clarks-EAN] No MDD file found in %s", MDD_DIR)

    # Extract brand from filename and look up Brand LOV ID dynamically
    brand_from_file = brand  # default
    brand_lov_id = brand_code  # fallback

    # Load input EAN source file
    ean_file = getattr(args, "input_file", None)
    if ean_file and Path(ean_file).exists():
        ean_file = Path(ean_file)
    else:
        ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file:
        raise FileNotFoundError(f"No EAN Source file in {EAN_DIR}")

    # Extract brand from filename (e.g., "0888-FF-CLARKS-EAN Source-..." → "CLARKS")
    filename_stem = ean_file.stem
    parts = filename_stem.split("-")
    if len(parts) >= 3:
        brand_token = parts[2].strip().upper()
        if brand_token:
            brand_from_file = brand_token.title()  # e.g. 'CLARKS' → 'Clarks'
            log.info(
                "[Clarks-EAN] Brand extracted from filename: '%s' → '%s'",
                brand_token, brand_from_file,
            )
    
    # Look up Brand LOV ID from MDD
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            brand_key    = brand_from_file.upper()   # e.g. "CLARKS"
            looked_up    = brand_lov.get(brand_key)  # e.g. "CKS"
            brand_lov_id = looked_up if looked_up else brand_code
            log.info(
                "[Clarks-EAN] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_key, brand_lov_id,
                "(FOUND)" if looked_up else f"(NOT FOUND — fallback to '{brand_code}')",
            )
            if not looked_up:
                log.warning(
                    "[Clarks-EAN] '%s' not in Brand LOV. "
                    "Available keys: %s",
                    brand_key, list(brand_lov.keys()),
                )
        else:
            log.warning("[Clarks-EAN] Brand LOV not found in MDD — using fallback brand_code='%s'", brand_code)
    else:
        log.warning("[Clarks-EAN] No MDD loaded — brand_lov_id defaulting to '%s'", brand_code)

    # Use the dynamic brand_lov_id instead of hardcoded brand_code
    brand_code = brand_lov_id

    # Verify brand mapping file exists (informational)
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if mapping_file:
        log.info("[Clarks-EAN] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[Clarks-EAN] Brand mapping file not found in %s", ATTR_DIR)

    loader = ClarksEANLoader(ean_file)
    if loader.df.empty:
        log.warning("[Clarks-EAN] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")

    generics = group_ean_rows(raw_rows, brand_code, loader)

    if not generics:
        log.warning("[Clarks-EAN] No valid generics produced (all rows missing required columns?)")
        return

    log.info("[Clarks-EAN] Total generics to write: %d", len(generics))

    # Generate XML
    xml_filename = f"{ean_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    log.info("[Clarks-EAN] Writing XML -> %s", xml_filename)
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
    log.info("[Clarks-EAN] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(f"=== CLARKS EAN SOURCE SUMMARY ===\n"
          f"  Input rows : {len(raw_rows)}\n"
          f"  Generics   : {generic_count}\n"
          f"  Variants   : {variant_count}\n"
          f"  XML size   : {file_kb} KB", flush=True)

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return xml_path


# ======================================================================
# MAIN
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clarks EAN Source -> Stibo XML")
    parser.add_argument("--brand",      default=BRAND_NAME)
    parser.add_argument("--brand_code", default=BRAND_CODE)
    parser.add_argument("--input_file", default=None)
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    run(args)
