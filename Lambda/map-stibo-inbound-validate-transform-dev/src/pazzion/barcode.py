"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Pazzion Barcode/EAN v1.0        ║
║   Barcode File → Stibo STEP XML (Generic + Variants)            ║
╚══════════════════════════════════════════════════════════════════╝

Input file pattern:
    CI1167 barcode.xlsx (or similar)

Column layout (Pazzion Barcode):
    Prod Code        -> Principal Style Code (used in KEY_InboundArticle)
    Colour           -> Principal Color Description
    Size             -> Principal Size (for Size LOV ID lookup)
    POS Code         -> Barcode (EAN/UPC)

Article structure:
  One Generic per unique (Prod Code + Colour) combination
  One Variant per row (size LOV ID from MDD lookup)
  Principal Barcode → Variant DC_Barcode DataContainer

KEY_InboundArticle formula:
    BrandCode (from filename) + ProdCode + Colour

KEY_InboundVariant formula:
    KEY_InboundArticle + Size LOV ID (from MDD Size Code LOV lookup)

Filtering:
    - Skip rows where Prod Code is empty
    - Skip summary rows (e.g., "BLACK Total", "2402-21 Total")
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
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", "/tmp/stibo_workdir_pazzion"))

BARCODE_DIR = BASE_DIR / "input" / "barcode"     # Barcode Excel files go here
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"  # Brand mapping file
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (BARCODE_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Pazzion"
BRAND_CODE = "PZZ"  # Default, will be extracted from filename if possible

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


# ======================================================================
# SIZE TRANSFORMATION
# ======================================================================

def _pazzion_size_code_fallback(size_val: str) -> str:
    """Fallback formula when size is not found in MDD Size Code LOV.

    Rules:
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
    alnum = re.sub(r"[^A-Z0-9]", "", s)
    if not alnum:
        return "XXX"
    
    if len(alnum) == 1:   return "00" + alnum       # M   → 00M
    if len(alnum) == 2:   return "0"  + alnum       # XL  → 0XL
    return alnum[:3]                                # ABC → ABC


def _pazzion_size_code(size_val: str, size_lov: dict | None = None) -> str:
    """Return a 3-character size code for use in KEY_Variant.

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

    # ── 2. Fallback formula ───────────────────────────────────────
    return _pazzion_size_code_fallback(s)


# ======================================================================
# BRAND CODE EXTRACTION FROM FILENAME
# ======================================================================

def _extract_brand_code_from_filename(filename: str) -> str:
    """Extract brand code from filename.
    
    Pattern examples:
      0888-FF-PAZZION PZZ-Barcode-SP2028-ID-1.xlsx → PZZ
      CI1167 barcode.xlsx → (use default PZZ)
    
    Logic:
      - Look for pattern: {BRAND_NAME} {CODE}- or -{CODE}-
      - If not found, use default BRAND_CODE
    """
    stem = Path(filename).stem.upper()
    
    # Try to find pattern like "PAZZION PZZ-" or "-PZZ-"
    patterns = [
        r'PAZZION\s+([A-Z]{3})',  # "PAZZION PZZ"
        r'-([A-Z]{3})-',          # "-PZZ-"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, stem)
        if match:
            return match.group(1)
    
    # Default
    return BRAND_CODE


# ======================================================================
# LOADERS
# ======================================================================

class MDDLoader:
    """Loads Core Attributes + LOVs from the MDD Excel."""

    def __init__(self, path: Path):
        self.path = path
        self.core: list[dict] = []
        self.lovs: dict[str, dict] = {}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        # Load Core Attributes
        if "Core Attributes" in wb.sheetnames:
            ws   = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
            hdr  = rows[1] if len(rows) > 1 else rows[0]
            col  = {str(h).strip(): i for i, h in enumerate(hdr) if h}
            for row in rows[2:]:
                attr_id = row[7] if len(row) > 7 else None
                if not attr_id:
                    continue
                attr_id = str(attr_id).strip()
                def _v(key, _row=row, _col=col):
                    idx = _col.get(key)
                    if idx is None or idx >= len(_row): return None
                    v = _row[idx]; return str(v).strip() if v else None
                self.core.append({
                    "id": attr_id, 
                    "name": _v("PIM Attribute Name"),
                    "cardinality": _v("Cardinality"), 
                    "lov_name": _v("Name of LOV")
                })
        
        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        self._load_size_lov(wb)   # Col G (description) → Col F (LOV ID)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.core), len(self.lovs))

    def _load_size_lov(self, wb):
        """Load Size Code LOV: Col G = description, Col F = Stibo LOV ID.
        Stored as lovs['Size Code'] = { description: lov_id }."""
        if "Size Code LOV" not in wb.sheetnames:
            return
        rows = list(wb["Size Code LOV"].iter_rows(values_only=True))
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

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames: return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[1:]:
            lov_name, _, val_name, val_id = (row + (None,)*4)[:4]
            if lov_name and val_name:
                self.lovs.setdefault(str(lov_name).strip(), {})[str(val_name).strip()] = (
                    str(val_id).strip() if val_id else str(val_name).strip())

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper(): continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2: continue
            display = sn.replace(" LOV","").replace("LOV ","").strip()
            for row in rows[1:]:
                if not row or len(row) < 2: continue
                code, name = row[0], row[1]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip())


class PazzionBarcodeLoader:
    """Loads Pazzion barcode/EAN source data."""

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Pazzion-Barcode] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        log.info("[Pazzion-Barcode] Available sheets: %s", wb.sheetnames)
        
        # Use first sheet
        sheet_name = wb.sheetnames[0]
        log.info("[Pazzion-Barcode] Using sheet: '%s'", sheet_name)
        
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        
        # Find header row - look for "Prod Code", "Colour", "Size", "POS Code" columns
        HEADER_SIGNALS = {"prod code", "colour", "size", "pos code"}
        hdr_idx = 0
        for i in range(min(15, len(rows))):
            if rows[i]:
                rv = {str(v).strip().lower() for v in rows[i] if v is not None}
                if rv & HEADER_SIGNALS: 
                    hdr_idx = i
                    break
        
        log.info("[Pazzion-Barcode] Header at row %d (0-indexed)", hdr_idx)
        header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[hdr_idx])]
        
        df = pd.DataFrame(rows[hdr_idx+1:], columns=header)

        # ── Forward-fill Prod Code + Colour ──────────────────────
        # The file uses a sparse/grouped layout where Prod Code and Colour
        # are only set on the FIRST row of each group; subsequent size rows
        # have None. We forward-fill both columns so every variant row
        # inherits the correct Prod Code and Colour before filtering.
        prod_code_col = self._find_column(header, ["Prod Code", "PROD CODE", "prod code", "Product Code"])
        colour_col    = self._find_column(header, ["Colour", "COLOUR", "colour", "Color", "COLOR"])

        if prod_code_col:
            # Replace empty-string / "None" / "nan" with NaN so ffill works
            df[prod_code_col] = df[prod_code_col].replace(
                {"": None, "None": None, "nan": None}
            )
            df[prod_code_col] = df[prod_code_col].ffill()

        if colour_col:
            df[colour_col] = df[colour_col].replace(
                {"": None, "None": None, "nan": None}
            )
            df[colour_col] = df[colour_col].ffill()

        # ── Filter rows ───────────────────────────────────────────
        # 1. Must have Prod Code value
        # 2. Skip summary rows (containing "Total" in Prod Code or Colour)
        if prod_code_col:
            df = df[
                df[prod_code_col].notna() &
                (~df[prod_code_col].astype(str).str.strip().isin(["", "None", "nan"])) &
                (~df[prod_code_col].astype(str).str.contains("Total", case=False, na=False))
            ]
        if colour_col:
            df = df[
                ~df[colour_col].astype(str).str.contains("Total", case=False, na=False)
            ]
        
        wb.close()
        self.df = df.reset_index(drop=True)
        self.headers = header
        log.info("[Pazzion-Barcode] %d rows loaded after filtering", len(self.df))
        if self.headers:
            log.info("[Pazzion-Barcode] Headers: %s", self.headers)

    def _find_column(self, headers: list, candidates: list) -> str | None:
        """Find column name from candidates."""
        for c in candidates:
            if c in headers:
                return c
            # Case-insensitive match
            for h in headers:
                if h.lower() == c.lower():
                    return h
        return None

    def find_column(self, *candidates) -> str | None:
        """Find column from candidates in loaded headers."""
        return self._find_column(self.headers, list(candidates))


# ======================================================================
# HELPERS
# ======================================================================

def _s(v) -> str:
    """Clean string value."""
    if v is None: return ""
    s = str(v).strip()
    return "" if s in ("None","nan","0","NaT") else s


def _fmt_barcode(barcode_val) -> str:
    """Format barcode/POS code."""
    if barcode_val is None: return ""
    try: 
        return str(int(float(str(barcode_val).strip())))
    except (ValueError, TypeError): 
        s = str(barcode_val).strip()
        return "" if s in ("None", "nan") else s


def _clean_prod_code(prod_code_val) -> str:
    """Clean prod code - remove spaces, special chars if needed."""
    if prod_code_val is None: return ""
    s = str(prod_code_val).strip()
    if s in ("None", "nan", ""):
        return ""
    return s


# ======================================================================
# MAPPER — groups flat rows into Generic + Variants dict
# ======================================================================

def group_barcode_rows(raw_rows: list[dict], brand_code: str, loader: PazzionBarcodeLoader, 
                       mdd=None) -> dict[str, dict]:
    """
    Groups barcode rows into generics.
    
    KEY_InboundArticle = BrandCode + ProdCode + Colour
    KEY_InboundVariant = KEY_InboundArticle + Size LOV ID (from MDD lookup)
    
    Each unique (ProdCode + Colour) = One Generic
    Each row = One Variant (with size and barcode)
    """
    result: dict[str, dict] = {}
    
    # Find column names
    prod_code_col = loader.find_column("Prod Code", "PROD CODE", "prod code", "Product Code", "Article")
    colour_col = loader.find_column("Colour", "COLOUR", "colour", "Color", "COLOR")
    size_col = loader.find_column("Size", "SIZE", "size")
    pos_code_col = loader.find_column("POS Code", "POS CODE", "pos code", "Barcode", "EAN", "UPC")
    
    if not prod_code_col:
        log.error("[Pazzion-Barcode] Cannot find Prod Code column")
        return result
    
    log.info("[Pazzion-Barcode] All available headers: %s", loader.headers)
    log.info("[Pazzion-Barcode] Using columns: Prod Code='%s', Colour='%s', Size='%s', POS Code='%s'", 
             prod_code_col, colour_col, size_col, pos_code_col)
    
    if not colour_col:
        log.warning("[Pazzion-Barcode] Colour column NOT FOUND - will use empty string!")
    
    if not size_col:
        log.warning("[Pazzion-Barcode] Size column NOT FOUND - sizes will default to '000'!")
    
    if not pos_code_col:
        log.warning("[Pazzion-Barcode] POS Code column NOT FOUND - barcodes will be empty!")
    
    # Get size LOV from MDD
    size_lov = mdd.lovs.get("Size Code") if mdd else None
    if size_lov:
        log.info("[Pazzion-Barcode] Size Code LOV loaded with %d entries", len(size_lov))
    else:
        log.warning("[Pazzion-Barcode] Size Code LOV not found in MDD - using fallback formulas")

    for row in raw_rows:
        # Get prod code
        prod_code_raw = _clean_prod_code(row.get(prod_code_col))
        if not prod_code_raw:
            continue
        
        # Get colour
        colour_raw = _s(row.get(colour_col)) if colour_col else ""
        colour_upper = colour_raw.upper()
        
        # Get size
        size_raw = _s(row.get(size_col)) if size_col else ""
        
        # Get size LOV ID from MDD lookup
        sap_size = _pazzion_size_code(size_raw, size_lov) if size_raw else "000"
        
        # Get POS Code (barcode)
        pos_code_str = _fmt_barcode(row.get(pos_code_col)) if pos_code_col else ""
        
        # Build generic code: BrandCode + ProdCode + Colour
        generic_code = f"{brand_code}{prod_code_raw}{colour_upper}"
        
        if generic_code not in result:
            # Create new generic
            result[generic_code] = {
                "prod_code": prod_code_raw,
                "colour": colour_raw,
                "colour_upper": colour_upper,
                "generic_code": generic_code,
                "brand_code": brand_code,
                "variants": {},
            }
        
        # Build variant code: generic_code + size LOV ID
        variant_code = f"{generic_code}{sap_size}"
        
        # Use size LOV ID as key to avoid duplicates
        variant_key = sap_size
        
        g = result[generic_code]
        if variant_key not in g["variants"]:
            g["variants"][variant_key] = {
                "size_raw": size_raw,
                "sap_size": sap_size,
                "barcode": pos_code_str,
                "variant_code": variant_code,
            }
    
    log.info("[Mapper] %d rows -> %d generics  %d variants",
             len(raw_rows), len(result), sum(len(g["variants"]) for g in result.values()))
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
    
    # AT_Barcode - the actual barcode value
    v1 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v1.set("AttributeID", "AT_Barcode")
    v1.text = barcode
    
    # AT_BarcodeType - P for Principal
    v2 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v2.set("AttributeID", "AT_BarcodeType")
    v2.set("ID", "P")
    
    # AT_MainEANIndicator - Y for Yes
    v3 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v3.set("AttributeID", "AT_MainEANIndicator")
    v3.set("ID", "Y")


# ======================================================================
# PRODUCT XML BUILDER
# ======================================================================

def build_product_xml(generic: dict) -> str:
    """Build XML for one Generic article with its Variants."""
    if not generic.get("prod_code"): 
        return ""

    # Generic Article element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")

    # KEY_InboundArticle
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_code"]

    # Empty Values element for generic
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # Variant Articles
    for var_key in sorted(generic["variants"].keys()):
        variant = generic["variants"][var_key]
        
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
    """Main entry point for Pazzion Barcode/EAN processing."""
    brand        = getattr(args, "brand",        BRAND_NAME)
    default_code = getattr(args, "brand_code",   BRAND_CODE)

    log.info("[Pazzion-Barcode] Starting: brand=%s", brand)

    # Load MDD
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Pazzion-Barcode] No MDD file found in %s", MDD_DIR)

    # Verify brand mapping file exists (not used for EAN processing)
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if mapping_file:
        log.info("[Pazzion-Barcode] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[Pazzion-Barcode] Brand mapping file not found in %s", ATTR_DIR)

    # Load input file
    barcode_file = next(BARCODE_DIR.glob("*.xls*"), None)
    if not barcode_file: 
        raise FileNotFoundError(f"No Barcode file in {BARCODE_DIR}")
    
    # Extract brand code from filename (like in order_form.py)
    brand_code = _extract_brand_code_from_filename(barcode_file.name)
    log.info("[Pazzion-Barcode] Brand code extracted from filename: %s", brand_code)
    
    loader = PazzionBarcodeLoader(barcode_file)
    
    if loader.df.empty:
        log.warning("[Pazzion-Barcode] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_barcode_rows(raw_rows, brand_code, loader, mdd=mdd)
    
    if not generics:
        log.warning("[Pazzion-Barcode] No valid generics produced")
        return

    log.info("[Pazzion-Barcode] Total generics to write: %d", len(generics))

    # Generate XML
    xml_filename = f"{barcode_file.stem}.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    log.info("[Pazzion-Barcode] Writing XML -> %s", xml_filename)
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
    log.info("[Pazzion-Barcode] XML written: %s  (%d KB)", xml_path.name, file_kb)
    
    print("═══ PAZZION BARCODE/EAN SUMMARY ════════════════════════════", flush=True)
    print(f"  Input file      : {barcode_file.name}",           flush=True)
    print(f"  Brand Code      : {brand_code}",                  flush=True)
    print(f"  Generics written: {generic_count}",               flush=True)
    print(f"  Variants written: {variant_count}",               flush=True)
    print(f"  Output XML      : {xml_path.name}",               flush=True)
    print(f"  File size       : {file_kb}KB",                   flush=True)
    print("═══════════════════════════════════════════════════════════", flush=True)

    if auditor:
        try:
            auditor.set_xml_uploads([str(xml_path)])
        except AttributeError:
            pass


if __name__ == "__main__":
    import argparse
    import types
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--brand", default=BRAND_NAME)
    parser.add_argument("--brand-code", dest="brand_code", default=BRAND_CODE)
    
    parsed_args = parser.parse_args()
    args = types.SimpleNamespace(**vars(parsed_args))
    
    run(args)
