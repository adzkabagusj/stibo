"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Crocs Order Form ACCs BAG v1.0       ║
║   Order Form ACCs BAG → Stibo STEP XML (Generic + Variants)          ║
╚══════════════════════════════════════════════════════════════════╝

Input file pattern:
    0888-SP-CROCS-Order Form ACCs BAG (MAA Sport) Inline-Multi-SP2029-KH-1

Column layout (Crocs Order Form FW):
    Style             -> Principal Style Code (used in KEY_InboundArticle)
    Size              -> Size value (e.g., M7W9 → M7/W9 for MDD lookup)
    UPC               -> Barcode (EAN/UPC)

Article structure:
  One Generic per unique Style
  One Variant per row (size LOV ID from MDD lookup)
  Principal Barcode → Variant DC_Barcode DataContainer

KEY_InboundArticle formula:
    BrandCode + PrincipalStyleCode (from Style column)

KEY_InboundVariant formula:
    KEY_InboundArticle + Size LOV ID (from MDD Size Code LOV lookup)

Size transformation:
    M7W9 → M7/W9 → lookup in MDD "Size Code LOV" (Col G description → Col F LOV ID)
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

EAN_DIR     = BASE_DIR / "input" / "order_form_accs_bag"  # Lambda downloads order_form_accs_bag here
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"     # Brand mapping file
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (EAN_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# BRAND CONSTANTS
# ======================================================================
BRAND_NAME = "Crocs"
BRAND_CODE = "CCR"

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

def _transform_size_value(size_val: str) -> str:
    """Transform Size format: M7W9 → M7/W9.
    
    Rules:
      M7W9    → M7/W9
      M10W12  → M10/W12
      10      → 10 (no change for numeric)
      XL      → XL (no change for alpha)
    """
    if not size_val:
        return ""
    
    s = str(size_val).strip().upper()
    
    # Pattern: M followed by digits, then W followed by digits
    # Examples: M7W9, M10W12, M5W7
    pattern = re.match(r'^(M\d+)(W\d+)$', s)
    if pattern:
        return f"{pattern.group(1)}/{pattern.group(2)}"
    
    return s


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
      M7/W9 → M7W   (special: take first 3 chars)
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
    
    # Remove slash for LOV ID (M7/W9 → M7W9 → M7W)
    alnum_no_slash = alnum.replace("/", "")
    
    if len(alnum_no_slash) == 1:   return "00" + alnum_no_slash       # M   → 00M
    if len(alnum_no_slash) == 2:   return "0"  + alnum_no_slash       # XL  → 0XL
    return alnum_no_slash[:3]                                          # ABC → ABC, M7W9 → M7W


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


class CrocsOrderFormACCsBAGLoader:
    """Loads Crocs Order Form ACCs BAG source data from specific sheets."""

    # Only read these sheets from the input Excel
    TARGET_SHEETS = ["ACCs", "BAG", "ACCs BAG"]

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheets_loaded: list[str] = []
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Crocs-OrderFormACCsBAG] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        log.info("[Crocs-OrderFormACCsBAG] Available sheets: %s", wb.sheetnames)
        log.info("[Crocs-OrderFormACCsBAG] Target sheets: %s", self.TARGET_SHEETS)
        
        all_dfs: list[pd.DataFrame] = []
        combined_headers: list[str] = []
        
        for sheet_name in self.TARGET_SHEETS:
            # Try exact match first, then case-insensitive
            target = None
            if sheet_name in wb.sheetnames:
                target = sheet_name
            else:
                for sn in wb.sheetnames:
                    if sn.lower() == sheet_name.lower():
                        target = sn
                        break
            
            if not target:
                log.warning("[Crocs-OrderFormACCsBAG] Sheet '%s' not found, skipping", sheet_name)
                continue
            
            log.info("[Crocs-OrderFormACCsBAG] Processing sheet: '%s'", target)
            ws = wb[target]
            rows = list(ws.iter_rows(values_only=True))
            
            if not rows:
                log.warning("[Crocs-OrderFormACCsBAG] Sheet '%s' is empty", target)
                continue
            
            # Find header row - look for "Style", "Size", "UPC" columns
            HEADER_SIGNALS = {"style", "size", "upc", "barcode"}
            hdr_idx = 0
            for i in range(min(15, len(rows))):
                if rows[i]:
                    rv = {str(v).strip().lower() for v in rows[i] if v is not None}
                    if rv & HEADER_SIGNALS: 
                        hdr_idx = i
                        break
            
            log.info("[Crocs-OrderFormACCsBAG] Sheet '%s': Header at row %d (0-indexed)", target, hdr_idx)
            header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[hdr_idx])]
            
            # Handle duplicate column names
            seen_cols: dict = {}
            for ci, col in enumerate(header):
                if col in seen_cols: 
                    seen_cols[col] += 1
                    header[ci] = f"{col}_{seen_cols[col]}"
                else: 
                    seen_cols[col] = 0
            
            df = pd.DataFrame(rows[hdr_idx+1:], columns=header)
            
            # Filter rows with valid Style
            style_col = self._find_column(header, ["Style", "STYLE", "style"])
            if style_col:
                df = df[df[style_col].notna() & (~df[style_col].astype(str).str.strip().isin(["","None","nan"]))]
            
            # Add source sheet column for tracking
            df["_source_sheet"] = target
            
            all_dfs.append(df)
            self.sheets_loaded.append(target)
            
            # Use first sheet's headers as reference
            if not combined_headers:
                combined_headers = header
            
            log.info("[Crocs-OrderFormACCsBAG] Sheet '%s': %d rows loaded", target, len(df))
        
        wb.close()
        
        # Combine all DataFrames
        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            self.headers = combined_headers
        else:
            self.df = pd.DataFrame()
            self.headers = []
        
        log.info("[Crocs-OrderFormACCsBAG] Total: %d rows from %d sheets (%s)", 
                 len(self.df), len(self.sheets_loaded), ", ".join(self.sheets_loaded) if self.sheets_loaded else "none")
        if self.headers:
            log.info("[Crocs-OrderFormACCsBAG] Headers: %s", self.headers[:15])
        else:
            log.warning("[Crocs-OrderFormACCsBAG] No headers loaded - no matching sheets found")

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


def _fmt_upc(upc_val) -> str:
    """Format UPC/EAN code."""
    if upc_val is None: return ""
    try: 
        return str(int(float(str(upc_val).strip())))
    except (ValueError, TypeError): 
        s = str(upc_val).strip()
        return "" if s in ("None", "nan") else s


def _clean_style(style_val) -> str:
    """Clean style code - remove spaces, special chars if needed."""
    if style_val is None: return ""
    s = str(style_val).strip()
    if s in ("None", "nan", ""):
        return ""
    return s


# ======================================================================
# MAPPER — groups flat rows into Generic + Variants dict
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str, season: str, 
                   country_code: str, loader: CrocsOrderFormACCsBAGLoader, 
                   mdd=None) -> dict[str, dict]:
    """
    Groups EAN rows into generics.
    
    KEY_InboundArticle = BrandCode + Style (from Style column)
    KEY_InboundVariant = KEY_InboundArticle + Size LOV ID (from MDD lookup)
    
    Each unique Style = One Generic
    Each row = One Variant (with size from Size column and barcode from UPC)
    """
    result: dict[str, dict] = {}
    
    # Find column names
    style_col = loader.find_column("Style", "STYLE", "style", "SKU")
    size_col = loader.find_column("Size", "SIZE", "size", "Grid Value")
    # Note: Excel may have duplicate "UPC" columns - first one has style code, second one (renamed to UPC_1) has actual barcode
    upc_col = loader.find_column("UPC_1", "UPC_2", "UPC", "upc", "UPC CODE", "Barcode", "EAN", "EAN CODE", "EAN Code", "EAN/UPC", "GTIN")
    
    if not style_col:
        log.error("[Crocs-OrderFormACCsBAG] Cannot find Style column")
        return result
    
    log.info("[Crocs-OrderFormACCsBAG] All available headers: %s", loader.headers)
    log.info("[Crocs-OrderFormACCsBAG] Using columns: Style='%s', Size='%s', UPC='%s'", 
             style_col, size_col, upc_col)
    
    if not size_col:
        log.warning("[Crocs-OrderFormACCsBAG] Size column NOT FOUND - sizes will default to '000'!")
    
    if not upc_col:
        log.warning("[Crocs-OrderFormACCsBAG] UPC column NOT FOUND - barcodes will be empty!")
    
    # Get size LOV from MDD
    size_lov = mdd.lovs.get("Size Code") if mdd else None
    if size_lov:
        log.info("[Crocs-OrderFormACCsBAG] Size Code LOV loaded with %d entries", len(size_lov))
    else:
        log.warning("[Crocs-OrderFormACCsBAG] Size Code LOV not found in MDD - using fallback formulas")

    for row in raw_rows:
        # Get style code
        style_raw = _clean_style(row.get(style_col))
        if not style_raw:
            continue
        
        # Get size and transform it
        size_raw = _s(row.get(size_col)) if size_col else ""
        size_transformed = _transform_size_value(size_raw)
        
        # Get size LOV ID from MDD lookup
        sap_size = _maa_size_code(size_transformed, size_lov) if size_transformed else "000"
        
        # Get UPC/barcode
        upc_str = _fmt_upc(row.get(upc_col)) if upc_col else ""
        
        # Build generic code: BrandCode + Style
        generic_code = f"{brand_code}{style_raw}"
        
        if generic_code not in result:
            # Create new generic
            result[generic_code] = {
                "article_no": style_raw,
                "style_code": style_raw,
                "generic_code": generic_code,
                "brand_code": brand_code,
                "division": "F",  # Footwear default for Crocs
                "source_sheet": _s(row.get("_source_sheet")),
                "variants": {},
            }
        
        # Build variant code: generic_code + size LOV ID
        variant_code = f"{generic_code}{sap_size}"
        
        # Use size as key to group same sizes together (avoid duplicates)
        # If same size appears multiple times with different UPCs, keep the first one
        variant_key = sap_size
        
        g = result[generic_code]
        if variant_key not in g["variants"]:
            g["variants"][variant_key] = {
                "size_raw": size_raw,
                "size_transformed": size_transformed,
                "sap_size": sap_size,
                "upc": upc_str,
                "variant_code": variant_code,
                "variant_desc": f"Size {size_transformed}" if size_transformed else f"Size {sap_size}",
            }
    
    # Build size ranges
    for g in result.values():
        g["size_range"] = ", ".join([v.get("size_raw", "") for v in g["variants"].values() if v.get("size_raw")])
    
    log.info("[Mapper] %d rows -> %d generics  %d variants",
             len(raw_rows), len(result), sum(len(g["variants"]) for g in result.values()))
    return result


# ======================================================================
# XML HELPERS
# ======================================================================

def _add_barcode_datacontainer(parent_el: ET.Element, upc: str) -> None:
    """Add DC_Barcode DataContainer with barcode attributes."""
    if not upc: 
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
    v1.text = upc
    
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

def build_product_xml(generic: dict, brand: str, brand_code: str, comp_code: str,
                      sbu: str, season_id: str, mdd=None) -> str:
    """Build XML for one Generic article with its Variants."""
    if not generic.get("article_no"): 
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

    # Variant Articles (all variants)
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

        # Add barcode DataContainer if UPC exists
        if variant.get("upc"):
            _add_barcode_datacontainer(v_el, variant["upc"])

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """Main entry point for Crocs Order Form ACCs BAG EAN processing."""
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SP29")
    country_code = getattr(args, "country_code", "KH")

    log.info("[Crocs-OrderFormACCsBAG] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    # Load MDD
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Crocs-OrderFormACCsBAG] No MDD file found in %s", MDD_DIR)

    # Verify brand mapping file exists (not used for EAN processing)
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if not mapping_file:
        for pattern in ["*brand*mapping*.xlsx", "*mapping*template*.xlsx", "*mapping*.xlsx"]:
            candidates = list(ATTR_DIR.glob(pattern))
            if candidates:
                mapping_file = candidates[0]
                break
    if mapping_file:
        log.info("[Crocs-OrderFormACCsBAG] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[Crocs-OrderFormACCsBAG] Brand mapping file not found in %s", ATTR_DIR)

    # Load input file
    ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file: 
        raise FileNotFoundError(f"No Order Form ACCs BAG file in {EAN_DIR}")
    
    loader = CrocsOrderFormACCsBAGLoader(ean_file)
    
    if loader.df.empty:
        log.warning("[Crocs-OrderFormACCsBAG] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")
    generics = group_ean_rows(raw_rows, brand_code, season, country_code, loader, mdd=mdd)
    
    if not generics:
        log.warning("[Crocs-OrderFormACCsBAG] No valid generics produced")
        return

    # Build season classification ID
    sp = season[:2].upper() if len(season) >= 2 else season
    sys_short = season[2:] if len(season) > 2 else ""
    sy = f"20{sys_short}" if len(sys_short) == 2 else sys_short
    season_id = f"CLH_{brand_code}_{sp}{sy}"

    # Generate XML
    xml_filename = f"{ean_file.stem}_CROCS_ORDERFORM_ACCS_BAG_EAN.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    log.info("[Crocs-OrderFormACCsBAG] Writing XML -> %s", xml_filename)
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
        
        log.info("[Crocs-OrderFormACCsBAG] Writing %d generics to XML", len(generics))

        # Use all generic products (no limit)
        limited_generics = list(generics.values())
        log.info("[Crocs-OrderFormACCsBAG] Processing all products: %d generics",
                 len(limited_generics))

        for generic in limited_generics:
            px = build_product_xml(generic, brand, brand_code, comp_code, sbu, season_id, mdd=mdd)
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])
        
        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[Crocs-OrderFormACCsBAG] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(f"=== CROCS ORDER FORM FW EAN SUMMARY ===\n"
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
    parser = argparse.ArgumentParser(description="Crocs Order Form ACCs BAG EAN -> Stibo XML")
    parser.add_argument("--brand",        default=BRAND_NAME)
    parser.add_argument("--brand_code",   default=BRAND_CODE)
    parser.add_argument("--comp_code",    default="0888")
    parser.add_argument("--sbu",          default="SP")
    parser.add_argument("--season",       default="SP29")
    parser.add_argument("--country_code", default="KH")
    args = parser.parse_args()
    
    logging.basicConfig(
        level=logging.INFO, 
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )
    run(args)
