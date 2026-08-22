"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Crocs Order Form JBZ EAN v1.0  ║
║   Order Form JBZ → Stibo STEP XML (Generic + Variants)         ║
╚══════════════════════════════════════════════════════════════════╝

Column layout (Crocs Order Form JBZ):
    Style             -> Principal Style Code (used in KEY_InboundArticle)
    UPC               -> Barcode (EAN/UPC)
    (other columns as needed)

Article structure:
  One Single Article per row (no variants)
  Principal Barcode → DC_Barcode DataContainer directly on article

KEY_InboundArticle formula:
    BrandCode + PrincipalStyleCode (from Style column)
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

EAN_DIR     = BASE_DIR / "input" / "order_form_jbz"  # Lambda downloads order_form_jbz here
MDD_DIR     = BASE_DIR / "input" / "mdd"
ATTR_DIR    = BASE_DIR / "input" / "attributes"      # Brand mapping file
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
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.core), len(self.lovs))

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


class CrocsOrderFormJBZLoader:
    """Loads Crocs Order Form JBZ source data from specific sheets."""

    # Only read these sheets from the input Excel
    TARGET_SHEETS = ["SP WHS", "KS"]

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheets_loaded: list[str] = []
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[Crocs-JBZ-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        
        log.info("[Crocs-JBZ-EAN] Available sheets: %s", wb.sheetnames)
        log.info("[Crocs-JBZ-EAN] Target sheets: %s", self.TARGET_SHEETS)
        
        all_dfs: list[pd.DataFrame] = []
        combined_headers: list[str] = []
        
        for sheet_name in self.TARGET_SHEETS:
            if sheet_name not in wb.sheetnames:
                log.warning("[Crocs-JBZ-EAN] Sheet '%s' not found, skipping", sheet_name)
                continue
            
            log.info("[Crocs-JBZ-EAN] Processing sheet: '%s'", sheet_name)
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            
            if not rows:
                log.warning("[Crocs-JBZ-EAN] Sheet '%s' is empty", sheet_name)
                continue
            
            # Find header row - look for "Style" or "UPC" columns
            HEADER_SIGNALS = {"style", "upc", "sku", "product"}
            hdr_idx = 0
            for i in range(min(15, len(rows))):
                if rows[i]:
                    rv = {str(v).strip().lower() for v in rows[i] if v is not None}
                    if rv & HEADER_SIGNALS: 
                        hdr_idx = i
                        break
            
            log.info("[Crocs-JBZ-EAN] Sheet '%s': Header at row %d (0-indexed)", sheet_name, hdr_idx)
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
            df["_source_sheet"] = sheet_name
            
            all_dfs.append(df)
            self.sheets_loaded.append(sheet_name)
            
            # Use first sheet's headers as reference
            if not combined_headers:
                combined_headers = header
            
            log.info("[Crocs-JBZ-EAN] Sheet '%s': %d rows loaded", sheet_name, len(df))
        
        wb.close()
        
        # Combine all DataFrames
        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            self.headers = combined_headers
        else:
            self.df = pd.DataFrame()
            self.headers = []
        
        log.info("[Crocs-JBZ-EAN] Total: %d rows from %d sheets (%s)", 
                 len(self.df), len(self.sheets_loaded), ", ".join(self.sheets_loaded) if self.sheets_loaded else "none")
        if self.headers:
            log.info("[Crocs-JBZ-EAN] Headers: %s", self.headers[:15])
        else:
            log.warning("[Crocs-JBZ-EAN] No headers loaded - no matching sheets found")

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
# MAPPER — groups flat rows into Single Articles
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str, season: str, 
                   country_code: str, loader: CrocsOrderFormJBZLoader, 
                   mdd=None) -> list[dict]:
    """
    Maps EAN rows to single articles (no variants).
    
    KEY_InboundArticle = BrandCode + Style + UPC (to ensure unique article per barcode)
    
    Each row = One Single Article with barcode DataContainer
    """
    result: list[dict] = []
    
    # Find column names
    style_col = loader.find_column("Style", "STYLE", "style", "SKU", "Product Code")
    # Note: Excel has duplicate "UPC" columns - first one has style code, second one (renamed to UPC_1) has actual barcode
    upc_col = loader.find_column("UPC_1", "UPC_2", "UPC", "upc", "UPC CODE", "Barcode", "EAN", "EAN CODE", "EAN Code", "EAN/UPC", "GTIN")
    
    if not style_col:
        log.error("[Crocs-JBZ-EAN] Cannot find Style column")
        return result
    
    log.info("[Crocs-JBZ-EAN] All available headers: %s", loader.headers)
    log.info("[Crocs-JBZ-EAN] Using columns: Style='%s', UPC='%s'", style_col, upc_col)
    
    if not upc_col:
        log.warning("[Crocs-JBZ-EAN] UPC column NOT FOUND - barcodes will be empty!")

    for row in raw_rows:
        # Get style code
        style_raw = _clean_style(row.get(style_col))
        if not style_raw:
            continue
        
        # Get UPC/barcode
        upc_str = _fmt_upc(row.get(upc_col)) if upc_col else ""
        
        # Build article code: BrandCode + Style
        article_code = f"{brand_code}{style_raw}"
        
        # Get source sheet for tracking
        source_sheet = row.get("_source_sheet", "Unknown")
        
        # Create single article
        article = {
            "article_no": style_raw,
            "style_code": style_raw,
            "article_code": article_code,
            "brand_code": brand_code,
            "division": "F",  # Footwear default for Crocs
            "upc": upc_str,
            "source_sheet": source_sheet,
        }
        
        result.append(article)
    
    log.info("[Mapper] %d rows -> %d single articles", len(raw_rows), len(result))
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

def build_product_xml(article: dict, brand: str, brand_code: str, comp_code: str,
                      sbu: str, season_id: str, mdd=None) -> str:
    """Build XML for one Single Article (no variants)."""
    if not article.get("article_no"): 
        return ""

    # Single Article element
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_SingleArticle")

    # KEY_InboundArticle
    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = article["article_code"]

    # Empty Values element
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    # Add barcode DataContainer directly to article if UPC exists
    if article.get("upc"):
        _add_barcode_datacontainer(g_el, article["upc"])

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    """Main entry point for Crocs Order Form JBZ EAN processing."""
    brand        = getattr(args, "brand",        BRAND_NAME)
    brand_code   = getattr(args, "brand_code",   BRAND_CODE)
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SP27")
    country_code = getattr(args, "country_code", "ID")

    log.info("[Crocs-JBZ-EAN] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    # Load MDD (optional)
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[Crocs-JBZ-EAN] No MDD file found in %s", MDD_DIR)

    # Verify brand mapping file exists (not used for EAN processing)
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if not mapping_file:
        for pattern in ["*brand*mapping*.xlsx", "*mapping*template*.xlsx", "*mapping*.xlsx"]:
            candidates = list(ATTR_DIR.glob(pattern))
            if candidates:
                mapping_file = candidates[0]
                break
    if mapping_file:
        log.info("[Crocs-JBZ-EAN] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[Crocs-JBZ-EAN] Brand mapping file not found in %s", ATTR_DIR)

    # Load input file
    ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file: 
        raise FileNotFoundError(f"No Order Form JBZ file in {EAN_DIR}")
    
    loader = CrocsOrderFormJBZLoader(ean_file)
    
    if loader.df.empty:
        log.warning("[Crocs-JBZ-EAN] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")
    articles = group_ean_rows(raw_rows, brand_code, season, country_code, loader, mdd=mdd)
    
    if not articles:
        log.warning("[Crocs-JBZ-EAN] No valid articles produced")
        return

    log.info("[Crocs-JBZ-EAN] Total articles to write: %d", len(articles))

    # Build season classification ID
    sp = season[:2].upper() if len(season) >= 2 else season
    sys_short = season[2:] if len(season) > 2 else ""
    sy = f"20{sys_short}" if len(sys_short) == 2 else sys_short
    season_id = f"CLH_{brand_code}_{sp}{sy}"

    # Generate XML
    xml_filename = f"{ean_file.stem}_CROCS_JBZ_EAN.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    article_count = 0

    log.info("[Crocs-JBZ-EAN] Writing XML -> %s", xml_filename)
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
        
        for article in articles:
            px = build_product_xml(article, brand, brand_code, comp_code, sbu, season_id, mdd=mdd)
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                article_count += 1
        
        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[Crocs-JBZ-EAN] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(f"=== CROCS ORDER FORM JBZ EAN SUMMARY ===\n"
          f"  Input rows   : {len(raw_rows)}\n"
          f"  Articles     : {article_count}\n"
          f"  XML size     : {file_kb} KB", flush=True)
    
    if auditor: 
        auditor.set_xml_uploads([str(xml_path)])
    
    return xml_path


# ======================================================================
# MAIN
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crocs Order Form JBZ EAN -> Stibo XML")
    parser.add_argument("--brand",        default=BRAND_NAME)
    parser.add_argument("--brand_code",   default=BRAND_CODE)
    parser.add_argument("--comp_code",    default="0888")
    parser.add_argument("--sbu",          default="SP")
    parser.add_argument("--season",       default="SP27")
    parser.add_argument("--country_code", default="ID")
    args = parser.parse_args()
    
    logging.basicConfig(
        level=logging.INFO, 
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)]
    )
    run(args)
