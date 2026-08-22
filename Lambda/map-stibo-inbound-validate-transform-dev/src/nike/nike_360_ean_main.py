"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Nike 360 EAN v1.0              ║
║   Nike 360 Order Form → Stibo STEP XML (Generic + Variants)    ║
╚══════════════════════════════════════════════════════════════════╝

Column layout (Nike 360 MAA SP 27 ORDER FORM):
    SKU NO.           -> Full SKU identifier (Style.Color.Size)
    CATEGORY          -> Product category (BASEBALL, BASKETBALL, etc.)
    PRODUCT LINE      -> Product line/type
    ITEM DESCRIPTION  -> Product description/model name
    SIZE              -> Size code
    COLOR CODE        -> 3-digit color code (e.g., '052, '010)
    STYLE COLOR       -> Color description
    MCQ               -> Minimum order quantity
    UPC CODE          -> Barcode (EAN/UPC)
    Price             -> Retail price
    QUANTITY          -> Order quantity

Article structure:
  One Generic per unique Style-Color combination
  One Variant per SIZE row
  Principal Barcode → Variant DC_Barcode DataContainer

Generic code: brand_code(3) + style_code(9) + color_code(3) → 15 chars
Variant code: generic_code + SAP-size(3) → 18 chars
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

EAN_DIR     = BASE_DIR / "input" / "linelist"   # Lambda downloads to linelist/
MDD_DIR     = BASE_DIR / "input" / "mdd"
MAPPING_DIR = BASE_DIR / "input" / "mapping"    # Brand mapping file (for validation)
XML_OUT_DIR = BASE_DIR / "output" / "xml"
LOG_DIR     = BASE_DIR / "output" / "logs"

for _d in (EAN_DIR, MDD_DIR, MAPPING_DIR, XML_OUT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')

# ======================================================================
# SIZE CODE LOOKUP — dynamic MDD first, then formula fallback
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
    if len(alnum) == 1:   return "00" + alnum       # M   → 00M
    if len(alnum) == 2:   return "0"  + alnum       # XL  → 0XL
    return alnum[:3]                                 # ABC → ABC


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

    # ── 2. Fallback formula ───────────────────────────────────────
    return _maa_size_code_fallback(s)


# ======================================================================
# LOADERS
# ======================================================================

class MDDLoader:
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
                self.core.append({"id": attr_id, "name": _v("PIM Attribute Name"),
                                  "cardinality": _v("Cardinality"), "lov_name": _v("Name of LOV")})
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


class Nike360EANLoader:
    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheet = ""
        self._load()

    def _load(self):
        log.info("[Nike360-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        target = wb.sheetnames[0]
        log.info("[Nike360-EAN] Using sheet: '%s'", target)
        ws   = wb[target]
        rows = list(ws.iter_rows(values_only=True))
        
        # Header is at row 1 (index 1)
        HEADER_SIGNALS = {"sku no.", "upc code", "color code"}
        hdr_idx = 1
        for i in range(min(10, len(rows))):
            rv = {str(v).strip().lower() for v in rows[i] if v is not None}
            if rv & HEADER_SIGNALS: hdr_idx = i; break
        
        log.info("[Nike360-EAN] Header at row %d (0-indexed)", hdr_idx)
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
        
        # Filter rows with valid SKU NO.
        if "SKU NO." in df.columns:
            df = df[df["SKU NO."].notna() & (~df["SKU NO."].astype(str).str.strip().isin(["","None","nan"]))]
        
        wb.close()
        self.df = df.reset_index(drop=True)
        self.sheet = target
        log.info("[Nike360-EAN] %d rows loaded", len(self.df))


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
        return str(upc_val).strip()

def _price_2dp(v) -> str:
    """Round to max 2 decimal places."""
    if v is None: return ""
    try:
        s = f"{float(str(v).strip()):.2f}".rstrip('0').rstrip('.')
        return s if s else ""
    except (ValueError, TypeError): 
        return ""


# ======================================================================
# MAPPER — groups flat rows into Generic + Variants dict
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str, season: str, country_code: str, mdd=None) -> dict[str, dict]:
    """
    Groups EAN rows into generics.
    SKU NO. format: J.101.2174.052.OS → Style.Color.Size
    Each Style-Color = One Generic
    Each SIZE = One Variant
    """
    result: dict[str, dict] = {}

    for row in raw_rows:
        sku_no = _s(row.get("SKU NO.") or "")
        if not sku_no: continue
        
        # Parse SKU: J.101.2174.052.OS → parts
        parts = sku_no.split(".")
        if len(parts) < 4:
            log.warning(f"Invalid SKU format: {sku_no}")
            continue
        
        # Color: Read directly from COLOR CODE column and remove apostrophe
        color_code_raw = _s(row.get("COLOR CODE") or "")
        color_code = color_code_raw.lstrip("'")  # Remove leading apostrophe: '052 → 052
        
        # Size: Read from SIZE column
        size_raw = _s(row.get("SIZE") or "")
        
        # Extract style (all parts except last - the size)
        style_parts = parts[:-1]
        style_code = ".".join(style_parts)  # e.g., J.101.2174.052
        
        style_color = f"{sku_no[:-len(parts[-1])-1]}-{color_code}"  # Unique key for grouping
        
        # Get other fields
        item_desc = _s(row.get("ITEM DESCRIPTION") or row.get("ITEM DESCRIPTION ") or "")
        category = _s(row.get("CATEGORY") or "")
        product_line = _s(row.get("PRODUCT LINE") or "")
        style_color_desc = _s(row.get("STYLE COLOR") or "")
        upc_str = _fmt_upc(row.get("UPC CODE") or row.get("UPC CODE "))
        price = _s(row.get("Price") or "")
        
        # Use MAA size code: MDD Size Code LOV first, then fallback formula
        size_lov = mdd.lovs.get("Size Code") if mdd else None
        sap_size = _maa_size_code(size_raw, size_lov)
        
        if style_color not in result:
            # Create generic
            # Formula: Brand Code + SKU (from SKU NO.) + Color Code (without apostrophe)
            # SKU includes all parts except size: J.101.2174.052
            sku_without_size = ".".join(parts[:-1])  # e.g., J.101.2174.052
            generic_code = f"{brand_code}{sku_without_size}{color_code}"
            
            result[style_color] = {
                "article_no": style_color,
                "style_code": sku_without_size,  # Full SKU without size
                "colour_code": color_code,
                "generic_code": generic_code,
                "model_name": item_desc,
                "brand_code": brand_code,
                "colour": style_color_desc,
                "gender_code": "U",
                "gender_raw": "Unisex",
                "age_code": "AA",
                "product_type": "Accessories",
                "division": "A",
                "art_category": "2",
                "article_type": "Inline",
                "line": category,
                "subline": product_line,
                "universe": "",
                "discipline": "",
                "collection": category,
                "fob": "",
                "fob_currency": "",
                "rrp": _price_2dp(price),
                "currency": "IDR",
                "coo": "",
                "made_in": "",
                "cbm": "",
                "shelf_date": "",
                "sap_style_code": style_code.replace('.', ''),
                "season": season,
                "country_code": country_code,
                "size_range": "",
                "variants": {},
            }
        
        g = result[style_color]
        variant_code = f"{g['generic_code']}{sap_size}"
        variant_desc = f"Size {size_raw}"[:40]
        
        g["variants"][sap_size] = {
            "size_raw": size_raw,
            "sap_size": sap_size,
            "upc": upc_str,
            "variant_code": variant_code,
            "variant_desc": variant_desc,
        }
    
    # Build size ranges
    for g in result.values():
        g["size_range"] = ", ".join([v["size_raw"] for v in g["variants"].values()])
    
    log.info("[Mapper] %d rows -> %d generics  %d variants",
             len(raw_rows), len(result), sum(len(g["variants"]) for g in result.values()))
    return result


# ======================================================================
# XML HELPERS
# ======================================================================

def _add_barcode_datacontainer(parent_el: ET.Element, upc: str) -> None:
    if not upc: return
    dc_root = ET.SubElement(parent_el, f"{{{STIBO_NS}}}DataContainers")
    mdc = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type","DC_Barcode")
    dc  = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc.set("Analyzer", "true")
    dcv = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")
    v1  = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v1.set("AttributeID","AT_Barcode")
    v1.text = upc
    v2  = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v2.set("AttributeID","AT_BarcodeType")
    v2.set("ID","P")
    v3  = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v3.set("AttributeID","AT_MainEANIndicator")
    v3.set("ID","Y")


# ======================================================================
# PRODUCT XML BUILDER
# ======================================================================

def build_product_xml(generic: dict, brand: str, brand_code: str, comp_code: str,
                      sbu: str, season_id: str, mdd=None) -> str:
    if not generic.get("article_no"): return ""

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", f"PPH_{generic.get('division','A')}-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID","KEY_InboundArticle")
    kv.text = generic["generic_code"]

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    for sap_size in sorted(generic["variants"].keys()):
        variant = generic["variants"][sap_size]
        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID","PRD_VariantArticle")

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID","KEY_InboundVariant")
        v_kv.text = variant["variant_code"]

        ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")

        if variant.get("upc"):
            _add_barcode_datacontainer(v_el, variant["upc"])

    return ET.tostring(g_el, encoding="unicode")


# ======================================================================
# ORCHESTRATOR
# ======================================================================

def run(args, auditor=None):
    brand        = getattr(args, "brand",        "Nike")
    brand_code   = getattr(args, "brand_code",   "NIK")
    comp_code    = getattr(args, "comp_code",    "0888")
    sbu          = getattr(args, "sbu",          "SP")
    season       = getattr(args, "season",       "SP27")
    country_code = getattr(args, "country_code", "ID")

    log.info("[Nike360-EAN] Starting: brand=%s code=%s season=%s", brand, brand_code, season)

    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if not mdd_file: 
        raise FileNotFoundError(f"No MDD file in {MDD_DIR}")
    mdd = MDDLoader(mdd_file)

    # Verify brand mapping file exists (not used for EAN processing)
    mapping_file = next(MAPPING_DIR.glob("*.xlsx"), None)
    if not mapping_file:
        for pattern in ["*brand*mapping*.xlsx", "*mapping*template*.xlsx", "*mapping*.xlsx"]:
            candidates = list(MAPPING_DIR.glob(pattern))
            if candidates:
                mapping_file = candidates[0]
                break
    if mapping_file:
        log.info("[Nike360-EAN] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[Nike360-EAN] Brand mapping file not found in %s", MAPPING_DIR)

    ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file: 
        raise FileNotFoundError(f"No Nike 360 EAN file in {EAN_DIR}")
    ean_loader = Nike360EANLoader(ean_file)
    
    if ean_loader.df.empty:
        log.warning("[Nike360-EAN] No data rows found")
        return

    raw_rows = ean_loader.df.to_dict("records")
    generics = group_ean_rows(raw_rows, brand_code, season, country_code, mdd=mdd)
    
    if not generics:
        log.warning("[Nike360-EAN] No valid generics produced")
        return

    sp = season[:2].upper() if len(season) >= 2 else season
    sys_short = season[2:] if len(season) > 2 else ""
    sy = f"20{sys_short}" if len(sys_short) == 2 else sys_short
    season_id = f"CLH_{brand_code}_{sp}{sy}"

    xml_filename = f"{ean_file.stem}_NIKE360_EAN.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    # TESTING: Limit to first 2 products only
    # generics_to_process = dict(list(generics.items())[:2])
    # log.info("[Nike360-EAN] TESTING MODE: Processing only first 2 products")
    
    log.info("[Nike360-EAN] Writing XML -> %s", xml_filename)
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
            px = build_product_xml(generic, brand, brand_code, comp_code, sbu, season_id, mdd=mdd)
            if px:
                f.write(f"    {_XMLNS_RE.sub('', px)}\n")
                generic_count += 1
                variant_count += len(generic["variants"])
        f.write("  </Products>\n</STEP-ProductInformation>\n")

    file_kb = xml_path.stat().st_size // 1024
    log.info("[Nike360-EAN] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(f"=== NIKE 360 EAN SUMMARY ===\n"
          f"  Input rows : {len(raw_rows)}\n  Generics   : {generic_count}\n"
          f"  Variants   : {variant_count}\n  XML size   : {file_kb} KB", flush=True)
    
    if auditor: 
        auditor.set_xml_uploads([str(xml_path)])
    return xml_path


# ======================================================================
# MAIN
# ======================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Nike 360 EAN -> Stibo XML")
    parser.add_argument("--brand",        default="Nike")
    parser.add_argument("--brand_code",   default="NIK")
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
