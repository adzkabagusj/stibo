"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — Steve Madden EAN Source v1.0    ║
║   EAN Source → Stibo STEP XML (Generic + Variants)              ║
╚══════════════════════════════════════════════════════════════════╝

Input file:
    STEVE MADDEN_EAN Source.xlsx   (sheet: "BARCODES")

Column layout (from MAA comments / SVM(Inline) brand mapping):
    STYLE          -> Principal Style Code   (AT_PrincipalStyleCode)
    COLOR CODE     -> Principal Color Code    (AT_PrincipalColorCode)
    SIZE           -> Principal Size          (→ Size LOV ID)
    BARCODES       -> Principal Barcode        (AT_Barcode / EAN)

Article structure:
  One Generic per unique STYLE + COLOR CODE
  One Variant per row (size LOV ID from MDD Size Code LOV lookup)
  Principal Barcode → Variant DC_Barcode DataContainer

KEY_InboundArticle formula:
    BrandCode + PrincipalStyleCode + PrincipalColorCode

KEY_InboundVariant formula:
    KEY_InboundArticle + Size LOV ID  (from MDD Size Code LOV lookup)

Size LOV ID logic:
    Reference: crocs/shipment_confirmation_main.py
    MDD "Size Code LOV" (Col G description → Col F LOV ID), else fallback
    zero-pad / 'H' half-size / alpha-pad formula.

Business rule:
    Rows with NO Color Code OR NO Size value are SKIPPED (not sent).
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
BRAND_NAME = "SteveMadden"
BRAND_CODE = "SVM"          # 3-letter SAP brand code for Steve Madden

# ======================================================================
# XML NAMESPACE
# ======================================================================
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z]+)?="[^"]*"')


# ======================================================================
# SIZE TRANSFORMATION  (reference: crocs/shipment_confirmation_main.py)
# ======================================================================

def _transform_size_value(size_val: str) -> str:
    """Normalise raw size text before MDD lookup.

    Rules:
      M7W9    → M7/W9
      M10W12  → M10/W12
      10      → 10   (numeric unchanged)
      XL      → XL   (alpha unchanged)
    """
    if not size_val:
        return ""

    s = str(size_val).strip().upper()

    # Pattern: M<digits>W<digits>  (e.g. M7W9, M10W12)
    pattern = re.match(r'^(M\d+)(W\d+)$', s)
    if pattern:
        return f"{pattern.group(1)}/{pattern.group(2)}"

    return s


def _maa_size_code_fallback(size_val: str) -> str:
    """Fallback formula when size is not found in MDD Size Code LOV.

    Rules (from brand spec):
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

    # Pure integer numeric → zero-pad to 3  (10→010, 7→007)
    if re.match(r'^\d+$', s):
        try:
            return str(int(s)).zfill(3)[:3]
        except ValueError:
            pass

    # Half-size numeric → 2-digit int part + 'H'  (6.5→06H, 10.5→10H)
    m = re.match(r'^(\d+)\.5$', s)
    if m:
        return str(int(m.group(1))).zfill(2)[:2] + "H"

    # Other decimal → zero-pad int part to 3 digits  (e.g. 6.0→060)
    m2 = re.match(r'^(\d+)\.\d+$', s)
    if m2:
        return str(int(m2.group(1))).zfill(3)[:3]

    # Alpha / alphanumeric → left-pad with '0' to reach 3 chars
    alnum = re.sub(r"[^A-Z0-9/]", "", s)
    if not alnum:
        return "XXX"

    # Remove slash for LOV ID (M7/W9 → M7W9 → M7W)
    alnum_no_slash = alnum.replace("/", "")

    if len(alnum_no_slash) == 1:   return "00" + alnum_no_slash       # M   → 00M
    if len(alnum_no_slash) == 2:   return "0"  + alnum_no_slash       # XL  → 0XL
    return alnum_no_slash[:3]                                          # ABC → ABC, M7W9 → M7W


def _maa_size_code(size_val: str, size_lov: dict | None = None) -> str | None:
    """Return a 3-character MAA size code for use in KEY_InboundVariant.

    Only returns LOV ID if found in MDD 'Size Code LOV' (Col G description → Col F LOV ID).
    Returns None if not found in LOV (variant will be skipped).
    """
    if not size_lov:
        return None

    s = size_val.strip().upper()

    # ── MDD Size Code LOV lookup (Col G → Col F) ───────────────
    if s in size_lov:
        return size_lov[s]
    
    stripped = s.lstrip("0") or "0"
    if stripped in size_lov:
        return size_lov[stripped]
    
    # Try without slash
    s_no_slash = s.replace("/", "")
    if s_no_slash in size_lov:
        return size_lov[s_no_slash]

    # Not found in LOV - return None to skip this variant
    return None


# ======================================================================
# MDD LOADER  (reference: crocs/shipment_confirmation_main.py)
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
        self._load_size_lov(wb)   # Col G (description) → Col F (LOV ID)
        wb.close()
        log.info("[MDD] %d LOVs", len(self.lovs))

    def _load_size_lov(self, wb):
        """Load Size Code LOV: Col G = description, Col F = Stibo LOV ID.
        Stored as lovs['Size Code'] = { description: lov_id }."""
        if "Size Code LOV" not in wb.sheetnames:
            log.warning("[MDD] 'Size Code LOV' sheet not found")
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
            if key.endswith('.'):
                key = key[:-1]                      # strip trailing dot
            if key and lid:
                size_map.setdefault(key, lid)       # first occurrence wins
        if size_map:
            self.lovs["Size Code"] = size_map
            log.info("[MDD] Size Code LOV (colF/G): %d entries", len(size_map))


# ======================================================================
# EAN SOURCE LOADER
# ======================================================================

class SteveMaddenEANLoader:
    """Loads Steve Madden EAN Source data from the 'BARCODES' tab."""

    TARGET_SHEET = "BARCODES"

    def __init__(self, path: Path):
        self.path = path
        self.df = pd.DataFrame()
        self.sheet = ""
        self.headers: list[str] = []
        self._load()

    def _load(self):
        log.info("[SVM-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        log.info("[SVM-EAN] Available sheets: %s", wb.sheetnames)

        target = None
        for sn in wb.sheetnames:
            if sn.strip().lower() == self.TARGET_SHEET.lower():
                target = sn
                break
        if not target:
            target = wb.sheetnames[0]
            log.warning("[SVM-EAN] '%s' sheet not found — using first sheet '%s'",
                        self.TARGET_SHEET, target)

        log.info("[SVM-EAN] Using sheet: '%s'", target)
        ws = wb[target]
        rows = list(ws.iter_rows(values_only=True))

        # Find header row — look for STYLE / COLOR CODE / SIZE / BARCODES
        HEADER_SIGNALS = {"style", "color code", "size", "barcodes", "barcode"}
        hdr_idx = 0
        for i in range(min(15, len(rows))):
            if rows[i]:
                rv = {str(v).strip().lower() for v in rows[i] if v is not None}
                if rv & HEADER_SIGNALS:
                    hdr_idx = i
                    break

        log.info("[SVM-EAN] Header at row %d (0-indexed)", hdr_idx)
        header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[hdr_idx])]

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
        log.info("[SVM-EAN] %d rows loaded", len(self.df))
        if self.headers:
            log.info("[SVM-EAN] Headers: %s", self.headers[:15])

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


def _color_code_3d(color_code: str) -> str:
    """Convert color code to 3-character format for the variant key."""
    cc = _s(color_code)
    if not cc:
        return ""
    digits = re.sub(r"[^0-9]", "", cc)
    if digits:
        return digits[:3].zfill(3)
    alpha = re.sub(r"[^A-Za-z0-9]", "", cc).upper()
    return alpha[:3].ljust(3, "0") if alpha else ""


# ======================================================================
# MAPPER — groups flat rows into Generic + Variants
# ======================================================================

def group_ean_rows(raw_rows: list[dict], brand_code: str,
                   loader: SteveMaddenEANLoader, mdd=None) -> dict[str, dict]:
    """
    Groups EAN rows into generics.

    KEY_InboundArticle = BrandCode + PrincipalStyleCode + PrincipalColorCode
    KEY_InboundVariant = KEY_InboundArticle + Size LOV ID (from MDD lookup)

    Business rule: rows with NO Color Code OR NO Size are SKIPPED.
    """
    result: dict[str, dict] = {}

    style_col   = loader.find_column("STYLE", "Style", "style")
    color_col   = loader.find_column("COLOR CODE", "Color Code", "COLORCODE",
                                     "color code", "COLOUR CODE", "Colour Code")
    size_col    = loader.find_column("SIZE", "Size", "size", "SIZE CODE", "Size Code")
    barcode_col = loader.find_column("BARCODES", "BARCODE", "Barcode", "EAN", "UPC")

    if not style_col:
        log.error("[SVM-EAN] Cannot find STYLE column. Available headers: %s",
                  loader.headers)
        return result

    if not color_col:
        log.error("[SVM-EAN] COLOR CODE column NOT FOUND — every row will be skipped. "
                  "Available headers: %s", loader.headers)
    if not size_col:
        log.error("[SVM-EAN] SIZE column NOT FOUND — every row will be skipped. "
                  "Available headers: %s", loader.headers)

    log.info("[SVM-EAN] Columns: STYLE='%s' COLOR CODE='%s' SIZE='%s' BARCODES='%s'",
             style_col, color_col, size_col, barcode_col)

    size_lov = mdd.lovs.get("Size Code") if mdd else None
    if size_lov:
        log.info("[SVM-EAN] Size Code LOV loaded with %d entries", len(size_lov))
    else:
        log.warning("[SVM-EAN] Size Code LOV not found — all variants will be skipped")

    skipped_no_color = skipped_no_size = skipped_no_size_lov = 0

    for row in raw_rows:
        style_raw = _s(row.get(style_col)) if style_col else ""
        if not style_raw:
            continue

        color_raw = _s(row.get(color_col)) if color_col else ""
        size_raw  = _s(row.get(size_col)) if size_col else ""

        # ── Business rule: skip rows with no Color Code OR no Size ──────
        if not color_raw:
            skipped_no_color += 1
            continue
        if not size_raw:
            skipped_no_size += 1
            continue

        color_code_3 = _color_code_3d(color_raw)
        if not color_code_3:
            skipped_no_color += 1
            continue

        # Size LOV ID - only proceed if found in MDD
        size_transformed = _transform_size_value(size_raw)
        size_lov_id = _maa_size_code(size_transformed, size_lov)
        
        # Skip variant if size not found in LOV
        if not size_lov_id:
            skipped_no_size_lov += 1
            continue

        barcode = _fmt_barcode(row.get(barcode_col)) if barcode_col else ""

        # KEY_InboundArticle = BrandCode + Style + ColorCode
        generic_code = f"{brand_code}{style_raw}{color_code_3}"

        if generic_code not in result:
            result[generic_code] = {
                "article_no":   style_raw,
                "style_code":   style_raw,
                "color_code":   color_raw,
                "color_code_3": color_code_3,
                "generic_code": generic_code,
                "brand_code":   brand_code,
                "variants":     {},
            }

        # KEY_InboundVariant = generic_code + Size LOV ID
        variant_code = f"{generic_code}{size_lov_id}"
        variant_key = size_lov_id

        g = result[generic_code]
        if variant_key not in g["variants"]:
            g["variants"][variant_key] = {
                "size_raw":         size_raw,
                "size_transformed": size_transformed,
                "size_lov_id":      size_lov_id,
                "barcode":          barcode,
                "variant_code":     variant_code,
            }

    log.info("[SVM-EAN] %d rows -> %d generics  %d variants "
             "(skipped: no-color=%d, no-size=%d, size-not-in-lov=%d)",
             len(raw_rows), len(result),
             sum(len(g["variants"]) for g in result.values()),
             skipped_no_color, skipped_no_size, skipped_no_size_lov)
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

    # Empty Values element for generic
    ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

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
    """Main entry point for Steve Madden EAN Source processing."""
    brand      = getattr(args, "brand",      BRAND_NAME)
    brand_code = getattr(args, "brand_code", BRAND_CODE)

    log.info("[SVM-EAN] Starting: brand=%s code=%s", brand, brand_code)

    # Load MDD (Size Code LOV)
    mdd = None
    mdd_file = next(MDD_DIR.glob("*.xlsx"), None)
    if mdd_file:
        mdd = MDDLoader(mdd_file)
    else:
        log.warning("[SVM-EAN] No MDD file found in %s", MDD_DIR)

    # Verify brand mapping file exists (informational)
    mapping_file = next(ATTR_DIR.glob("*.xlsx"), None)
    if mapping_file:
        log.info("[SVM-EAN] Brand mapping file validated: %s", mapping_file.name)
    else:
        log.warning("[SVM-EAN] Brand mapping file not found in %s", ATTR_DIR)

    # Load input EAN source file
    ean_file = getattr(args, "input_file", None)
    if ean_file and Path(ean_file).exists():
        ean_file = Path(ean_file)
    else:
        ean_file = next(EAN_DIR.glob("*.xls*"), None)
    if not ean_file:
        raise FileNotFoundError(f"No EAN Source file in {EAN_DIR}")

    loader = SteveMaddenEANLoader(ean_file)
    if loader.df.empty:
        log.warning("[SVM-EAN] No data rows found")
        return

    raw_rows = loader.df.to_dict("records")

    generics = group_ean_rows(raw_rows, brand_code, loader, mdd=mdd)

    if not generics:
        log.warning("[SVM-EAN] No valid generics produced (all rows missing color/size?)")
        return

    log.info("[SVM-EAN] Total generics to write: %d", len(generics))

    # Generate XML
    xml_filename = f"{ean_file.stem}_SVM_EAN.xml"
    xml_path     = XML_OUT_DIR / xml_filename
    export_time  = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = variant_count = 0

    log.info("[SVM-EAN] Writing XML -> %s", xml_filename)
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
    log.info("[SVM-EAN] XML written: %s  (%d KB)", xml_path.name, file_kb)
    print(f"=== STEVE MADDEN EAN SOURCE SUMMARY ===\n"
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
    parser = argparse.ArgumentParser(description="Steve Madden EAN Source -> Stibo XML")
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
