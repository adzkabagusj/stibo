"""
╔══════════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — New Balance Licensed Order Sheet v1.0 ║
║  UPC / Master-UPC List → Stibo STEP XML (Generic + Variants)         ║
╚══════════════════════════════════════════════════════════════════════╝

Source file
  • Single input file  : "Master_UPC_List-{Season}_Apparel_Official_Price_List.xlsx"
                         (or any file routed to /tmp/stibo_workdir/input/ean_source/)
  • Active sheet       : first sheet (or the one containing the UPC data)
  • Header row         : row 0 (0-based, the file has NO extra title rows)
  • Data starts        : row 1

Column layout (0-indexed, confirmed from file):
    col[0]  ITEM_NUMBER   → Principal Style Code (item / style identifier)
    col[1]  SKU           → {ITEM_NUMBER}_{COLOR_CODE}_{SIZE}  (full variant key)
    col[2]  SKU UPC       → {ITEM_NUMBER}_{COLOR_CODE}         (generic key from source)
    col[3]  UK Size       → UK size string (present for some footwear rows)
    col[4]  UPC#          → Principal Barcode / EAN value
    col[5]  Width Description → Width label (usually empty for apparel)
    col[6]  SKU Size      → Principal Size Code (e.g. S, M, L, XL, OSZ …)
    col[7]  WIDTH         → Width code (usually empty for apparel)
    col[8]  SKU Color Code → Principal Color Code (e.g. BK, AS1, WT …)

Article structure
  • One Generic  per unique (ITEM_NUMBER, SKU_COLOR_CODE) combination
  • One Variant  per row (i.e. per unique ITEM_NUMBER + COLOR + SIZE)
  • Principal Barcode (AT_PrincipalBarcode) lives on the Variant

Generic code  : brand_code(3) + ITEM_NUMBER (up to 9 chars) + SKU Color Code (full, alphanumeric)
                e.g. NEWAC0048UBK
Variant code  : brand_code(3) + ITEM_NUMBER (up to 9 chars) + color_code (full, alphanumeric) + size_lov_id (3 chars)
                e.g. NEWAC0048UBK005

Generic Description  : brand_code + " " + ITEM_NUMBER + " " + COLOR_CODE
Variant Description  : generic_desc + " " + SAP_SIZE_CODE

Color code (variant) : full alphanumeric SKU_COLOR_CODE (no length restriction)
Size LOV ID (variant): looked up from MDD Size Code LOV; fallback to zero-left-padded 3-char code

────────────────────────────────────────────────────────────────────────
ENVIRONMENT VARIABLES (inherited from lambda_function.py)
────────────────────────────────────────────────────────────────────────
    LAMBDA_TMP_DIR   Base working directory  (default: /tmp/stibo_workdir)
"""

import re
import sys
import logging
import argparse
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import os

# ─────────────────────────────────────────────────────────────────────────────
# Directory layout
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR      = BASE_DIR / "input"
ORDER_DIR = INPUT_DIR / "ordersheet_licensed"   # routed by lambda_function.py
MDD_DIR        = INPUT_DIR / "mdd"
ATTR_DIR       = INPUT_DIR / "attributes"
NAMING_DIR     = INPUT_DIR / "naming"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for _d in [ORDER_DIR, MDD_DIR, ATTR_DIR, NAMING_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
_log_path = LOG_DIR / f"ordersheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(_log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def _parse_metadata_from_input_filename(stem: str) -> dict:
    normalized = stem.replace("_", "-")
    parts = re.split(r"\s*-\s*", normalized)
    season = ""
    season_idx = None
    for i, p in enumerate(parts):
        tok = p.strip()
        if re.match(r"^(SS|FW|AW|HO|AL|CO|WN)\d{2}(\d{2})?$", tok, re.IGNORECASE):
            season = tok.upper()
            season_idx = i
            break

    article_type = ""
    for p in parts:
        token = p.strip().lower()
        if "inline" in token:
            article_type = "Inline"
            break
        if "license" in token or "licensed" in token:
            article_type = "License"
            break
        if "sse" in token:
            article_type = "SSE"
            break

    country = ""
    if season_idx is not None:
        trailing = parts[season_idx + 1:]
        for p in trailing:
            tok = p.strip()
            m = re.match(r"^([A-Z]{2,3})(?:\d+)?$", tok, re.IGNORECASE)
            if m and not tok.isdigit():
                country = m.group(1).upper()
                break

    # Fallback: if season is not detected or country wasn't found in trailing
    # tokens, use the last country-like token from filename (e.g. "MY-1").
    _NON_COUNTRY_TOKENS = {
        "SP", "AP", "FW", "SS", "AW", "HO", "AL", "GN",
        "AC", "EQ", "NEW", "NIK", "ADI", "SMI", "ALD",
        "CRO", "LOT", "BIR", "NB", "MULTI", "MONO",
    }

    if not country:
        for p in reversed(parts):
            tok = p.strip().upper()
            if (
                re.match(r'^[A-Z]{2,3}$', tok)
                and not tok.isdigit()
                and tok not in _NON_COUNTRY_TOKENS
            ):
                country = tok
                break

    return {"season": season, "article_type": article_type, "country": country}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — LOOKUP TABLES
# ══════════════════════════════════════════════════════════════════════════════

LOV_BRAND: dict[str, str] = {
    "NEW": "NEW BALANCE",
    "ADI": "ADIDAS",
    "NIK": "NIKE",
    "SMI": "SMIGGLE",
    "ALD": "ALDO",
    "CRO": "CROCS",
    "LOT": "LOTTO",
    "BIR": "BIRKENSTOCK",
}

LOV_SEASON: dict[str, str] = {
    "SS": "Spring-Summer",
    "FW": "Fall-Winter",
    "AL": "All Season",
    "HO": "Holiday",
    "AW": "Autumn-Winter",
}

LOV_SBU: dict[str, str] = {
    "SP": "Sports",   "FQ": "Footlocker",
    "FL": "Fashion Footwear", "SM": "Smiggle",
}

LOV_AGE: dict[str, str] = {
    "AD": "Adults", "CH": "Children",
    "IN": "Infant", "AA": "All Ages", "JR": "Junior",
}

LOV_GENDER: dict[str, str] = {
    "M": "Male", "F": "Female", "U": "Unisex",
}

LOV_COUNTRY_ORIGIN: dict[str, str] = {
    "CN": "China",     "VN": "Vietnam",   "ID": "Indonesia",
    "KH": "Cambodia",  "BD": "Bangladesh","IN": "India",
    "MY": "Malaysia",  "TH": "Thailand",  "PK": "Pakistan",
    "LK": "Sri Lanka", "JO": "Jordan",    "SG": "Singapore",
}
LOV_COUNTRY: dict[str, str] = {
    "ID": "Indonesia",
    "SG": "Singapore",
    "MY": "Malaysia",
    "TH": "Thailand",
    "PH": "Philippines",
    "VN": "Vietnam",
    "KH": "Cambodia",
    "IN": "India",
}

LOV_COMPANY_CODE: dict[str, str] = {
    "0888": "PT. Map Aktif Adiperkasa",
    "0886": "PT. MAP FTL Adiperkasa",
    "0882": "Magna Management Asia",
}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STIBO XML CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace("", STIBO_NS)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — LOADERS
# ══════════════════════════════════════════════════════════════════════════════

class MDDLoader:
    """Loads Core Attributes + LOVs from the MDD Excel workbook."""

    def __init__(self, path: Path):
        self.path       = path
        self.attributes: dict[str, dict] = {}
        self.lovs:       dict[str, dict] = {}   # {lov_name: {display_name: id_code}}
        self._load()

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        wb   = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
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
                "cardinality":  _v("Cardinality \\n(Conditional / Mandatory / Optional)"),
                "validation":   _v("Validation Base Type"),
                "multi_valued": _v("Multi Valued"),
                "lov_name":     _v("Name of LOV\\n(Only if Validation Base Type = LOV)"),
                "max_chars":    _v("Max Characters\\n<Only if Validation Type = Text / Numeric Text / URL>"),
                "group":        _v("PIM Attribute Group\\n(As Configured in STEP)"),
            }

        self._load_simple_lovs(wb)
        self._load_named_lov_sheets(wb)
        wb.close()
        log.info("[MDD] %d attributes | %d LOVs", len(self.attributes), len(self.lovs))

    def _load_simple_lovs(self, wb):
        if "Simple LOVs" not in wb.sheetnames:
            return
        for row in list(wb["Simple LOVs"].iter_rows(values_only=True))[2:]:
            lov_name, _, val_name, val_id = (row + (None,) * 4)[:4]
            if lov_name and val_name:
                raw_id = str(val_id).strip() if val_id else str(val_name).strip()
                if raw_id.isdigit():
                    raw_id = raw_id.zfill(3)
                self.lovs.setdefault(str(lov_name).strip(), {})[
                    str(val_name).strip()
                ] = raw_id

    def _load_named_lov_sheets(self, wb):
        for sn in wb.sheetnames:
            if "LOV" not in sn.upper():
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            display = sn.replace(" LOV", "").replace("LOV ", "").strip()

            # ── For Size Code LOV: Col F = Size Code (ID), Col G = Size Code Description (value) ──
            if "SIZE" in sn.upper() and "CODE" in sn.upper():
                code_col = 5  # Column F (0-indexed)
                name_col = 6  # Column G (0-indexed)
                log.info("[MDD] Sheet '%s' → Size Code LOV detected → code_col=F(5), name_col=G(6)", sn)
            else:
                # ── Detect header row to find correct col indices for other LOV sheets ──
                header_row = rows[0]
                code_col = None
                name_col = None
                for i, h in enumerate(header_row):
                    if h is None:
                        continue
                    h_str = str(h).strip().lower()
                    if h_str in ("size code", "color code", "code"):
                        code_col = i
                    elif h_str in ("size code description", "color code description", "description", "name"):
                        name_col = i

                # Fallback to col 0/1 if header detection fails
                if code_col is None:
                    code_col = 0
                if name_col is None:
                    name_col = 1

                log.info("[MDD] Sheet '%s' → code_col=%d  name_col=%d", sn, code_col, name_col)

            for row in rows[1:]:
                if not row or len(row) <= max(code_col, name_col):
                    continue
                code = row[code_col]
                name = row[name_col]
                if name:
                    self.lovs.setdefault(display, {})[str(name).strip()] = (
                        str(code).strip() if code else str(name).strip()
                    )

            log.info("[MDD] LOV '%s' → %d entries loaded", display, len(self.lovs.get(display, {})))


class NamingConventionLoader:
    """Loads naming convention rules — used only for output filename building."""

    def __init__(self, path: Path):
        self.path  = path
        self.rules: dict[str, str] = {}
        self._load()

    def _load(self):
        log.info("[Naming] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        for sheet in ["Approved", "Sheet1"]:
            if sheet not in wb.sheetnames:
                continue
            for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
                if not row or not row[0]:
                    continue
                brand   = str(row[0]).strip().upper()
                pattern = row[2] if len(row) > 2 else None
                if pattern and "[" in str(pattern):
                    self.rules[brand] = str(pattern).strip()
            break
        wb.close()
        log.info("[Naming] %d brand rules loaded", len(self.rules))

    def build_filename(
        self,
        brand: str,
        comp_code: str,
        sbu: str,
        file_type: str,
        multi_mono: str,
        season: str,
        seq: int,
    ) -> str:
        return f"{comp_code}-{sbu}-{brand}-{file_type}-{multi_mono}-{season}-{seq}.xml"


class NBOrderSheetLoader:
    """
    Loads the New Balance Master UPC / Order Sheet file.

    Expected columns (header in row 0, data from row 1):
        [0]  ITEM_NUMBER     → Principal Style Code
        [1]  SKU             → {ITEM_NUMBER}_{COLOR}_{SIZE}
        [2]  SKU UPC         → {ITEM_NUMBER}_{COLOR}
        [3]  UK Size         → UK size (optional, footwear)
        [4]  UPC#            → barcode / EAN value
        [5]  Width Description → width label (usually empty for apparel)
        [6]  SKU Size        → Principal Size Code
        [7]  WIDTH           → width code (usually empty)
        [8]  SKU Color Code  → Principal Color Code

    Returns a pandas-style list of dicts for easy iteration.
    """

    EXPECTED_COLS = {
        "ITEM_NUMBER", "SKU", "SKU UPC", "UPC#", "SKU Size", "SKU Color Code",
    }

    def __init__(self, path: Path):
        self.path  = path
        self.rows: list[dict] = []
        self._load()

    def _load(self):
        log.info("[OrderSheet] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        # Prefer the first sheet; the UPC report sheet name may vary
        target = wb.sheetnames[0]
        log.info("[OrderSheet] Using sheet: '%s'", target)
        ws = wb[target]

        raw_rows = list(ws.iter_rows(values_only=True))
        if not raw_rows:
            log.error("[OrderSheet] Empty sheet — nothing to load.")
            wb.close()
            return

        # Header is always row 0
        header = [str(h).strip() if h is not None else f"col_{i}"
                  for i, h in enumerate(raw_rows[0])]

        missing = self.EXPECTED_COLS - set(header)
        if missing:
            log.warning("[OrderSheet] Missing expected columns: %s", missing)

        for raw in raw_rows[1:]:
            row = {header[i]: v for i, v in enumerate(raw) if i < len(header)}
            item = _s(row.get("ITEM_NUMBER"))
            if not item:
                continue
            upc = row.get("UPC#")
            if upc is None:
                continue
            self.rows.append(row)

        wb.close()
        log.info("[OrderSheet] %d data rows loaded", len(self.rows))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _s(v) -> str:
    """Safely stringify any cell value; returns '' for empties / NaN sentinels."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT", "N/A", "none") else s


def _size_to_sap_code(size_raw: str, mdd: "MDDLoader | None" = None) -> str:
    """
    Map a raw principal size string to a 3-char SAP size code.
    Looks up in MDD Size Code LOV, then fallback to zero-prefix padding.
    """
    key = size_raw.strip().upper()
    
    # 1. Try MDD Size Code LOV lookup
    if mdd:
        size_lov = mdd.lovs.get("Size Code", {})
        if size_lov:
            # Direct match
            if key in size_lov:
                lov_id = size_lov[key]
                if lov_id:
                    return str(lov_id).strip()
            # Case-insensitive match
            for desc, code in size_lov.items():
                if desc.strip().upper() == key:
                    if code:
                        return str(code).strip()
    
    # 2. Fallback: strip non-alphanumeric, take first 3, left-pad with 0
    cleaned = re.sub(r"[^A-Z0-9]", "", key)[:3]
    return cleaned.rjust(3, "0")


def _lookup_size_from_mdd(size_raw: str, mdd: "MDDLoader | None") -> str:
    """
    Look up AT_Size LOV ID from MDD 'Size Code LOV' sheet.
    Matches size_raw against Size Code Description (col G),
    returns the Size Code (col F) as the LOV ID.
    Falls back to zero-left-padded 3-char code if not found.
    """
    if not mdd:
        log.warning("[Size LOV] No MDD available — using fallback for size: %s", size_raw)
        cleaned = re.sub(r"[^A-Z0-9]", "", size_raw.strip().upper())[:3]
        return cleaned.rjust(3, "0")

    size_lov = mdd.lovs.get("Size Code", {})
    if not size_lov:
        log.warning("[Size LOV] 'Size Code' LOV not found in MDD — using fallback for size: %s", size_raw)
        cleaned = re.sub(r"[^A-Z0-9]", "", size_raw.strip().upper())[:3]
        return cleaned.rjust(3, "0")

    key = size_raw.strip()

    # 1. Direct match on description
    if key in size_lov:
        log.debug("[Size LOV] Direct match: '%s' → '%s'", key, size_lov[key])
        return size_lov[key]

    # 2. Case-insensitive match
    key_upper = key.upper()
    for desc, code in size_lov.items():
        if desc.strip().upper() == key_upper:
            log.debug("[Size LOV] Case-insensitive match: '%s' → '%s'", key, code)
            return code

    # 3. Fallback: strip non-alphanumeric, take first 3, left-pad with 0
    log.warning("[Size LOV] No match for '%s' — using zero-padded fallback", size_raw)
    cleaned = re.sub(r"[^A-Z0-9]", "", key_upper)[:3]
    return cleaned.rjust(3, "0")



def _color_to_sap_token(color_code: str) -> str:
    """
    Derive SAP Color token from the principal color code (full length, alphanumeric only).
    e.g. "BK" → "BK", "AS1" → "AS1", "WT" → "WT", "CUSTOM" → "CUSTOM"
    """
    raw = re.sub(r"[^A-Z0-9]", "", color_code.strip().upper())
    return raw


def _build_generic_base_code(brand_code: str, item_number: str) -> str:
    """
    3 chars brand_code + up to 9 chars of ITEM_NUMBER (alphanumeric only).
    e.g. "NEW" + "AC0048U" → "NEWAC0048U"
    """
    clean = re.sub(r"[^A-Z0-9]", "", item_number.upper())[:9]
    return f"{brand_code}{clean}"


def _build_generic_code(brand_code: str, item_number: str, color_code: str, is_footwear: bool = False) -> str:
    """
    Generic inbound code:
    For footwear: brand_code + ITEM_NUMBER[:9]
    For apparel/accessories: brand_code + ITEM_NUMBER[:9] + direct SKU Color Code
    """
    base = _build_generic_base_code(brand_code, item_number)
    if is_footwear:
        return base
    color_clean = re.sub(r"[^A-Z0-9]", "", color_code.upper())
    return f"{base}{color_clean}"


def _build_variant_code(generic_base_code: str, color_token: str, sap_size: str) -> str:
    """brand_code + ITEM_NUMBER[:9] + color_token(full) + sap_size(3)."""
    return f"{generic_base_code}{color_token}{sap_size}"


def _build_generic_description(brand_code: str, item_number: str, color_code: str) -> str:
    """
    MAA Generic Description — max 40 chars.
    Format: '{brand_code} {ITEM_NUMBER} {COLOR_CODE}'
    e.g. 'NEW AC0048U BK'
    """
    desc = f"{brand_code} {item_number.upper()} {color_code.upper()}"
    return desc[:40]


def _build_variant_description(generic_desc: str, size_display: str) -> str:
    """Variant Description — generic_desc + ' ' + size_display, max 40 chars."""
    return f"{generic_desc} {size_display}"[:40]


def _fmt_upc(upc_val) -> str:
    """Return UPC# as a clean string (int cells lose decimal point)."""
    if upc_val is None:
        return ""
    try:
        return str(int(float(str(upc_val).strip())))
    except (ValueError, TypeError):
        return str(upc_val).strip()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — MAPPER
# ══════════════════════════════════════════════════════════════════════════════

def map_ordersheet_rows(
    rows: list[dict],
    brand_code: str,
    season: str,
    mdd: "MDDLoader | None" = None,
) -> dict[str, dict]:
    """
    Group order-sheet rows into a nested dict:

        result[generic_key] = {
            "item_number": str,
            "color_code":  str,
            "color_token": str,        # 3-char SAP token
            "generic_code": str,       # brand_code + item_number[:9] (+ color_code if apparel)
            "generic_base_code": str,  # brand_code + item_number[:9]
            "generic_desc": str,
            "variants": {
                sap_size_code: {
                    "sku":         str,
                    "size_raw":    str,
                    "sap_size":    str,
                    "upc":         str,
                    "uk_size":     str,
                    "variant_code": str,
                    "variant_desc": str,
                }
            }
        }

    generic_key = "{item_number}_{color_code}"  (matches SKU UPC column)
    """
    result: dict[str, dict] = {}

    for row in rows:
        item_number = _s(row.get("ITEM_NUMBER"))
        color_code  = _s(row.get("SKU Color Code"))
        size_raw    = _s(row.get("SKU Size"))
        sku         = _s(row.get("SKU"))
        sku_upc     = _s(row.get("SKU UPC"))
        upc_val     = row.get("UPC#")
        uk_size     = _s(row.get("UK Size"))
        country     = _s(
            row.get("Country")
            or row.get("Country of Origin")
            or row.get("COO")
            or row.get("COUNTRY")
            or ""
        )

        if not item_number or not color_code or not size_raw:
            log.debug("  Skipping incomplete row: %s", row)
            continue

        upc_str     = _fmt_upc(upc_val)
        color_token = _color_to_sap_token(color_code)
        sap_size    = _size_to_sap_code(size_raw, mdd=mdd)
        generic_key = sku_upc or f"{item_number}_{color_code}"

        if generic_key not in result:
            is_footwear = any(_s(row.get(k)) != "" for k in ["WIDTH", "Width", "Width Description", "WIDTH DESCRIPTION"])
            gen_base_code = _build_generic_base_code(brand_code, item_number)
            gen_code = _build_generic_code(brand_code, item_number, color_code, is_footwear=is_footwear)
            gen_desc = _build_generic_description(brand_code, item_number, color_code)
            result[generic_key] = {
                "item_number":   item_number,
                "color_code":    color_code,
                "color_token":   color_token,
                "generic_code":  gen_code,
                "generic_base_code": gen_base_code,
                "generic_desc":  gen_desc,
                "season":        season,
                "brand_code":    brand_code,
                "variants":      {},
            }

        g = result[generic_key]
        # De-duplicate on sap_size; last row wins if duplicated (data issue)
        var_code = _build_variant_code(g["generic_base_code"], color_token, sap_size)
        var_desc = _build_variant_description(g["generic_desc"], size_raw)

        g["variants"][sap_size] = {
            "sku":          sku,
            "size_raw":     size_raw,
            "sap_size":     sap_size,
            "upc":          upc_str,
            "uk_size":      uk_size,
            "country":      country, 
            "variant_code": var_code,
            "variant_desc": var_desc,
        }

    log.info(
        "[Mapper] %d input rows → %d generics  %d total variants",
        len(rows),
        len(result),
        sum(len(g["variants"]) for g in result.values()),
    )
    return result


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _val(
    parent:  ET.Element,
    attr_id: str,
    value:   str = "",
    id_val:  str = "",
) -> ET.Element | None:
    clean_id  = str(id_val).strip()  if id_val  else ""
    clean_val = str(value).strip()   if value   else ""
    if clean_id in ("", "None", "nan"):
        clean_id = ""
    if clean_val in ("", "None", "nan"):
        clean_val = ""
    if not clean_id and not clean_val:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if clean_id:
        el.set("ID", clean_id)
    if clean_val:
        el.text = clean_val
    return el


def _multival(parent: ET.Element, attr_id: str, id_val: str) -> None:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v  = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", id_val)


def _add_barcode_datacontainer(parent_el: ET.Element, upc: str, variant_code: str) -> None:
    """Write DC_Barcode DataContainer block for a variant's EAN/UPC value."""
    if not upc:
        return

    dc_root = ET.SubElement(parent_el, f"{{{STIBO_NS}}}DataContainers")

    mdc = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")

    dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc.set("Type",     "DC_Barcode")
    dc.set("Analyzer", "true")

    dc_vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

    v_bc = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_bc.set("AttributeID", "AT_Barcode")
    v_bc.text = upc

    v_bt = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_bt.set("AttributeID", "AT_BarcodeType")
    v_bt.set("ID", "P")

    v_main = ET.SubElement(dc_vals, f"{{{STIBO_NS}}}Value")
    v_main.set("AttributeID", "AT_MainEANIndicator")
    v_main.set("ID", "Y")

    

# ── Generic-level placeholder attributes (written empty for now)
_GENERIC_PLACEHOLDERS = [
    "AT_PrincipalStyleDescription",
    "AT_PrincipalColorCode",
    "AT_PrincipalColorName",
    "AT_PrincipalGenderDescription",
    "AT_PrincipalAgeDescription",
    "AT_Collection1",            "AT_Collection2",
    "AT_SPUGrouping",            "AT_SPUGroupingName",
    "AT_Franchise",
    "AT_EComProductNameEN",      "AT_EComProductNameID",
    "AT_EComProductNamePH",      "AT_EComProductNameMY",
    "AT_EComProductNameVN",      "AT_EComProductNameTH",
    "AT_EComProductNameKH",
    "AT_ShortDescriptionEN",     "AT_ShortDescriptionID",
    "AT_ShortDescriptionPH",     "AT_ShortDescriptionMY",
    "AT_ShortDescriptionVN",     "AT_ShortDescriptionTH",
    "AT_ShortDescriptionKH",
    "AT_LongDescriptionEN",      "AT_LongDescriptionID",
    "AT_LongDescriptionPH",      "AT_LongDescriptionMY",
    "AT_LongDescriptionVN",      "AT_LongDescriptionTH",
    "AT_LongDescriptionKH",
    "AT_CareInstructionEN",      "AT_CareInstructionID",
    "AT_CareInstructionPH",      "AT_CareInstructionMY",
    "AT_CareInstructionVN",      "AT_CareInstructionTH",
    "AT_CareInstructionKH",
    "AT_TechnologyUsed",
    "AT_CertificateNumber",
    "AT_ProductWeight",
    "AT_ProductLengthWidthHeight",
    "AT_PackagingLength",        "AT_PackagingWidth",
    "AT_PackagingHeight",        "AT_PackagingWeight",
    "AT_ImagesSource",
    "AT_EstimatedLandedCost",
    "AT_ExchangeRate",
    "AT_FOBCurrency",
    "AT_FOB",
    "AT_OriginalPrice",
    "AT_CurrentPrice",
    "AT_RetailPriceCurrency",
    "AT_HSCode",
    "AT_Royalty",
    "AT_FreightCost",
    "AT_MarketingFee",
    "AT_SGS",
    "AT_HaddadOfficeCharge",
    "AT_HaddadOfficeChargeDeduction",
    "AT_Others",
    "AT_Commission",
    "AT_LaunchingDate",
    "AT_Createdon",
    "AT_Updatedon",
    "AT_Updatedby",
    "AT_SportsCategoryEN",
    "AT_EComAgesCategory",
    "AT_CountryOrigin",
    "AT_MainVendorIdentification",
]

# ── Variant-level placeholder attributes
_VARIANT_PLACEHOLDERS = [
    "AT_EComSize",
    "AT_CountrySize",
    "AT_FGBarcode",
    "AT_InternalBarcode",
    # "AT_MainEANIndicator",
]


def _add_generic_values(
    vals_el:    ET.Element,
    generic:    dict,
    brand_name: str,
    comp_code:  str,
    sbu:        str,
    mdd:        "MDDLoader | None" = None,
) -> None:
    """Write all Generic-level <Value> elements."""
    written: set[str] = set()

    def _w(attr_id, value="", id_val=""):
        if attr_id not in written:
            _val(vals_el, attr_id, value=value, id_val=id_val)
            written.add(attr_id)

    def _mw(attr_id, id_val):
        if attr_id not in written:
            _multival(vals_el, attr_id, id_val)
            written.add(attr_id)

    b_code = generic["brand_code"]

    # ── System / org ─────────────────────────────────────────────
    _mw("AT_SBU",         sbu)
    _mw("AT_CompanyCode", comp_code)

    # ── Brand ────────────────────────────────────────────────────
    _w("AT_Brand",      id_val=b_code)
    _w("AT_BrandGroup", id_val=LOV_BRAND.get(b_code, b_code))

    # ── Principal identifiers ────────────────────────────────────
    item_no_clean = re.sub(r"[^A-Z0-9]", "", generic["item_number"].upper())[:9]
    _w("AT_PrincipalStyleCode", item_no_clean)
    _w("AT_SAPStyleCode",       item_no_clean)
    _w("AT_InboundGenericCode", generic["generic_code"])

    # ── SAP Color (derived from color token) ─────────────────────
    color_token = generic["color_token"]
    # Look up colour ID from MDD Color Code LOV
    colour_id = ""
    if mdd:
        color_lov = mdd.lovs.get("Color Code", {})
        # Try direct color code lookup (e.g. "BK", "AS1")
        colour_id = (
            color_lov.get(generic["color_code"], "")
            or color_lov.get(generic["color_code"].upper(), "")
        )
    if not colour_id:
        colour_id = color_token
    if colour_id.isdigit():
        colour_id = colour_id.zfill(3)
    _w("AT_Color", id_val=colour_id)

    # ── Principal Color Code (from source) ───────────────────────
    # _w("AT_PrincipalColorCode", generic["color_code"])

    # ── SAP Article Category = Generic (1) ───────────────────────
    _w("AT_SAPArticleCategory", id_val="1")

    # ── Article type = License ───────────────────────────────────
    _w("AT_BYArticleType", id_val=generic.get("article_type_from_filename") or "License")

    # ── System indicators ────────────────────────────────────────
    _w("AT_BYIndicator",   id_val="Y")
    _w("AT_SAPIndicator",  id_val="N")
    _w("AT_EcomIndicator", "")

    # ── UOM ──────────────────────────────────────────────────────
    _w("AT_UOM", id_val="EA")

    # ── Season ───────────────────────────────────────────────────
    season    = generic.get("season_from_filename") or generic.get("season", "")
    sea_code  = season[:2].upper() if len(season) >= 2 else season
    year_match = re.search(r"20\d{2}", season)
    year_str   = year_match.group() if year_match else (
        f"20{season[2:]}" if len(season) > 2 else ""
    )
    _w("AT_Season", id_val=sea_code)
    _w("AT_SeasonYear", year_str)
    country_code = generic.get("country_from_filename", "")
    if country_code:
        _w("AT_Country", id_val=country_code)
    else:
        log.warning(
            "[Generic %s] AT_Country is empty — no country token found in "
            "filename or args.country_code", generic["generic_code"]
        )
    # Do not emit empty placeholder attributes.


def _add_variant_values(
    vals_el:  ET.Element,
    generic:  dict,
    variant:  dict,
    mdd:      "MDDLoader | None" = None,
) -> None:
    """Write Variant-level <Value> elements — one per size/barcode row."""

    def _w(attr_id, value="", id_val=""):
        _val(vals_el, attr_id, value=value, id_val=id_val)

    # ── Key-defining attributes (required for Stibo KEY resolution) ──
    size_lov_id = _lookup_size_from_mdd(variant["size_raw"], mdd)
    _w("AT_Size", id_val=size_lov_id)

    # ── Variant identity ─────────────────────────────────────────
    _w("AT_InboundVariantCode", variant["variant_code"])


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — CLASSIFICATIONS BLOCK
# ══════════════════════════════════════════════════════════════════════════════

def build_classifications(
    brand: str,
    brand_code: str,
    comp_code: str,
    sbu: str,
    season_code: str,
) -> ET.Element:
    """Build the <Classifications> block (season hierarchy)."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix     = season_code[:2].upper() if len(season_code) >= 2 else season_code
    sea_year_short = season_code[2:] if len(season_code) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    full_sea_code  = f"{sea_prefix}{sea_year}"
    season_id      = f"CLH_{brand_code}_{full_sea_code}"
    batches_parent = f"CLH_{brand.title().replace(' ', '')}Batches"

    season_label_map = {
        "SS": "Spring Summer", "FW": "Fall Winter",
        "AW": "Autumn Winter", "HO": "Holiday",
    }
    sea_name       = season_label_map.get(sea_prefix, sea_prefix)
    season_display = f"{brand} {sea_name} {sea_year}".strip()
    season_short   = f"{sea_prefix} {sea_year}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_parent)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — PRODUCT XML BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_product_xml(
    generic:    dict,
    brand:      str,
    brand_code: str,
    comp_code:  str,
    sbu:        str,
    season_id:  str,
    mdd:        "MDDLoader | None" = None,
) -> str:
    """
    Build the full <Product> XML fragment for ONE generic article,
    including all its <Product> Variant child elements.

    Generic structure
    -----------------
    <Product UserTypeID="PRD_GenericArticle" ParentID="PPH_A-TempSubCat">
      <KeyValue KeyID="KEY_InboundArticle">{generic_code}</KeyValue>
      <Name>{generic_desc}</Name>
      <ClassificationReference ClassificationID="CLH_{Brand}Articles"
                                Type="CPL_Merchandiser"/>
      <ClassificationReference ClassificationID="{season_id}UA"
                                Type="CPL_UnConfirmedForSeason"/>
      <Values>
        … generic-level attributes …
      </Values>

      <!-- one per size/EAN row -->
      <Product UserTypeID="PRD_VariantArticle" ParentID="{generic_code}">
        # <KeyValue KeyID="KEY_Variant">{variant_code}</KeyValue>
        <KeyValue KeyID="KEY_InboundVariant">{variant_code}</KeyValue>
        <Name>{variant_desc}</Name>
        <Values>
          … variant-level attributes …
        </Values>
      </Product>
    </Product>
    """
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_A-TempSubCat")   # Apparel division (Licensed)

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = generic["generic_code"]

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = generic["generic_desc"]

    # ── Classification references ────────────────────────────────
    brand_title = brand.title().replace(" ", "")
    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand_title}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Generic-level Values ─────────────────────────────────────
    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, generic, brand, comp_code, sbu, mdd=mdd)

    # ── Variant sub-products ─────────────────────────────────────
    # Sort by sap_size for deterministic output
    for sap_size in sorted(generic["variants"].keys()):
        variant = generic["variants"][sap_size]

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")
        # ParentID intentionally omitted — Stibo uses XML nesting to resolve parent
        # v_el.set("ParentID",   generic["generic_code"])

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        # v_kv.set("KeyID", "KEY_Variant")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant["variant_code"]

        v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
        _add_variant_values(v_vals, generic, variant,mdd=mdd)

        # ── Barcode DataContainer ─────────────────────────────
        if variant.get("upc"):
            _add_barcode_datacontainer(v_el, variant["upc"], variant["variant_code"])

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — VALIDATOR
# ══════════════════════════════════════════════════════════════════════════════

def validate_generic(generic: dict, mdd: MDDLoader) -> list[str]:
    """Check mandatory Generic-level attributes."""
    warns = []
    gcode = generic["generic_code"]

    mandatory_checks = {
        "item_number":  "AT_PrincipalStyleCode",
        "color_code":   "AT_PrincipalColorCode",
        "brand_code":   "AT_Brand",
        "generic_code": "AT_InboundGenericCode",
    }
    for field, at_id in mandatory_checks.items():
        val = generic.get(field, "")
        if not val or str(val).strip() in ("", "None", "nan"):
            warns.append(f"[{gcode}] MISSING mandatory: {at_id} (field={field})")

    return warns


def validate_variant(variant: dict, gcode: str) -> list[str]:
    """Check mandatory Variant-level attributes."""
    warns = []
    vcode = variant.get("variant_code", "?")

    mandatory_variant = {
        "sap_size": "AT_Size",
        "upc":      "AT_PrincipalBarcode",
    }
    for field, at_id in mandatory_variant.items():
        val = variant.get(field, "")
        if not val or str(val).strip() in ("", "None", "nan"):
            warns.append(f"[{gcode}/{vcode}] MISSING mandatory: {at_id} (field={field})")

    return warns


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════════════════

def run(args, auditor=None):
    """
    Main entry point — called by lambda_function.py as:
        etl_ean_source.run(args, auditor=auditor)

    args must have:
        brand       str   e.g. "New Balance"
        brand_code  str   e.g. "NEW"
        comp_code   str   e.g. "0888"
        sbu         str   e.g. "SP"
        season      str   e.g. "FW26"
        seq         int   e.g. 1
        file_type   str   e.g. "EAN Source"
        multi_mono  str   e.g. "Multi"
    """
    log.info(
        "Pipeline args → brand=%s  brand_code=%s  comp_code=%s  "
        "sbu=%s  season=%s  seq=%s  file_type=%s",
        args.brand, args.brand_code, args.comp_code,
        args.sbu, args.season, args.seq,
        getattr(args, "file_type", "EAN Source"),
    )

    def _first(d: Path, ext: str = "*.xlsx") -> Path | None:
        files = list(d.glob(ext))
        return files[0] if files else None

    mdd_f     = _first(MDD_DIR)
    naming_f  = _first(NAMING_DIR)
    os_files  = list(ORDER_DIR.glob("*.xlsx"))

    # ── Mandatory file checks ────────────────────────────────────
    for label, val in [("MDD", mdd_f), ("Order Sheet", os_files)]:
        if not val:
            log.error("No %s file found — aborting.", label)
            raise FileNotFoundError(f"No {label} file found in expected input directory.")

    mdd    = MDDLoader(mdd_f)
    naming = NamingConventionLoader(naming_f) if naming_f else None

    # ── Season classification ID ─────────────────────────────────
    sea_prefix     = args.season[:2].upper() if len(args.season) >= 2 else args.season
    sea_year_short = args.season[2:] if len(args.season) > 2 else ""
    sea_year       = f"20{sea_year_short}" if len(sea_year_short) == 2 else sea_year_short
    season_id      = f"CLH_{args.brand_code}_{sea_prefix}{sea_year}"

    all_warnings: list[str] = []

    for os_path in os_files:
        log.info("─── Processing Order Sheet: %s ───", os_path.name)
        filename_meta = _parse_metadata_from_input_filename(os_path.stem)

        loader = NBOrderSheetLoader(os_path)
        if not loader.rows:
            log.warning("[OrderSheet] Empty — skipping %s", os_path.name)
            continue

        # ── Map rows into generic/variant dict ───────────────────
        generics = map_ordersheet_rows(loader.rows, args.brand_code, args.season, mdd=mdd)
        for _g in generics.values():
            _g["article_type_from_filename"] = filename_meta.get("article_type", "")
            _g["season_from_filename"]       = filename_meta.get("season", "")
            _g["country_from_filename"]      = (
                filename_meta.get("country", "")
                or getattr(args, "country_code", "")   # ← fallback from lambda args
            )
        if not generics:
            log.warning("[OrderSheet] No valid generics produced — skipping.")
            continue

        # ── DEV LIMIT: remove or set to None to process all products ──
        # DEV_PRODUCT_LIMIT = 5
        # if DEV_PRODUCT_LIMIT:
        #     generics = dict(list(generics.items())[:DEV_PRODUCT_LIMIT])
        #     log.warning("[DEV] Product limit active — processing %d generic(s) only.", DEV_PRODUCT_LIMIT)

        # ── Validate ─────────────────────────────────────────────
        for gkey, generic in generics.items():
            all_warnings.extend(validate_generic(generic, mdd))
            for variant in generic["variants"].values():
                all_warnings.extend(validate_variant(variant, generic["generic_code"]))

        # ── Output filename: keep exactly same as input stem ─────
        # ── Output filename: derive from args, not from os_path ─────
        country_code = filename_meta.get("country", "")

        country_part = f"-{country_code}" if country_code else ""

        out_name = (
            f"{args.comp_code}-{args.sbu}-NEW BALANCE-"
            f"{getattr(args, 'file_type', 'EAN Source')}-"
            f"{getattr(args, 'multi_mono', 'Multi')}-"
            f"{args.season}"
            f"{country_part}-"
            f"{args.seq}.xml"
        )
        out_path = XML_OUT_DIR / out_name

        # ── Build Classification block ───────────────────────────
        cls_el  = build_classifications(
            args.brand, args.brand_code, args.comp_code, args.sbu, args.season,
        )
        cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
        del cls_el

        # ── Stream XML to file ───────────────────────────────────
        export_time   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        generic_count = 0
        variant_count = 0

        log.info("Writing XML → %s", out_name)
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
            for generic in generics.values():
                if not generic.get("item_number"):
                    continue
                product_xml = build_product_xml(
                    generic,
                    args.brand,
                    args.brand_code,
                    args.comp_code,
                    args.sbu,
                    season_id,
                    mdd=mdd,
                )
                product_xml = _XMLNS_RE.sub("", product_xml)
                f.write(f"    {product_xml}\n")
                del product_xml

                generic_count += 1
                variant_count += len(generic["variants"])

            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        file_kb = out_path.stat().st_size // 1024
        log.info("✓ XML written → %s  (%d KB)", out_path, file_kb)

        print("═══ ARTICLE SUMMARY (NEW BALANCE LICENSED ORDER SHEET) ════", flush=True)
        print(f"  Input rows     : {len(loader.rows)}",    flush=True)
        print(f"  Generics       : {generic_count}",       flush=True)
        print(f"  Variants       : {variant_count}",       flush=True)
        print(f"  XML file size  : {file_kb} KB",          flush=True)
        print("═══════════════════════════════════════════════════════════", flush=True)

        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    # ── Validation report ────────────────────────────────────────
    rpt_path = LOG_DIR / f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(rpt_path, "w") as f:
        f.write(f"Run: {datetime.now()}\nTotal warnings: {len(all_warnings)}\n\n")
        f.write("\n".join(all_warnings) if all_warnings else "✓ No issues found.")

    if all_warnings:
        log.warning("%d validation warnings → %s", len(all_warnings), rpt_path)
    else:
        log.info("✓ All articles passed validation → %s", rpt_path)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — CLI (local testing)
# ══════════════════════════════════════════════════════════════════════════════

def main():
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — NB Licensed Order Sheet v1.0"
    )
    p.add_argument("--brand",       default="New Balance")
    p.add_argument("--brand-code",  default="NEW")
    p.add_argument("--comp-code",   default="0888")
    p.add_argument("--sbu",         default="SP")
    p.add_argument("--season",      default="FW26")
    p.add_argument("--seq",         default=1, type=int)
    p.add_argument("--file-type",   default="EAN Source")
    p.add_argument("--multi-mono",  default="Multi")
    run(p.parse_args())


if __name__ == "__main__":
    main()
