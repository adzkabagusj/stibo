"""
STIBO INBOUND XML GENERATOR -- Impulse EAN Update v1.0
"SUM ORDER" sheet (Sofsole linelist Excel) -> Stibo STEP XML

Purpose:
    Generates Variant-level STEP XML with barcode data.
    Companion to implus/linelist_main_source_sofsole.py (which sends
    Generic-only -- no variants, no barcodes).

VERIFIED against the real Sofsole EAN source file ("1. IPL_Article
Attributes Source_Sofsole.xlsx", sheet "SUM ORDER", barcode column "EAN")
-- sheet/column choice confirmed. Column detection stays dynamic
(header-name matched, same approach as linelist_main_source_sofsole.py's
`_find_col`) rather than fixed-index, so it tolerates minor column
reordering in future file deliveries.

MAPPED ATTRIBUTES (per query tracker -- Sofsole row):
    File          : 1. IPL_Article Attributes Source_Sofsole
    Size          : SIZE
    Principal Barcode : EAN   (single barcode, unlike Harbinger's UPC+EAN)

    Generic (PRD_GenericArticle):
        KEY_InboundArticle  = brand_code + implus_eu
        <Values/>            (intentionally empty — EAN update only)

    Variant (PRD_VariantArticle):
        KEY_InboundVariant  = KEY_InboundArticle + size_code (3-char)
        <Values/>            (intentionally empty)
        DC_Barcode DataContainer (single entry, same shape as Balega's):
            AT_Barcode           = EAN value
            AT_BarcodeType       = P  (ID-only LOV)
            AT_MainEANIndicator  = Y  (ID-only LOV)

KEY_InboundArticle formula (must match linelist_main_source_sofsole.py's
generic_key so variants attach to the same generic products):
    brand_code + implus_eu   (IMPLUS EU ITEM # / SKU column)

KEY_InboundVariant formula:
    KEY_InboundArticle + _size_code(size)

Input sheet: "SUM ORDER" (fallback: active sheet)
Column detection: dynamic, by header name (same signals as
linelist_main_source_sofsole.py):
    Category            -> (not used in EAN update XML)
    IMPLUS EU ITEM #     -> implus_eu  (style/item key)
    UPC                  -> (not used -- Sofsole barcode is EAN only)
    EAN                  -> barcode
    SIZE                 -> size -> size_code for KEY_InboundVariant
    COLOR                -> color (grouping only)

A row counts as a data row if it has a non-empty EAN **or** UPC value
(matches linelist_main_source_sofsole.py's row detection), even though
only EAN is written to AT_Barcode -- keeps banner-row filtering identical
between the two Sofsole modules.
"""

from __future__ import annotations

import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import openpyxl

# ══════════════════════════════════════════════════════════════════════════════
# STIBO / STEP XML CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z0-9]+)?="[^"]*"')

# Header signal words used to detect the header row in "SUM ORDER" sheet —
# identical set to linelist_main_source_sofsole.py so both modules agree
# on which row is the header.
_HEADER_SIGNALS = {"category", "implus eu", "implus eu item", "upc", "ean", "product description"}

# ══════════════════════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════════════════════
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("implus.ean_update_source_sofsole")


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class EANRow:
    row_num:        int
    major_category: str = ""   # -> style-grouping key component (usually "" — see load_ean_sheet)
    style_desc:     str = ""   # -> style-grouping key component (DESCRIPTION column)
    implus_eu:      str = ""
    color:          str = ""
    size:           str = ""
    ean:            str = ""


@dataclass
class VariantInfo:
    size_raw:     str
    size_code:    str
    ean:          str
    variant_key:  str   # KEY_InboundVariant value


@dataclass
class GenericGroup:
    generic_key:  str           # KEY_InboundArticle value
    implus_eu:    str
    variants:     dict[str, VariantInfo] = field(default_factory=dict)
    # key = size_code (deduplication); last EAN wins on collision


# ══════════════════════════════════════════════════════════════════════════════
# SIZE CODE FORMATTER (identical to ean_update_source_balega.py)
# ══════════════════════════════════════════════════════════════════════════════
_SIZE_OVERRIDES: dict[str, str] = {
    "ONE SIZE": "ONS",
    "ONE-SIZE": "ONS",
    "OS":       "ONS",
    "ONESIZE":  "ONS",
    "FREE":     "ONS",
    "FREE SIZE": "ONS",
    "NS":       "NSZ",
    "NO SIZE":  "NSZ",
}

# Word-size aliases — hardcoded per user direction (2026-08-17): spelled-out
# sizes on Implus EAN files (Small/Medium/Large, etc.) collapse to their
# standard letter abbreviation BEFORE the normal padding rules run, e.g.
# "Small" -> "S" -> "00S" (same padding a literal "S" already gets),
# "X-Large" -> "XL" -> "0XL", "XX-Large" -> "XXL" (already 3 chars).
_SIZE_WORD_ALIASES: dict[str, str] = {
    "SMALL":              "S",
    "MEDIUM":             "M",
    "LARGE":              "L",
    "X-SMALL":            "XS",
    "XSMALL":             "XS",
    "EXTRA SMALL":        "XS",
    "EXTRA-SMALL":        "XS",
    "X-LARGE":            "XL",
    "XLARGE":             "XL",
    "EXTRA LARGE":        "XL",
    "EXTRA-LARGE":        "XL",
    "XX-LARGE":           "XXL",
    "XXLARGE":            "XXL",
    "2X LARGE":           "XXL",
    "2X-LARGE":           "XXL",
    "2XLARGE":            "XXL",
    "DOUBLE EXTRA LARGE": "XXL",
}


def _size_code(size_val: str) -> str:
    """
    Convert a raw size string to a 3-char MAA size code for use in
    KEY_InboundVariant.

    Priority:
      0. Word-size alias (Small/Medium/Large/X-Small/X-Large/XX-Large, etc.)
         → standard letter abbreviation, THEN the normal rules below
         (e.g. "Small" → "S" → "00S"; "X-Large" → "XL" → "0XL")
      1. Override table (ONE SIZE → ONS, etc.)
      2. Pure integer → zero-padded to 3 digits (e.g. "9" → "009")
      3. Half-size decimal (e.g. "9.5") → "09H"
      4. Other decimal → leading integer zero-padded (e.g. "12.0" → "012")
      5. 1-char alpha → "00X"  (e.g. "M" → "00M")
      6. 2-char alpha → "0XY"  (e.g. "XL" → "0XL")
      7. 3+ char alpha/alnum → first 3 chars uppercased (e.g. "X-Large" → "XLA")
    """
    if not size_val:
        return "MSC"
    s = str(size_val).strip().upper()
    if not s:
        return "MSC"

    # 0. Word-size alias — spelled-out sizes collapse to their letter form
    # before anything else runs.
    s = _SIZE_WORD_ALIASES.get(s, s)

    # 1. Override table
    if s in _SIZE_OVERRIDES:
        return _SIZE_OVERRIDES[s]

    # 2. Pure integer
    if re.match(r"^\d+$", s):
        try:
            return str(int(s)).zfill(3)[:3]
        except ValueError:
            pass

    # 3. Half-size (e.g. 9.5)
    m = re.match(r"^(\d+)\.5$", s)
    if m:
        return str(int(m.group(1))).zfill(2)[:2] + "H"

    # 4. Other decimal (e.g. 9.0, 12.33)
    m2 = re.match(r"^(\d+)\.\d+$", s)
    if m2:
        return str(int(m2.group(1))).zfill(3)[:3]

    # 5-7. Alpha / alnum
    alnum = re.sub(r"[^A-Z0-9]", "", s)
    if not alnum:
        return "MSC"
    if len(alnum) == 1:
        return "00" + alnum
    if len(alnum) == 2:
        return "0" + alnum
    return alnum[:3]


# ══════════════════════════════════════════════════════════════════════════════
# STRING / BARCODE HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _fmt_barcode(val) -> str:
    """
    Return barcode as a clean string.

    Only run the float→int cleanup (stripping a trailing ".0") on actual
    numeric cell values; string cells are returned as-is so a
    text-formatted barcode with significant leading zeros (seen on
    Harbinger's UPC column) survives untouched. See
    ean_update_source_harbinger.py for the bug this guards against.
    """
    if val is None:
        return ""
    if isinstance(val, str):
        s = val.strip()
        return "" if s in ("None", "nan", "NaT") else s
    try:
        f = float(val)
        return str(int(f))
    except (ValueError, TypeError):
        s = str(val).strip()
        return "" if s in ("None", "nan") else s


# ══════════════════════════════════════════════════════════════════════════════
# DIRECTORY HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _get_dirs():
    base     = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))
    in_dir   = base / "input"
    ll_dir   = in_dir / "linelist"
    out_dir  = base / "output"
    xml_dir  = out_dir / "xml"
    log_dir  = out_dir / "logs"
    for d in (in_dir, ll_dir, out_dir, xml_dir, log_dir):
        d.mkdir(parents=True, exist_ok=True)
    return base, in_dir, ll_dir, out_dir, xml_dir, log_dir


def _find_input_files() -> list[Path]:
    _, _, ll_dir, *_ = _get_dirs()
    files: list[Path] = []
    if ll_dir.exists():
        files.extend(
            p for p in ll_dir.iterdir()
            if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
        )
    return sorted(files)


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN FINDER (identical to linelist_main_source_sofsole.py)
# ══════════════════════════════════════════════════════════════════════════════
def _find_col(header: list[str], *candidates: str) -> Optional[int]:
    """
    Return the 0-based index of the first header cell that matches any
    candidate (case-insensitive, stripped). Returns None if none match.
    """
    normalised = [str(h).strip().lower() for h in header]
    for candidate in candidates:
        target = candidate.strip().lower()
        for i, h in enumerate(normalised):
            if h == target or target in h:
                return i
    return None


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL READER  (same sheet/header detection as linelist_main_source_sofsole.py)
# ══════════════════════════════════════════════════════════════════════════════
def load_ean_sheet(path: Path) -> list[EANRow]:
    """
    Read EAN / barcode rows from the Sofsole linelist Excel.

    Sheet: "SUM ORDER" (fallback: active sheet).
    Header detection: first row (within the first 30) whose cells match
    at least 2 of _HEADER_SIGNALS — identical to the linelist parser.
    Columns are then resolved dynamically by header name.
    """
    log.info("[EAN] Loading: %s", path.name)
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True)
    except Exception as exc:
        log.error("[EAN] Cannot open file: %s", exc)
        return []

    ws = wb["SUM ORDER"] if "SUM ORDER" in wb.sheetnames else wb.active
    log.info("[EAN] Using sheet: '%s'", ws.title)

    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not all_rows:
        log.warning("[EAN] Sheet is empty")
        return []

    header_row_idx = -1
    for i, row in enumerate(all_rows[:30]):
        if not row:
            continue
        cell_vals = {str(v).strip().lower() for v in row if v is not None}
        matches = sum(1 for sig in _HEADER_SIGNALS if any(sig in cv for cv in cell_vals))
        if matches >= 2:
            header_row_idx = i
            break

    if header_row_idx == -1:
        log.warning("[EAN] Could not detect header row — defaulting to row index 19 (row 20)")
        header_row_idx = 19

    header = [str(h).strip() if h is not None else "" for h in all_rows[header_row_idx]]
    log.info("[EAN] Header row at index %d: %s", header_row_idx, header[:20])

    c_category   = _find_col(header, "Category", "Major Category", "CATEGORY")
    c_implus_eu  = _find_col(header, "IMPLUS EU ITEM #", "IMPLUS EU ITEM", "IMPLUS EU", "SKU", "ITEM #", "ITEM NO")
    c_upc        = _find_col(header, "UPC", "UPC CODE")
    c_ean        = _find_col(header, "EAN", "EAN CODE", "EAN/UPC")
    c_style_desc = _find_col(header, "DESCRIPTION", "PRODUCT DESCRIPTION", "STYLE DESCRIPTION", "STYLE NAME", "PRODUCE NAME")
    c_size       = _find_col(header, "SIZE", "PRINCIPAL SIZE")
    c_color      = _find_col(header, "COLOR", "COLOUR", "COLOR NAME")

    log.info(
        "[EAN] Column indices → category=%s  implus_eu=%s  upc=%s  ean=%s  style_desc=%s  size=%s  color=%s",
        c_category, c_implus_eu, c_upc, c_ean, c_style_desc, c_size, c_color,
    )

    def _cell(row, idx) -> str:
        if idx is None or idx >= len(row) or row[idx] is None:
            return ""
        s = str(row[idx]).strip()
        return "" if s in ("None", "nan") else s

    rows: list[EANRow] = []
    current_l1 = ""
    for row_idx, row in enumerate(all_rows[header_row_idx + 1:], start=header_row_idx + 2):
        if not row or all(v is None for v in row):
            continue  # completely blank row

        cat_val = _cell(row, c_category)
        b_val   = _cell(row, c_implus_eu)   # used both as SKU and sub-category banner

        if "TOTAL" in cat_val.upper() or "TOTAL" in b_val.upper():
            continue

        ean_val = _fmt_barcode(row[c_ean]) if c_ean is not None and c_ean < len(row) else ""
        upc_val = _fmt_barcode(row[c_upc]) if c_upc is not None and c_upc < len(row) else ""
        is_data_row = bool(ean_val) or bool(upc_val)
        if not is_data_row:
            # Banner / section-header row (e.g. "INSOLES", "PERFORM") —
            # update the L1 tracker exactly like linelist_main_source_sofsole.py
            # so the style-grouping key below stays in sync with it.
            if cat_val and cat_val != current_l1:
                current_l1 = cat_val
            continue

        implus_eu = b_val
        if not implus_eu:
            continue

        if not ean_val:
            # Tracker mapping is EAN-only for Sofsole -- rows with a UPC
            # but no EAN have nothing to write to AT_Barcode.
            log.debug("[EAN] Row %d has UPC but no EAN — skipped (Sofsole barcode = EAN only)", row_idx)
            continue

        style_desc = _cell(row, c_style_desc) or b_val

        rows.append(EANRow(
            row_num=row_idx,
            major_category=current_l1,
            style_desc=style_desc,
            implus_eu=implus_eu,
            color=_cell(row, c_color),
            size=_cell(row, c_size),
            ean=ean_val,
        ))

    log.info("[EAN] Parsed %d data rows", len(rows))
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# GROUPING — rows → GenericGroup dict
# ══════════════════════════════════════════════════════════════════════════════
def group_rows(rows: list[EANRow], brand_code: str) -> dict[str, GenericGroup]:
    """
    Group EANRow list into GenericGroup objects.

    IMPORTANT — this must mirror linelist_main_source_sofsole.py's
    group_generics() *exactly*, not the simpler "1 row = 1 generic"
    scheme ean_update_source_balega.py / _harbinger.py use. On the real
    Sofsole "SUM ORDER" sheet (no Category/Color columns), several rows
    with DIFFERENT IMPLUS EU ITEM # values but the SAME DESCRIPTION
    (e.g. "Airr" @ sizes 36-38/39-41/42-44/45-46, four distinct SKUs)
    get merged by the linelist parser into ONE generic article, keyed off
    only the FIRST row's IMPLUS EU ITEM #. If this script instead gave
    every row its own generic key, the barcode variants it writes would
    attach to KEY_InboundArticle values the linelist run never created.

    Style-grouping key: (major_category.upper() | style_desc.upper())
    Color sub-key:      COLOR value (or "NC" if absent)
    Generic key:        brand_code + <first row's implus_eu for that
                         style_key/color combination>
    Variant key:         generic_key + size_code (3-char, from THIS row's
                         own size — sizes stay per-row even though the
                         generic key doesn't)

    On size_code collision within a generic the last row encountered wins
    (same approach as ean_update_source_balega.py / _harbinger.py).
    """
    style_groups: dict[str, dict[str, GenericGroup]] = {}

    for row in rows:
        style_desc = row.style_desc or f"UNKNOWN_ROW_{row.row_num}"
        style_key  = f"{row.major_category.upper()}|{style_desc.upper()}"
        color      = row.color or "NC"

        color_map = style_groups.setdefault(style_key, {})
        if color not in color_map:
            color_map[color] = GenericGroup(
                generic_key=f"{brand_code}{row.implus_eu}",
                implus_eu=row.implus_eu,
            )

        g  = color_map[color]
        sc = _size_code(row.size)
        g.variants[sc] = VariantInfo(
            size_raw=row.size,
            size_code=sc,
            ean=row.ean,
            variant_key=f"{g.generic_key}{sc}",
        )

    generics: dict[str, GenericGroup] = {}
    for color_map in style_groups.values():
        for g in color_map.values():
            generics[g.generic_key] = g

    log.info(
        "[Group] %d rows → %d style keys → %d generics, %d variants total",
        len(rows),
        len(style_groups),
        len(generics),
        sum(len(g.variants) for g in generics.values()),
    )
    return generics


# ══════════════════════════════════════════════════════════════════════════════
# XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _add_barcode_datacontainer(parent: ET.Element, ean: str) -> None:
    """
    Append DC_Barcode DataContainer to a Variant product element.
    Single barcode entry (EAN) — Sofsole's tracker mapping is EAN-only,
    unlike Harbinger's two-barcode (UPC + EAN) requirement.

        <DataContainers>
          <MultiDataContainer Type="DC_Barcode">
            <DataContainer Analyzer="true">
              <Values>
                <Value AttributeID="AT_Barcode">3700006360159</Value>
                <Value AttributeID="AT_BarcodeType" ID="P"/>
                <Value AttributeID="AT_MainEANIndicator" ID="Y"/>
              </Values>
            </DataContainer>
          </MultiDataContainer>
        </DataContainers>
    """
    if not ean:
        return

    dc_root = ET.SubElement(parent, f"{{{STIBO_NS}}}DataContainers")
    mdc     = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")
    dc      = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    dc.set("Analyzer", "true")
    dcv     = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

    v1 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v1.set("AttributeID", "AT_Barcode")
    v1.text = ean

    v2 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v2.set("AttributeID", "AT_BarcodeType")
    v2.set("ID", "P")

    v3 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
    v3.set("AttributeID", "AT_MainEANIndicator")
    v3.set("ID", "Y")


def _build_generic_xml(group: GenericGroup) -> str:
    """
    Build XML string for one Generic article with all its Variants.

        <Product UserTypeID="PRD_GenericArticle">
          <KeyValue KeyID="KEY_InboundArticle">IPL360159</KeyValue>
          <Values/>
          <Product UserTypeID="PRD_VariantArticle">
            <KeyValue KeyID="KEY_InboundVariant">IPL360159SMA</KeyValue>
            <Values/>
            <DataContainers>...</DataContainers>
          </Product>
          ...
        </Product>
    """
    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = group.generic_key

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")  # intentionally empty

    for variant in group.variants.values():
        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant.variant_key

        ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")  # intentionally empty

        _add_barcode_datacontainer(v_el, variant.ean)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════════════════
# XML WRITER
# ══════════════════════════════════════════════════════════════════════════════
def build_xml(generics: dict[str, GenericGroup], out_path: Path) -> None:
    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = 0
    variant_count = 0

    # ══════════════════════════════════════════════════════════════════════
    # TEST LIMITER — Set TEST_MODE = False for production
    # ══════════════════════════════════════════════════════════════════════
    TEST_MODE = True
    PRODUCT_LIMIT = 2
    VARIANT_LIMIT = 3

    log.info("[EAN] Writing STEP XML → %s", out_path.name)
    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(
            f'<STEP-ProductInformation '
            f'xmlns="{STIBO_NS}" '
            f'xmlns:xsi="{STIBO_XSI}" '
            f'xsi:schemaLocation="{STIBO_SCHEMA}" '
            f'ExportTime="{export_time}" '
            f'ExportContext="Context1" '
            f'WorkspaceID="Main" '
            f'UseContextLocale="false">\n'
        )
        f.write("  <Products>\n")

        for group in generics.values():
            if not group.variants:
                continue
            if TEST_MODE:
                if generic_count >= PRODUCT_LIMIT:
                    break
                if variant_count + len(group.variants) > VARIANT_LIMIT:
                    continue  # would overflow the variant budget — try the next group
            xml_str = _build_generic_xml(group)
            xml_str = _XMLNS_RE.sub("", xml_str)
            f.write(f"    {xml_str}\n")
            generic_count += 1
            variant_count += len(group.variants)

        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")

    if TEST_MODE:
        log.info("[XML] TEST MODE: limited to %d product(s), %d variant(s)", PRODUCT_LIMIT, VARIANT_LIMIT)

    log.info(
        "[EAN] Done — %d generics, %d variants written → %s",
        generic_count, variant_count, out_path.name,
    )


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run(args, auditor=None) -> tuple[list[dict], Optional[Path]]:
    """
    Main entry point — called by lambda_function.py.

    args attributes used:
        brand_code   (str)  default "IPL"
        input_file   (str)  optional: explicit path to linelist Excel
    """
    brand_code = getattr(args, "brand_code", "IPL") or "IPL"
    log.info("[EAN-Update] Starting — brand_code=%s", brand_code)

    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        file_path = Path(input_file_arg)
    else:
        files = _find_input_files()
        if not files:
            log.warning("[EAN-Update] No linelist .xlsx/.xlsm found — nothing to process.")
            return [], None
        file_path = max(files, key=lambda p: p.stat().st_mtime)

    log.info("[EAN-Update] Processing file: %s", file_path.name)

    rows = load_ean_sheet(file_path)
    if not rows:
        log.warning("[EAN-Update] No valid rows parsed from %s", file_path.name)
        return [], None

    generics = group_rows(rows, brand_code)
    if not generics:
        log.warning("[EAN-Update] No valid generics after grouping.")
        return [], None

    _, _, _, _, xml_dir, _ = _get_dirs()
    stem     = file_path.stem
    xml_path = xml_dir / f"{stem}_EAN_Update.xml"

    build_xml(generics, xml_path)

    if auditor:
        auditor.set_xml_uploads([str(xml_path)])

    return [], xml_path


# ══════════════════════════════════════════════════════════════════════════════
# CLI (for local testing)
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Sofsole EAN Update → STEP XML")
    ap.add_argument("--brand-code",  default="IPL", dest="brand_code")
    ap.add_argument("--input-file",  default=None,  dest="input_file",
                    help="Path to Sofsole linelist Excel (optional)")
    run(ap.parse_args())
