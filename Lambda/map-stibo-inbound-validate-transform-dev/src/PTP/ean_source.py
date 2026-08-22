"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — PTP EAN Source  v1.0             ║
║   Barcode (UPC / innerpack) data → Stibo STEP XML                 ║
║   (Structure mirrors k_swiss/ean_source.py /                      ║
║    dr_marten/inline_ean_source.py)                                 ║
╚══════════════════════════════════════════════════════════════════╝

Purpose:
    Generates Variant-level STEP XML carrying barcode data ONLY.
    Companion to PTP/order_form_usd_fob.py, which sends the
    Generic-level article data (no barcodes) from a separate upload.

Source file: PTP DOES upload a separate "PTP_EAN Source" file (own S3
key / own trigger, e.g. "...PTP_EAN Source MAA Sport Inline-...xlsx" —
see PTP/lambda_function.py's "ean_source" file type, checked before
"orderform" so this upload never gets misclassified as the main order
form). Read from its own input/ean_source directory — NOT shared with
order_form_usd_fob.py's input/orderform anymore (that coupling used to
make an ean_source upload also regenerate a full, bogus duplicate
order-form XML — see PTP/lambda_function.py's module docstring).
Structurally the EAN source file is a column-for-column duplicate of
the main order form minus the SIZE SHORT/SIZE LONG columns — same
header anchors, same UPC CODE / INNERPACK BARCODE columns — so the
loader/column-detection logic below works unchanged against either
file; only the upload/trigger routing changed.

Column mapping — updated 2026-08-21 (PTP now sends TWO barcodes per
variant, mirroring reebok/ean_main.py's SKU/UPC two-DataContainer
pattern; supersedes the original single-barcode mapping):
    Principal Style Code -> UPC CODE               (join-key source)
    Size                  -> hardcoded default "000" — no longer read
                              from a source column; every variant uses
                              the same fixed 3-char size code.
    Principal Barcode     -> TWO "P" (piece) AT_BarcodeType entries,
                              each its own <DataContainer> inside the
                              same DC_Barcode MultiDataContainer:
                                Entry 1: UPC CODE column
                                         -> AT_MainEANIndicator="Y"
                                Entry 2: INNERPACK BARCODE column
                                         -> AT_MainEANIndicator="N"
                              order_form_usd_fob.py no longer writes
                              AT_PrincipalBarcode at all — barcode data
                              is owned exclusively by this file.

KEY_InboundArticle formula (per MAA's mapping, literal):
    brand_code[:3].upper() + UPC_CODE      e.g. "PTP9345164010917"

This matches order_form_usd_fob.py's own generic key exactly —
that file's AT_Generic/AT_PrincipalStyleCode mapping (sheet rows 13 and
23) also resolves to "UPC Code", not NAME, confirmed with the team
(2026-08-17). Both files independently build
brand_prefix + UPC_CODE, so barcodes from this file attach to the
correct Generic articles order_form_usd_fob.py creates.

KEY_InboundVariant formula (fixed size code, per 2026-08-21 spec):
    KEY_InboundArticle + "000"   e.g. "PTP9345164010917000"

Brand: read per-row from the BRAND column (PTP or BAHE), same as
order_form_usd_fob.py — NOT a fixed brand-code default. See
order_form_usd_fob.py's run() for the matching per-row RNA/brand fix.

XML shape mirrors dr_marten/k_swiss EAN: Generic <Values/> empty,
Variant <Values/> empty, barcodes go into a DC_Barcode DataContainer.
"""

from __future__ import annotations

import argparse
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

log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# PATHS  (own dedicated upload — input/ean_source, NOT order_form_usd_fob.py's
# input/orderform; see PTP/lambda_function.py's "ean_source" file type)
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).resolve().parent)))

EAN_DIR     = BASE_DIR / "input" / "ean_source"
XML_OUT_DIR = BASE_DIR / "output" / "xml"

for _d in (EAN_DIR, XML_OUT_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# BRAND CONSTANTS  (default only — real brand read per-row from BRAND column)
# ══════════════════════════════════════════════════════════════════════════════
BRAND_NAME_DEFAULT = "PTP"
BRAND_CODE_DEFAULT = "PTP"

# Generic/style-code brand prefix override — MUST stay in sync with
# order_form_usd_fob.py's REPORTING_BRAND_CODE_OVERRIDES (this file's
# KEY_InboundArticle has to match that file's exactly, per this module's
# own docstring). Confirmed with the team (2026-08-19): the prefix is
# always "PTP", BAHE rows included — NOT a "BH" prefix and NOT a naive
# first-3-chars slice ("BAH").
GENERIC_PREFIX_OVERRIDES = {
    "BAHE": "PTP",
}

# ══════════════════════════════════════════════════════════════════════════════
# STIBO / STEP XML CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns(?::[a-z0-9]+)?="[^"]*"')

# Same header-detection anchors as order_form_usd_fob.py's
# PTPOrderFormLoader, so this locks onto the same header row/sheet
# regardless of which of the "same file" variants gets uploaded.
HEADER_ANCHORS = {"display name", "colour", "upc code", "customer buy price", "hs code"}


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class VariantInfo:
    size_raw:        str
    size_code:       str
    upc_barcode:      str
    innerpack_barcode: str
    variant_key:      str   # KEY_InboundVariant value


@dataclass
class GenericGroup:
    generic_key:  str            # KEY_InboundArticle value
    upc_code:     str            # UPC CODE (article_code, no brand prefix)
    brand_prefix: str
    variants:     dict[str, VariantInfo] = field(default_factory=dict)
    # key = size_code (deduplication); last row wins on collision


# ══════════════════════════════════════════════════════════════════════════════
# SIZE CODE  (hardcoded — PTP EAN source no longer carries real size data;
# every variant uses the same fixed 3-char code, per 2026-08-21 spec)
# ══════════════════════════════════════════════════════════════════════════════
SIZE_CODE_DEFAULT = "000"


# ══════════════════════════════════════════════════════════════════════════════
# STRING HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _fmt_barcode(val) -> str:
    """Return a barcode value as a clean digit string (strips .0 from floats)."""
    if val is None:
        return ""
    try:
        f = float(str(val).strip())
        return str(int(f))
    except (ValueError, TypeError):
        s = str(val).strip()
        return "" if s in ("None", "nan") else s


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL READER
# ══════════════════════════════════════════════════════════════════════════════
def _find_header_row(ws) -> Optional[int]:
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        cells = [str(v).strip().lower() for v in row if v and str(v).strip()]
        matches = sum(1 for anchor in HEADER_ANCHORS if anchor in cells)
        if matches >= 2:
            return i
    return None


class EANSourceLoader:
    """Loads the PTP order form — style/brand/UPC/innerpack rows."""

    def __init__(self, path: Path):
        self.path = path
        self.headers: list[str] = []
        self.rows: list[dict] = []
        self._load()

    def _load(self) -> None:
        log.info("[PTP-EAN] Loading: %s", self.path.name)
        wb = openpyxl.load_workbook(str(self.path), data_only=True, read_only=True)

        target_sheet = None
        target_hdr_idx = None
        target_headers: list[str] = []

        for sn in wb.sheetnames:
            ws = wb[sn]
            hdr_idx = _find_header_row(ws)
            if hdr_idx is not None:
                candidate_row = list(ws.iter_rows(min_row=hdr_idx + 1, max_row=hdr_idx + 1, values_only=True))[0]
                target_sheet, target_hdr_idx = sn, hdr_idx
                target_headers = [str(v).strip() if v is not None else "" for v in candidate_row]
                break

        if not target_sheet:
            log.error("[PTP-EAN] No recognisable header row found in any sheet of %s", self.path.name)
            wb.close()
            return

        log.info("[PTP-EAN] Using sheet '%s', header row %d: %s", target_sheet, target_hdr_idx + 1, target_headers)
        self.headers = target_headers

        col_idx = {h.strip().upper(): i for i, h in enumerate(target_headers) if h}

        def _col(*aliases) -> Optional[int]:
            for alias in aliases:
                if alias.upper() in col_idx:
                    return col_idx[alias.upper()]
            return None

        name_col       = _col("NAME")
        brand_col      = _col("BRAND")
        upc_col        = _col("UPC CODE", "UPC")
        innerpack_col  = _col("INNERPACK BARCODE", "INNER PACK BARCODE", "INNERPACK_BARCODE")

        log.info(
            "[PTP-EAN] Columns — Name='%s' Brand='%s' UPC='%s' Innerpack='%s'",
            name_col, brand_col, upc_col, innerpack_col,
        )
        if upc_col is None:
            log.error("[PTP-EAN] UPC CODE column NOT FOUND — no join key / barcode available, nothing will be generated")
        if innerpack_col is None:
            log.warning("[PTP-EAN] INNERPACK BARCODE column NOT FOUND — second barcode entry will be empty")

        def _cell(row, idx):
            if idx is None or idx >= len(row):
                return None
            return row[idx]

        rows_out: list[dict] = []
        ws = wb[target_sheet]
        for row in ws.iter_rows(min_row=target_hdr_idx + 2, values_only=True):
            if row is None:
                continue

            upc_raw = _fmt_barcode(_cell(row, upc_col))
            if not upc_raw:
                continue  # no usable join key / barcode on this row

            rows_out.append({
                "name":        _s(_cell(row, name_col)),
                "brand":       _s(_cell(row, brand_col)),
                "upc":         upc_raw,
                "innerpack":   _fmt_barcode(_cell(row, innerpack_col)),
            })

        wb.close()
        self.rows = rows_out
        log.info("[PTP-EAN] Parsed %d data rows", len(self.rows))


# ══════════════════════════════════════════════════════════════════════════════
# GROUPING — rows → GenericGroup dict
# ══════════════════════════════════════════════════════════════════════════════
def group_rows(rows: list[dict], brand_code_default: str) -> dict[str, GenericGroup]:
    """
    Group flat EAN rows into GenericGroup objects.

    Generic key (KEY_InboundArticle) : brand_prefix + UPC CODE
        — per MAA's mapping, matching order_form_usd_fob.py's own
        UPC-based generic key exactly (see module docstring).
    Variant key (KEY_InboundVariant) : generic_key + SIZE_CODE_DEFAULT
        ("000", fixed — see module docstring; PTP EAN source no longer
        carries real size data).

    Brand prefix is read per-row from the BRAND column (PTP/BAHE),
    falling back to brand_code_default only when a row has no BRAND
    value — matching order_form_usd_fob.py's per-row brand handling.

    On variant_key collision within a generic (i.e. duplicate UPC rows),
    the last row encountered wins (same approach as k_swiss/dr_marten
    EAN sources).
    """
    generics: dict[str, GenericGroup] = {}

    for row in rows:
        row_brand    = (row["brand"] or brand_code_default).strip().upper()
        brand_prefix = GENERIC_PREFIX_OVERRIDES.get(row_brand, row_brand[:3])
        upc_code     = row["upc"]

        generic_key = f"{brand_prefix}{upc_code}"
        sc          = SIZE_CODE_DEFAULT
        variant_key = f"{generic_key}{sc}"

        if generic_key not in generics:
            generics[generic_key] = GenericGroup(
                generic_key=generic_key,
                upc_code=upc_code,
                brand_prefix=brand_prefix,
            )

        g = generics[generic_key]
        g.variants[sc] = VariantInfo(
            size_raw=sc,
            size_code=sc,
            upc_barcode=row["upc"],
            innerpack_barcode=row["innerpack"],
            variant_key=variant_key,
        )

    log.info(
        "[PTP-EAN] %d rows → %d generics, %d variants total",
        len(rows), len(generics), sum(len(g.variants) for g in generics.values()),
    )
    return generics


# ══════════════════════════════════════════════════════════════════════════════
# XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _add_barcode_datacontainer(parent: ET.Element, upc_barcode: str, innerpack_barcode: str = "") -> None:
    """
    Append DC_Barcode DataContainer to a Variant product element — TWO
    "P" (piece) AT_BarcodeType entries, each its own <DataContainer>
    inside the same MultiDataContainer, mirroring reebok/ean_main.py's
    SKU/UPC two-entry pattern:
      Entry 1 (UPC CODE):        AT_Barcode=upc_barcode,        AT_MainEANIndicator="Y"
      Entry 2 (INNERPACK BARCODE): AT_Barcode=innerpack_barcode, AT_MainEANIndicator="N"

        <DataContainers>
          <MultiDataContainer Type="DC_Barcode">
            <DataContainer Analyzer="true">
              <Values>
                <Value AttributeID="AT_Barcode">9345164010917</Value>
                <Value AttributeID="AT_BarcodeType" ID="P"/>
                <Value AttributeID="AT_MainEANIndicator" ID="Y"/>
              </Values>
            </DataContainer>
            <DataContainer Analyzer="true">
              <Values>
                <Value AttributeID="AT_Barcode">19345164010918</Value>
                <Value AttributeID="AT_BarcodeType" ID="P"/>
                <Value AttributeID="AT_MainEANIndicator" ID="N"/>
              </Values>
            </DataContainer>
          </MultiDataContainer>
        </DataContainers>
    """
    if not upc_barcode and not innerpack_barcode:
        return

    dc_root = ET.SubElement(parent, f"{{{STIBO_NS}}}DataContainers")
    mdc     = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
    mdc.set("Type", "DC_Barcode")

    for barcode, is_main in ((upc_barcode, "Y"), (innerpack_barcode, "N")):
        if not barcode:
            continue
        dc  = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
        dc.set("Analyzer", "true")
        dcv = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

        v1 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
        v1.set("AttributeID", "AT_Barcode")
        v1.text = barcode

        v2 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
        v2.set("AttributeID", "AT_BarcodeType")
        v2.set("ID", "P")

        v3 = ET.SubElement(dcv, f"{{{STIBO_NS}}}Value")
        v3.set("AttributeID", "AT_MainEANIndicator")
        v3.set("ID", is_main)


def _build_generic_xml(group: GenericGroup) -> str:
    """
    Build XML string for one Generic article with all its Variants.

        <Product UserTypeID="PRD_GenericArticle">
          <KeyValue KeyID="KEY_InboundArticle">PTP9345164010917</KeyValue>
          <Values/>
          <Product UserTypeID="PRD_VariantArticle">
            <KeyValue KeyID="KEY_InboundVariant">PTP9345164010917000</KeyValue>
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

        _add_barcode_datacontainer(v_el, variant.upc_barcode, variant.innerpack_barcode)

    return ET.tostring(g_el, encoding="unicode")


# ══════════════════════════════════════════════════════════════════════════════
# XML WRITER
# ══════════════════════════════════════════════════════════════════════════════
def build_xml(generics: dict[str, GenericGroup], out_path: Path) -> tuple[int, int]:
    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    generic_count = 0
    variant_count = 0

    log.info("[PTP-EAN] Writing STEP XML → %s", out_path.name)
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
            xml_str = _build_generic_xml(group)
            xml_str = _XMLNS_RE.sub("", xml_str)
            f.write(f"    {xml_str}\n")
            generic_count += 1
            variant_count += len(group.variants)

        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")

    log.info(
        "[PTP-EAN] Done — %d generics, %d variants written → %s",
        generic_count, variant_count, out_path.name,
    )
    return generic_count, variant_count


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run(args, auditor=None):
    """
    Main entry point — called by PTP/lambda_function.py for the
    "ean_source" trigger, off its own dedicated upload (see this file's
    module docstring — no longer shared with order_form_usd_fob.py).

    args attributes used:
        brand_code (str)  default "PTP" — fallback only; real brand is
                          read per-row from BRAND
        input_file  (str)  optional: explicit path to the EAN source file
    """
    brand_code_default = getattr(args, "brand_code", BRAND_CODE_DEFAULT) or BRAND_CODE_DEFAULT
    log.info("[PTP-EAN] Starting — brand_code(default)=%s", brand_code_default)

    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        ean_file = Path(input_file_arg)
    else:
        candidates = sorted(
            (p for p in EAN_DIR.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")),
            key=lambda p: p.stat().st_mtime,
        ) if EAN_DIR.exists() else []
        if not candidates:
            log.warning("[PTP-EAN] No EAN source .xlsx/.xlsm found in %s — nothing to process.", EAN_DIR)
            return [], None
        ean_file = candidates[-1]

    log.info("[PTP-EAN] Processing file: %s", ean_file.name)

    loader = EANSourceLoader(ean_file)
    if not loader.rows:
        log.warning("[PTP-EAN] No valid rows parsed from %s", ean_file.name)
        return [], None

    generics = group_rows(loader.rows, brand_code_default)
    if not generics:
        log.warning("[PTP-EAN] No valid generics after grouping.")
        return [], None

    xml_path = XML_OUT_DIR / f"{ean_file.stem}_EAN.xml"
    generic_count, variant_count = build_xml(generics, xml_path)

    print(
        f"=== PTP EAN SUMMARY ===\n"
        f"  Input file : {ean_file.name}\n"
        f"  Input rows : {len(loader.rows)}\n"
        f"  Generics   : {generic_count}\n"
        f"  Variants   : {variant_count}\n"
        f"  XML file   : {xml_path.name}",
        flush=True,
    )

    if auditor:
        try:
            auditor.set_xml_uploads([str(xml_path)])
        except AttributeError:
            pass

    return [], xml_path


# ══════════════════════════════════════════════════════════════════════════════
# CLI (for local testing)
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    ap = argparse.ArgumentParser(description="PTP EAN Source → STEP XML")
    ap.add_argument("--brand-code", default=BRAND_CODE_DEFAULT, dest="brand_code")
    ap.add_argument("--input-file", default=None,                dest="input_file",
                     help="Path to PTP EAN source Excel (optional)")
    run(ap.parse_args())
