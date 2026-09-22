"""
excel_images.py — extract images placed *in cells* from an .xlsx workbook
=========================================================================
Excel 365 "Insert → Pictures → Place in Cell" stores the picture as a rich value,
which openpyxl cannot read (the cell comes back as "#VALUE!").  This module walks
the OOXML parts directly, using only the standard library:

    xl/worksheets/sheetN.xml          <c r="F2" t="e" vm="1">          cell → value metadata (1-based)
    xl/metadata.xml                   valueMetadata bk → rc v → futureMetadata XLRICHVALUE bk → rvb i
    xl/richData/rdrichvalue.xml       rv i → field values; structure s tells which one is the image
    xl/richData/rdrichvaluestructure.xml  key "_rvRel:LocalImageIdentifier"
    xl/richData/richValueRel.xml      rel index → r:id
    xl/richData/_rels/richValueRel.xml.rels  r:id → ../media/imageN.jpeg

Usage:
    result = extract_in_cell_images(path)
    result.images[excel_row].data      # image bytes for that recap row
    result.issues                      # [{excel_row, supp_art, code, severity, message}]
"""
from __future__ import annotations

import posixpath
import re
import struct
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_RICHDATA = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
NS_RICHVALUEREL = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"

SUPPORTED_FORMATS = ("jpeg", "png", "gif", "webp")   # formats Amazon Bedrock accepts
MAX_IMAGE_BYTES = 3_750_000                           # Bedrock Converse image limit
MIN_RECOMMENDED_PX = 500

DEFAULT_SHEETS = ("LOT", "LOTTO")
DEFAULT_IMAGE_COLUMNS = ("Image", "Thumbnail Image")
DEFAULT_KEY_COLUMNS = ("Supp Art #", "Supp Art#", "Supplier Article", "Article")


@dataclass
class CellImage:
    excel_row: int
    cell: str
    supp_art: str
    media_path: str
    format: str | None
    width: int | None
    height: int | None
    data: bytes = field(repr=False)


@dataclass
class ExtractionResult:
    sheet: str
    header_row: int | None
    images: dict[int, CellImage]
    issues: list[dict]

    @property
    def errors(self) -> list[dict]:
        return [i for i in self.issues if i["severity"] == "error"]


# ── helpers ──────────────────────────────────────────────────────────────────

def _norm(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().upper()


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _read_xml(zf: zipfile.ZipFile, part: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(part))
    except KeyError:
        return None


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, dict]:
    """Relationships of *part*: {Id: {"type", "target" (resolved zip path)}}."""
    rels_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    root = _read_xml(zf, rels_part)
    if root is None:
        return {}
    base = posixpath.dirname(part)
    return {
        rel.get("Id"): {
            "type": rel.get("Type", ""),
            "target": posixpath.normpath(posixpath.join(base, rel.get("Target", "")))
            if not rel.get("Target", "").startswith("/") else rel.get("Target").lstrip("/"),
        }
        for rel in root.iter(_q(NS_PKG_REL, "Relationship"))
    }


def _split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref or "")
    return (m.group(1), int(m.group(2))) if m else ("", 0)


def detect_format(data: bytes) -> str | None:
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    return None


def image_size(data: bytes, fmt: str | None) -> tuple[int | None, int | None]:
    """Pixel dimensions read from the file header (no Pillow needed)."""
    try:
        if fmt == "png":
            return struct.unpack(">II", data[16:24])
        if fmt == "gif":
            return struct.unpack("<HH", data[6:10])
        if fmt == "jpeg":
            i = 2
            while i + 9 < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    height, width = struct.unpack(">HH", data[i + 5:i + 9])
                    return width, height
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
        if fmt == "webp":
            chunk = data[12:16]
            if chunk == b"VP8 ":
                w, h = struct.unpack("<HH", data[26:30])
                return w & 0x3FFF, h & 0x3FFF
            if chunk == b"VP8L":
                bits = int.from_bytes(data[21:25], "little")
                return (bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1
            if chunk == b"VP8X":
                return int.from_bytes(data[24:27], "little") + 1, int.from_bytes(data[27:30], "little") + 1
    except (struct.error, IndexError):
        pass
    return None, None


# ── rich value chain ─────────────────────────────────────────────────────────

def _rich_value_images(zf: zipfile.ZipFile, workbook_part: str) -> dict[int, str]:
    """{value-metadata index (1-based, as in the cell vm attribute): media zip path}."""
    by_type = {rel["type"].rsplit("/", 1)[-1].lower(): rel["target"] for rel in _rels(zf, workbook_part).values()}
    metadata_part = by_type.get("sheetmetadata", "xl/metadata.xml")
    rich_value_part = by_type.get("rdrichvalue", "xl/richData/rdrichvalue.xml")
    structure_part = by_type.get("rdrichvaluestructure", "xl/richData/rdrichvaluestructure.xml")
    rel_part = by_type.get("richvaluerel", "xl/richData/richValueRel.xml")

    metadata = _read_xml(zf, metadata_part)
    rich_values = _read_xml(zf, rich_value_part)
    structures = _read_xml(zf, structure_part)
    rel_list = _read_xml(zf, rel_part)
    if metadata is None or rich_values is None or structures is None or rel_list is None:
        return {}

    type_names = [t.get("name") for t in metadata.iter(_q(NS_MAIN, "metadataType"))]
    future = next((f for f in metadata.iter(_q(NS_MAIN, "futureMetadata")) if f.get("name") == "XLRICHVALUE"), None)
    if future is None or "XLRICHVALUE" not in type_names:
        return {}
    rich_type_index = type_names.index("XLRICHVALUE") + 1  # rc/@t is 1-based
    future_to_rv = []
    for bk in future.findall(_q(NS_MAIN, "bk")):
        rvb = bk.find(f".//{_q(NS_RICHDATA, 'rvb')}")
        future_to_rv.append(int(rvb.get("i")) if rvb is not None else None)

    image_key_position = {}
    for s_index, s in enumerate(structures.findall(_q(NS_RICHDATA, "s"))):
        keys = [k.get("n") for k in s.findall(_q(NS_RICHDATA, "k"))]
        if "_rvRel:LocalImageIdentifier" in keys:
            image_key_position[s_index] = keys.index("_rvRel:LocalImageIdentifier")
    rv_to_rel = []
    for rv in rich_values.findall(_q(NS_RICHDATA, "rv")):
        position = image_key_position.get(int(rv.get("s", "-1")))
        values = [v.text for v in rv.findall(_q(NS_RICHDATA, "v"))]
        rv_to_rel.append(int(values[position]) if position is not None and position < len(values) else None)

    rel_ids = [rel.get(_q(NS_REL, "id")) for rel in rel_list.findall(_q(NS_RICHVALUEREL, "rel"))]
    rel_targets = _rels(zf, rel_part)

    images = {}
    value_metadata = metadata.find(_q(NS_MAIN, "valueMetadata"))
    for vm_index, bk in enumerate(value_metadata.findall(_q(NS_MAIN, "bk")) if value_metadata is not None else [], 1):
        rc = bk.find(_q(NS_MAIN, "rc"))
        if rc is None or int(rc.get("t", 0)) != rich_type_index:
            continue
        future_index = int(rc.get("v", -1))
        rv_index = future_to_rv[future_index] if 0 <= future_index < len(future_to_rv) else None
        rel_index = rv_to_rel[rv_index] if rv_index is not None and rv_index < len(rv_to_rel) else None
        rel_id = rel_ids[rel_index] if rel_index is not None and rel_index < len(rel_ids) else None
        if rel_id in rel_targets:
            images[vm_index] = rel_targets[rel_id]["target"]
    return images


# ── public API ───────────────────────────────────────────────────────────────

def extract_in_cell_images(
    path: str | Path,
    sheet_names: tuple[str, ...] = DEFAULT_SHEETS,
    image_columns: tuple[str, ...] = DEFAULT_IMAGE_COLUMNS,
    key_columns: tuple[str, ...] = DEFAULT_KEY_COLUMNS,
) -> ExtractionResult:
    """Images placed in the image column of the recap sheet, keyed by Excel row number.

    Every row with a non-blank key (Supp Art #) must carry exactly one in-cell image;
    rows without one are reported as ``missing_image`` errors.
    """
    with zipfile.ZipFile(path) as zf:
        workbook_part = "xl/workbook.xml"
        workbook = _read_xml(zf, workbook_part)
        sheets = [(s.get("name"), s.get(_q(NS_REL, "id")))
                  for s in workbook.iter(_q(NS_MAIN, "sheet"))]
        wanted = {_norm(n) for n in sheet_names}
        sheet_name, sheet_rid = next(((n, r) for n, r in sheets if _norm(n) in wanted), sheets[0])
        sheet_part = _rels(zf, workbook_part)[sheet_rid]["target"]

        shared_root = _read_xml(zf, "xl/sharedStrings.xml")
        shared = ["".join(t.text or "" for t in si.iter(_q(NS_MAIN, "t")))
                  for si in shared_root.findall(_q(NS_MAIN, "si"))] if shared_root is not None else []

        vm_images = _rich_value_images(zf, workbook_part)

        rows: dict[int, dict[str, dict]] = {}
        sheet_root = _read_xml(zf, sheet_part)
        for row_position, row in enumerate(sheet_root.iter(_q(NS_MAIN, "row")), 1):
            row_number = int(row.get("r", row_position))
            cells = rows.setdefault(row_number, {})
            for col_position, cell in enumerate(row.findall(_q(NS_MAIN, "c")), 1):
                column, _ = _split_ref(cell.get("r", ""))
                cell_type = cell.get("t", "n")
                if cell_type == "inlineStr":
                    text = "".join(t.text or "" for t in cell.iter(_q(NS_MAIN, "t")))
                else:
                    v = cell.find(_q(NS_MAIN, "v"))
                    text = v.text if v is not None else ""
                    if cell_type == "s" and text:
                        text = shared[int(text)]
                cells[column or f"#{col_position}"] = {"value": text or "", "vm": cell.get("vm"),
                                                       "ref": cell.get("r", "")}

        # Header row: first of the top 10 rows that has both a key column and an image column.
        key_names, image_names = {_norm(n) for n in key_columns}, {_norm(n) for n in image_columns}
        header_row = key_col = image_col = None
        for row_number in sorted(rows)[:10]:
            by_name = {_norm(c["value"]): col for col, c in rows[row_number].items() if c["value"]}
            key_col = next((by_name[n] for n in by_name if n in key_names), None)
            image_col = next((by_name[n] for n in by_name if n in image_names), None)
            if key_col and image_col:
                header_row = row_number
                break

        issues: list[dict] = []
        images: dict[int, CellImage] = {}
        if header_row is None:
            issues.append({"excel_row": None, "supp_art": "", "code": "header_not_found", "severity": "error",
                           "message": f"no header with {key_columns[0]!r} and {image_columns[0]!r} in sheet {sheet_name!r}"})
            return ExtractionResult(sheet_name, None, images, issues)

        seen_media: dict[str, int] = {}
        for row_number in sorted(r for r in rows if r > header_row):
            cells = rows[row_number]
            supp_art = (cells.get(key_col) or {}).get("value", "").strip()
            image_cell = cells.get(image_col) or {}
            media_path = vm_images.get(int(image_cell["vm"])) if image_cell.get("vm") else None

            def issue(code, severity, message):
                issues.append({"excel_row": row_number, "supp_art": supp_art, "code": code,
                               "severity": severity, "message": message})

            if not supp_art:
                if media_path:
                    issue("image_without_article", "warning", "image found in a row without Supp Art #")
                continue
            if not media_path:
                found = image_cell.get("value", "")
                issue("missing_image", "error",
                      f"no in-cell image in {image_col}{row_number}" + (f" (cell value {found!r})" if found else ""))
                continue

            data = zf.read(media_path)
            fmt = detect_format(data)
            width, height = image_size(data, fmt)
            if fmt not in SUPPORTED_FORMATS:
                issue("unsupported_image", "error",
                      f"image format {fmt or posixpath.splitext(media_path)[1]} is not one of {', '.join(SUPPORTED_FORMATS)}")
                continue
            if len(data) > MAX_IMAGE_BYTES:
                issue("unsupported_image", "error", f"image is {len(data):,} bytes; maximum is {MAX_IMAGE_BYTES:,}")
                continue
            if width and height and max(width, height) < MIN_RECOMMENDED_PX:
                issue("low_resolution", "warning", f"image is {width}x{height}px; at least {MIN_RECOMMENDED_PX}px recommended")
            if media_path in seen_media:
                issue("shared_image", "warning", f"same image as Excel row {seen_media[media_path]}")
            seen_media.setdefault(media_path, row_number)

            images[row_number] = CellImage(row_number, image_cell.get("ref") or f"{image_col}{row_number}",
                                           supp_art, media_path, fmt, width, height, data)

        return ExtractionResult(sheet_name, header_row, images, issues)
