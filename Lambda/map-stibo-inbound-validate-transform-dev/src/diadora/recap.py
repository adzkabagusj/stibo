"""
STIBO inbound XML generator for Diadora Recap files.

The Recap flow is intentionally metadata-driven:
  * Recap source columns are discovered from the workbook header row.
  * Source-to-attribute mappings are discovered from the DIADORA tab in the
    Attribute List workbook.
  * STEP AttributeIDs and LOV metadata are resolved from the MDD Core
    Attributes tab and LOV sheets.

Only direct Recap mappings are emitted. Derived, manual, AI/image-based, and
formula-style mappings are logged and skipped.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl


BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR = BASE_DIR / "input"
RECAP_DIR = INPUT_DIR / "recap"
MDD_DIR = INPUT_DIR / "mdd"
ATTR_DIR = INPUT_DIR / "attributes"

OUTPUT_DIR = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR = OUTPUT_DIR / "logs"

for d in [RECAP_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"recap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.FileHandler(log_path), logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


STIBO_NS = "http://www.stibosystems.com/step"
STIBO_XSI = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


DIRECT_ALIAS_TO_RECAP_HEADER = {
    "supplier article": "Supp Art #",
    "supplier art": "Supp Art #",
    "supplier article number": "Supp Art #",
    "article": "Supp Art #",
    "color code": "Color Code",
    "size": "Size Range",
    "sizes": "Size Range",
    "fob": "FOB Price",
    "fob price": "FOB Price",
}

ATTRIBUTE_TO_RECAP_HEADER = {
    "fob": "FOB Price",
}

DERIVED_MAPPING_MARKERS = (
    "ai ",
    "ai-",
    "ai might",
    "based on",
    "defined from images",
    "defined by images",
    "from images",
    "image",
    "manual input",
    "manual fill",
    "mapping based",
    "mapping to",
    "mapping between",
    "by logic",
    "formula",
    "translation",
)

DERIVED_ATTRIBUTE_NAMES = {
    "sap age",
    "sap gender",
    "by age",
    "by gender",
    "sap size",
    "sap color",
}

# Diadora Recap Category → Sports Category EN mapping
DIA_RECAP_CATEGORY_TO_SPORTS_CAT: dict[str, str] = {
    "RUNNING":            "Running",
    "RUNNING SSE":        "Running",
    "TRAIL":              "Running",
    "SANDALS":            "Lifestyle / Casual",
    "CASUAL":             "Lifestyle / Casual",
    "LIFESTYLE":          "Lifestyle / Casual",
    "KIDS":               "Lifestyle / Casual",
    "FUTSAL":             "Soccer",
    "FITNESS":            "Fitness / Training",
    "ACT RUNNING":        "Running",
    "ACT SOCCER":         "Soccer",
    "ACT TENNIS":         "Tennis / Padel",
    "LIFESTYLE SPORTSWEAR": "Lifestyle / Casual",
}

LOV_AGE = {
    "AD": "Adults", "CH": "Children", "IN": "Infant",
    "AA": "All Ages", "JR": "Junior", "K": "Kids",
}

LOV_BY_AGE = {
    "ADULT": "Adult", "ADULTS": "Adult", "AD": "Adult",
    "JUNIOR": "Junior", "CH": "Children",
    "CHILDREN": "Children", "CHILD": "Children",
    "KIDS": "Kids", "K": "Kids",
    "ALL AGES": "All Ages", "AA": "All Ages",
}

LOV_GENDER = {"M": "Male", "F": "Female", "U": "Unisex"}

DIA_GENDER_MAP = {
    "MALE": "Male", "FEMALE": "Female", "UNISEX": "Unisex",
    "BOYS": "Male", "GIRLS": "Female", "MEN": "Male", "WOMEN": "Female"
}


RECAP_HEADER_HINTS = {
    "brand",
    "season",
    "supplier",
    "supp art #",
    "color",
    "color code",
    "artco",
    "category",
    "outsole material",
    "upper material",
    "gender",
    "size range",
    "article type",
    "fob price",
    "outsole",
    "etd date",
    "eta date",
    "sample status",
    "remarks",
    "proto comment",
    "revise sample",
    "pps comment",
}


def _norm(value) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def _compact_norm(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return "" if text.lower() in {"none", "nan"} else text



def _fmt_date(v) -> str:
    from datetime import datetime
    if isinstance(v, datetime):
        return v.strftime('%d-%m-%Y')
    if hasattr(v, 'strftime'):
        return v.strftime('%d-%m-%Y')
    raw = _clean_text(v)
    if not raw:
        return raw
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(raw, fmt).strftime('%d-%m-%Y')
        except (ValueError, TypeError):
            pass
    return raw

def _parse_metadata_from_input_filename(stem: str) -> dict[str, str]:
    """Parse season/article type/country tokens from the source filename."""
    parts = re.split(r"\s*-\s*", stem)
    season = ""
    season_idx = None
    for idx, part in enumerate(parts):
        token = part.strip()
        if re.match(r"^[A-Z]{2}\d{2,4}$", token, re.IGNORECASE):
            season = token.upper()
            season_idx = idx
            break

    file_type_token = parts[3].strip() if len(parts) > 3 else ""
    article_type = ""
    for candidate in ("Inline", "License", "SSE"):
        if candidate.lower() in file_type_token.lower():
            article_type = candidate
            break

    country = ""
    if season_idx is not None:
        for part in parts[season_idx + 1:]:
            token = part.strip()
            if re.match(r"^[A-Z]{2,3}$", token, re.IGNORECASE) and not token.isdigit():
                country = token.upper()
                break

        if not country:
            for idx, part in enumerate(parts):
                if idx in (0, 1, 2, 3, 4, season_idx):
                    continue
                token = part.strip()
                if re.match(r"^[A-Z]{2,3}$", token, re.IGNORECASE) and not token.isdigit():
                    country = token.upper()
                    break

    return {"season": season, "article_type": article_type, "country": country}


def _is_blank(value) -> bool:
    return _clean_text(value) == ""


def _split_attr_ids(raw_id: str) -> list[str]:
    ids = [p.strip() for p in re.split(r"[\s,;/]+", raw_id or "") if p.strip()]
    return [p for p in ids if p.startswith("AT_")]


def _first_file(directory: Path, pattern: str = "*.xlsx") -> Path | None:
    files = sorted(directory.glob(pattern))
    return files[0] if files else None


def _find_sheet_name(wb, desired: str) -> str:
    desired_norm = _norm(desired)
    for sheet in wb.sheetnames:
        if _norm(sheet) == desired_norm:
            return sheet
    for sheet in wb.sheetnames:
        if desired_norm in _norm(sheet):
            return sheet
    raise ValueError(f"Cannot find sheet matching '{desired}' in {wb.sheetnames}")


def _row_values(row) -> list:
    return [cell.value for cell in row]


def _find_header_row(rows: list[list], required_terms: set[str]) -> tuple[int, dict[str, int]]:
    best_idx = -1
    best_score = 0
    best_columns: dict[str, int] = {}

    for idx, row in enumerate(rows):
        columns: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            text = _clean_text(value)
            if text:
                columns[text] = col_idx
        normalized = {_norm(h) for h in columns}
        score = len(normalized & {_norm(t) for t in required_terms})
        if score > best_score:
            best_idx = idx
            best_score = score
            best_columns = columns

    if best_idx < 0 or best_score == 0:
        raise ValueError("Cannot locate a Recap header row")

    return best_idx, best_columns


def _column_lookup(columns: dict[str, int]) -> dict[str, tuple[str, int]]:
    return {_norm(name): (name, idx) for name, idx in columns.items() if _norm(name)}


@dataclass
class AttributeMeta:
    name: str
    attr_ids: list[str]
    validation: str
    multi_valued: str
    lov_name: str

    @property
    def attr_id(self) -> str:
        return self.attr_ids[0] if self.attr_ids else ""

    @property
    def is_lov(self) -> bool:
        return "lov" in _norm(self.validation) or "list of values" in _norm(self.validation)

    @property
    def is_multi(self) -> bool:
        return _norm(self.multi_valued).startswith("yes")


class MDDLoader:
    def __init__(self, path: Path):
        self.path = path
        self.attributes_by_name: dict[str, AttributeMeta] = {}
        self.lovs: dict[str, dict[str, str]] = {}
        self._load()

    def _load(self) -> None:
        log.info("[MDD] Loading %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            ws = wb[_find_sheet_name(wb, "Core Attributes")]
            rows = [_row_values(row) for row in ws.iter_rows()]
            header_idx = self._find_core_header_idx(rows)
            col = self._header_map(rows[header_idx])

            for row in rows[header_idx + 1:]:
                name = self._get(row, col, "PIM Attribute Name")
                raw_id = self._get(row, col, "PIM Attribute ID")
                if not name or not raw_id:
                    continue
                attr_ids = _split_attr_ids(raw_id)
                if not attr_ids:
                    continue
                meta = AttributeMeta(
                    name=name,
                    attr_ids=attr_ids,
                    validation=self._get(row, col, "Validation Base Type"),
                    multi_valued=self._get(row, col, "Multi Valued"),
                    lov_name=self._get(row, col, "Name of LOV") or "",
                )
                self.attributes_by_name[_norm(name)] = meta

            self._load_lovs(wb)
        finally:
            wb.close()
        log.info("[MDD] Loaded %d attributes and %d LOV groups", len(self.attributes_by_name), len(self.lovs))

    @staticmethod
    def _find_core_header_idx(rows: list[list]) -> int:
        for idx, row in enumerate(rows[:20]):
            normalized = {_norm(v) for v in row if v}
            if "pim attribute id" in normalized and any(v.startswith("pim attribute name") for v in normalized):
                return idx
        raise ValueError("Cannot locate MDD Core Attributes header row")

    @staticmethod
    def _header_map(header: list) -> dict[str, int]:
        out: dict[str, int] = {}
        for idx, value in enumerate(header):
            text = _clean_text(value)
            if not text:
                continue
            out[_norm(text)] = idx
            out[_norm(text.split("(")[0])] = idx
            out[_norm(text.split("\n")[0])] = idx
        return out

    @staticmethod
    def _get(row: list, col: dict[str, int], key: str) -> str:
        idx = col.get(_norm(key))
        if idx is None:
            matches = [i for name, i in col.items() if _norm(key) in name]
            idx = matches[0] if matches else None
        if idx is None or idx >= len(row):
            return ""
        return _clean_text(row[idx])

    def resolve_attribute(self, attribute_name: str) -> AttributeMeta | None:
        key = _norm(attribute_name)
        if key in self.attributes_by_name:
            return self.attributes_by_name[key]

        simplified = self._simplified_attr_name(attribute_name)
        if simplified in self.attributes_by_name:
            return self.attributes_by_name[simplified]

        attr_words = set(simplified.split())
        candidates = []
        for name_key, meta in self.attributes_by_name.items():
            name_words = set(self._simplified_attr_name(name_key).split())
            if attr_words and (attr_words <= name_words or name_words <= attr_words):
                candidates.append(meta)

        unique = {candidate.attr_id: candidate for candidate in candidates if candidate.attr_id}
        if len(unique) == 1:
            return next(iter(unique.values()))
        if candidates:
            log.warning(
                "[MDD] Ambiguous Attribute List name '%s' matched %s",
                attribute_name,
                [c.name for c in candidates[:5]],
            )
        return None

    @staticmethod
    def _simplified_attr_name(name: str) -> str:
        words = [
            w for w in _norm(name).split()
            if w not in {"code", "description", "id", "en"}
        ]
        return " ".join(words)

    def _load_lovs(self, wb) -> None:
        if "Simple LOVs" in wb.sheetnames:
            ws = wb["Simple LOVs"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                lov_name = _clean_text(row[0] if len(row) > 0 else "")
                display = _clean_text(row[2] if len(row) > 2 else "")
                value_id = _clean_text(row[3] if len(row) > 3 else "")
                if lov_name and display:
                    self._add_lov_value(lov_name, display, value_id or display)

        for sheet_name in wb.sheetnames:
            if "lov" not in _norm(sheet_name):
                continue
            if sheet_name == "Simple LOVs":
                continue
            ws = wb[sheet_name]
            lov_name = re.sub(r"\s*lov\s*$", "", sheet_name, flags=re.I).strip()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 2:
                    continue
                code = _clean_text(row[0])
                display = _clean_text(row[1])
                if display:
                    self._add_lov_value(lov_name, display, code or display)
                    if code:
                        self._add_lov_value(lov_name, code, code)

    def _add_lov_value(self, lov_name: str, display: str, value_id: str) -> None:
        lov_key = _norm(lov_name)
        value_id = value_id.zfill(3) if value_id.isdigit() else value_id
        values = self.lovs.setdefault(lov_key, {})
        values[_norm(display)] = value_id
        values[_compact_norm(display)] = value_id
        values[_norm(value_id)] = value_id
        values[_compact_norm(value_id)] = value_id

    def resolve_lov_id(self, meta: AttributeMeta, raw_value: str) -> str:
        value = _clean_text(raw_value)
        if not value:
            return ""

        lov_keys = [_norm(meta.lov_name), _norm(meta.name)]
        for key in list(lov_keys):
            if key.endswith(" lov"):
                lov_keys.append(key[:-4])

        for lov_key in lov_keys:
            lov = self.lovs.get(lov_key)
            if not lov:
                continue
            for lookup in (_norm(value), _compact_norm(value)):
                if lookup in lov:
                    return lov[lookup]
            prefix = re.match(r"^([A-Za-z]{2,3})\d{2,4}$", value)
            if prefix:
                code = prefix.group(1)
                for lookup in (_norm(code), _compact_norm(code)):
                    if lookup in lov:
                        return lov[lookup]

        log.warning(
            "[LOV] No LOV ID found for %s=%r; using raw value as ID",
            meta.attr_id,
            value,
        )
        return value


@dataclass
class MappingEntry:
    attribute_name: str
    attribute_id: str
    source_header: str
    source_col_idx: int
    meta: AttributeMeta
    reason: str


class AttributeListLoader:
    def __init__(self, path: Path, mdd: MDDLoader, brand: str = "DIADORA"):
        self.path = path
        self.mdd = mdd
        self.brand = brand

    def build_recap_mappings(self, recap_columns: dict[str, int]) -> list[MappingEntry]:
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet_name = self._find_brand_sheet(wb)
            ws = wb[sheet_name]
            rows = [_row_values(row) for row in ws.iter_rows()]
        finally:
            wb.close()

        header_idx = self._find_attr_header_idx(rows)
        header = rows[header_idx]
        col = {str(v).strip(): idx for idx, v in enumerate(header) if _clean_text(v)}
        licensed_cols = self._licensed_regular_columns(rows, header_idx)
        recap_lookup = _column_lookup(recap_columns)

        mappings: list[MappingEntry] = []
        seen: set[tuple[str, str]] = set()

        for row in rows[header_idx + 1:]:
            attribute_name = _clean_text(row[1] if len(row) > 1 else "")
            if not attribute_name:
                continue

            source_headers = self._source_headers_from_logic(attribute_name, row, col, recap_lookup)
            if not source_headers:
                source_headers = self._source_headers_from_attribute_name(
                    attribute_name,
                    row,
                    recap_lookup,
                    licensed_cols,
                )
            if not source_headers:
                continue

            meta = self.mdd.resolve_attribute(attribute_name)
            if not meta:
                log.info("[Mapping] Skipping %s: no MDD AttributeID", attribute_name)
                continue

            for source_header, reason in source_headers:
                source_col_idx = recap_columns[source_header]
                key = (meta.attr_id, source_header)
                if key in seen:
                    continue
                seen.add(key)
                mappings.append(
                    MappingEntry(
                        attribute_name=attribute_name,
                        attribute_id=meta.attr_id,
                        source_header=source_header,
                        source_col_idx=source_col_idx,
                        meta=meta,
                        reason=reason,
                    )
                )

        log.info("[Mapping] Built %d Recap mappings", len(mappings))
        for entry in mappings:
            log.info(
                "[Mapping] %s -> %s from Recap column %s (%s)",
                entry.attribute_name,
                entry.attribute_id,
                entry.source_header,
                entry.reason,
            )
        return mappings

    def _find_brand_sheet(self, wb) -> str:
        brand_norm = _norm(self.brand)
        for sheet in wb.sheetnames:
            if _norm(sheet) == brand_norm:
                return sheet
        for sheet in wb.sheetnames:
            if brand_norm in _norm(sheet) and "v4" not in _norm(sheet):
                return sheet
        raise ValueError(f"Cannot find Attribute List sheet for {self.brand}")

    @staticmethod
    def _find_attr_header_idx(rows: list[list]) -> int:
        for idx, row in enumerate(rows[:40]):
            normalized = [_norm(v) for v in row]
            if "attributes" in normalized and "field name mapping logic" in normalized:
                return idx
        raise ValueError("Cannot locate Attribute List header row")

    @staticmethod
    def _licensed_regular_columns(rows: list[list], header_idx: int) -> set[int]:
        group_row = rows[header_idx - 2] if header_idx >= 2 else []
        end = len(rows[header_idx])
        start = None
        for idx, value in enumerate(group_row):
            text = _norm(value)
            if "licensed" in text and "sse" not in text:
                start = idx
                break
        if start is None:
            return set()
        for idx in range(start + 1, len(group_row)):
            text = _norm(group_row[idx])
            if "licensed" in text and "sse" in text:
                end = idx
                break
        return set(range(start, end))

    def _source_headers_from_logic(
        self,
        attribute_name: str,
        row: list,
        col: dict[str, int],
        recap_lookup: dict[str, tuple[str, int]],
    ) -> list[tuple[str, str]]:
        logic_idx = col.get("Field Name / Mapping Logic")
        logic = _clean_text(row[logic_idx] if logic_idx is not None and logic_idx < len(row) else "")
        if not logic:
            return []

        if _norm(attribute_name) in DERIVED_ATTRIBUTE_NAMES:
            log.info("[Mapping] Skipping derived attribute: %s", attribute_name)
            return []

        logic_lower = logic.lower()
        if "recap" not in logic_lower and "supplier article" not in logic_lower:
            return []
        if any(marker in logic_lower for marker in DERIVED_MAPPING_MARKERS):
            log.info("[Mapping] Skipping derived/manual logic: %s", logic[:140])
            return []

        found: list[tuple[str, str]] = []
        for match in re.finditer(r"column\s*(?::|')\s*'?([^/'\n]+)'?", logic, re.I):
            candidate = _clean_text(match.group(1))
            source_header = self._resolve_recap_header(candidate, recap_lookup)
            if source_header:
                found.append((source_header, "mapping_logic"))

        return self._dedupe_sources(found)

    def _source_headers_from_attribute_name(
        self,
        attribute_name: str,
        row: list,
        recap_lookup: dict[str, tuple[str, int]],
        licensed_cols: set[int],
    ) -> list[tuple[str, str]]:
        if licensed_cols:
            licensed_values = [
                _norm(row[idx]) for idx in licensed_cols
                if idx < len(row) and _clean_text(row[idx])
            ]
            if not any(value in {"a", "d", "done"} for value in licensed_values):
                return []

        attr_norm = _norm(attribute_name)
        candidates = [attribute_name]
        if attr_norm in ATTRIBUTE_TO_RECAP_HEADER:
            candidates.insert(0, ATTRIBUTE_TO_RECAP_HEADER[attr_norm])

        found = []
        for candidate in candidates:
            source_header = self._resolve_recap_header(candidate, recap_lookup, exact_only=True)
            if source_header:
                found.append((source_header, "attribute_name"))
        return self._dedupe_sources(found)

    @staticmethod
    def _resolve_recap_header(
        candidate: str,
        recap_lookup: dict[str, tuple[str, int]],
        exact_only: bool = False,
    ) -> str:
        candidate_norm = _norm(candidate)
        if candidate_norm in DIRECT_ALIAS_TO_RECAP_HEADER:
            aliased = DIRECT_ALIAS_TO_RECAP_HEADER[candidate_norm]
            if _norm(aliased) in recap_lookup:
                return recap_lookup[_norm(aliased)][0]

        if candidate_norm in recap_lookup:
            return recap_lookup[candidate_norm][0]

        if exact_only:
            return ""

        for alias_norm, aliased in DIRECT_ALIAS_TO_RECAP_HEADER.items():
            if candidate_norm.startswith(f"{alias_norm} ") and _norm(aliased) in recap_lookup:
                return recap_lookup[_norm(aliased)][0]

        prefix_matches = [
            (norm_header, header)
            for norm_header, (header, _) in recap_lookup.items()
            if candidate_norm.startswith(f"{norm_header} ")
        ]
        if prefix_matches:
            return sorted(prefix_matches, key=lambda item: len(item[0]), reverse=True)[0][1]

        compact = _compact_norm(candidate)
        compact_matches = [
            header for norm_header, (header, _) in recap_lookup.items()
            if _compact_norm(norm_header) == compact
        ]
        if len(compact_matches) == 1:
            return compact_matches[0]
        return ""

    @staticmethod
    def _dedupe_sources(values: list[tuple[str, str]]) -> list[tuple[str, str]]:
        out: list[tuple[str, str]] = []
        seen: set[str] = set()
        for source_header, reason in values:
            if source_header and source_header not in seen:
                out.append((source_header, reason))
                seen.add(source_header)
        return out


class RecapWorkbook:
    def __init__(self, path: Path, expected_sheet: str = ""):
        self.path = path
        self.expected_sheet = expected_sheet
        self.sheet_name = ""
        self.header_row_idx = -1
        self.columns: dict[str, int] = {}
        self.rows: list[dict[str, object]] = []
        self._load()

    def _load(self) -> None:
        log.info("[Recap] Loading %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            best = None
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = [_row_values(row) for row in ws.iter_rows()]
                try:
                    header_idx, columns = _find_header_row(rows[:50], RECAP_HEADER_HINTS)
                except ValueError:
                    continue
                score = len({_norm(h) for h in columns} & {_norm(h) for h in RECAP_HEADER_HINTS})
                
                # Priority logic
                if self.expected_sheet and sheet_name.upper() == self.expected_sheet.upper():
                    score += 1000
                elif sheet_name.lower().startswith("sheet") and score > 0:
                    score -= 0.5

                if best is None or score > best[0]:
                    best = (score, sheet_name, rows, header_idx, columns)

            if best is None:
                raise ValueError("Cannot find Recap data sheet")

            _, self.sheet_name, rows, self.header_row_idx, self.columns = best
            headers_by_idx = {idx: name for name, idx in self.columns.items()}
            for raw_row in rows[self.header_row_idx + 1:]:
                row_data: dict[str, object] = {}
                for idx, header in headers_by_idx.items():
                    row_data[header] = raw_row[idx] if idx < len(raw_row) else None
                if self._is_data_row(row_data):
                    self.rows.append(row_data)
        finally:
            wb.close()

        log.info(
            "[Recap] Sheet=%s header_row=%d data_rows=%d columns=%s",
            self.sheet_name,
            self.header_row_idx + 1,
            len(self.rows),
            list(self.columns),
        )

    @staticmethod
    def _is_data_row(row_data: dict[str, object]) -> bool:
        meaningful = [
            value for key, value in row_data.items()
            if _norm(key) != "no" and not _is_blank(value)
        ]
        return bool(meaningful)


def _derive_generic_article_code(
    row: dict[str, object],
    brand_code: str,
    row_num: int | None = None,
) -> str:
    article_type_raw = _clean_text(row.get("Article Type"))
    article_type_map = {
        "REGULAR": "R",
        "SSE": "X",
        "LICENSE": "L",
        "LICENSED": "L",
    }
    article_type_code = article_type_map.get(article_type_raw.upper(), "R")  # Default to R

    season_raw = _clean_text(row.get("Season")).upper()
    season_prefix = season_raw[:2]
    season_code_map = {
        "SS": "S",
        "FW": "F",
        "AW": "F",  # Autumn Winter = Fall Winter
        "HO": "H",
    }
    season_code = season_code_map.get(season_prefix, season_prefix[:1] if season_prefix else "S")
    
    _yr_match = re.search(r"(\d{2})$", season_raw)
    season_year = _yr_match.group(1) if _yr_match else "00"

    supplier_article = _clean_text(row.get("Supp Art #"))
    supplier_suffix = supplier_article[-3:].rjust(3, "0") if supplier_article else "000"

    color_code = _clean_text(row.get("Color Code"))
    if not color_code:
        # Fallback to color name to ensure unique generic codes
        color_name = _clean_text(row.get("Color")).upper()
        words = re.findall(r'[A-Z0-9]+', color_name)
        if len(words) > 1:
            color_code = words[0][0] + words[-1][0]
        elif len(color_name) >= 2:
            color_code = color_name[0] + color_name[-1]
        else:
            color_code = color_name.ljust(2, "0")
        
    color_suffix = color_code[-2:].rjust(2, "0") if color_code else "00"

    missing = []
    if not article_type_code:
        missing.append(f"Article Type={article_type_raw!r}")
    if not season_code:
        missing.append(f"Season={season_raw!r}")
    if not season_year:
        missing.append("Season year")
    if not supplier_suffix:
        missing.append(f"Supp Art #={supplier_article!r}")
    if not color_suffix:
        missing.append(f"Color Code={color_code!r}")

    if missing:
        row_label = f" row {row_num}" if row_num is not None else ""
        log.warning(
            "[Recap] Cannot derive generic article code%s: %s",
            row_label,
            ", ".join(missing),
        )
        return ""

    return f"{brand_code}{article_type_code}{season_code}{season_year}{supplier_suffix}{color_suffix}"


def _product_key(row: dict[str, object], brand_code: str, row_num: int) -> str:
    return _derive_generic_article_code(row, brand_code, row_num)


def _product_name(row: dict[str, object], key: str) -> str:
    parts = [
        _clean_text(row.get("Supp Art #")),
        _clean_text(row.get("Color")),
        _clean_text(row.get("Category")),
    ]
    return " - ".join([p for p in parts if p]) or key


def _write_value(parent: ET.Element, entry: MappingEntry, raw_value: object, mdd: MDDLoader) -> None:
    value = _clean_text(raw_value)
    if not value:
        return

    if entry.meta.is_lov:
        lov_id = mdd.resolve_lov_id(entry.meta, value)
        if not lov_id:
            return
        if entry.meta.is_multi:
            mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
            mv.set("AttributeID", entry.attribute_id)
            child = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
            child.set("ID", lov_id)
        else:
            value_el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
            value_el.set("AttributeID", entry.attribute_id)
            value_el.set("ID", lov_id)
        return

    value_el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    value_el.set("AttributeID", entry.attribute_id)
    value_el.text = value


class RNALoader:
    """
    Loads the 'Source Mapping Related RNA' tab from the Attributes List workbook.
    Maps (Country Name, Company Code, SBU, Brand Code) -> Brand Type + Brand Category.
    """

    RNA_SHEET_KEYWORDS = ["RNA", "SOURCE MAPPING"]

    @staticmethod
    def _norm_country(v: str) -> str:
        return (v or "").strip().upper()

    @staticmethod
    def _norm_comp_code(v: str) -> str:
        s = (v or "").strip()
        if not s:
            return ""
        if s.endswith(".0"):
            s = s[:-2]
        s2 = s.lstrip("0")
        return s2 if s2 else "0"

    def __init__(self, path: Path):
        self.path = path
        self.lookup: dict[tuple, dict] = {}
        self._load()

    def _load(self):
        log.info("[RNA] Loading from: %s", self.path.name)
        import openpyxl
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames if any(kw in s.upper() for kw in self.RNA_SHEET_KEYWORDS)),
            None,
        )
        if not sheet_name:
            log.warning("[RNA] 'Source Mapping Related RNA' tab not found in %s", self.path.name)
            wb.close()
            return

        log.info("[RNA] Using sheet: '%s'", sheet_name)
        rows = list(wb[sheet_name].iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, r in enumerate(rows)
             if any(isinstance(v, str) and "COUNTRY" in v.upper() for v in r if v)),
            None,
        )
        if hdr_idx is None:
            log.warning("[RNA] Cannot find header row in sheet '%s'", sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col = {str(h).strip().upper(): i for i, h in enumerate(hdr) if h}

        def _find(*candidates) -> int | None:
            for c in candidates:
                if c.upper() in col:
                    return col[c.upper()]
            return None

        c_country = _find("COUNTRY", "COUNTRY NAME")
        c_comp    = _find("COMPANY CODE", "COMP CODE", "COMPCODE", "COMP_CODE")
        c_sbu     = _find("SBU")
        c_bcode   = _find("BRANDCODE", "BRAND CODE", "REPORTING BRAND CODE MAPPED")
        c_btype   = _find("BRANDTYPE_DETAIL", "BRAND TYPE", "AT_BRANDTYPE", "BRANDTYPE")
        c_bcat    = _find("BRANDCATEGORY", "BRAND CATEGORY", "AT_BRANDCATEGORY", "BRANDCATEGORY")
        c_bgroup  = _find("BRANDGROUP", "BRAND GROUP")

        def _cell(row, idx):
            if idx is None or idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()

        for row in rows[hdr_idx + 1:]:
            country = _cell(row, c_country)
            if not country:
                continue
            key = (
                self._norm_country(country),
                self._norm_comp_code(_cell(row, c_comp)),
                _cell(row, c_sbu).upper(),
                _cell(row, c_bcode).upper(),
            )
            self.lookup[key] = {
                "brand_type":     _cell(row, c_btype),
                "brand_category": _cell(row, c_bcat),
                "brand_group":    _cell(row, c_bgroup),
            }
        wb.close()


def _write_multi_value(
    values_el: ET.Element,
    written: set[str],
    attr_id: str,
    id_val: str,
) -> None:
    if attr_id in written:
        return
    clean_id = _clean_text(id_val)
    if not clean_id:
        return
    mv = ET.SubElement(values_el, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    child = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    child.set("ID", clean_id)
    written.add(attr_id)
def _write_simple_value(
    values_el: ET.Element,
    written: set[str],
    attr_id: str,
    value: str,
    *,
    id_val: str | None = None,
) -> None:
    if attr_id in written:
        return
    clean_value = _clean_text(value)
    if not clean_value and not id_val:
        return
    value_el = ET.SubElement(values_el, f"{{{STIBO_NS}}}Value")
    value_el.set("AttributeID", attr_id)
    if id_val:
        value_el.set("ID", id_val)
    if clean_value:
        value_el.text = clean_value
    written.add(attr_id)


def _build_product_xml(
    row: dict[str, object],
    mappings: list[MappingEntry],
    mdd: MDDLoader,
    rna: RNALoader,
    brand: str,
    brand_code: str,
    season_id: str,
    row_num: int,
    filename_meta: dict[str, str],
    comp_code: str,
    sbu: str,
) -> str:
    key = _product_key(row, brand_code, row_num)
    if not key:
        return ""

    product = ET.Element(f"{{{STIBO_NS}}}Product")
    product.set("UserTypeID", "PRD_GenericArticle")
    product.set("ParentID", "PPH_F-TempSubCat")

    key_value = ET.SubElement(product, f"{{{STIBO_NS}}}KeyValue")
    key_value.set("KeyID", "KEY_InboundArticle")
    key_value.text = key

    ET.SubElement(product, f"{{{STIBO_NS}}}Name").text = _product_name(row, key)

    brand_title = brand.title().replace(" ", "")
    cr_merch = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand_title}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    if season_id:
        cr_unconfirmed = ET.SubElement(product, f"{{{STIBO_NS}}}ClassificationReference")
        cr_unconfirmed.set("ClassificationID", f"{season_id}UA")
        cr_unconfirmed.set("Type", "CPL_UnConfirmedForSeason")

    values = ET.SubElement(product, f"{{{STIBO_NS}}}Values")
    written: set[str] = set()
    
    # ── Core identifiers ─────────────────────────────────────────
    _write_simple_value(values, written, "AT_InboundGenericCode", key)
    _write_simple_value(values, written, "AT_Generic", key)
    
    supp_art = _clean_text(row.get("Supp Art #"))
    if supp_art:
        _write_simple_value(values, written, "AT_PrincipalStyleCode", key)
        # SAP Style Code is the generic code (key) without the 3-char brand prefix
        _write_simple_value(values, written, "AT_SAPStyleCode", key[3:])
        
    color = _clean_text(row.get("Color"))
    if color:
        _write_simple_value(values, written, "AT_PrincipalColorName", color)
        
    color_code = _clean_text(row.get("Color Code"))
    if color_code:
        _write_simple_value(values, written, "AT_PrincipalColorCode", color_code)
        
    _write_simple_value(values, written, "AT_Brand", "", id_val=brand_code)
    _write_simple_value(values, written, "AT_BrandGroup", "", id_val=brand.upper())

    # ── Derived manual mappings ──
    c_code_for_rna = (filename_meta.get("country") or "").upper()
    country_name_map = {"ID": "INDONESIA", "PH": "PHILIPPINES", "VN": "VIETNAM", "TH": "THAILAND", "MY": "MALAYSIA", "SG": "SINGAPORE"}
    country_name = country_name_map.get(c_code_for_rna, c_code_for_rna)
    
    brand_cats = rna.lookup.get((
        RNALoader._norm_country(country_name),
        RNALoader._norm_comp_code(comp_code),
        sbu.upper(),
        brand_code.upper()
    ), {})
    b_type = brand_cats.get("brand_type", "LICENCE")
    b_cat = brand_cats.get("brand_category", "LICENCE")

    _write_simple_value(values, written, "AT_BrandType", "", id_val=b_type)
    _write_simple_value(values, written, "AT_BrandCategory", "", id_val=b_cat)
    _write_multi_value(values, written, "AT_CompanyCode", comp_code)
    _write_multi_value(values, written, "AT_SBU", sbu)
    
    if row.get("Division"):
        _write_simple_value(values, written, "AT_PrincipalMerchandiseHierarchyL1", row.get("Division"))
    if row.get("MD Category"):
        _write_simple_value(values, written, "AT_PrincipalMerchandiseHierarchyL2", row.get("MD Category"))
    if row.get("Image"):
        _write_simple_value(values, written, "AT_ThumbnailImage", row.get("Image"))
    if row.get("ETA DATE"):
        _write_simple_value(values, written, "AT_IncomingMonth", _fmt_date(row.get("ETA DATE")))
        
    _write_simple_value(values, written, "AT_SAPArticleCategory", "", id_val="1")
    _write_simple_value(values, written, "AT_UOM", "", id_val="EA")
    _write_simple_value(values, written, "AT_PricingDistributionChannel", "", id_val="01")
    
    c_code = (filename_meta.get("country") or "").upper()
    if c_code:
        cur_map = {"ID": "IDR", "PH": "PHP", "VN": "VND", "TH": "THB", "MY": "MYR", "SG": "SGD"}
        if c_code in cur_map:
            _write_simple_value(values, written, "AT_RetailPriceCurrency", "", id_val=cur_map[c_code])
            
    gender_raw = _clean_text(row.get("Gender")).upper()
    gender_code_raw = _clean_text(row.get("Gender Code"))
    if gender_raw:
        _write_simple_value(values, written, "AT_PrincipalGenderDescription", gender_raw.title())
        sap_gender = DIA_GENDER_MAP.get(gender_raw, "")
        if not sap_gender:
            sap_gender = LOV_GENDER.get(gender_raw[:1], "")
        if sap_gender:
            _write_simple_value(values, written, "AT_Gender", sap_gender, id_val=sap_gender[:1].upper())
            _write_simple_value(values, written, "AT_BYGender", sap_gender, id_val=sap_gender[:1].upper())
            
    if gender_code_raw:
        _write_simple_value(values, written, "AT_PrincipalGenderCode", gender_code_raw)
        
    age_group_val = _clean_text(row.get("Age Group")).title()
    if age_group_val:
        _write_simple_value(values, written, "AT_PrincipalAgeDescription", age_group_val)
        
        reverse_lov_age = {v.upper(): k for k, v in LOV_AGE.items()}
        sap_age_lov_id = reverse_lov_age.get(age_group_val.upper(), "AD")
        _write_simple_value(values, written, "AT_SAPAge", age_group_val, id_val=sap_age_lov_id)
        
        by_age_lov_value = LOV_BY_AGE.get(age_group_val.upper(), age_group_val)
        _write_simple_value(values, written, "AT_BYAge", by_age_lov_value, id_val=by_age_lov_value.upper())

    for entry in mappings:
        if entry.attribute_id in written:
            continue
        raw_value = row.get(entry.source_header)
        if _is_blank(raw_value):
            continue
        _write_value(values, entry, raw_value, mdd)
        written.add(entry.attribute_id)

    season_raw = (filename_meta.get("season") or "").upper()
    season_code = season_raw[:2] if len(season_raw) >= 2 else ""
    season_year = ""
    year_match = re.search(r"20\d{2}", season_raw)
    if year_match:
        season_year = year_match.group(0)
    else:
        short_match = re.search(r"^[A-Z]{2}(\d{2})$", season_raw)
        if short_match:
            season_year = f"20{short_match.group(1)}"

    article_type = filename_meta.get("article_type") or ""
    country_code = (filename_meta.get("country") or "").upper()

    _write_simple_value(values, written, "AT_BYArticleType", article_type, id_val=article_type or None)
    _write_simple_value(values, written, "AT_Season", season_code, id_val=season_code or None)
    _write_simple_value(values, written, "AT_SeasonYear", season_year)
    _write_simple_value(values, written, "AT_Country", "", id_val=country_code)
    _write_simple_value(values, written, "AT_SAPProductFlag", "", id_val="A")
    _write_simple_value(values, written, "AT_MaterialType", "", id_val="ZINA")

    # ── Sports Category ──────────────────────────────────────────
    category_raw = _clean_text(row.get("Collection") or row.get("Category"))
    if category_raw:
        sports_cat_label = DIA_RECAP_CATEGORY_TO_SPORTS_CAT.get(category_raw.upper(), "")
        if sports_cat_label:
            hardcoded_lov = {
                "RUNNING": "01", "SOCCER": "02", "BASKETBALL": "03",
                "TENNIS / PADEL": "04", "FITNESS / TRAINING": "05", "SWIMMING": "06",
                "LIFESTYLE / CASUAL": "07", "OUTDOOR": "08", "ACTION SPORTS": "09",
                "OTHER": "10", "NOT APPLICABLE": "11",
            }
            sports_cat_id = hardcoded_lov.get(sports_cat_label.upper())
            if not sports_cat_id:
                sports_lov = mdd.lovs.get("sports category", {})
                sports_cat_id = sports_lov.get(_norm(sports_cat_label), "")
                if not sports_cat_id:
                    sports_cat_id = sports_lov.get(_compact_norm(sports_cat_label), "")
                if sports_cat_id and sports_cat_id.isdigit() and len(sports_cat_id) == 1:
                    sports_cat_id = f"0{sports_cat_id}"
            
            if sports_cat_id:
                _write_simple_value(values, written, "AT_SportsCategoryEN", "", id_val=sports_cat_id)

    return _XMLNS_RE.sub("", ET.tostring(product, encoding="unicode"))


def _season_classification_id(brand_code: str, season: str) -> str:
    season = _clean_text(season).upper()
    if not season:
        return ""
    prefix = season[:2]
    year = season[2:]
    if len(year) == 2:
        year = f"20{year}"
    return f"CLH_{brand_code}_{prefix}{year}"


def _output_filename(source_name: str) -> str:
    return f"{Path(source_name).stem}.xml"


def run(args, auditor=None) -> None:
    mdd_file = _first_file(MDD_DIR)
    attr_file = _first_file(ATTR_DIR)
    recap_files = sorted(RECAP_DIR.glob("*.xlsx"))

    for label, value in [
        ("MDD", mdd_file),
        ("Attributes List", attr_file),
        ("Recap", recap_files),
    ]:
        if not value:
            raise FileNotFoundError(f"No {label} file found in expected input directory")

    mdd = MDDLoader(mdd_file)
    attr_loader = AttributeListLoader(attr_file, mdd, brand="DIADORA")
    rna = RNALoader(attr_file)

    total_written = 0
    for recap_file in recap_files:
        filename_meta = _parse_metadata_from_input_filename(recap_file.stem)
        recap = RecapWorkbook(recap_file, expected_sheet=args.brand_code)
        mappings = attr_loader.build_recap_mappings(recap.columns)
        if not mappings:
            log.warning("[Recap] No usable direct mappings found for %s", recap_file.name)

        out_path = XML_OUT_DIR / _output_filename(recap_file.name)
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        season_for_classification = filename_meta.get("season") or args.season
        season_id = _season_classification_id(args.brand_code, season_for_classification)

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
                f' UseContextLocale="false">\n'
            )
            f.write("  <Products>\n")
            written = 0
            for idx, row in enumerate(recap.rows, start=1):  # TEST MODE: limit to first 5 rows
                product_xml = _build_product_xml(
                    row,
                    mappings,
                    mdd,
                    rna,
                    args.brand,
                    args.brand_code,
                    season_id,
                    idx,
                    filename_meta,
                    args.comp_code,
                    args.sbu,
                )
                if not product_xml:
                    continue
                f.write(f"    {product_xml}\n")
                written += 1
            f.write("  </Products>\n")
            f.write("</STEP-ProductInformation>\n")

        total_written += written
        log.info("[Recap] XML written: %s (%d products)", out_path, written)
        if auditor:
            auditor.set_xml_uploads([str(out_path)])

    log.info("[Recap] Completed. Products written: %d", total_written)


def main() -> None:
    parser = argparse.ArgumentParser(description="Diadora Recap STEP XML generator")
    parser.add_argument("--brand", default="Diadora")
    parser.add_argument("--brand-code", default="DIA")
    parser.add_argument("--comp-code", default="0888")
    parser.add_argument("--sbu", default="SP")
    parser.add_argument("--season", default="FW26")
    parser.add_argument("--seq", default=1, type=int)
    parser.add_argument("--file-type", default="Recap Sample")
    parser.add_argument("--multi-mono", default="Multi")
    parser.add_argument("--country-code", default="")
    run(parser.parse_args())


if __name__ == "__main__":
    main()
