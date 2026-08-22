"""
adidas/dtb_main.py
==================
DTB (Digital Trade Book) Ecommerce Enrichment pipeline for Adidas Inline.

Flow:
    1. Load ADIDAS brand sheet from Attributes List file.
       For every attribute row, read 'Attributes' (the PIM attribute name)
       and 'Field Name / Mapping Logic' (the cell text).
       Inclusion rule  → cell contains "dtb"  AND  does NOT contain "linelist".

    2. Cross-reference those filtered attribute rows against the actual
       column headers present in the DTB English sheet.
       For each DTB column header, do a case-insensitive substring search
       across all filtered mapping-logic cells.
       If exactly ONE attribute row matches a given DTB column → keep it.
       If MULTIPLE DTB columns feed the SAME attribute → skip (pending
       senior guidance; logged as warning).

    3. For each surviving (dtb_col → attr_name) pair, look up the
       MDD 'Core Attributes' sheet to retrieve:
           • PIM Attribute ID  (AT_xxx)  — column index 7
           • Validation Base Type        — column 'Validation Base Type'
       Drop any pair where the attr_name is absent from the MDD.

    4. Load the first sheet ('English') of the DTB file.
       Build one dict per article row (keyed on Article No.).

    5. Build Stibo STEP XML:
           <Products>
             <Product UserTypeID="PRD_GenericArticle" …>
               <KeyValue KeyID="KEY_InboundArticle">ADI{article_no}</KeyValue>
               <Name>…</Name>
               <Values>
                 <Value AttributeID="{AT_xxx}">{value}</Value>
                 …
               </Values>
             </Product>
             …
           </Products>
       Only attributes with a non-empty value are written.
       Empty articles (zero values written) are skipped.

    6. Write XML to output/xml/ ; caller uploads to S3.

Called from adidas/lambda_function.py:
    from adidas.dtb_main import run_dtb
    run_dtb(args, auditor=auditor, tmp_dir=Path(TMP_WORKDIR))

args must have: brand, brand_code, comp_code, sbu, season, seq, multi_mono
"""

import logging
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pandas as pd

log = logging.getLogger("dtb_main")

BASE_DIR    = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))
INPUT_DIR   = BASE_DIR / "input"
DTB_DIR     = INPUT_DIR / "dtb"
MDD_DIR     = INPUT_DIR / "mdd"
ATTR_DIR    = INPUT_DIR / "attributes"
OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for _d in [DTB_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')

_MDD_AT_ID_COL_IDX       = 7   # confirmed: PIM Attribute ID is at 0-based index 7
_MDD_HEADER_ROW_IDX      = 1   # headers are at row index 1 (row 2 in Excel)
_MDD_DATA_START_ROW_IDX  = 2   # data starts at row index 2

# ──────────────────────────────────────────────────────────────────
# Special column rules — columns whose names don't match the
# attribute name in the Attributes List, so the dynamic substring
# search cannot discover them automatically.
#
# source_norms  : normalised DTB column name(s) to find (all required)
# mdd_attr      : exact PIM Attribute Name to look up in MDD
# compose       : fn(list[str]) → str  — builds the final value
# ──────────────────────────────────────────────────────────────────
SPECIAL_COLUMN_RULES: list[dict] = [
    {
        "rule_id":      "article_weight",
        "source_norms": ["article weight"],
        "mdd_attr":     "Product Weight",
        "multivalue":   False,
        "compose":      lambda vals: vals[0],
    },
    {
        "rule_id":      "dimensions_lwh",
        "source_norms": ["article length", "article width", "article height"],
        "mdd_attr":     "Product Length X Width X Height ( in CM )",
        "multivalue":   False,
        "compose":      lambda vals: f"{vals[0]} X {vals[1]} X {vals[2]} CM",
    },
    {
        # Care Instruction EN is a MultiValue in Stibo — all 6 DTB care columns feed it.
        # Add/update source_norms if the 6th column name differs from what's listed here.
        "rule_id":      "care_instruction_en",
        "source_norms": [
            "bleaching instruction",
            "extra care instructions",
            "ironing instruction",
            "professional care instructions",
            "washing instruction",
        ],
        "mdd_attr":     "Care Instruction EN",
        "multivalue":   True,
        "compose":      None,
    },
]


def _s(v) -> str:
    """Safely convert any value to a clean non-empty string, or return ''."""
    if v is None:
        return ""
    try:
        if pd.isnull(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _norm(text: str) -> str:
    """Normalise a string for fuzzy matching: lowercase + collapse whitespace."""
    return " ".join(text.strip().split()).lower()


def _norm_lov(text: str) -> str:
    """Normalise for LOV key lookups: uppercase, strip all non-alphanumeric."""
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def _is_numeric_only(s: str) -> bool:
    """True only if string is a plain integer with no leading zeros (invalid as LOV ID)."""
    stripped = s.strip()
    if not stripped:
        return False
    if len(stripped) > 1 and stripped[0] == "0":
        return False
    return bool(re.match(r"^\d+$", stripped))


# Valid Stibo gender IDs — M, F, U only
_STIBO_GENDER_VALID = {"F", "M", "U"}

DIVISION_PARENT_MAP = {
    "ACCESSORIES": "E", "FOOTWEAR": "F", "APPAREL": "A",
    "EQUIPMENT": "Q", "HARDWARE": "Q", "TOYS": "T",
}


# DTB column substring candidates for defining-attribute detection.
# _find_col_val() checks whether a candidate appears as a substring of
# the actual column header (case-insensitive). Order matters — first
# match wins, so list more-specific names before generic ones.
_DTB_COLOR_CANDIDATES       = ["b2b base color", "base color", "colorway name",
                                "colour description", "color description",
                                "colour", "color"]
_DTB_DIVISION_CANDIDATES    = ["b2b product division", "product division", "division"]
_DTB_BIZ_SEGMENT_CANDIDATES = ["b2b business segment", "business segment", "segment"]


# ══════════════════════════════════════════════════════════════════
# SECTION 1 — LOADERS
# ══════════════════════════════════════════════════════════════════

class MDDPIMLookup:
    """
    Reads the MDD Excel file and builds two things from it:

    1. Core Attributes sheet → AT_xxx ID reverse lookup
       (used for dynamic enrichment attribute mapping)

    2. All LOV sheets → lookup tables for brand, gender, age, season,
       color, company code, SBU, article category.
       All values are loaded from the MDD — nothing hardcoded.

    Core Attributes column layout (0-based):
        col 5  → PIM Attribute Name
        col 7  → PIM Attribute ID  (AT_xxx)
        col 9  → Validation Base Type

    LOV sheet column layout (0-based, consistent across sheets):
        col 0  → Value ID / Code
        col 1  → Label / Description
        (col 2  → SAP code where applicable, e.g. Gender LOV)
    """

    def __init__(self, path: Path):
        self.path    = path
        self.by_name: dict[str, dict] = {}   # pim_attr_name → {at_id, validation_type}

        # All populated from MDD sheets — no default values here
        self._color_desc_to_id: dict[str, str]   = {}  # norm(description) → LOV ID
        self._color_id_to_desc: dict[str, str]   = {}  # norm(LOV ID)      → description
        self._brand:            dict[str, str]   = {}  # norm(code)        → label
        self._gender:           dict[str, tuple] = {}  # norm(key)         → (stibo_id, label)
        self._age:              dict[str, tuple] = {}  # norm(key)         → (stibo_code, label)
        self._season:           dict[str, str]   = {}  # code.upper()      → label
        self._article_cat:      dict[str, str]   = {}  # code              → label
        self._company_code:     dict[str, str]   = {}  # code              → label
        self._sbu:              dict[str, str]   = {}  # code              → label

        self._load()

    # ── sheet row iterator ────────────────────────────────────────────────────

    def _sheet_rows(self, wb, sheet_name: str, skip: int = 1):
        if sheet_name not in wb.sheetnames:
            log.warning("[MDD] Sheet not found: %s", sheet_name)
            return
        for i, row in enumerate(wb[sheet_name].iter_rows(values_only=True)):
            if i < skip:
                continue
            if any(c is not None for c in row):
                yield row

    # ── loaders ───────────────────────────────────────────────────────────────

    def _load(self):
        log.info("[MDD] Loading: %s", self.path.name)
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        except Exception as exc:
            log.error("[MDD] Cannot open file: %s", exc)
            return

        self._load_core_attributes(wb)
        self._load_colors(wb)
        self._load_brands(wb)
        self._load_gender(wb)
        self._load_age(wb)
        self._load_season(wb)
        self._load_article_category(wb)
        self._load_company_code(wb)
        self._load_sbu(wb)
        wb.close()

        log.info(
            "[MDD] Loaded — attrs:%d  colors:%d  brands:%d  genders:%d  ages:%d  seasons:%d",
            len(self.by_name), len(self._color_desc_to_id),
            len(self._brand), len(self._gender),
            len(self._age), len(self._season),
        )

    def _load_core_attributes(self, wb):
        try:
            ws   = wb["Core Attributes"]
            rows = list(ws.iter_rows(values_only=True))
        except Exception as exc:
            log.warning("[MDD] Core Attributes error: %s", exc)
            return

        hdr_row = rows[_MDD_HEADER_ROW_IDX]
        col: dict[str, int] = {}
        for i, h in enumerate(hdr_row):
            if h:
                col[str(h).split("\n")[0].strip()] = i
                col[str(h).strip()] = i

        def _v(row, key, fallback_idx=None):
            idx = col.get(key, fallback_idx)
            if idx is None or idx >= len(row):
                return None
            v = row[idx]
            return str(v).strip() if v else None

        loaded = 0
        for row in rows[_MDD_DATA_START_ROW_IDX:]:
            if not row or len(row) <= _MDD_AT_ID_COL_IDX:
                continue
            at_id = row[_MDD_AT_ID_COL_IDX]
            if not at_id:
                continue
            at_id    = str(at_id).strip()
            pim_name = _v(row, "PIM Attribute Name", fallback_idx=5)
            if not pim_name:
                continue
            validation = _v(row, "Validation Base Type", fallback_idx=9) or "Text"
            self.by_name[_norm(pim_name)] = {
                "at_id":           at_id,
                "validation_type": validation,
                "pim_name":        pim_name,
            }
            loaded += 1
        log.info("[MDD] %d core PIM attributes indexed", loaded)

    def _load_colors(self, wb):
        """Color Code LOV: col0 = LOV ID (code), col1 = description."""
        for row in self._sheet_rows(wb, "Color Code LOV"):
            lov_id = _s(row[0]) if len(row) > 0 else ""
            desc   = _s(row[1]) if len(row) > 1 else ""
            if not lov_id:
                continue
            lov_id = str(lov_id).strip()
            desc   = desc.strip().upper()
            self._color_id_to_desc[_norm_lov(lov_id)] = desc
            if desc:
                self._color_desc_to_id[_norm_lov(desc)] = lov_id

        # SIS SM Color Description LOV: col0 = LOV ID, col2 = description
        for row in self._sheet_rows(wb, "SIS SM Color Description LOV"):
            lov_id = _s(row[0]) if len(row) > 0 else ""
            desc   = _s(row[2]) if len(row) > 2 else ""
            if not lov_id or not desc:
                continue
            key = _norm_lov(desc)
            if key not in self._color_desc_to_id:
                self._color_desc_to_id[key] = str(lov_id).strip()

        log.info("[MDD] Color LOV: %d descriptions indexed", len(self._color_desc_to_id))

    def _load_brands(self, wb):
        """Brand LOV: col0 = brand code, col1 = brand label."""
        for row in self._sheet_rows(wb, "Brand LOV"):
            code  = _s(row[0]) if len(row) > 0 else ""
            label = _s(row[1]) if len(row) > 1 else ""
            if code:
                self._brand[_norm_lov(code)] = label
        log.info("[MDD] Brand LOV: %d brands indexed", len(self._brand))

    def _load_gender(self, wb):
        """Gender LOV: col0 = label (Male/Female/Unisex), col1 = Stibo ID (M/F/U),
        col2 = SAP gender code (may differ, e.g. SAP Female = W)."""
        if "Gender LOV" not in wb.sheetnames:
            log.warning("[MDD] Sheet not found: Gender LOV")
            return
        rows = list(wb["Gender LOV"].iter_rows(values_only=True))
        for row in rows[2:]:                  # row 0 empty, row 1 = headers
            label    = _s(row[0]) if len(row) > 0 else ""
            stibo_id = _s(row[1]) if len(row) > 1 else ""
            sap_code = _s(row[2]) if len(row) > 2 else ""
            if not label and not stibo_id:
                continue
            sid = stibo_id if stibo_id in _STIBO_GENDER_VALID else ""
            if not sid:
                continue
            entry = (sid, label or stibo_id)
            # Index by label, Stibo ID, and SAP code — all from MDD
            for key in filter(None, [label, stibo_id, sap_code]):
                self._gender[_norm_lov(key)] = entry
        log.info("[MDD] Gender LOV: %d keys indexed", len(self._gender))

    def _load_age(self, wb):
        """Age LOV: col0 = label (Adults/Children/All Ages),
        col1 = Stibo code (AD/CH/AA), col2 = SAP age code."""
        for row in self._sheet_rows(wb, "Age LOV"):
            label      = _s(row[0]) if len(row) > 0 else ""
            stibo_code = _s(row[1]) if len(row) > 1 else ""
            sap_code   = _s(row[2]) if len(row) > 2 else ""
            if not label and not stibo_code:
                continue
            entry = (stibo_code, label)
            # Index by label, Stibo code, and SAP code — all from MDD
            for key in filter(None, [label, stibo_code, sap_code]):
                self._age[_norm_lov(key)] = entry
        log.info("[MDD] Age LOV: %d keys indexed", len(self._age))

    def _load_season(self, wb):
        """Season LOV: col0 = code (SS/FW/AL…), col1 = label."""
        for row in self._sheet_rows(wb, "Season LOV"):
            code  = _s(row[0]).strip() if len(row) > 0 else ""
            label = _s(row[1])         if len(row) > 1 else ""
            if code:
                self._season[code.strip().upper()] = label
        log.info("[MDD] Season LOV: %d seasons indexed", len(self._season))

    def _load_article_category(self, wb):
        """Article Category LOV: col0 = code, col1 = label."""
        for row in self._sheet_rows(wb, "Article Category LOV"):
            code  = _s(row[0]) if len(row) > 0 else ""
            label = _s(row[1]) if len(row) > 1 else ""
            if code:
                self._article_cat[str(code).strip()] = label.strip()
        log.info("[MDD] Article Category LOV: %d categories indexed", len(self._article_cat))

    def _load_company_code(self, wb):
        """Company Code LOV: col0 = code, col1 = label."""
        for row in self._sheet_rows(wb, "Company Code LOV"):
            code  = _s(row[0]) if len(row) > 0 else ""
            label = _s(row[1]) if len(row) > 1 else ""
            if code:
                self._company_code[str(code).strip()] = label
        log.info("[MDD] Company Code LOV: %d codes indexed", len(self._company_code))

    def _load_sbu(self, wb):
        """SBU LOV: col0 = code, col1 = label."""
        for row in self._sheet_rows(wb, "SBU LOV"):
            code  = _s(row[0]) if len(row) > 0 else ""
            label = _s(row[1]) if len(row) > 1 else ""
            if code:
                self._sbu[code.strip()] = label
        log.info("[MDD] SBU LOV: %d codes indexed", len(self._sbu))

    # ── flexible lookup ───────────────────────────────────────────────────────

    def _flexible_lookup(self, table: dict, raw: str):
        """Try exact → prefix → substring match against a normalised-key dict."""
        if not raw:
            return None
        key = _norm_lov(raw)
        if key in table:
            return table[key]
        for stored in table:
            if stored.startswith(key) and stored != key:
                return table[stored]
        for stored in table:
            if key.startswith(stored) and stored:
                return table[stored]
        return None

    # ── public lookup: core attribute ID ─────────────────────────────────────

    def lookup(self, attribute_name: str) -> dict | None:
        """Return mapping entry for a given PIM attribute name, or None if absent.

        Also tries with spaces stripped from inside parentheses so that
        '( in CM )' (attribute list) matches '(in CM)' (MDD) and vice-versa.
        """
        key = _norm(attribute_name)
        entry = self.by_name.get(key)
        if entry:
            return entry
        # Normalise paren spacing: "( in cm )" → "(in cm)"
        key_paren = re.sub(r'\(\s+', '(', re.sub(r'\s+\)', ')', key))
        return self.by_name.get(key_paren) if key_paren != key else None

    # ── public lookup: LOV values — all return None when not found in MDD ────

    def brand_label(self, code: str) -> str | None:
        """Return brand label from MDD, or None if code not in Brand LOV."""
        return self._flexible_lookup(self._brand, code)

    def gender(self, raw: str) -> tuple | None:
        """Return (stibo_id, label) from MDD Gender LOV, or None if not found."""
        result = self._flexible_lookup(self._gender, raw)
        if not result:
            log.warning("[MDD] No gender match for %r in MDD — attribute will be skipped", raw)
        return result

    def age(self, raw: str) -> tuple | None:
        """Return (stibo_code, label) from MDD Age LOV, or None if not found."""
        result = self._flexible_lookup(self._age, raw)
        if not result:
            log.warning("[MDD] No age match for %r in MDD — attribute will be skipped", raw)
        return result

    def season_label(self, code: str) -> str | None:
        """Return season label from MDD, or None if code not in Season LOV."""
        return self._season.get(code.strip().upper())

    def article_category(self, code: str) -> str | None:
        """Return article category label from MDD, or None if code not found."""
        return self._article_cat.get(str(code).strip())

    def company_code_label(self, code: str) -> str | None:
        """Return company code label from MDD, or None if code not found."""
        return self._company_code.get(str(code).strip())

    def sbu_label(self, code: str) -> str | None:
        """Return SBU label from MDD, or None if code not found."""
        return self._sbu.get(code.strip())

    def color_to_lov_id(self, raw: str) -> str:
        """Return Stibo Color LOV ID by matching description against MDD Color Code LOV.
        Returns '' if no match found — caller should then omit AT_Color."""
        if not raw:
            return ""
        raw_key = _norm_lov(raw)

        # 1. Direct description match
        if raw_key in self._color_desc_to_id:
            result = self._color_desc_to_id[raw_key]
            if result and not _is_numeric_only(result):
                return result

        # 2. Two-word prefix (e.g. "Core Black" → "COREBLACK")
        words = raw.strip().split()
        if len(words) >= 2:
            two = _norm_lov(f"{words[0]} {words[1]}")
            if two in self._color_desc_to_id:
                result = self._color_desc_to_id[two]
                if result and not _is_numeric_only(result):
                    return result

        # 3. First word only
        if words:
            first = _norm_lov(words[0])
            if first in self._color_desc_to_id:
                result = self._color_desc_to_id[first]
                if result and not _is_numeric_only(result):
                    return result

        # 4. Substring match against all description keys
        for k, v in self._color_desc_to_id.items():
            if raw_key in k or k in raw_key:
                if v and not _is_numeric_only(v):
                    return v

        # 5. The raw value itself may be a valid LOV ID
        if raw_key in self._color_id_to_desc:
            cleaned = re.sub(r"[^A-Z0-9/]", "", raw.upper())
            if cleaned and not _is_numeric_only(cleaned):
                return cleaned

        log.debug("[MDD] No color LOV match for %r", raw)
        return ""


class AdidasAttributeMapLoader:
    """
    Reads the ADIDAS sheet from the Attributes List file.

    For each attribute row, reads:
        'Attributes'                  → pim_attribute_name  (join key to MDD)
        'Field Name / Mapping Logic'  → free-text cell

    Inclusion rule:
        cell_text contains "dtb"       (case-insensitive)
        AND does NOT contain "linelist" (case-insensitive)

    After filtering, exposes build_dtb_column_map(dtb_headers) which:
        1. For each DTB column header, searches all filtered rows to see if
           that header appears as a substring inside the mapping-logic text.
        2. Keeps only (dtb_col → attr_name) pairs where exactly ONE row matched
           and exactly ONE DTB column feeds that attr_name.
        3. Returns {dtb_col_header: attr_name}.
    """

    def __init__(self, path: Path):
        self.path      = path
        self.attr_rows: list[dict] = []
        self._load()

    def _load(self):
        log.info("[AttrMap] Loading ADIDAS sheet from: %s", self.path.name)
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)

        sheet_name = next(
            (s for s in wb.sheetnames
             if _norm(s) == "adidas" and "v4" not in s.lower()),
            None,
        ) or next(
            (s for s in wb.sheetnames
             if "adidas" in s.lower() and "v4" not in s.lower()),
            None,
        )
        if not sheet_name:
            log.error("[AttrMap] ADIDAS sheet not found in %s", self.path.name)
            wb.close()
            return

        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        hdr_idx = next(
            (i for i, row in enumerate(rows)
             if any(str(c or "").strip() == "Attributes" for c in (row or [])[:3])),
            None,
        )
        if hdr_idx is None:
            log.error("[AttrMap] Cannot locate header row ('Attributes' column) in sheet '%s'",
                      sheet_name)
            wb.close()
            return

        hdr = rows[hdr_idx]
        col: dict[str, int] = {str(h).strip(): i for i, h in enumerate(hdr) if h}

        attr_col    = col.get("Attributes", 0)
        mapping_col = col.get("Field Name / Mapping Logic") or next(
            (i for i, h in enumerate(hdr)
             if h and "mapping" in str(h).lower() and "logic" in str(h).lower()),
            None,
        )
        if mapping_col is None:
            log.error("[AttrMap] Cannot locate 'Field Name / Mapping Logic' column")
            wb.close()
            return

        for row in rows[hdr_idx + 1:]:
            if not row:
                continue
            max_needed = max(attr_col, mapping_col)
            if len(row) <= max_needed:
                continue

            attr_name     = str(row[attr_col] or "").strip()
            mapping_logic = str(row[mapping_col] or "").strip()
            if not attr_name or not mapping_logic:
                continue

            ml_lower = mapping_logic.lower()
            if "dtb" not in ml_lower:
                continue
            # Include rows even when linelist is also mentioned — many attributes
            # are sourced from both Linelist AND DTB; excluding them drops valid mappings.

            self.attr_rows.append({
                "attr_name":     attr_name,
                "mapping_logic": mapping_logic,
            })

        wb.close()
        log.info("[AttrMap] %d DTB attribute rows discovered in sheet '%s'",
                 len(self.attr_rows), sheet_name)

    def build_dtb_column_map(self, dtb_headers: list[str]) -> list[tuple[str, str]]:
        """
        Cross-reference DTB file column headers against filtered attribute rows.

        Returns list of (dtb_col_header, attr_name) pairs.
        - One DTB column can feed multiple attributes — all pairs are kept.
        - If multiple DTB columns feed the same attribute, the first one by header
          order wins (pick-first rule); logs a warning for the losers.

        Matching uses exact word-boundary regex against the mapping-logic cell.
        The DTB column name must appear verbatim in the mapping logic — no B2B
        prefix stripping. If the attribute list says "b2b gender", only "B2B Gender"
        matches; if it says "gender", only a plain "Gender" column matches.
        """
        if not self.attr_rows:
            log.warning("[AttrMap] No DTB attribute rows — map will be empty")
            return []

        col_to_attrs: dict[str, list[str]] = defaultdict(list)

        for dtb_col in dtb_headers:
            col_lower = _norm(dtb_col)
            if not col_lower:
                continue

            # Exact word-boundary match only — no B2B prefix stripping.
            # If mapping logic says "b2b gender", DTB must have "B2B Gender".
            # If mapping logic says "gender", DTB must have "Gender".
            pat = re.compile(r'\b' + re.escape(col_lower) + r'\b')

            for row in self.attr_rows:
                ml = _norm(row["mapping_logic"])
                if pat.search(ml):
                    col_to_attrs[dtb_col].append(row["attr_name"])

        # For each attribute, the first DTB column (by header order) wins
        attr_to_winner: dict[str, str] = {}
        for dtb_col in dtb_headers:
            for attr_name in col_to_attrs.get(dtb_col, []):
                if attr_name not in attr_to_winner:
                    attr_to_winner[attr_name] = dtb_col
                else:
                    log.warning(
                        "[AttrMap] Attribute '%s' already claimed by DTB col '%s' — "
                        "ignoring '%s' (pick-first; update Attributes List to resolve)",
                        attr_name, attr_to_winner[attr_name], dtb_col,
                    )

        pairs = [(dtb_col, attr_name) for attr_name, dtb_col in attr_to_winner.items()]
        log.info("[AttrMap] Final map: %d (dtb_col → attr) pairs", len(pairs))
        return pairs


class DTBLoader:
    """
    Reads the first sheet of the DTB Excel file (the 'English' product sheet).

    Row 0  → column headers
    Row 1+ → article data (one row per article; Article No. is the unique key)

    Stores:
        self.headers  : list[str]   — raw column header names from row 0
        self.articles : list[dict]  — one dict per valid article row
    """

    _ARTICLE_NO_CANDIDATES = frozenset(
        ["article no.", "article no", "article number", "article_no"]
    )

    def __init__(self, path: Path):
        self.path     = path
        self.headers:  list[str]  = []
        self.articles: list[dict] = []
        self._load()

    def _load(self):
        log.info("[DTB] Loading: %s", self.path.name)
        wb         = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        sheet_name = wb.sheetnames[0]
        log.info("[DTB] Using first sheet: '%s'", sheet_name)
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            log.error("[DTB] File is empty")
            return

        self.headers = [
            str(h).strip() if h is not None else f"__col_{i}__"
            for i, h in enumerate(rows[0])
        ]

        article_no_idx = next(
            (i for i, h in enumerate(self.headers)
             if h.lower().strip() in self._ARTICLE_NO_CANDIDATES),
            None,
        )
        if article_no_idx is None:
            log.error("[DTB] Cannot find Article No. column. Headers seen: %s",
                      self.headers[:20])
            return

        seen_articles: set[str] = set()
        for row in rows[1:]:
            raw_no = _s(row[article_no_idx]) if article_no_idx < len(row) else ""
            if not raw_no:
                continue
            if raw_no in seen_articles:
                log.debug("[DTB] Duplicate Article No. '%s' — skipping subsequent row", raw_no)
                continue
            seen_articles.add(raw_no)

            row_dict = {
                self.headers[i]: (row[i] if i < len(row) else None)
                for i in range(len(self.headers))
            }
            self.articles.append(row_dict)

        log.info("[DTB] %d unique articles loaded from '%s'", len(self.articles), sheet_name)


# ══════════════════════════════════════════════════════════════════
# SECTION 2 — RESOLVER
# ══════════════════════════════════════════════════════════════════

def resolve_attribute_map(
    pairs:      list[tuple[str, str]],
    mdd_lookup: MDDPIMLookup,
) -> list[dict]:
    """
    Final resolution step:

        (dtb_col_header, attr_name)  →  MDD  →  {at_id, validation_type, …}

    Returns list of resolved mappings. Pairs whose attr_name is absent from
    the MDD are logged as warnings and skipped; the rest continue.
    """
    resolved: list[dict] = []

    for dtb_col, attr_name in pairs:
        entry = mdd_lookup.lookup(attr_name)
        if entry:
            resolved.append({
                "dtb_col":         dtb_col,
                "at_id":           entry["at_id"],
                "validation_type": entry["validation_type"],
                "attr_name":       attr_name,
                "pim_name":        entry["pim_name"],
            })
            log.info(
                "[Resolve] '%s'  →  attr='%s'  →  %s  (%s)",
                dtb_col, attr_name, entry["at_id"], entry["validation_type"],
            )
        else:
            log.warning(
                "[Resolve] Attr '%s' (from DTB col '%s') NOT found in MDD — skipping",
                attr_name, dtb_col,
            )

    log.info("[Resolve] %d / %d pairs fully resolved to AT_xxx IDs",
             len(resolved), len(pairs))
    return resolved


# ══════════════════════════════════════════════════════════════════
# SECTION 3 — XML BUILDER
# ══════════════════════════════════════════════════════════════════

# ── XML element helpers ───────────────────────────────────────────────────────

def _val(parent, attr_id: str, value: str = "", id_val: str = ""):
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val  else ""
    clean_val = str(value).strip()  if value   else ""
    if clean_id and clean_id not in ("None", "nan", ""):
        el.set("ID", clean_id)
    elif clean_val and clean_val not in ("None", "nan", ""):
        el.text = clean_val
    return el


def _val_required(parent, attr_id: str, value: str = "", id_val: str = "",
                  fallback: str = ""):
    clean_id  = str(id_val).strip() if id_val  else ""
    clean_val = str(value).strip()  if value   else ""
    if not clean_id and not clean_val:
        clean_val = fallback
    return _val(parent, attr_id, clean_val, clean_id)


def _multival(parent, attr_id: str, id_val: str, label: str = ""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))


def _find_col_val(row_dict: dict, candidates: list[str]) -> str:
    """Return first non-empty value whose column header contains a candidate (case-insensitive).
    Ordered: more-specific candidates listed first in each list."""
    row_lower = {k.lower(): v for k, v in row_dict.items()}
    for candidate in candidates:
        cand_lower = candidate.lower()
        # exact key match
        if cand_lower in row_lower:
            val = _s(row_lower[cand_lower])
            if val:
                return val
        # candidate is substring of an actual column header
        for key in row_lower:
            if cand_lower in key:
                val = _s(row_lower[key])
                if val:
                    return val
    return ""


def _parse_season_dtb(season_raw: str) -> tuple[str, str, str]:
    """Parse season codes with 2-digit or 4-digit years.
    'FW25' → ('FW', '25', '2025')
    'CO2029' → ('CO', '29', '2029')
    Returns ('', '', '') on failure."""
    s = (season_raw or "").strip().upper()
    # Try 4-digit year first (e.g. CO2029, SS2026)
    m = re.match(r"^([A-Z]{2})(\d{4})$", s)
    if m:
        prefix    = m.group(1)
        full_year = m.group(2)
        short_year = full_year[2:]
        return (prefix, short_year, full_year)
    # Fall back to 2-digit year (e.g. FW25, SM26)
    m = re.match(r"^([A-Z]{2})(\d{2})$", s)
    if not m:
        return ("", "", "")
    prefix, short_year = m.group(1), m.group(2)
    full_year = f"20{short_year}"
    return (prefix, short_year, full_year)


def _get_division_code(division: str) -> str:
    """Map division string to single-letter parent code; returns 'X' if unmapped."""
    if not division:
        return "X"
    key = re.sub(r"[^A-Z]", "", division.upper())
    return DIVISION_PARENT_MAP.get(key, "X")


# ── defining-attribute writer ─────────────────────────────────────────────────

def _add_generic_values_dtb(
    vals_el,
    row_dict:    dict,
    mdd:         MDDPIMLookup,
    brand_code:  str,
    brand_name:  str,
    comp_code:   str,
    sbu:         str,
    article_no:  str,
    article_name: str,
    season_raw:  str,
    args = None,  # ← ADD THIS PARAMETER
) -> set:
    """Write all structural/defining attributes to vals_el.
    Returns the set of AT_xxx IDs written so the enrichment loop can skip duplicates.
    When MDD lookup returns None: log warning + write empty element (no wrong data sent)."""
    written: set[str] = set()

    # SBU
    sbu_lbl = mdd.sbu_label(sbu)
    if sbu_lbl is None:
        log.warning("[DTB] SBU code %r not in MDD — AT_SBU written without label", sbu)
        sbu_lbl = ""
    if sbu:
        _multival(vals_el, "AT_SBU", sbu, sbu_lbl)
    else:
        _val(vals_el, "AT_SBU")
    written.add("AT_SBU")

    # Company Code
    cc_lbl = mdd.company_code_label(comp_code)
    if cc_lbl is None:
        log.warning("[DTB] Company code %r not in MDD — AT_CompanyCode written without label", comp_code)
        cc_lbl = ""
    if comp_code:
        _multival(vals_el, "AT_CompanyCode", comp_code, cc_lbl)
    else:
        _val(vals_el, "AT_CompanyCode")
    written.add("AT_CompanyCode")

    # Brand
    b_lbl = mdd.brand_label(brand_code)
    if b_lbl is None:
        log.warning("[DTB] Brand code %r not in MDD Brand LOV — AT_Brand written without label", brand_code)
        b_lbl = brand_name or brand_code
    _val(vals_el, "AT_Brand",      b_lbl, id_val=brand_code)
    _val(vals_el, "AT_BrandGroup", b_lbl, id_val=b_lbl)
    written.update({"AT_Brand", "AT_BrandGroup"})

    # Color — use "B2B Base Color" description → MDD Color Code LOV
    # AT_PrincipalColorCode is NOT sent: attribute list says Source N/A for this field.
    # color_raw is still needed below for AT_Color LOV lookup.
    color_raw    = _find_col_val(row_dict, _DTB_COLOR_CANDIDATES)

    color_lov_id = mdd.color_to_lov_id(color_raw) if color_raw else ""
    if color_lov_id:
        _val(vals_el, "AT_Color", color_lov_id, id_val=color_lov_id)
    else:
        if color_raw:
            log.warning("[DTB] No color LOV match for %r (article %s) — AT_Color skipped",
                        color_raw, article_no)
        _val(vals_el, "AT_Color")
    written.add("AT_Color")

    # AT_PrincipalColorName — from "B2B Color1" column
    color_name_raw = _find_col_val(row_dict, ["b2b color1", "b2b color 1", "color1", "color 1"])
    if color_name_raw:
        _val(vals_el, "AT_PrincipalColorName", color_name_raw)
        written.add("AT_PrincipalColorName")

    # Season
    sea_prefix, short_year, full_year = _parse_season_dtb(season_raw)
    if sea_prefix and full_year:
        sea_label = mdd.season_label(sea_prefix)
        if sea_label is None:
            log.warning("[DTB] Season code %r not in MDD Season LOV", sea_prefix)
            sea_label = sea_prefix
        _val(vals_el, "AT_Season", sea_label, id_val=sea_prefix)
        _val(vals_el, "AT_SeasonYear", full_year)  # ← Changed: use _val instead of _val_required
        written.update({"AT_Season", "AT_SeasonYear"})
    else:
        if season_raw:
            log.warning("[DTB] Cannot parse season %r — AT_Season/AT_SeasonYear skipped", season_raw)
        # ← DON'T write empty elements — just skip

    # Generic key — computed from brand + article number
    _val(vals_el, "AT_InboundGenericCode", f"{brand_code}{article_no}")
    written.add("AT_InboundGenericCode")
    _val(vals_el, "AT_PrincipalStyleCode", article_no)
    _val(vals_el, "AT_SAPStyleCode", article_no)
    _val(vals_el, "AT_PrincipalStyleDescription", article_name or article_no)
    written.update({"AT_PrincipalStyleCode", "AT_SAPStyleCode", "AT_PrincipalStyleDescription"})

    # BYArticleType — from filename metadata (Inline / License / SSE)
    _LOV_BY_ARTICLE_TYPE = {
        "Inline": "Inline", "License": "License", "Sse": "SSE",
        "Licensed": "License",
    }
    at_raw   = getattr(args, 'article_type_from_filename', '') or ''
    at_norm  = at_raw.capitalize()
    at_label = _LOV_BY_ARTICLE_TYPE.get(at_norm, at_norm) or at_norm
    _val(vals_el, "AT_BYArticleType", at_label, id_val=at_norm)
    written.add("AT_BYArticleType")


    # ── Country (from args if available) ──────────────────────────────────────
    # NEW: Add country attribute from filename metadata
    country_code = getattr(args, 'country_code', '') if hasattr(args, 'country_code') else ''
    if country_code:
        _val(vals_el, "AT_Country", country_code, id_val=country_code)
        written.add("AT_Country")

    return written


# ── classifications builder ───────────────────────────────────────────────────

def _build_classifications_dtb(
    brand:      str,
    brand_code: str,
    comp_code:  str,
    sbu:        str,
    season_raw: str,
    mdd:        MDDPIMLookup,
) -> tuple:
    """Build <Classifications> block. Returns (cls_element, season_id)."""
    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")

    sea_prefix, short_year, full_year = _parse_season_dtb(season_raw)
    if not sea_prefix or not full_year:
        log.error(
            "[DTB] Cannot parse season %r — expected format 'FW25' / 'SM26'. "
            "Classifications block will be empty.", season_raw,
        )
        return ET.Element(f"{{{STIBO_NS}}}Classifications"), ""

    season_id      = f"CLH_{brand_code}_{sea_prefix}{full_year}"
    batches_parent = f"CLH_{brand.capitalize()}Batches"

    sea_label_str  = mdd.season_label(sea_prefix) or sea_prefix
    season_display = f"{brand} {sea_label_str} {full_year}".strip()
    season_short   = f"{sea_prefix} {full_year}".strip()

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID",         season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID",   batches_parent)
    season_cls.set("update",     "true")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_display

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID",         f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    confirmed.set("update",     "true")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID",         f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    unconfirmed.set("update",     "true")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root, season_id


def _exact_col_val(row_dict: dict, *col_names: str) -> str:
    """Return value from the first column whose name matches exactly (case-insensitive).
    No substring search — avoids 'B2C Copy' accidentally matching 'B2C Copy Short'."""
    row_lower = {k.lower(): v for k, v in row_dict.items()}
    for col in col_names:
        val = _s(row_lower.get(col.lower(), ""))
        if val:
            return val
    return ""


def _format_value(raw: str, validation_type: str) -> str:
    """
    Conservative formatting: return the raw string for all validation types.
    Stibo accepts text for all attribute types in an upsert context.
    Extend with per-type rules (date normalisation, numeric stripping, etc.)
    once the Stibo team confirms expected formats.
    """
    return raw


def build_dtb_xml(
    articles:         list[dict],
    resolved_map:     list[dict],
    article_no_col:   str,
    article_name_col: str | None,
    brand_code:       str,
    brand_name:       str,
    comp_code:        str,
    sbu:              str,
    season_raw:       str,
    mdd:              MDDPIMLookup,
    season_id:        str,
    args,
    out_path:         Path,
) -> int:
    """Write Stibo STEP XML with Classifications + Products to out_path.
    Returns count of articles written."""
    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written     = 0
    skipped     = 0

    # Build Classifications block once for the whole file
    cls_el, resolved_season_id = _build_classifications_dtb(
        brand      = brand_name,
        brand_code = brand_code,
        comp_code  = comp_code,
        sbu        = sbu,
        season_raw = season_raw,
        mdd        = mdd,
    )
    if season_id:
        resolved_season_id = season_id
    classifications_xml = _XMLNS_RE.sub(
        "", ET.tostring(cls_el, encoding="unicode")
    )

    # Determine division + segment IDs for ClassificationReferences
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
            f' UseContextLocale="false"'
            f' update="true">\n\n'
        )

        # Classifications block
        # f.write(f"  {classifications_xml}\n\n")
        f.write("  <Products>\n")

        # ── TEST MODE: limit to 5 articles ──────────────────────────
        # articles = articles[:5]
        for row_dict in articles:
            article_no = _s(row_dict.get(article_no_col, ""))
            if not article_no:
                skipped += 1
                continue

            article_name = (
                _s(row_dict.get(article_name_col, ""))
                if article_name_col else ""
            ) or article_no

            key_article = f"{brand_code}{article_no}"

            # Division / segment for ClassificationReferences
            division_raw = _find_col_val(row_dict, _DTB_DIVISION_CANDIDATES)
            biz_seg_raw  = _find_col_val(row_dict, _DTB_BIZ_SEGMENT_CANDIDATES)
            div_letter   = _get_division_code(division_raw)
            parent_id    = f"PPH_{div_letter}-TempSubCat"
            div_id       = re.sub(r"[^A-Z0-9]", "", division_raw.upper())[:2] if division_raw else "X"
            seg_id       = re.sub(r"[^A-Z0-9]", "", biz_seg_raw.upper())[:4]  if biz_seg_raw  else ""

            g_el = ET.Element(f"{{{STIBO_NS}}}Product")
            g_el.set("UserTypeID", "PRD_GenericArticle")
            g_el.set("ParentID",   parent_id)
            g_el.set("update",     "true")

            kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
            kv.set("KeyID", "KEY_InboundArticle")
            kv.text = key_article

            # ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = article_name

            # Four ClassificationReferences — same as tdd/backlog
            # brand_cap = brand_name.capitalize() if brand_name else brand_code.capitalize()
            # for cls_id, cpl_type in [
            #     (f"CLH_{brand_cap}Articles",             "CPL_Merchandiser"),
            #     (f"MA_{brand_code}_{div_id}_{seg_id}", "CPL_BYHierarchy"),
            #     (f"CLH_{div_id}",                      "CPL_SAPHierarchy"),
            #     (f"{resolved_season_id}UA",              "CPL_UnConfirmedForSeason"),
            # ]:
            #     cr = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
            #     cr.set("ClassificationID", cls_id)
            #     cr.set("Type", cpl_type)

            vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

            # Write defining/structural attributes first
            already_written = _add_generic_values_dtb(
                vals_el      = vals_el,
                row_dict     = row_dict,
                mdd          = mdd,
                brand_code   = brand_code,
                brand_name   = brand_name,
                comp_code    = comp_code,
                sbu          = sbu,
                article_no   = article_no,
                article_name = article_name,
                season_raw   = season_raw,
                args         = args,  # ← ADD THIS LINE
            )

            # Short/Long description: always use explicit column names + B2B→B2C fallback.
            # Written BEFORE the enrichment loop so the loop skips these AT_IDs regardless
            # of which column the dynamic mapper happened to assign (pick-first can get it wrong
            # when the mapping-logic cell mentions multiple column names).
            enrichment_written: set[str] = set()
            for at_id, b2b_col, b2c_col in [
                ("AT_ShortDescriptionEN", "B2B Copy Short", "B2C Copy Short"),
                ("AT_LongDescriptionEN",  "B2B Copy",       "B2C Copy"),
            ]:
                if at_id in already_written:
                    continue
                val = _exact_col_val(row_dict, b2b_col)
                if not val:
                    val = _exact_col_val(row_dict, b2c_col)
                    if val:
                        log.debug("[DTB] %s: %r empty — using %r (article %s)",
                                  at_id, b2b_col, b2c_col, article_no)
                if val:
                    val_el = ET.SubElement(vals_el, f"{{{STIBO_NS}}}Value")
                    val_el.set("AttributeID", at_id)
                    val_el.text = val
                    enrichment_written.add(at_id)

            # Special column rules — compose multi-source attributes (e.g. L×W×H, care labels).
            # Run BEFORE the enrichment loop so their AT_IDs are marked written and the loop
            # cannot overwrite them with a wrong single-column value.
            for rule in SPECIAL_COLUMN_RULES:
                mdd_entry = mdd.lookup(rule["mdd_attr"])
                if not mdd_entry:
                    log.warning("[DTB] SPECIAL_RULE '%s': attr '%s' not in MDD — skipping",
                                rule["rule_id"], rule["mdd_attr"])
                    continue
                at_id = mdd_entry["at_id"]
                if at_id in already_written or at_id in enrichment_written:
                    continue

                if rule.get("multivalue"):
                    # Collect all non-empty values from source columns → MultiValue element
                    non_empty = [
                        _find_col_val(row_dict, [col])
                        for col in rule["source_norms"]
                    ]
                    non_empty = [v for v in non_empty if v]
                    if non_empty:
                        mv = ET.SubElement(vals_el, f"{{{STIBO_NS}}}MultiValue")
                        mv.set("AttributeID", at_id)
                        for v in non_empty:
                            ve = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
                            ve.text = v
                        enrichment_written.add(at_id)
                else:
                    src_vals = [_find_col_val(row_dict, [col]) for col in rule["source_norms"]]
                    if all(src_vals):
                        composed = rule["compose"](src_vals)
                        val_el = ET.SubElement(vals_el, f"{{{STIBO_NS}}}Value")
                        val_el.set("AttributeID", at_id)
                        val_el.text = composed
                        enrichment_written.add(at_id)

            # Write remaining DTB enrichment attributes, skipping already-written ones
            _skip_ids = {"AT_ShortDescriptionEN", "AT_LongDescriptionEN"}
            for rule in SPECIAL_COLUMN_RULES:
                entry = mdd.lookup(rule["mdd_attr"])
                if entry:
                    _skip_ids.add(entry["at_id"])
            for mapping in resolved_map:
                at_id = mapping["at_id"]
                if at_id in already_written or at_id in enrichment_written or at_id in _skip_ids:
                    continue
                raw_val = _s(row_dict.get(mapping["dtb_col"]))
                if not raw_val:
                    continue
                formatted = _format_value(raw_val, mapping["validation_type"])
                val_el    = ET.SubElement(vals_el, f"{{{STIBO_NS}}}Value")
                val_el.set("AttributeID", at_id)
                val_el.text = formatted
                enrichment_written.add(at_id)

            product_xml = _XMLNS_RE.sub("", ET.tostring(g_el, encoding="unicode"))
            f.write(f"    {product_xml}\n")
            written += 1

        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")

    log.info("[XML] Articles written=%d  skipped=%d", written, skipped)
    return written


# ══════════════════════════════════════════════════════════════════
# SECTION 4 — ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════

def run(args, auditor=None, tmp_dir: Path = None):
    """
    Main entry point for DTB ecommerce enrichment ETL.

    args must have: brand, brand_code, comp_code, sbu, season, seq, multi_mono
    """
    if tmp_dir:
        dtb_dir     = tmp_dir / "input" / "dtb"
        mdd_dir     = tmp_dir / "input" / "mdd"
        attr_dir    = tmp_dir / "input" / "attributes"
        xml_out_dir = tmp_dir / "output" / "xml"
        log_dir     = tmp_dir / "output" / "logs"
    else:
        dtb_dir     = DTB_DIR
        mdd_dir     = MDD_DIR
        attr_dir    = ATTR_DIR
        xml_out_dir = XML_OUT_DIR
        log_dir     = LOG_DIR

    for _d in [dtb_dir, mdd_dir, attr_dir, xml_out_dir, log_dir]:
        _d.mkdir(parents=True, exist_ok=True)

    def first_xlsx(d: Path) -> Path | None:
        files = list(d.glob("*.xlsx"))
        return files[0] if files else None

    dtb_f  = first_xlsx(dtb_dir)
    mdd_f  = first_xlsx(mdd_dir)
    attr_f = first_xlsx(attr_dir)

    required = {"DTB": dtb_f, "MDD": mdd_f, "Attributes List": attr_f}
    for label, val in required.items():
        if not val:
            log.error("[DTB-ETL] %s file missing — aborting", label)
            if auditor:
                auditor.set_etl_status("error", error=f"{label} file not found")
            return

    log.info("[DTB-ETL] Starting  brand=%s  season=%s  seq=%s",
             args.brand, args.season, args.seq)

    # ── 1. MDD — build PIM attribute name → AT_xxx lookup ────────────────────
    try:
        mdd_lookup = MDDPIMLookup(mdd_f)
        if auditor:
            auditor.record_loader("mdd_pim_lookup", status="ok",
                                  attributes=len(mdd_lookup.by_name))
    except Exception as exc:
        log.error("[DTB-ETL] MDD load failed: %s", exc, exc_info=True)
        if auditor:
            auditor.record_loader("mdd_pim_lookup", status="error", error=str(exc))
            auditor.set_etl_status("error", error=f"MDD load failed: {exc}")
        return

    # ── 2. DTB file ───────────────────────────────────────────────────────────
    try:
        dtb = DTBLoader(dtb_f)
        dtb.articles

        if not dtb.articles or not dtb.headers:
            log.error("[DTB-ETL] DTB file produced no usable data — aborting")
            if auditor:
                auditor.record_loader("dtb_loader", status="error", error="empty")
                auditor.set_etl_status("error", error="DTB file empty or unparseable")
            return
        if auditor:
            auditor.record_loader("dtb_loader", status="ok",
                                  articles=len(dtb.articles),
                                  columns=len(dtb.headers))
    except Exception as exc:
        log.error("[DTB-ETL] DTB load failed: %s", exc, exc_info=True)
        if auditor:
            auditor.record_loader("dtb_loader", status="error", error=str(exc))
            auditor.set_etl_status("error", error=f"DTB load failed: {exc}")
        return

    # ── 3. Attributes List → DTB column map ──────────────────────────────────
    try:
        attr_loader = AdidasAttributeMapLoader(attr_f)
        dtb_pairs   = attr_loader.build_dtb_column_map(dtb.headers)
        if auditor:
            auditor.record_loader("attr_map_loader", status="ok",
                                  dtb_attr_rows=len(attr_loader.attr_rows),
                                  mapped_columns=len(dtb_pairs))
    except Exception as exc:
        log.error("[DTB-ETL] Attribute map load failed: %s", exc, exc_info=True)
        if auditor:
            auditor.record_loader("attr_map_loader", status="error", error=str(exc))
            auditor.set_etl_status("error", error=f"Attribute map failed: {exc}")
        return

    if not dtb_pairs:
        log.warning("[DTB-ETL] Zero DTB columns matched attribute rows — "
                    "check Attributes List 'Field Name / Mapping Logic' content")
        if auditor:
            auditor.set_etl_status("ok", error="no_dtb_columns_matched")
        return

    # ── 4. Resolve dtb_col → attr_name → AT_xxx via MDD ─────────────────────
    resolved_map = resolve_attribute_map(dtb_pairs, mdd_lookup)

    if not resolved_map:
        log.warning("[DTB-ETL] No DTB columns resolved to AT_xxx IDs — "
                    "verify MDD PIM Attribute Names match Attributes List entries")
        if auditor:
            auditor.set_etl_status("ok", error="no_attributes_resolved")
        return

    # ── 5. Locate Article No. and Article Name column headers ─────────────────
    article_no_col = next(
        (h for h in dtb.headers
         if h.strip().lower() in DTBLoader._ARTICLE_NO_CANDIDATES),
        None,
    )
    if not article_no_col:
        log.error("[DTB-ETL] Article No. column not found in DTB headers — aborting")
        if auditor:
            auditor.set_etl_status("error", error="Article No. column missing from DTB")
        return

    article_name_col = next(
        (h for h in dtb.headers if "article name" in h.strip().lower()),
        None,
    )

    # ── 6. Output filename: match input filename exactly ──────────────────────
    dtb_input_file = dtb_f  # Path object from first_xlsx(dtb_dir)
    input_stem = dtb_input_file.stem  
    out_name = f"{input_stem}.xml"
    out_path = xml_out_dir / out_name

    # Pre-build season_id — requires 'FWXX' / 'SMXX' format (e.g. 'FW25', 'SM26')
    sea_prefix, _, full_year = _parse_season_dtb(getattr(args, "season", "") or "")
    if not sea_prefix or not full_year:
        log.error(
            "[DTB-ETL] args.season=%r cannot be parsed — expected format 'FW25' / 'SM26'. "
            "Pass the season with a 2-digit year suffix.", getattr(args, "season", "")
        )
        if auditor:
            auditor.set_etl_status("error", error=f"Unparseable season: {getattr(args, 'season', '')!r}")
        return
    season_id = f"CLH_{args.brand_code}_{sea_prefix}{full_year}"

    log.info(
        "[DTB-ETL] Building XML → %s  (%d articles  ×  %d resolved attributes)  season=%s",
        out_name, len(dtb.articles), len(resolved_map), season_id,
    )

    # ── 7. Build XML ──────────────────────────────────────────────────────────
    try:
        written = build_dtb_xml(
            articles         = dtb.articles,
            resolved_map     = resolved_map,
            article_no_col   = article_no_col,
            article_name_col = article_name_col,
            brand_code       = args.brand_code,
            brand_name       = getattr(args, "brand", args.brand_code),
            comp_code        = getattr(args, "comp_code", ""),
            sbu              = getattr(args, "sbu", ""),
            season_raw       = getattr(args, "season", ""),
            mdd              = mdd_lookup,
            season_id        = season_id,
            args             = args,
            out_path         = out_path,
        )
    except Exception as exc:
        log.error("[DTB-ETL] XML build failed: %s", exc, exc_info=True)
        if auditor:
            auditor.set_etl_status("error", error=f"XML build failed: {exc}")
        return

    file_kb = out_path.stat().st_size // 1024 if out_path.exists() else 0

    log.info(
        "[DTB-ETL] ✓  XML → %s  (%d KB)  |  articles written=%d  skipped=%d",
        out_name, file_kb, written, len(dtb.articles) - written,
    )

    print("═══ DTB ENRICHMENT SUMMARY ════════════════════════════", flush=True)
    print(f"  DTB articles in file  : {len(dtb.articles)}", flush=True)
    print(f"  Resolved attributes   : {len(resolved_map)}", flush=True)
    print(f"  Articles written      : {written}", flush=True)
    print(f"  Articles skipped      : {len(dtb.articles) - written}", flush=True)
    print(f"  XML file              : {out_name}  ({file_kb} KB)", flush=True)
    print("═══════════════════════════════════════════════════════", flush=True)

    if auditor:
        auditor.set_etl_counts(
            mapped      = len(dtb.articles),
            written     = written,
            skipped     = len(dtb.articles) - written,
            variants    = 0,
            xml_file    = out_name,
            xml_size_kb = file_kb,
        )
        auditor.set_etl_status("ok")
