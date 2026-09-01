"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  STIBO INBOUND XML GENERATOR — implus Sofsole Line List v1.0               ║
║  Sub-brand source: "IPL_Article Attributes Source_Sofsole.xlsx"            ║
║  → Stibo STEP XML                                                           ║
╚══════════════════════════════════════════════════════════════════════════════╝

Reference: Brand Mapping File → "IMPLUS" tab
  Col A  Stibo attribute name
  Col B  Stibo attribute ID
  Col C  LOV → send as ID  /  Text → send as value
  Col G  "Direct from Principal"        → include in XML (Col J = source column name)
         "Formula in system" + Col J    → include in XML (Col J = source column name)
         otherwise                      → skip / system-managed
  Col J  Column name in source Excel
  Col K  Mapping logic

Source file  : Sofsole "SUM ORDER" sheet Excel (.xlsx / .xlsm)
Sheet name   : "SUM ORDER"

Attribute → source column mapping (per IMPLUS brand mapping tab):
──────────────────────────────────────────────────────────────────────────────
Stibo Attribute ID                      Type  Source Column / Logic
──────────────────────────────────────────────────────────────────────────────
AT_InboundGenericCode                   Text  brand_code + IMPLUS EU ITEM #  (formula)
AT_PrincipalStyleCode                   Text  IMPLUS EU ITEM #
AT_PrincipalStyleDescription            Text  L1 + L2 + DESCRIPTION + SIZE  (formula)
AT_PrincipalColorName                   Text  COLOR
AT_PrincipalColorCode                   Text  COLOR  (no separate code column)
AT_PrincipalSize                        Text  SIZE
AT_PrincipalMerchandiseHierarchyL1      Text  CATEGORY  (Col A top-level banner)
AT_PrincipalMerchandiseHierarchyL2      Text  sub-category banner (Col B section header)
AT_CountryOrigin                        LOV   COUNTRY OF ORIGIN  → MDD Country Origin LOV
AT_FOB                                  Text  DISTRIBUTOR PRICE / COST
AT_FOBCurrency                          LOV   USD  (default)
AT_Brand                                LOV   brand_name → MDD Brand LOV
AT_BrandGroup                           LOV   RNA lookup
AT_BrandType                            LOV   RNA lookup
AT_BrandStatus                          LOV   A  (default Active)
AT_Season                               LOV   season_prefix from filename
AT_SeasonYear                           Text  season_year from filename
AT_Country                              LOV   country_code from filename
AT_CompanyCode                          LOV   comp_code from filename
AT_SBU                                  LOV   sbu from filename
AT_RetailPriceCurrency                  LOV   derived from country_code
AT_Gender                               LOV   U  (default Unisex)
AT_SAPAge                               LOV   AD (default Adults)
AT_BYGender                             LOV   U  (default Unisex)
AT_BYAge                                LOV   ADULT (default)
AT_BYArticleType                        LOV   Inline (default)
AT_PackDetails                          LOV   Single (default)
AT_SAPProductFlag                       LOV   A  (default Intercompany)
AT_MaterialType                         LOV   ZINA (default Intercompany)
AT_SAPArticleCategory                   LOV   01 (default Generic/Variant)
AT_UOM                                  LOV   EA (default)
AT_BCI                                  LOV   COMMERCIAL (default)
AT_ArticleStatus                        LOV   A  (default Active)
AT_PricingDistributionChannel           LOV   01 (default)
──────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import openpyxl

# ── Reuse every brand-agnostic helper from the reference implementation ──────
try:
    from implus.linelist_main_source_balega import (
        STIBO_NS,
        _val,
        _keyval,
        _multival,
        open_step_xml,
        close_step_xml,
        _get_dirs,
        _find_input_files,
        _s,
        MDDLoader,
        _find_mdd_file,
        BRAND_LOV_MAP,
        _resolve_brand_group,
        _resolve_brand_type,
        _resolve_pack_details,
        _resolve_coo,
        build_classifications,
        RNALoader,
        _find_rna_source_file,
        _COUNTRY_MAP,
    )
except ImportError:  # pragma: no cover — allows direct `python implus/linelist_main_source_sofsole.py`
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from implus.linelist_main_source_balega import (
        STIBO_NS,
        _val,
        _keyval,
        _multival,
        open_step_xml,
        close_step_xml,
        _get_dirs,
        _find_input_files,
        _s,
        MDDLoader,
        _find_mdd_file,
        BRAND_LOV_MAP,
        _resolve_brand_group,
        _resolve_brand_type,
        _resolve_pack_details,
        _resolve_coo,
        build_classifications,
        RNALoader,
        _find_rna_source_file,
        _COUNTRY_MAP,
    )

log = logging.getLogger("implus.linelist_main_source_sofsole")


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

# Retail price currency per destination country code
_RETAIL_CURRENCY_MAP: dict[str, str] = {
    "ID": "IDR",
    "PH": "PHP",
    "TH": "THB",
    "SG": "SGD",
    "MY": "MYR",
    "VN": "VND",
    "KH": "USD",
}

# Header signal words used to detect the header row in "SUM ORDER" sheet
# (Col A = Category, Col B = IMPLUS EU ITEM # or similar)
_HEADER_SIGNALS = {"category", "implus eu", "implus eu item", "upc", "ean", "product description"}


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class SofsoleRow:
    """One data row from the Sofsole SUM ORDER sheet."""
    row_num:           int
    major_category:    str = ""   # Col A top-level "CATEGORY" banner → AT_PrincipalMerchandiseHierarchyL1
    sub_category:      str = ""   # Col B section-header text          → AT_PrincipalMerchandiseHierarchyL2
    implus_us:         str = ""   # IMPLUS US ITEM #
    implus_eu:         str = ""   # IMPLUS EU ITEM # / SKU             → AT_PrincipalStyleCode
    upc:               str = ""   # UPC barcode
    ean:               str = ""   # EAN barcode
    style_desc:        str = ""   # PRODUCT DESCRIPTION                → AT_PrincipalStyleDescription
    size:              str = ""   # SIZE                               → AT_PrincipalSize
    color:             str = ""   # COLOR                              → AT_PrincipalColorName / AT_PrincipalColorCode
    country_of_origin: str = ""   # COUNTRY OF ORIGIN                  → AT_CountryOrigin (LOV)
    status:            str = ""   # STATUS
    distributor_price: str = ""   # DISTRIBUTOR PRICE / COST           → AT_FOB
    case_pack_qty:     str = ""   # CASE PACK QTY


@dataclass
class SofsoleGenericGroup:
    """One Generic Article = one (major_category + style description + color) combination."""
    style_key:     str            # internal grouping key
    style_desc:    str
    hierarchy_l1:  str            # AT_PrincipalMerchandiseHierarchyL1
    hierarchy_l2:  str = ""       # AT_PrincipalMerchandiseHierarchyL2
    hierarchy_l3:  str = ""       # AT_PrincipalMerchandiseHierarchyL3
    color:         str = ""       # AT_PrincipalColorName
    color_code:    str = ""       # AT_PrincipalColorCode
    implus_eu:     str = ""       # first variant SKU → anchors KEY_InboundArticle & AT_PrincipalStyleCode
    variants:      list[SofsoleRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN FINDER
# ══════════════════════════════════════════════════════════════════════════════

def _find_col(header: list[str], *candidates: str) -> int | None:
    """
    Return the 0-based index of the first header cell that matches any candidate
    (case-insensitive, whitespace-normalised — collapses newlines/repeated spaces).
    Returns None if none match.
    """
    normalised = [re.sub(r"\s+", " ", str(h)).strip().lower() for h in header]
    for candidate in candidates:
        target = re.sub(r"\s+", " ", candidate).strip().lower()
        for i, h in enumerate(normalised):
            if h == target or target in h:
                return i
    return None


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL LOADER
# ══════════════════════════════════════════════════════════════════════════════

def _is_black_fill(cell, log_ctx: str = "") -> bool:
    """
    True if a cell's background fill is black/near-black.

    Used to tell apart the two banner levels in Column A, which has no
    dedicated "Category" header — the top-level Category banner row (e.g.
    "INSOLES") is black-filled, while sub-category banner rows (e.g.
    "PERFORM", "SUPPORT") use other colors (yellow, green, ...).

    Handles all 3 ways Excel can store a fill color: explicit RGB, a theme
    color (e.g. "Black, Text 1" = theme index 1), or a legacy indexed color.
    """
    try:
        fill = cell.fill
        if fill is None or fill.patternType is None:
            log.info("[Sofsole][fill]%s no pattern fill → not black", log_ctx)
            return False

        fg      = fill.fgColor
        c_type  = getattr(fg, "type", None)
        rgb     = getattr(fg, "rgb", None)
        theme   = getattr(fg, "theme", None)
        tint    = getattr(fg, "tint", None)
        indexed = getattr(fg, "indexed", None)
        log.info(
            "[Sofsole][fill]%s pattern=%s type=%s rgb=%s theme=%s tint=%s indexed=%s",
            log_ctx, fill.patternType, c_type, rgb, theme, tint, indexed,
        )

        is_black = False

        # 1. Explicit RGB (e.g. "FF000000")
        if isinstance(rgb, str) and len(rgb) >= 6:
            hex6 = rgb[-6:]
            try:
                r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
                is_black = r < 40 and g < 40 and b < 40
            except ValueError:
                pass

        # 2. Theme color — theme index 1 = "Black, Text 1" in the default
        #    Office theme (0=Background1/white, 1=Text1/black, ...).
        if not is_black and c_type == "theme" and theme == 1 and (not tint or abs(tint) < 0.1):
            is_black = True

        # 3. Legacy indexed palette — 8 = black, 64 = automatic (often black text/fill)
        if not is_black and indexed in (8, 64):
            is_black = True

        log.info("[Sofsole][fill]%s → is_black=%s", log_ctx, is_black)
        return is_black
    except Exception as exc:
        log.info("[Sofsole][fill]%s error checking fill: %s", log_ctx, exc)
        return False


def load_linesheet(path: Path) -> list[SofsoleRow]:
    """
    Load the "SUM ORDER" sheet from the Sofsole source Excel.

    Header row detection:
      Scan the first 30 rows; the header row is the first one whose cells
      contain at least two of the _HEADER_SIGNALS keywords.

    Data row detection:
      A row is a valid data row if it has a non-empty EAN **or** UPC value.
      Rows without EAN/UPC are treated as banner rows: Column A holds the
      banner text, and its fill color tells L1 (Category, black-filled)
      apart from L2 (sub-category, any other colored fill) — this sheet has
      no dedicated "Category" column, so a real Category column (if found)
      still takes priority when present.
    """
    log.info("[Sofsole] Loading excel: %s", path.name)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:
        log.error("[Sofsole] Cannot open workbook: %s", exc)
        return []

    # ── Select sheet ────────────────────────────────────────────────────────
    if "SUM ORDER" in wb.sheetnames:
        ws = wb["SUM ORDER"]
        log.info("[Sofsole] Using sheet: 'SUM ORDER'")
    else:
        ws = wb.active
        log.warning("[Sofsole] 'SUM ORDER' sheet not found — using active sheet: '%s'", ws.title)

    all_rows = list(ws.iter_rows(values_only=True))

    if not all_rows:
        log.warning("[Sofsole] Sheet is empty")
        wb.close()
        return []

    # ── Locate header row ────────────────────────────────────────────────────
    header_row_idx = -1
    for i, row in enumerate(all_rows[:30]):
        if not row:
            continue
        cell_vals = {str(v).strip().lower() for v in row if v is not None}
        # Check how many signal words are present
        matches = sum(1 for sig in _HEADER_SIGNALS if any(sig in cv for cv in cell_vals))
        if matches >= 2:
            header_row_idx = i
            break

    if header_row_idx == -1:
        log.warning("[Sofsole] Could not detect header row — defaulting to row index 19 (row 20)")
        header_row_idx = 19

    header = [str(h).strip() if h is not None else "" for h in all_rows[header_row_idx]]
    log.info("[Sofsole] Header row at index %d: %s", header_row_idx, header[:20])

    # ── Map column indices from header names (per IMPLUS brand mapping tab) ──
    #   Col J in brand mapping = source column name to use for each attribute
    c_category   = _find_col(header, "Category", "Major Category", "CATEGORY")
    c_implus_eu  = _find_col(header, "IMPLUS EU ITEM #", "IMPLUS EU ITEM", "IMPLUS EU", "SKU", "ITEM #", "ITEM NO")
    c_upc        = _find_col(header, "UPC", "UPC CODE")
    c_implus_us  = _find_col(header, "IMPLUS US ITEM #", "IMPLUS US ITEM", "IMPLUS US")
    c_ean        = _find_col(header, "EAN", "EAN CODE", "EAN/UPC")
    c_style_desc = _find_col(header, "DESCRIPTION", "PRODUCT DESCRIPTION", "STYLE DESCRIPTION", "STYLE NAME", "PRODUCE NAME")
    c_size       = _find_col(header, "SIZE", "PRINCIPAL SIZE")
    c_color      = _find_col(header, "COLOR", "COLOUR", "COLOR NAME")
    c_coo        = _find_col(header, "COUNTRY OF ORIGIN", "COUNTRY", "SOURCE COUNTRY", "ORIGIN")
    c_status     = _find_col(header, "STATUS")
    c_dist_price = _find_col(header, "DISTRIBUTOR  PRICE (USD)")
    c_case_pack  = _find_col(header, "CASE PACK QTY", "CASE PACK", "PACK QTY")

    log.info(
        "[Sofsole] Column indices → category=%s  implus_eu=%s  upc=%s  ean=%s  "
        "style_desc=%s  size=%s  color=%s  coo=%s  dist_price=%s",
        c_category, c_implus_eu, c_upc, c_ean,
        c_style_desc, c_size, c_color, c_coo, c_dist_price,
    )

    def _cell(row, idx) -> str:
        if idx is None or idx >= len(row) or row[idx] is None:
            return ""
        return _s(row[idx])

    rows: list[SofsoleRow] = []
    current_l1 = ""
    current_l2 = ""

    for row_idx, row in enumerate(all_rows[header_row_idx + 1:], start=header_row_idx + 2):
        if not row or all(v is None for v in row):
            continue  # completely blank row

        cat_val = _cell(row, c_category)
        b_val   = _cell(row, c_implus_eu)   # used both as SKU and sub-category banner

        # Skip total / summary rows
        if "TOTAL" in cat_val.upper() or "TOTAL" in b_val.upper():
            continue

        # Update L1 whenever a Category value is present on the row — merged
        # Category cells only carry a value on their first row, which may
        # itself be a data row rather than a separate banner row.
        if cat_val and cat_val != current_l1:
            current_l1 = cat_val
            current_l2 = ""      # reset sub-category on L1 change

        ean_val = _cell(row, c_ean)
        upc_val = _cell(row, c_upc)
        is_data_row = bool(ean_val) or bool(upc_val)

        if not is_data_row:
            # Treat as banner row. If a dedicated Category column already set
            # L1 above, Column A is just the sub-category (L2). Otherwise
            # (this sheet has no Category column) use Column A's fill color:
            # black-filled = top-level Category (L1), any other color = L2.
            if b_val:
                if not cat_val and c_implus_eu is not None:
                    banner_cell = ws.cell(row=row_idx, column=c_implus_eu + 1)
                    log_ctx = f" row={row_idx} text='{b_val}'"
                    if _is_black_fill(banner_cell, log_ctx):
                        if b_val != current_l1:
                            current_l1 = b_val
                            current_l2 = ""
                            log.info("[Sofsole] L1 set → '%s' (row %d)", current_l1, row_idx)
                        continue
                current_l2 = b_val
                log.info("[Sofsole] L2 set → '%s' (row %d, L1='%s')", current_l2, row_idx, current_l1)
            continue

        style_desc = _cell(row, c_style_desc) or b_val

        rows.append(SofsoleRow(
            row_num=row_idx,
            major_category=current_l1,
            sub_category=current_l2,
            implus_us=_cell(row, c_implus_us),
            implus_eu=_cell(row, c_implus_eu),
            upc=upc_val,
            ean=ean_val,
            style_desc=style_desc,
            size=_cell(row, c_size),
            color=_cell(row, c_color),
            country_of_origin=_cell(row, c_coo),
            status=_cell(row, c_status),
            distributor_price=_cell(row, c_dist_price),
            case_pack_qty=_cell(row, c_case_pack),
        ))

    wb.close()
    with_l1 = sum(1 for r in rows if r.major_category)
    log.info(
        "[Sofsole] Parsed %d valid data rows (%d with L1 set, %d missing L1)",
        len(rows), with_l1, len(rows) - with_l1,
    )
    if rows[:3]:
        for r in rows[:3]:
            log.info(
                "[Sofsole] Sample row %d: L1='%s' L2='%s' desc='%s'",
                r.row_num, r.major_category, r.sub_category, r.style_desc,
            )
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# GROUPER  —  flat rows → Generic articles (one generic per style + color)
# ══════════════════════════════════════════════════════════════════════════════

def group_generics(rows: list[SofsoleRow]) -> dict[str, dict[str, SofsoleGenericGroup]]:
    """
    Group rows into { style_key → { color → SofsoleGenericGroup } }.

    Grouping key:  IMPLUS EU ITEM # (falls back to IMPLUS US ITEM #, then row_num)
    Color key:     COLOR value (or "NC" if absent)

    Each IMPLUS EU ITEM # in this source is a distinct article (one number
    per size break), so it is the grouping key — the same number that
    anchors AT_PrincipalStyleCode and KEY_InboundArticle (brand_code + implus_eu).
    Grouping by style description instead would wrongly collapse every size
    break of the same style into one product.
    """
    groups: dict[str, dict[str, SofsoleGenericGroup]] = defaultdict(dict)

    for r in rows:
        style_desc = r.style_desc or f"UNKNOWN_ROW_{r.row_num}"
        style_key  = r.implus_eu or r.implus_us or f"UNKNOWN_ROW_{r.row_num}"
        color      = r.color or "NC"

        if color not in groups[style_key]:
            groups[style_key][color] = SofsoleGenericGroup(
                style_key=style_key,
                style_desc=style_desc,
                hierarchy_l1=r.major_category,
                hierarchy_l2=r.sub_category,
                hierarchy_l3=style_desc,
                color=r.color,
                color_code=r.color,   # no separate color-code column in this source
                implus_eu=r.implus_eu,
            )
        groups[style_key][color].variants.append(r)

    log.info(
        "[Sofsole] Grouped %d rows → %d style keys / %d generic(s)",
        len(rows),
        len(groups),
        sum(len(cm) for cm in groups.values()),
    )
    return groups


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILDER  —  one Generic Article element per SofsoleGenericGroup
# ══════════════════════════════════════════════════════════════════════════════

def _build_generic_product(group: SofsoleGenericGroup, cfg: dict) -> ET.Element:
    """
    Build a <Product UserTypeID="PRD_GenericArticle"> element for one
    Sofsole generic article.

    Attribute mapping follows the IMPLUS brand mapping tab:
      • "Direct from Principal" + Col J column name → reads from SofsoleRow
      • "Formula in system" + Col J column name     → computed / derived value
      • LOV attributes → sent as ID  (col C = "LOV")
      • Text attributes → sent as text value (col C = "Text")
    """
    brand_code   = cfg.get("brand_code",   "IPL")
    brand_name   = cfg.get("brand_name",   "IMPLUS")
    sea_prefix   = cfg.get("season_prefix", "SM")
    sea_year     = cfg.get("season_year") or str(datetime.now().year)
    country_code = cfg.get("country_code", "ID")
    mdd: Optional[MDDLoader] = cfg.get("mdd")

    # ── KEY_InboundArticle (formula: brand_code + IMPLUS EU ITEM #) ──────────
    # Col G = "Formula in system", Col J = "IMPLUS EU ITEM #"
    generic_key = f"{brand_code}{group.implus_eu}"

    first_var = group.variants[0] if group.variants else None

    # ── AT_PrincipalStyleDescription formula: L1 + L2 + input DESCRIPTION + SIZE ─
    combined_style_desc = " ".join(
        part for part in (
            group.hierarchy_l1, group.hierarchy_l2, group.style_desc,
            first_var.size if first_var else "",
        ) if part
    ).strip()

    # ── Product element ───────────────────────────────────────────────────────
    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_E-TempSubCat")   # Sofsole = accessories/insoles

    _keyval(gen_el, "KEY_InboundArticle", generic_key)

    name_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name")
    name_el.text = combined_style_desc

    # ── Classification References ─────────────────────────────────────────────
    # Parent-level: Merchandiser classification (brand articles)
    cr_merch = ET.SubElement(gen_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", f"CLH_{brand_name}Articles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # Product-level: Season classification (unconfirmed for season)
    season_id = f"CLH_{brand_code}_{sea_prefix}{sea_year}"
    cr_unconf = ET.SubElement(gen_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")

    # ── AT_InboundGenericCode  (formula: brand_code + IMPLUS EU ITEM #) ──────
    _val(vals_el, "AT_InboundGenericCode", generic_key)

    # ── AT_PrincipalStyleCode  (Direct from Principal — IMPLUS EU ITEM #) ────
    _val(vals_el, "AT_PrincipalStyleCode", group.implus_eu)

    # ── AT_PrincipalStyleDescription  (formula: L1 + L2 + DESCRIPTION + SIZE) ─
    _val(vals_el, "AT_PrincipalStyleDescription", combined_style_desc)

    # ── AT_PrincipalColorName  (Direct from Principal — COLOR) ───────────────
    # DISABLED: Not sending AT_PrincipalColorName per user request
    # if group.color:
    #     _val(vals_el, "AT_PrincipalColorName", group.color)

    # ── AT_PrincipalColorCode  (Direct from Principal — COLOR, no separate code col) ─
    # DISABLED: Not sending AT_PrincipalColorCode per user request
    # if group.color_code:
    #     _val(vals_el, "AT_PrincipalColorCode", group.color_code)

    # ── AT_PrincipalSize  (Direct from Principal — SIZE, first variant) ──────
    if first_var and first_var.size:
        _val(vals_el, "AT_PrincipalSize", first_var.size)

    # ── AT_PrincipalMerchandiseHierarchyL1  (Direct from Principal — CATEGORY) ─
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", group.hierarchy_l1)

    # ── AT_PrincipalMerchandiseHierarchyL2  (Direct from Principal — sub-category) ─
    if group.hierarchy_l2:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", group.hierarchy_l2)

    # ── AT_PrincipalMerchandiseHierarchyL3  (formula: L1 + L2 + DESCRIPTION + SIZE) ─
    if combined_style_desc:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", combined_style_desc)

    # ── AT_CountryOrigin  (LOV default = HK) ──────────────────────────────────
    _val(vals_el, "AT_CountryOrigin", "", id_val="HK")

    # ── AT_FOB  (Direct from Principal — DISTRIBUTOR PRICE / COST) ───────────
    if first_var and first_var.distributor_price:
        _val(vals_el, "AT_FOB", first_var.distributor_price)

    # ── AT_FOBCurrency  (LOV default = USD) ──────────────────────────────────
    _val(vals_el, "AT_FOBCurrency", "", id_val="USD")

    # ── AT_Brand  (LOV — brand_code in cfg IS the resolved MDD LOV ID) ───────
    _val(vals_el, "AT_Brand", "", id_val=brand_code)

    # ── AT_BrandType  (LOV — raw from RNA → MDD AT_BrandType LOV lookup) ─────
    # Mirrors Clarks: RNA BRANDTYPE_DETAIL → look up in MDD Brand Type LOV sheet
    # (Col A = LOV ID, Col B = display). Fallback: hardcoded map → raw text.
    brand_type_raw = cfg.get("brand_type", "")
    if brand_type_raw:
        bt_lov_id = (mdd.lovs.get("AT_BrandType", {}) if mdd else {}).get(brand_type_raw.upper(), "")
        if bt_lov_id:
            _val(vals_el, "AT_BrandType", "", id_val=bt_lov_id)
        else:
            bt_fallback = _resolve_brand_type(brand_type_raw)
            if bt_fallback:
                _val(vals_el, "AT_BrandType", "", id_val=bt_fallback)
            else:
                _val(vals_el, "AT_BrandType", brand_type_raw)
        log.info("[Sofsole] AT_BrandType: raw='%s' → id='%s'", brand_type_raw, bt_lov_id or _resolve_brand_type(brand_type_raw) or brand_type_raw)

    # ── AT_BrandStatus  (LOV — raw from RNA → MDD AT_BrandStatus LOV lookup) ─
    # Mirrors Clarks: RNA BRAND STATUS → look up in MDD Brand Status LOV sheet.
    brand_status_raw = cfg.get("brand_status", "")
    if brand_status_raw:
        bs_lov_id = (mdd.lovs.get("AT_BrandStatus", {}) if mdd else {}).get(brand_status_raw.upper(), "")
        if bs_lov_id:
            _val(vals_el, "AT_BrandStatus", "", id_val=bs_lov_id)
        else:
            _val(vals_el, "AT_BrandStatus", brand_status_raw)
        log.info("[Sofsole] AT_BrandStatus: raw='%s' → id='%s'", brand_status_raw, bs_lov_id or brand_status_raw)
    else:
        # Default Active when RNA has no Brand Status value
        _val(vals_el, "AT_BrandStatus", "", id_val="A")

    # ── AT_BrandGroup  (LOV — raw from RNA → MDD AT_BrandGroup LOV lookup) ───
    # Mirrors Clarks: RNA BRANDGROUP → look up in MDD Brand Group LOV sheet.
    brand_grp_raw = cfg.get("brand_group", "")
    if brand_grp_raw:
        bg_lov_id = (mdd.lovs.get("AT_BrandGroup", {}) if mdd else {}).get(brand_grp_raw.upper(), "")
        if bg_lov_id:
            _val(vals_el, "AT_BrandGroup", "", id_val=bg_lov_id)
        else:
            bg_fallback = _resolve_brand_group(brand_grp_raw)
            if bg_fallback:
                _val(vals_el, "AT_BrandGroup", "", id_val=bg_fallback)
            else:
                _val(vals_el, "AT_BrandGroup", brand_grp_raw)
        log.info("[Sofsole] AT_BrandGroup: raw='%s' → id='%s'", brand_grp_raw, bg_lov_id or _resolve_brand_group(brand_grp_raw) or brand_grp_raw)

    # ── AT_CompanyCode / AT_SBU  (MultiValue LOV — from filename metadata) ───
    _multival(vals_el, "AT_CompanyCode", cfg.get("comp_code", "0888"))
    _multival(vals_el, "AT_SBU",         cfg.get("sbu",       "FF"))

    # ── AT_Country  (LOV — from filename metadata) ────────────────────────────
    _val(vals_el, "AT_Country", "", id_val=country_code)

    # ── AT_RetailPriceCurrency  (LOV — derived from country_code, Col K logic) ─
    retail_currency = _RETAIL_CURRENCY_MAP.get(country_code, "")
    if retail_currency:
        _val(vals_el, "AT_RetailPriceCurrency", "", id_val=retail_currency)

    # ── AT_Season  (LOV — from filename) ─────────────────────────────────────
    _val(vals_el, "AT_Season", "", id_val=sea_prefix)

    # ── AT_SeasonYear  (Text — from filename) ─────────────────────────────────
    _val(vals_el, "AT_SeasonYear", sea_year)

    # ── AT_Gender  (LOV default = U / Unisex) ────────────────────────────────
    _val(vals_el, "AT_Gender", "", id_val="U")

    # ── AT_SAPAge  (LOV default = AD / Adults) ────────────────────────────────
    _val(vals_el, "AT_SAPAge", "", id_val="AD")

    # ── AT_BYGender  (LOV default = U / Unisex) ──────────────────────────────
    by_gender_raw = cfg.get("by_gender", "Unisex")
    by_gender_id = (
        "F" if "FEMALE" in by_gender_raw.upper() else
        "M" if "MALE"   in by_gender_raw.upper() else
        "U"
    )
    _val(vals_el, "AT_BYGender", "", id_val=by_gender_id)

    # ── AT_BYAge  (LOV default = ADULT) ──────────────────────────────────────
    by_age_id = "ADULT"
    if mdd:
        by_age_id = (
            mdd.lookup("BYAGE",    "Adult") or
            mdd.lookup("BY AGE",   "Adult") or
            mdd.lookup("AGE",      "Adults") or
            "ADULT"
        )
    _val(vals_el, "AT_BYAge", "", id_val=by_age_id)

    # ── AT_BYArticleType  (LOV — from metadata, default Inline) ──────────────
    _val(vals_el, "AT_BYArticleType", "", id_val=cfg.get("article_type", "Inline"))

    # ── AT_PackDetails  (LOV default = Single/S) ─────────────────────────────
    pack_details_raw = cfg.get("pack_details", "Single")
    pack_details_id  = _resolve_pack_details(pack_details_raw, mdd)
    _val(vals_el, "AT_PackDetails", "", id_val=pack_details_id)

    # ── AT_SAPProductFlag  (LOV default = A / Intercompany) ──────────────────
    _val(vals_el, "AT_SAPProductFlag", "", id_val="A")

    # ── AT_MaterialType  (LOV default = ZINA / Intercompany) ─────────────────
    _val(vals_el, "AT_MaterialType", "", id_val="ZINA")

    # ── AT_SAPArticleCategory  (LOV default = 01 / Generic–Variant) ──────────
    _val(vals_el, "AT_SAPArticleCategory", "", id_val="01")

    # ── AT_UOM  (LOV default = EA) ────────────────────────────────────────────
    _val(vals_el, "AT_UOM", "", id_val="EA")

    # ── AT_BCI  (LOV default = COMMERCIAL) ───────────────────────────────────
    _val(vals_el, "AT_BCI", "", id_val="COMMERCIAL")

    # ── AT_ArticleStatus  (LOV default = A / Active) ─────────────────────────
    _val(vals_el, "AT_ArticleStatus", "", id_val="A")

    # ── AT_PricingDistributionChannel  (LOV default = 01) ────────────────────
    _val(vals_el, "AT_PricingDistributionChannel", "", id_val="01")

    return gen_el


# ══════════════════════════════════════════════════════════════════════════════
# XML ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════════════════

def build_xml(rows: list[SofsoleRow], out_xml_path: Path, cfg: dict) -> None:
    """Write the full STEP-ProductInformation XML for all Sofsole generics."""
    groups = group_generics(rows)

    brand_code   = cfg.get("brand_code",   "IPL")
    brand_name   = cfg.get("brand_name",   "IMPLUS")
    sea_prefix   = cfg.get("season_prefix", "SM")
    sea_year     = cfg.get("season_year")  or str(datetime.now().year)
    export_time  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    season_list  = [(sea_prefix, sea_year)]

    log.info("[Sofsole] Writing STEP XML → %s", out_xml_path.name)
    with open(out_xml_path, "w", encoding="utf-8") as f:
        open_step_xml(f, export_time)

        # Classifications block
        cls_elem = build_classifications(brand_code, brand_name, season_list)
        cls_str  = ET.tostring(cls_elem, encoding="utf-8").decode("utf-8")
        cls_str  = cls_str.replace("ns0:", "").replace(":ns0", "")
        f.write(f"  {cls_str}\n")

        f.write("  <Products>\n")

        total_generics = 0
        for color_dict in groups.values():
            for color, grp in color_dict.items():
                gen_elem = _build_generic_product(grp, cfg)
                xml_str  = ET.tostring(gen_elem, encoding="utf-8").decode("utf-8")
                xml_str  = xml_str.replace("ns0:", "").replace(":ns0", "")
                xml_str  = re.sub(r'\s*xmlns="[^"]+"', "", xml_str)
                f.write(f"    {xml_str}\n")
                total_generics += 1

        f.write("  </Products>\n")
        close_step_xml(f)

    log.info(
        "[Sofsole] XML written: %s  (%d generic(s),  %d KB)",
        out_xml_path.name,
        total_generics,
        out_xml_path.stat().st_size // 1024,
    )


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT  — called by implus/lambda_function.py
# ══════════════════════════════════════════════════════════════════════════════

def run(args, auditor=None) -> tuple[list[dict], Path | None]:
    """
    Main ETL entry point for the Sofsole sub-brand.

    Called by lambda_function.py after all required files are downloaded to
    /tmp/stibo_workdir_implus/.
    """
    log.info("[Sofsole] ETL started — brand=%s", getattr(args, "brand", "IMPLUS"))

    # ── Resolve input file ───────────────────────────────────────────────────
    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        file_path = Path(input_file_arg)
    else:
        files = _find_input_files()
        if not files:
            log.warning("[Sofsole] No linelist .xlsx/.xlsm files found — nothing to process.")
            return [], None
        file_path = max(files, key=lambda p: p.stat().st_mtime)

    log.info("[Sofsole] Processing file: %s", file_path.name)

    base, in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir = _get_dirs()

    # ── Load MDD LOVs ─────────────────────────────────────────────────────────
    mdd: Optional[MDDLoader] = None
    mdd_path = _find_mdd_file(mdd_dir)
    if mdd_path:
        try:
            mdd = MDDLoader(mdd_path)
        except Exception as exc:
            log.warning("[Sofsole][MDD] Load failed (%s) — using hardcoded fallback maps", exc)
    else:
        log.info("[Sofsole][MDD] No MDD file in %s — using hardcoded fallback maps", mdd_dir)

    # ── Args from filename metadata (injected by lambda_function.py) ─────────
    brand_code   = getattr(args, "brand_code",   "IPL")    or "IPL"
    brand_name   = getattr(args, "brand",        "IMPLUS") or "IMPLUS"
    comp_code    = getattr(args, "comp_code",    "0888")   or "0888"
    sbu          = getattr(args, "sbu",          "FF")     or "FF"
    country_code = getattr(args, "country_code", "ID")     or "ID"
    season       = getattr(args, "season",       "SM27")   or "SM27"  # default season

    # ── Parse filename for metadata extraction ───────────────────────────────
    # Pattern: 0888-FF-IMPLUS-IPL_Article Attributes Source_Sofsole-SM27-ID-1.xlsx
    #   parts[0]  = comp code   (e.g. '0888')
    #   parts[1]  = SBU         (e.g. 'FF')
    #   parts[2]  = brand name  (e.g. 'IMPLUS')
    #   parts[-3] = season code (e.g. 'SM27')
    #   parts[-2] = country id  (e.g. 'ID')
    filename_stem   = file_path.stem
    stem_parts      = filename_stem.split("-")
    
    # Extract comp_code from filename
    if len(stem_parts) >= 1 and stem_parts[0].strip():
        comp_code = stem_parts[0].strip()
        log.info("[Sofsole] Comp code from filename: '%s'", comp_code)
    
    # Extract SBU from filename
    if len(stem_parts) >= 2 and stem_parts[1].strip():
        sbu = stem_parts[1].strip().upper()
        log.info("[Sofsole] SBU from filename: '%s'", sbu)

    # ── Resolve brand_lov_id from input filename → MDD BrandLOV ─────────────
    # Pattern: parts[2] of dash-split stem = brand token (e.g. "IMPLUS").
    # MDD BrandLOV: display name ("IMPLUS") → LOV ID ("IPL").
    # Fallback chain: MDD lookup → hardcoded BRAND_LOV_MAP → brand_code default.
    brand_from_file = brand_name  # default — overwritten below if parseable
    if len(stem_parts) >= 3:
        brand_token = stem_parts[2].strip().upper()
        if brand_token:
            brand_from_file = brand_token.title()   # e.g. "Implus"
            log.info("[Sofsole] Brand token from filename: '%s' → '%s'", brand_token, brand_from_file)

    brand_lov_id = brand_code  # fallback
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            looked_up    = brand_lov.get(brand_from_file.upper())
            brand_lov_id = looked_up if looked_up else BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.info(
                "[Sofsole] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_from_file.upper(), brand_lov_id,
                "(FOUND in MDD)" if looked_up else "(NOT FOUND in MDD — using fallback)",
            )
        else:
            brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.warning("[Sofsole] BrandLOV sheet not in MDD — using BRAND_LOV_MAP fallback: '%s'", brand_lov_id)
    else:
        brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
        log.warning("[Sofsole] No MDD loaded — brand_lov_id='%s' from BRAND_LOV_MAP", brand_lov_id)

    # ── Extract season code from filename ────────────────────────────────────
    # Mirrors Steve Madden logic: search for season pattern (2 letters + 2-4 digits)
    # Pattern examples: 'SM27', 'SS27', 'SP2028', 'FW2027'
    # Normalize 4-digit years to 2-digit format (e.g., 'SP2028' → 'SP28')
    filename_season = season  # fallback to args default
    season_found = False
    
    # Search all parts for season pattern
    for part in stem_parts:
        if re.match(r'^[A-Z]{2}\d{2,4}$', part.strip(), re.IGNORECASE):
            sea_num = part.strip()[2:]
            sea_year_2d = sea_num[-2:] if len(sea_num) == 4 else sea_num
            filename_season = f"{part.strip()[:2].upper()}{sea_year_2d}"
            season_found = True
            log.info("[Sofsole] Season extracted from filename: '%s' (from '%s')", filename_season, part.strip())
            break
    
    if not season_found:
        # Fallback: regex anywhere in stem
        season_match = re.search(r'([A-Z]{2})(\d{2,4})', filename_stem, re.IGNORECASE)
        if season_match:
            sea_num = season_match.group(2)
            sea_year_2d = sea_num[-2:] if len(sea_num) == 4 else sea_num
            filename_season = f"{season_match.group(1).upper()}{sea_year_2d}"
            log.info("[Sofsole] Season extracted (fallback regex): '%s'", filename_season)
    
    season = filename_season  # use filename-extracted season

    # ── Parse season into prefix and year ────────────────────────────────────
    # Mirrors Clarks logic: extract year from season string
    # e.g., "SM27" → sea_prefix="SM", sea_year="2027"
    # e.g., "SS2027" → sea_prefix="SS", sea_year="2027"
    sea_prefix = season[:2].upper() if len(season) >= 2 else season
    
    # Extract year from season (e.g., "SS27" → "2027", "SS2027" → "2027")
    year_match = re.search(r"20\d{2}", season)
    if not year_match:
        short_match = re.search(r"\d{2}$", season)
        sea_year = f"20{short_match.group()}" if short_match else ""
    else:
        sea_year = year_match.group()
    
    # Fallback to args or current year if parsing fails
    if not sea_year:
        sea_year = getattr(args, "season_year", None) or str(datetime.now().year)
    
    log.info("[Sofsole] Parsed season: '%s' → prefix='%s' year='%s'", season, sea_prefix, sea_year)

    # ── Extract country ID from filename ─────────────────────────────────────
    # Pattern: parts[-2] = country id (e.g. 'ID', 'MY')
    filename_country_id = ""
    if len(stem_parts) >= 2:
        candidate = stem_parts[-2].strip().upper()
        if re.match(r'^[A-Z]{2}$', candidate):
            filename_country_id = candidate
            log.info("[Sofsole] Country ID from filename: '%s'", filename_country_id)
    if filename_country_id:
        country_code = filename_country_id  # prefer filename over args default

    # ── Brand Type / Category / Group / Status from RNA (brand mapping file) ──
    brand_type   = getattr(args, "brand_type",     "") or ""
    brand_cat    = getattr(args, "brand_category", "") or ""
    brand_group  = getattr(args, "brand_group",    "") or ""
    brand_status = getattr(args, "brand_status",   "") or ""

    if not brand_type or not brand_cat:
        rna_path = _find_rna_source_file(attr_dir)
        if rna_path:
            try:
                rna          = RNALoader(rna_path)
                country_name = _COUNTRY_MAP.get(country_code.upper(), country_code.upper())
                result = rna.get(country_name, comp_code, sbu, brand_lov_id)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_lov_id)
                if not result.get("brand_type") and not result.get("brand_category"):
                    result = rna.get_fuzzy(country_name, sbu, brand_code)
                brand_type   = brand_type   or result.get("brand_type",     "")
                brand_cat    = brand_cat    or result.get("brand_category", "")
                brand_group  = brand_group  or result.get("brand_group",    "")
                brand_status = brand_status or result.get("brand_status",   "")
                log.info(
                    "[Sofsole] RNA → country=%s comp=%s sbu=%s brand=%s → "
                    "type='%s' cat='%s' group='%s' status='%s'",
                    country_name, comp_code, sbu, brand_lov_id,
                    brand_type, brand_cat, brand_group, brand_status,
                )
            except Exception as exc:
                log.warning("[Sofsole] RNA lookup failed: %s", exc)
        else:
            log.warning("[Sofsole] No RNA source workbook in %s", attr_dir)

    cfg = {
        "brand_code":     brand_lov_id,
        "brand_name":     brand_from_file,
        "sbu":            sbu,
        "comp_code":      comp_code,
        "country_code":   country_code,
        "season_prefix":  sea_prefix,
        "season_year":    sea_year,
        "article_type":   getattr(args, "article_type",  "Inline") or "Inline",
        "brand_type":     brand_type,
        "brand_category": brand_cat,
        "brand_group":    brand_group,
        "brand_status":   brand_status,
        "mdd":            mdd,
    }

    rows = load_linesheet(file_path)
    if not rows:
        log.warning("[Sofsole] No valid rows parsed — nothing to write.")
        return [], None

    xml_path = xml_dir / f"{file_path.stem}.xml"
    build_xml(rows, xml_path, cfg)

    if auditor and xml_path.exists():
        auditor.set_xml_uploads([str(xml_path)])

    return [], xml_path


# ══════════════════════════════════════════════════════════════════════════════
# CLI  (local test)
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse
    import logging as _logging

    _logging.basicConfig(
        level=_logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[_logging.StreamHandler(sys.stdout)],
    )

    ap = argparse.ArgumentParser(description="implus Sofsole Line List → STEP XML")
    ap.add_argument("--brand",        default="IMPLUS")
    ap.add_argument("--brand-code",   default="IPL",   dest="brand_code")
    ap.add_argument("--comp-code",    default="0888",  dest="comp_code")
    ap.add_argument("--sbu",          default="FF")
    ap.add_argument("--country-code", default="ID",    dest="country_code")
    ap.add_argument("--input-file",   default=None,    dest="input_file")
    run(ap.parse_args())
