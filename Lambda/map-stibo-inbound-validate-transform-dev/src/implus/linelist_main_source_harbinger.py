"""
STIBO INBOUND XML GENERATOR -- implus Line List (Inline) v1.0
Sub-brand source: "1. IPL_Article Attributes Source_Harbinger.xlsx" -> Stibo STEP XML

Reference implementation: implus/linelist_main.py (Balega source format).
This module reuses every brand-agnostic helper from linelist_main.py (STEP XML
writer, MDD/RNA loaders, LOV maps) and only re-implements the pieces that are
specific to the Harbinger source workbook's layout.

Source file format : implus "SUM ORDER" order-form Excel (.xlsx / .xlsm)
Sheet name          : "SUM ORDER"

Harbinger sheet layout (row 20 = header):
    Col A  Category                              -> AT_PrincipalMerchandiseHierarchyL1
    Col B  IMPLUS EU ITEM #                       -> per-SKU id / AT_PrincipalStyleCode
           (also carries the sub-category banner text, e.g. "POWER GLOVES",
            "NYLON BELTS", on rows with no UPC/EAN -> AT_PrincipalMerchandiseHierarchyL2)
    Col C  UPC
    Col D  IMPLUS EU ITEM #
    Col E  EAN
    Col F  EAN / UPC (duplicate of E)
    Col G  PRODUCT PICTURE
    Col H  PRODUCT DESCRIPTION                    -> AT_PrincipalStyleDescription
    Col I  SIZE                                   -> AT_PrincipalSize
    Col J  COLOR                                  -> AT_PrincipalColorName / AT_PrincipalColorCode
    Col K  PRODUCT DESCRIPTION / SIZE / COLOUR     (fallback description)
    Col L  COUNTRY OF ORIGIN                      -> AT_CountryOrigin
    Col M  STATUS
    Col N  Outer Carton Qty
    Col O  CASE PACK QTY
    Col P  MSRP ASIA
    Col Q  ORDER QTY
    Col R  DISTRIBUTOR PRICE                      -> AT_FOB
    Col S  CARTON QTY
    Col T  SUM

---------------------------------------------------------------------------------
Stibo Attribute ID                | Field Name in Harbinger Source
---------------------------------------------------------------------------------
AT_PrincipalStyleCode             | IMPLUS EU ITEM # (first variant of the style/color group)
AT_PrincipalStyleDescription      | L1 + L2 + PRODUCT DESCRIPTION + SIZE (space-joined, per Stibo team)
AT_PrincipalColorName             | COLOR
AT_PrincipalColorCode             | Not sent (mapping logic N/A for Harbinger, per Stibo team)
AT_FOB                            | DISTRIBUTOR PRICE
AT_FOBCurrency                    | USD (Default)
AT_PrincipalMerchandiseHierarchyL1| Category (col A)
AT_PrincipalMerchandiseHierarchyL2| Sub-category banner text (col B)
AT_PrincipalMerchandiseHierarchyL3| L1 + L2 + PRODUCT DESCRIPTION + SIZE (same formula as StyleDescription)
AT_CountryOrigin                  | HK (Default)
AT_Gender                         | Unisex (Default -- U)
AT_SAPAge                         | Adults (Default -- AD)
AT_BYGender                       | Unisex (Default)
AT_BYAge                          | Adults (Default)
AT_SAPProductFlag                 | A (Default: Intercompany)
AT_MaterialType                   | ZINA (Default: Intercompany)
AT_SAPArticleCategory             | 01 (Default: Generic / Variant)
AT_UOM                            | EA (Default: EA)
AT_ArticleType                    | Inline (Default)
AT_BCI                            | Commercial (Default)
AT_ArticleStatus                  | Active (Default -- MDD LOV ID)
---------------------------------------------------------------------------------
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

# ── Reuse every brand-agnostic helper from the reference implementation ─────
try:
    from implus.linelist_main_source_balega import (
        STIBO_NS,
        _val,
        _keyval,
        _multival,
        _clf,
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
        build_classifications,
        RNALoader,
        _find_rna_source_file,
        _COUNTRY_MAP,
    )
except ImportError:  # pragma: no cover - allows `python implus/linelist_main_two.py`
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
        _clf,
        MDDLoader,
        _find_mdd_file,
        BRAND_LOV_MAP,
        _resolve_brand_group,
        _resolve_brand_type,
        _resolve_pack_details,
        build_classifications,
        RNALoader,
        _find_rna_source_file,
        _COUNTRY_MAP,
    )

log = logging.getLogger("implus.linelist_main_two")


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class HarbingerRow:
    row_num:           int
    major_category:    str = ""   # Col A "Category" -> Hierarchy L1
    category:           str = ""  # Sub-category banner text -> Hierarchy L2
    implus_us:          str = ""  # Col B "IMPLUS EU ITEM #" (per-SKU id)
    implus_eu:          str = ""  # Col D "IMPLUS EU ITEM #"
    upc:                str = ""
    ean:                str = ""
    style_desc:         str = ""  # Col H "PRODUCT DESCRIPTION"
    color:              str = ""
    size:               str = ""
    country_of_origin:  str = ""
    status:             str = ""
    distributor_price:  str = ""


@dataclass
class HarbingerGenericGroup:
    style_key:         str
    style_desc:        str
    hierarchy_l1:      str
    hierarchy_l2:      str = ""
    color:             str = ""
    color_code:        str = ""
    implus_us:         str = ""   # first variant's US SKU id -- anchors the generic key
    implus_eu:         str = ""   # first variant's EU SKU id -- used for AT_PrincipalStyleCode
    variants:          list[HarbingerRow] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL LOADER
# ══════════════════════════════════════════════════════════════════════════════
def load_linesheet(path: Path) -> list[HarbingerRow]:
    log.info("[Harbinger] Loading excel file: %s", path.name)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        log.error("Failed to open excel: %s", e)
        return []

    ws = wb["SUM ORDER"] if "SUM ORDER" in wb.sheetnames else wb.active

    # ── Locate header row: Col A == "Category", Col B contains "IMPLUS US" ──
    header_row_idx = -1
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        cell_0 = _s(row[0]) if len(row) > 0 else ""
        cell_1 = _s(row[1]).replace("\n", " ") if len(row) > 1 else ""
        if cell_0.lower() == "category" and "implus us" in cell_1.lower():
            header_row_idx = row_idx
            break

    if header_row_idx == -1:
        log.warning("Could not find header row. Defaulting to row 20.")
        header_row_idx = 20

    rows: list[HarbingerRow] = []
    current_l1 = ""
    current_l2 = ""

    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, values_only=True), start=header_row_idx + 1):
        cell_0  = _s(row[0])  if len(row) > 0  else ""   # Category (L1)
        cell_1  = _s(row[1])  if len(row) > 1  else ""   # IMPLUS EU ITEM # / banner text
        cell_2  = _s(row[2])  if len(row) > 2  else ""   # UPC
        cell_3  = _s(row[3])  if len(row) > 3  else ""   # IMPLUS EU ITEM #
        cell_4  = _s(row[4])  if len(row) > 4  else ""   # EAN
        cell_7  = _s(row[7])  if len(row) > 7  else ""   # PRODUCT DESCRIPTION
        cell_8  = _s(row[8])  if len(row) > 8  else ""   # SIZE
        cell_9  = _s(row[9])  if len(row) > 9  else ""   # COLOR
        cell_10 = _s(row[10]) if len(row) > 10 else ""   # PRODUCT DESCRIPTION / SIZE / COLOUR (fallback)
        cell_11 = _s(row[11]) if len(row) > 11 else ""   # COUNTRY OF ORIGIN
        cell_12 = _s(row[12]) if len(row) > 12 else ""   # STATUS
        cell_17 = _s(row[17]) if len(row) > 17 else ""   # DISTRIBUTOR PRICE

        if not cell_0 and not cell_1:
            continue  # fully blank row (footer / spacer)

        if "TOTAL" in cell_0.upper():
            continue

        # New major category -> reset sub-category until the next banner row sets it
        if cell_0 and cell_0 != current_l1:
            current_l1 = cell_0
            current_l2 = ""

        # Section-banner row (sub-category title): no UPC and no EAN present
        is_data_row = bool(cell_4) or bool(cell_2)
        if not is_data_row:
            if cell_1 and cell_1.upper() != "TOTAL":
                current_l2 = cell_1
            continue

        style_desc = cell_7 or cell_10

        rows.append(HarbingerRow(
            row_num=row_idx,
            major_category=current_l1,
            category=current_l2,
            implus_us=cell_1,
            implus_eu=cell_3,
            upc=cell_2,
            ean=cell_4,
            style_desc=style_desc,
            color=cell_9,
            size=cell_8,
            country_of_origin=cell_11,
            status=cell_12,
            distributor_price=cell_17,
        ))

    log.info("[Harbinger] Parsed %d valid rows", len(rows))
    return rows


def group_generics(rows: list[HarbingerRow]) -> dict[str, dict[str, HarbingerGenericGroup]]:
    # Grouping key = IMPLUS EU ITEM # (falls back to IMPLUS US ITEM #, then
    # row_num). Each item number is a distinct article (one number per size
    # break) and is the same value that anchors KEY_InboundArticle /
    # AT_PrincipalStyleCode (brand_code + implus_eu). Grouping by style
    # description instead would wrongly collapse every size break of the
    # same style into one product.
    groups: dict[str, dict[str, HarbingerGenericGroup]] = defaultdict(dict)
    for r in rows:
        style_desc = r.style_desc or f"UNKNOWN_{r.row_num}"
        style_key = r.implus_eu or r.implus_us or f"UNKNOWN_{r.row_num}"
        color = r.color or "NC"

        if color not in groups[style_key]:
            groups[style_key][color] = HarbingerGenericGroup(
                style_key=style_key,
                style_desc=style_desc,
                hierarchy_l1=r.major_category,
                hierarchy_l2=r.category,
                color=r.color,
                color_code=r.color,   # no dedicated color-code column in this source
                implus_us=r.implus_us,
                implus_eu=r.implus_eu,
            )
        groups[style_key][color].variants.append(r)

    return groups


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILD (ONLY GENERIC PRODUCTS)
# ══════════════════════════════════════════════════════════════════════════════
def _build_generic_product(group: HarbingerGenericGroup, cfg: dict) -> ET.Element:
    brand_code   = cfg.get("brand_code", "IPL")
    brand_name   = cfg.get("brand_name", "IMPLUS")
    sea_prefix   = cfg.get("season_prefix", "SM")
    sea_year     = cfg.get("season_year") or str(datetime.now().year)
    country_code = cfg.get("country_code", "ID")
    mdd: Optional[MDDLoader] = cfg.get("mdd")

    generic_key = f"{brand_code}{group.implus_eu}"

    # First variant data (for FOB, CountryOrigin)
    first_var = group.variants[0] if group.variants else None

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_E-TempSubCat")  # implus = accessories/socks

    _keyval(gen_el, "KEY_InboundArticle", generic_key)

    name_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Name")
    name_el.text = f"{brand_code} {group.style_desc}".strip()
    
     # ── Classification References ────────────────────────────────────────────
    _clf(gen_el, "CLH_ImplusArticles", "CPL_Merchandiser")        
    _clf(gen_el, f"CLH_{brand_code}_{sea_prefix}{sea_year}UA", "CPL_UnConfirmedForSeason")

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")

    # ── Brand Attributes ──────────────────────────────────────────────────────
    b_lov_id = BRAND_LOV_MAP.get(brand_code.upper(), BRAND_LOV_MAP.get(brand_name.upper(), brand_code))
    _val(vals_el, "AT_Brand",        "", id_val=b_lov_id)
    b_grp_id = _resolve_brand_group(brand_name)
    _val(vals_el, "AT_BrandGroup",   "", id_val=b_grp_id)
    _val(vals_el, "AT_BrandStatus",  "", id_val="A")

    raw_brand_type = (cfg.get("brand_type") or "MAA").strip()
    bt_id = _resolve_brand_type(raw_brand_type)
    if bt_id:
        _val(vals_el, "AT_BrandType", id_val=bt_id)
    else:
        _val(vals_el, "AT_BrandType", value=raw_brand_type)

    # ── RNA / Portal Attributes ───────────────────────────────────────────────
    _multival(vals_el, "AT_SBU",         cfg.get("sbu", "FF"))
    _multival(vals_el, "AT_CompanyCode", cfg.get("comp_code", "0888"))
    _val(vals_el, "AT_Country",          "", id_val=country_code)

    _RETAIL_CURRENCY_MAP: dict[str, str] = {
        "ID": "IDR", "PH": "PHP", "TH": "THB",
        "SG": "SGD", "MY": "MYR", "VN": "VND", "KH": "USD",
    }
    retail_currency = _RETAIL_CURRENCY_MAP.get(country_code, "")
    if retail_currency:
        _val(vals_el, "AT_RetailPriceCurrency", "", id_val=retail_currency)

    # ── Season ────────────────────────────────────────────────────────────────
    _val(vals_el, "AT_Season",     "", id_val=sea_prefix)
    _val(vals_el, "AT_SeasonYear", sea_year)

    # ── Inbound Code ─────────────────────────────────────────────────────────
    _val(vals_el, "AT_InboundGenericCode", generic_key)

    # ── Principal Attributes ─────────────────────────────────────────────────
    # AT_PrincipalStyleDescription / AT_PrincipalMerchandiseHierarchyL3: per
    # Stibo team feedback, Harbinger builds both from the same formula --
    # L1 + L2 + Product Description + Size (space-joined, skipping empty
    # parts). Size comes from the first variant, same source as
    # AT_PrincipalSize below.
    l1_l2_desc = " ".join(
        p for p in (
            group.hierarchy_l1, group.hierarchy_l2, group.style_desc,
            first_var.size if first_var else "",
        ) if p
    )
    _val(vals_el, "AT_PrincipalStyleCode",        group.implus_eu)
    _val(vals_el, "AT_PrincipalStyleDescription", l1_l2_desc)
    # AT_PrincipalColorCode: per Stibo team feedback, mapping logic is N/A
    # for Harbinger -- do not send this attribute.
    if group.color:
        _val(vals_el, "AT_PrincipalColorName", group.color)
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL1", group.hierarchy_l1)
    if group.hierarchy_l2:
        _val(vals_el, "AT_PrincipalMerchandiseHierarchyL2", group.hierarchy_l2)
    _val(vals_el, "AT_PrincipalMerchandiseHierarchyL3", l1_l2_desc)

    # ── From First Variant ────────────────────────────────────────────────────
    # AT_CountryOrigin: per Stibo team feedback, default to HK (Hong Kong)
    # for Harbinger, same as the other implus sub-brand sources.
    _val(vals_el, "AT_CountryOrigin", "", id_val="HK")

    if first_var and first_var.size:
        _val(vals_el, "AT_PrincipalSize", first_var.size)

    if first_var and first_var.distributor_price:
        _val(vals_el, "AT_FOB", first_var.distributor_price)

    # ── Defaults (from MDD) ───────────────────────────────────────────────────
    _val(vals_el, "AT_FOBCurrency",    "", id_val="USD")
    _val(vals_el, "AT_SAPProductFlag", "", id_val="A")
    _val(vals_el, "AT_MaterialType",   "", id_val="ZINA")
    _val(vals_el, "AT_SAPArticleCategory", "", id_val="01")
    _val(vals_el, "AT_UOM",            "", id_val="EA")

    _val(vals_el, "AT_Gender",  "", id_val="U")
    _val(vals_el, "AT_SAPAge",  "", id_val="AD")

    by_gender_raw = cfg.get("by_gender", "Unisex")
    by_gender_id = "F" if "FEMALE" in by_gender_raw.upper() else ("M" if "MALE" in by_gender_raw.upper() else "U")
    _val(vals_el, "AT_BYGender", "", id_val=by_gender_id)

    by_age_id = "ADULT"
    if mdd:
        by_age_id = (mdd.lookup("BYAGE", "Adult")
                     or mdd.lookup("BY AGE", "Adult")
                     or mdd.lookup("AGE", "Adults")
                     or "ADULT")
    _val(vals_el, "AT_BYAge",    "", id_val=by_age_id)

    _val(vals_el, "AT_BYArticleType", "", id_val=cfg.get("article_type", "Inline"))

    pack_details_raw = cfg.get("pack_details", "Single")
    pack_details_id = _resolve_pack_details(pack_details_raw, mdd)
    _val(vals_el, "AT_PackDetails", "", id_val=pack_details_id)

    _val(vals_el, "AT_BCI", "", id_val="COMMERCIAL")

    _val(vals_el, "AT_ArticleStatus", "", id_val="A")

    _val(vals_el, "AT_PricingDistributionChannel", "", id_val="01")

    return gen_el

def build_xml(rows: list[HarbingerRow], out_xml_path: Path, cfg: dict) -> None:
    groups = group_generics(rows)

    fallback_year   = cfg.get("season_year") or str(datetime.now().year)
    fallback_prefix = cfg.get("season_prefix", "SM")
    season_list = [(fallback_prefix, fallback_year)]

    # ══════════════════════════════════════════════════════════════════════════
    # TEST LIMITER — Set TEST_MODE = False for production
    # ══════════════════════════════════════════════════════════════════════════
    TEST_MODE = False
    LIMIT = 5

    if TEST_MODE:
        limited = {}
        count = 0
        for sc, cm in groups.items():
            if count >= LIMIT:
                break
            limited[sc] = {}
            for color, grp in cm.items():
                if count >= LIMIT:
                    break
                limited[sc][color] = grp
                count += 1
        groups = limited
        log.info("[XML] TEST MODE: Limited to %d generic product(s)", count)

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    brand_code = cfg.get("brand_code", "IPL")
    brand_name = cfg.get("brand_name", "IMPLUS")

    log.info("[Harbinger] Writing STEP XML → %s", out_xml_path.name)
    with open(out_xml_path, "w", encoding="utf-8") as f:
        open_step_xml(f, export_time)

        cls_elem = build_classifications(brand_code, brand_name, season_list)
        cls_str  = ET.tostring(cls_elem, encoding="utf-8").decode("utf-8")
        cls_str  = cls_str.replace("ns0:", "").replace(":ns0", "")
        f.write(f"  {cls_str}\n")

        f.write("  <Products>\n")

        total_generics = 0
        for color_dict in groups.values():
            for color, group in color_dict.items():
                gen_elem = _build_generic_product(group, cfg)
                xml_str = ET.tostring(gen_elem, encoding="utf-8").decode("utf-8")
                xml_str = xml_str.replace("ns0:", "").replace(":ns0", "")
                xml_str = re.sub(r'\s*xmlns="[^"]+"', '', xml_str)
                f.write(f"    {xml_str}\n")
                total_generics += 1

        f.write("  </Products>\n")
        close_step_xml(f)

    log.info("[Harbinger] XML generated successfully: %d Generic product nodes.", total_generics)


def run(args, auditor=None) -> tuple[list[dict], Path | None]:
    log.info("[implus-Harbinger] ETL started — brand=%s", getattr(args, "brand", "IMPLUS"))

    input_file_arg = getattr(args, "input_file", None)
    if input_file_arg and Path(input_file_arg).exists():
        file_path = Path(input_file_arg)
    else:
        files = _find_input_files()
        if not files:
            log.warning("[Harbinger] No linelist .xlsx/.xlsm files found — nothing to process.")
            return [], None
        file_path = max(files, key=lambda p: p.stat().st_mtime)

    log.info("[Harbinger] Processing file: %s", file_path.name)

    base, in_dir, ll_dir, mdd_dir, attr_dir, out_dir, xml_dir, log_dir = _get_dirs()

    # ── Load MDD LOVs ────────────────────────────────────────────────────────
    mdd: Optional[MDDLoader] = None
    mdd_path = _find_mdd_file(mdd_dir)
    if mdd_path:
        try:
            mdd = MDDLoader(mdd_path)
        except Exception as exc:
            log.warning("[MDD] Load failed (%s) — using hardcoded fallback maps", exc)
    else:
        log.info("[MDD] No MDD file found in %s — using hardcoded fallback maps", mdd_dir)

    brand_code   = getattr(args, "brand_code",   "IPL")    or "IPL"
    brand_name   = getattr(args, "brand",        "IMPLUS") or "IMPLUS"
    comp_code    = getattr(args, "comp_code",    "0888")   or "0888"
    sbu          = getattr(args, "sbu",          "FF")     or "FF"
    country_code = getattr(args, "country_code", "ID")     or "ID"

    # ── Resolve brand_lov_id from input filename → MDD BrandLOV ─────────────
    brand_from_file = brand_name
    filename_stem   = file_path.stem
    stem_parts      = filename_stem.split("-")
    if len(stem_parts) >= 3:
        brand_token = stem_parts[2].strip().upper()
        if brand_token:
            brand_from_file = brand_token.title()
            log.info("[Harbinger] Brand token from filename: '%s' → '%s'", brand_token, brand_from_file)

    brand_lov_id = brand_code  # fallback
    if mdd:
        brand_lov = mdd.lovs.get("BrandLOV") or mdd.lovs.get("AT_Brand", {})
        if brand_lov:
            looked_up    = brand_lov.get(brand_from_file.upper())
            brand_lov_id = looked_up if looked_up else BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.info(
                "[Harbinger] Brand LOV lookup: key='%s' → LOV ID='%s' %s",
                brand_from_file.upper(), brand_lov_id,
                "(FOUND in MDD)" if looked_up else "(NOT FOUND in MDD — using fallback)",
            )
        else:
            brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
            log.warning("[Harbinger] BrandLOV sheet not in MDD — using BRAND_LOV_MAP fallback: '%s'", brand_lov_id)
    else:
        brand_lov_id = BRAND_LOV_MAP.get(brand_from_file.upper(), brand_code)
        log.warning("[Harbinger] No MDD loaded — brand_lov_id='%s' from BRAND_LOV_MAP", brand_lov_id)

    # Also derive country_id from filename (2nd-to-last dash token)
    filename_country_id = ""
    if len(stem_parts) >= 2:
        candidate = stem_parts[-2].strip().upper()
        if re.match(r'^[A-Z]{2}$', candidate):
            filename_country_id = candidate
            log.info("[Harbinger] Country ID from filename: '%s'", filename_country_id)
    if filename_country_id:
        country_code = filename_country_id

    # ── Brand Type / Brand Category from RNA ─────────────────────────────────
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
                    "[Harbinger] RNA → country=%s comp=%s sbu=%s brand=%s → "
                    "type='%s' cat='%s' group='%s' status='%s'",
                    country_name, comp_code, sbu, brand_lov_id,
                    brand_type, brand_cat, brand_group, brand_status,
                )
            except Exception as e:
                log.warning("[Harbinger] RNA lookup failed: %s", e)
        else:
            log.warning("[Harbinger] No RNA source workbook found in %s", attr_dir)

    cfg = {
        "brand_code":    brand_lov_id,
        "brand_name":    brand_from_file,
        "sbu":           sbu,
        "comp_code":     comp_code,
        "country_code":  country_code,
        "season_prefix": getattr(args, "season_prefix", None) or "SM",
        "season_year":   getattr(args, "season_year", None) or None,
        "article_type":  getattr(args, "article_type", "Inline") or "Inline",
        "brand_type":    brand_type,
        "brand_category": brand_cat,
        "brand_group":   brand_group,
        "brand_status":  brand_status,
        "mdd":           mdd,    # MDDLoader instance (or None)
    }

    rows = load_linesheet(file_path)
    if not rows:
        log.warning("[Harbinger] No valid rows parsed.")
        return [], None

    xml_path = xml_dir / f"{file_path.stem}.xml"
    build_xml(rows, xml_path, cfg)

    if auditor and xml_path:
        auditor.set_xml_uploads([str(xml_path)])

    return [], xml_path


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="implus (Harbinger) Line List → STEP XML")
    ap.add_argument("--brand",        default="IMPLUS")
    ap.add_argument("--brand-code",   default="IPL", dest="brand_code")
    ap.add_argument("--comp-code",    default="0888", dest="comp_code")
    ap.add_argument("--sbu",          default="FF")
    ap.add_argument("--country-code", default="ID", dest="country_code")
    ap.add_argument("--input-file",   default=None, dest="input_file")
    run(ap.parse_args())
