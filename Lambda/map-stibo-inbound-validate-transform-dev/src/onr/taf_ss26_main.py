"""
╔══════════════════════════════════════════════════════════════════╗
║   STIBO INBOUND XML GENERATOR — ON Running Packaging List TAF   ║
║   taf_ss26_main.py                                              ║
║   Linesheet: "Packaging List TAF - ON SS26 (APR 26).xlsx"       ║
║                                                                  ║
║   File Type : Packaging List TAF  (formerly TAF)                ║
║   Trigger   : SAP Article Creation                              ║
║   Endpoint  : EAN Update                                        ║
║                                                                  ║
║   SS26 Packaging List TAF File Layout:                           ║
║     Row 1–5  : empty                                            ║
║     Row 6    : header                                            ║
║     Row 7+   : data  (one row = one EAN + one US size)          ║
║                                                                  ║
║   Header columns:                                                ║
║     ACCOUNT | LAUNCH DATE | ALLOCATION MONTH | EAN | SKU |      ║
║     CATEGORY | PRODUCT | STYLE CODE | GENDER | STYLE NAME |     ║
║     COLOUR | RETAIL PRICE | SIZE | QTY | TOTAL | WAREHOUSE      ║
║                                                                  ║
║   KEY_InboundVariant logic (MAA Article Code max 18 digits):     ║
║     GenericCode (max 12) + ColorCode3d (3) + SizeCode3d (3)     ║
║     e.g.  ONR3WG100348 + IVO + 06H  =  ONR3WG100348IVO06H      ║
║                                                                  ║
║   EAN barcode → AT_Barcode on every PRD_VariantArticle           ║
╚══════════════════════════════════════════════════════════════════╝
"""

from __future__ import annotations

import logging
import math
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import openpyxl


# ══════════════════════════════════════════════════════════════════════════════
# STIBO XML HELPERS — inlined (previously from deleted onr._taf_xml_common)
# ══════════════════════════════════════════════════════════════════════════════
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace('', STIBO_NS)

LOV_AGE: dict[str, str] = {
    "AD": "Adults",
    "CH": "Children",
    "I":  "Infants",
    "Y":  "Youth",
}

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')


def strip_xmlns(xml_str: str) -> str:
    return _XMLNS_RE.sub("", xml_str)


def _val(parent, attr_id: str, value: str = "", id_val: str = ""):
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if id_val:
        clean_id = str(id_val).strip()
        if clean_id and clean_id not in ("None", "nan"):
            el.set("ID", clean_id)
    if value:
        clean_val = str(value).strip()
        if clean_val and clean_val not in ("None", "nan"):
            el.text = clean_val
    return el


def _multival(parent, attr_id: str, id_val: str, label: str = ""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    if label:
        v.text = str(label)
    return mv


def _keyval(parent, key_id: str, value: str):
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}KeyValue")
    el.set("KeyID", key_id)
    el.text = value
    return el


def _clf(parent, class_id: str, ref_type: str):
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}ClassificationReference")
    el.set("ClassificationID", class_id)
    el.set("Type", ref_type)
    return el


def derive_season(season: str) -> tuple[str, str]:
    """'SS26' -> ('SS', '2026'); 'FW2026' -> ('FW', '2026')."""
    s = (season or "").strip().upper()
    prefix = s[:2] if len(s) >= 2 else s
    tail   = s[2:]
    year   = tail if len(tail) == 4 else (f"20{tail}" if len(tail) == 2 else tail)
    return prefix, year


def season_id_full(brand_code: str, season: str) -> str:
    """Stibo Classification ID prefix for the season hierarchy."""
    prefix, year = derive_season(season)
    return f"CLH_{brand_code}_{prefix}{year}"


def build_classifications(brand_name: str, brand_code: str, season: str) -> ET.Element:
    """Season / Confirmed / Unconfirmed Classifications block."""
    prefix, year = derive_season(season)
    sid = season_id_full(brand_code, season)
    label_map = {"SS": "Spring Summer", "FW": "Fall Winter", "AW": "Fall Winter"}
    season_label = label_map.get(prefix, prefix)

    root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    season_cls = ET.SubElement(root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", sid)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", f"CLH_{brand_name.title().replace(' ', '')}Batches")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = (
        f"{brand_name} {season_label} {year}"
    )

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{sid}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{prefix} {year} Confirmed Articles"
    )

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{sid}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{prefix} {year} Unconfirmed Articles"
    )
    return root


def derive_gender(raw: str, mdd: dict | None = None) -> tuple[str, str, str]:
    """Raw gender token -> (sap_gender_code, label, sap_age_code). Prefers MDD Gender LOV."""
    key = (raw or "").strip().upper()[:1] or "U"
    _VALID_AGE = {"AD", "CH", "I", "Y"}
    if mdd and isinstance(mdd.get("gender"), dict):
        entry = mdd["gender"].get(key)
        if entry:
            sap_age_raw = entry.get("sap_age") or "AD"
            sap_age = sap_age_raw if sap_age_raw in _VALID_AGE else "AD"  # normalize "A" → "AD"
            return (
                entry.get("sap_gender")      or key,
                entry.get("sap_gender_desc") or key,
                sap_age,
            )
    mapping = {
        "M": ("M", "Male",   "AD"),
        "W": ("F", "Female", "AD"),
        "F": ("F", "Female", "AD"),
        "U": ("U", "Unisex", "AD"),
        "K": ("U", "Unisex", "CH"),
        "Y": ("U", "Unisex", "CH"),
        "I": ("U", "Unisex", "CH"),
    }
    return mapping.get(key, mapping["U"])


def derive_franchise(style_name: str, franchise_lov: dict | None = None) -> str:
    """Return full style name for AT_Franchise (emitted as uppercase text value, not a LOV)."""
    return style_name.strip() if style_name else ""


def derive_ecom_name(style_name: str) -> str:
    if not style_name:
        return ""
    return style_name.strip().title()[:80]


def fmt_launch_date(raw) -> str:
    """Normalize to dd-Mon-yyyy lowercase (e.g. '03-jun-2026'); pass through if unknown format."""
    if not raw:
        return ""
    if isinstance(raw, datetime):
        return raw.strftime("%d-%b-%Y").lower()
    s = str(raw).strip()
    if not s:
        return ""
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%b-%Y").lower()
        except ValueError:
            continue
    return s


def open_step_xml(f, export_time: str):
    f.write(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
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


def close_step_xml(f):
    f.write("</STEP-ProductInformation>\n")


def add_generic_values(vals_el, ctx: dict):
    """Emit the minimal Crocs orderform-style generic values."""
    def put(attr_id, value="", id_val=""):
        if not (value or id_val):
            return
        _val(vals_el, attr_id, value, id_val=id_val)

    # put("AT_CountryOrigin", id_val=ctx.get("coo", ""))


# ══════════════════════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
MDD_DIR      = INPUT_DIR / "mdd"
ATTR_DIR     = INPUT_DIR / "attributes"

OUTPUT_DIR  = BASE_DIR / "output"
XML_OUT_DIR = OUTPUT_DIR / "xml"
LOG_DIR     = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, MDD_DIR, ATTR_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("onr.taf_ss26_main")
if not log.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSITE MAPPINGS (kept local — these are not simple LOVs)
# ══════════════════════════════════════════════════════════════════════════════
COMP_CODE_MAP: dict[tuple[str, str], str] = {
    ("ID", "SP"): "0888", ("ID", "FF"): "0888", ("ID", "FL"): "0888",
    ("ID", "GO"): "0888", ("ID", "CH"): "0888", ("ID", "SD"): "0888",
    ("ID", "IZ"): "0888", ("ID", "LS"): "0888", ("ID", "GU"): "0888",
    ("ID", "GI"): "0510", ("ID", "PY"): "0190",
    ("PH", "SP"): "0803", ("PH", "FF"): "0803", ("PH", "FL"): "0803",
    ("PH", "GO"): "0803", ("PH", "SD"): "0803", ("PH", "PR"): "0803",
    ("PH", "IF"): "0803",
    ("SG", "SP"): "0883", ("SG", "FF"): "0883", ("SG", "FL"): "0883",
    ("SG", "GO"): "0883", ("SG", "OU"): "0883", ("SG", "OF"): "0883",
    ("MY", "SP"): "0885", ("MY", "FF"): "0885", ("MY", "FL"): "0885",
    ("MY", "GO"): "0885", ("MY", "OM"): "0885", ("MY", "JD"): "0885",
    ("TH", "SP"): "9991", ("TH", "FF"): "9991", ("TH", "FL"): "9991",
    ("TH", "CH"): "9991", ("TH", "TW"): "9991", ("TH", "TA"): "9991",
    ("VN", "SP"): "0882", ("VN", "FF"): "0882", ("VN", "FL"): "0882",
    ("VN", "CH"): "0882", ("VN", "VT"): "0882",
    ("KH", "SP"): "0881", ("KH", "FF"): "0881", ("KH", "FL"): "0881",
    ("KH", "K7"): "0881",
}

COUNTRY_DEFAULT_COMP: dict[str, str] = {
    "ID": "0888", "PH": "0803", "SG": "0883",
    "MY": "0885", "TH": "9991", "VN": "0882", "KH": "0881",
}

# CATEGORY token (selection, vertical) → (merch_L1, merch_L2, _spare, sports_category)
_ONR_SAP_MAP: dict[tuple[str, str], tuple[str, str, str, str]] = {
    ("Shoes",   "Performance All Day"):         ("Footwear", "Sports Shoes",            "Running",  "Lifestyle"),
    ("Shoes",   "Performance Outdoor"):         ("Footwear", "Sports Shoes",            "Outdoor",  "Outdoor"),
    ("Shoes",   "Performance Running"):         ("Footwear", "Sports Shoes",            "Running",  "Running"),
    ("Shoes",   "Performance Training"):        ("Footwear", "Sports Shoes",            "Running",  "Fitness"),
    ("Shoes",   "Performance Tennis"):          ("Footwear", "Sports Shoes",            "Tennis",   "Tennis"),
    ("Shoes",   "Shoes Performance Running"):   ("Footwear", "Sports Shoes",            "Running",  "Running"),
    ("Apparel", "Apparel Performance Running"): ("Apparel",  "Sports Running Apparel",  "Short",    "Running"),
}

# ON Running Franchise LOV mapping — style name first word → Stibo LOV_Franchise ID
ONR_FRANCHISE_MAP: dict[str, str] = {
    "CLOUDMONSTER":  "Cloud",
    "CLOUDRUNNER":   "Cloud",
    "CLOUDFLOW":     "Cloud",
    "CLOUDSURFER":   "Cloud",
    "CLOUDBOOM":     "Cloud",
    "CLOUDSWIFT":    "Cloud",
    "CLOUDSPARK":    "Cloud",
    "CLOUDECLIPSE":  "Cloud",
    "CLOUDGO":       "Cloud",
    "CLOUD":         "Cloud",
    "CLOUDULTRA":    "Cloud Ultra",
    "CLOUDAWAY":     "Cloud",
    "CLOUDACE":      "Cloud",
    "CLOUDBOOM":     "Cloud",
    "CLOUDSTRATUS":  "Cloud",
    "CLOUDVENTURE":  "Cloud",
}

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
AT = {
    "CompanyCode":           "AT_CompanyCode",
    "SBU":                   "AT_SBU",
    "Country":               "AT_Country",
    "BYIndicator":           "AT_BYIndicator",
    "SAPIndicator":          "AT_SAPIndicator",
    "EcomIndicator":         "AT_EcomIndicator",
    "SAPAge":                "AT_SAPAge",
    "GenericDesc":           "AT_InboundGenericCode",
    "VariantDesc":           "AT_InboundGenericCode",
    "Season":                "AT_Season",
    "Year":                  "AT_Year",
    "Gender":                "AT_Gender",
    "Brand":                 "AT_Brand",
    "BrandCode":             "AT_BrandCode",
    "StyleCode":             "AT_PrincipalStyleCode",
    "StyleDesc":             "AT_PrincipalStyleDescription",
    "PrincipalGender":       "AT_PrincipalGender",
    "PrincipalAge":          "AT_PrincipalAge",
    "PrincipalSize":         "AT_PrincipalSize",
    "CombinationStyleColor": "AT_CombinationStyleColor",
    "Variant":               "AT_Variant",
    "SAPStyleCode":          "AT_SAPStyleCode",
    "PrincipalColorCode":    "AT_PrincipalColorCode",
    "ColorDesc":             "AT_PrincipalColorDescription",
    "ColorCode":             "AT_ColorCode",
    "Color":                 "AT_Color",
    "ColorDescription":      "AT_ColorDescription",
    "MAAColor":              "AT_MAAColor",
    "SizeCode":              "AT_SizeCode",
    "Size":                  "AT_Size",
    "SizeDescription":       "AT_SizeDescription",
    "MAASize":               "AT_MAASize",
    "EUSizeCode":            "AT_EUSizeCode",
    "USSizeCode":            "AT_USSizeCode",
    "SAPSizeCode":           "AT_SAPSize",
    "ArticleType":           "AT_ArticleType",
    "UOM":                   "AT_UOM",
    "MaterialType":          "AT_SAPArticleType",
    "HierarchyL1":           "AT_PrincipalMerchandiseHierarchyL1",
    "HierarchyL2":           "AT_PrincipalMerchandiseHierarchyL2",
    "RetailPrice":           "AT_OriginalPrice",
    "LaunchDate":            "AT_LaunchDate",
    "InboundGeneric":        "AT_InboundGenericCode",
    "DealersAccount":        "AT_DealersAccount",
    "Barcode":               "AT_Barcode",  # MDD canonical for EAN / barcode
    "AllocationMonth":       "AT_AllocationMonth",
    "Warehouse":             "AT_Warehouse",
    "Remarks":               "AT_Remarks",
}


# ══════════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class SizeLine:
    """One data row from TAF SS26 = one EAN + one US size"""
    ean:     str
    us_size: str
    qty:     Optional[int]
    row_num: int


@dataclass
class Article:
    """
    Grouped by (style_code, colour) = one PRD_Variant.
    Contains multiple SizeLines (one per EAN/size row).
    """
    account:      str
    launch_date:  Optional[str]
    alloc_month:  str
    sku:          str
    category:     str
    product:      str
    style_code:   str
    gender:       str
    style_name:   str
    colour:       str
    retail_price: Optional[float]
    warehouse:    str
    remarks:      str
    sizes:        list[SizeLine] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _s(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _parse_date(val) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%d.%m.%Y")
    s = str(val).strip()
    if re.match(r'\d{1,2}[./-]\d{1,2}[./-]\d{2,4}', s):
        return s
    return None


HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "ACCOUNT": ("ACCOUNT", "DEALERS ACCOUNT", "DEALER ACCOUNT", "CUSTOMER", "CUSTOMER ACCOUNT"),
    "LAUNCH DATE": ("LAUNCH DATE", "LAUNCHING DATE", "LAUNCH", "INTRO DATE"),
    "ALLOCATION MONTH": ("ALLOCATION MONTH", "ALLOC MONTH", "ALLOCATION"),
    "EAN": ("EAN", "EAN CODE", "EAN NO", "EAN NUMBER", "BARCODE", "BAR CODE", "UPC", "UPC CODE", "GTIN"),
    "SKU": ("SKU", "SKU CODE", "ITEM NUMBER", "ITEM NO", "ARTICLE NUMBER", "ARTICLE NO"),
    "CATEGORY": ("CATEGORY", "SELECTION"),
    "PRODUCT": ("PRODUCT", "PRODUCT TYPE", "DIVISION"),
    "STYLE CODE": ("STYLE CODE", "STYLE NO", "STYLE NUMBER", "ARTICLE CODE", "MODEL CODE"),
    "GENDER": ("GENDER", "SAP GENDER"),
    "STYLE NAME": ("STYLE NAME", "MODEL NAME", "PRODUCT NAME", "ITEM NAME"),
    "COLOUR": ("COLOUR", "COLOR", "COLOUR NAME", "COLOR NAME", "COLORWAY"),
    "RETAIL PRICE": ("RETAIL PRICE", "RRP", "MSRP", "SEA MSRP"),
    "SIZE": ("SIZE", "US SIZE", "SIZE US"),
    "QTY": ("QTY", "QUANTITY", "ORDER QTY"),
    "TOTAL": ("TOTAL", "TOTAL QTY"),
    "WAREHOUSE": ("WAREHOUSE", "WH", "WHSE"),
    "REMARKS": ("REMARKS", "REMARK", "NOTES"),
}

PREFERRED_LINESHEET_SHEETS = ("Packaging List TAF", "TAF", "FTW", "FOOTWEAR", "Sheet1")


def _norm_header(val) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\u00a0", " ").strip().upper())


def _resolve_header(row) -> tuple[dict[str, int], dict[str, int]]:
    raw_hdr: dict[str, int] = {}
    for idx, val in enumerate(row):
        key = _norm_header(val)
        if key and key not in raw_hdr:
            raw_hdr[key] = idx

    hdr: dict[str, int] = {}
    for canonical, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias in raw_hdr:
                hdr[canonical] = raw_hdr[alias]
                break
        if canonical in hdr:
            continue
        for raw_name, idx in raw_hdr.items():
            if any(alias in raw_name for alias in aliases):
                hdr[canonical] = idx
                break
    return hdr, raw_hdr


# ══════════════════════════════════════════════════════════════════════════════
# MDD LOADER
# ══════════════════════════════════════════════════════════════════════════════
def load_mdd(path: Path) -> dict:
    """
    Load from MDD:
      color_desc_to_code : COLOUR description (upper) → 3-char MAA color code
      us_to_sap          : US size string → SAP size code (3-digit)
      brand              : brand_code → brand_name
    """
    log.info("[MDD] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    def _ms(v):
        return "" if v is None else str(v).strip()

    # ── Color Code LOV ────────────────────────────────────────────────────
    # col0 = code, col1 = description
    color_desc_to_code: dict[str, str] = {}
    if "Color Code LOV" in wb.sheetnames:
        for row in wb["Color Code LOV"].iter_rows(min_row=2, values_only=True):
            code = _ms(row[0])
            desc = _ms(row[1]).upper()
            if code and desc:
                color_desc_to_code[desc] = code

    # ── Size Code LOV ─────────────────────────────────────────────────────
    # col0 = Stibo LOV ID  (used for AT_Size attribute — must match LOV_SizeCode exactly)
    # col5 = SAP 3-digit code  (used for 18-char variant ID construction)
    # col6 = US size description  (matched against linesheet SIZE column)
    us_to_sap: dict[str, str] = {}        # US size desc → SAP 3-char code (variant ID)
    us_to_size_lov: dict[str, str] = {}   # US size desc → Stibo LOV ID  (AT_Size attr)
    if "Size Code LOV" in wb.sheetnames:
        for row in wb["Size Code LOV"].iter_rows(min_row=2, values_only=True):
            if len(row) < 7:
                continue
            lov_id   = _ms(row[0])   # Stibo LOV ID — as-is, no zero-padding
            sap_code = _ms(row[5])   # SAP 3-digit code
            desc_raw = row[6]        # US size description
            if not sap_code or desc_raw is None:
                continue
            if sap_code.isdigit():
                sap_code = sap_code.zfill(3)
            elif len(sap_code) == 2 and sap_code[0].isdigit() and sap_code[1].upper() == "H":
                # half-size SAP codes: "6H" → "06H", "7H" → "07H", etc.
                sap_code = f"0{sap_code[0]}H"
            desc_str = str(desc_raw).strip()
            # try numeric mapping
            try:
                f = float(desc_raw)
                if not math.isfinite(f):
                    raise ValueError
                us_to_sap[str(f)] = sap_code
                if f == int(f):
                    us_to_sap[str(int(f))] = sap_code
                # LOV ID: store as-is from col0 (e.g. "5" not "005")
                if lov_id:
                    us_to_size_lov[str(f)] = lov_id
                    if f == int(f):
                        us_to_size_lov[str(int(f))] = lov_id
            except (ValueError, TypeError, OverflowError):
                pass
            # always store raw string
            if desc_str:
                us_to_sap[desc_str] = sap_code
                if lov_id:
                    us_to_size_lov[desc_str] = lov_id

    # ── Brand LOV ─────────────────────────────────────────────────────────
    brand: dict[str, str] = {}
    if "Brand LOV" in wb.sheetnames:
        for row in wb["Brand LOV"].iter_rows(min_row=2, values_only=True):
            k, v = _ms(row[0]), _ms(row[1])
            if k and v:
                brand[k] = v

    # ── Simple key/value LOV sheets (code → label) ───────────────────────────
    def _kv_sheet(name: str) -> dict[str, str]:
        out: dict[str, str] = {}
        if name in wb.sheetnames:
            for row in wb[name].iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                k = _ms(row[0])
                v = _ms(row[1]) if len(row) > 1 else ""
                if k and v:
                    out[k] = v
        return out

    country   = _kv_sheet("Country LOV")
    sbu       = _kv_sheet("SBU LOV")
    company   = _kv_sheet("Company Code LOV")

    # ── Gender LOV ───────────────────────────────────────────────────────────────────
    # raw_key (M/W/U/K) → {sap_gender, sap_gender_desc, sap_age, sap_age_desc}
    gender: dict[str, dict] = {}
    if "Gender LOV" in wb.sheetnames:
        for row in wb["Gender LOV"].iter_rows(min_row=2, values_only=True):
            cells = [_ms(c) for c in (row or ())]
            if not cells or not cells[0]:
                continue
            code  = cells[0].upper()[:1]
            gender[code] = {
                "sap_gender":      cells[2] if len(cells) > 2 and cells[2] else code,
                "sap_gender_desc": cells[1] if len(cells) > 1 else "",
                "sap_age":         cells[3] if len(cells) > 3 else "",
                "sap_age_desc":    cells[4] if len(cells) > 4 else "",
            }

    wb.close()
    log.info(
        "[MDD] Colors:%d | SizeMappings:%d | SizeLOVs:%d | Brands:%d | Country:%d | SBU:%d | Company:%d | Gender:%d",
        len(color_desc_to_code), len(us_to_sap), len(us_to_size_lov),
        len(brand), len(country), len(sbu), len(company), len(gender),
    )
    return {
        "color_desc_to_code": color_desc_to_code,
        "us_to_sap":          us_to_sap,
        "us_to_size_lov":     us_to_size_lov,
        "brand":              brand,
        "country":            country,
        "sbu":                sbu,
        "company":            company,
        "gender":             gender,
    }


# ══════════════════════════════════════════════════════════════════════════════
# COLOR & SIZE CODE RESOLVERS
# ══════════════════════════════════════════════════════════════════════════════
def resolve_color_code(colour: str, color_desc_to_code: dict) -> str:
    """
    Map COLOUR column → 3-char MAA color code.
    e.g. "Ivory Seedling" → "IVO"

    Strategy:
      1. Exact match (uppercase)
      2. First word match
      3. Any word match
      4. Fallback → "000" (NO COLOR)
    """
    if not colour:
        return "000"
    upper = colour.upper().strip()
    if upper in color_desc_to_code:
        return color_desc_to_code[upper]
    words = upper.split()
    for word in words:
        if word in color_desc_to_code:
            return color_desc_to_code[word]
    return "000"


def resolve_size_code(us_size: str, us_to_sap: dict) -> tuple[str, bool]:
    """
    Map US SIZE column → (SAP size code, is_mdd_mapped).
    e.g. "6.5" → ("06H", True), "7" → ("7", True)

    Returns is_mdd_mapped=False when falling back to zero-padded digits.
    AT_Size (LOV field) must only be emitted when is_mdd_mapped=True.
    """
    if not us_size:
        return "000", False
    key = us_size.strip()
    if key in us_to_sap:
        return us_to_sap[key], True
    try:
        f = float(key)
        for try_key in [str(f), str(int(f)) if f == int(f) else None]:
            if try_key and try_key in us_to_sap:
                return us_to_sap[try_key], True
    except (ValueError, TypeError):
        pass
    digits = re.sub(r"[^0-9]", "", key)
    return (digits.zfill(3)[:3] if digits else "000"), False


def us_size_to_stibo_lov(us_size: str) -> str:
    """
    Derive Stibo LOV_SizeCode ID from a US shoe size string.

    Whole sizes : '6'   → '6',   '10'  → '10'
    Half sizes  : '6.5' → '06H', '9.5' → '09H'

    MDD col0 may contain supplier-internal codes (e.g. '0N6', '#06') that do
    NOT match LOV_SizeCode.  This derivation uses the US size directly, which
    always yields a valid Stibo LOV value for standard shoe sizes.
    Returns '' if the string cannot be parsed as a numeric size.
    """
    try:
        f = float(us_size.strip())
        whole = int(f)
        if f == whole:
            return str(whole)        # e.g. '6', '7', '10'
        else:
            return f"{whole:02d}H"   # e.g. '06H', '07H', '09H'
    except (ValueError, TypeError):
        return ""


# ══════════════════════════════════════════════════════════════════════════════
# LINESHEET LOADER  (SS26 TAF format)
# ══════════════════════════════════════════════════════════════════════════════
def load_linesheet(path: Path) -> list[Article]:
    """
    SS26 TAF layout:
      Rows 1–5 : empty
      Row  6   : header
      Row  7+  : data — one row = one EAN + one size

    Groups rows by (STYLE CODE, COLOUR) → one Article per variant.
    Skips subtotal/summary rows (no EAN).
    """
    log.info("[Linesheet] Loading: %s", path.name)
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)

    # ── Find sheet + header row ─────────────────────────────────────────────
    candidate_sheets = [s for s in PREFERRED_LINESHEET_SHEETS if s in wb.sheetnames]
    candidate_sheets.extend(s for s in wb.sheetnames if s not in candidate_sheets)

    sheet_name = ""
    rows = []
    hdr_idx = None
    hdr: dict[str, int] = {}
    best_seen: tuple[str, int, list[str]] | None = None

    for candidate in candidate_sheets:
        sheet_rows = list(wb[candidate].iter_rows(values_only=True))
        for i, row in enumerate(sheet_rows):
            candidate_hdr, raw_hdr = _resolve_header(row)
            if not raw_hdr:
                continue
            if best_seen is None or len(raw_hdr) > len(best_seen[2]):
                best_seen = (candidate, i + 1, list(raw_hdr.keys())[:20])
            if "STYLE CODE" in candidate_hdr and "EAN" in candidate_hdr:
                sheet_name = candidate
                rows = sheet_rows
                hdr_idx = i
                hdr = candidate_hdr
                break
        if hdr_idx is not None:
            break

    wb.close()

    if hdr_idx is None:
        if best_seen:
            log.error(
                "[Linesheet] Header row not found in %s. Best candidate: sheet='%s' row=%d columns=%s",
                path.name, best_seen[0], best_seen[1], best_seen[2],
            )
        else:
            log.error("[Linesheet] Header row not found in %s", path.name)
        return []

    log.info("[Linesheet] Using sheet '%s' | Header at row %d | columns: %s",
             sheet_name, hdr_idx + 1, list(hdr.keys()))

    def g(row, col: str):
        idx = hdr.get(col.upper())
        if idx is None:
            return None
        return row[idx] if idx < len(row) else None

    # ── Parse data rows ────────────────────────────────────────────────────
    articles: dict[tuple, Article] = {}

    for i, row in enumerate(rows[hdr_idx + 1:], start=hdr_idx + 2):
        if all(v is None for v in row):
            continue

        ean        = _s(g(row, "EAN"))
        style_code = _s(g(row, "STYLE CODE"))
        colour     = _s(g(row, "COLOUR"))

        # skip subtotal / summary rows (TOTAL, DISC, NETT)
        if not ean or not style_code:
            continue

        key = (style_code, colour)

        if key not in articles:
            price = None
            try:
                rp = g(row, "RETAIL PRICE")
                if rp is not None:
                    price = float(rp)
            except (ValueError, TypeError):
                pass

            gender_raw = _s(g(row, "GENDER")).strip().upper()
            gender     = gender_raw[:1] if gender_raw else "U"

            articles[key] = Article(
                account      = _s(g(row, "ACCOUNT")),
                launch_date  = _parse_date(g(row, "LAUNCH DATE")),
                alloc_month  = _s(g(row, "ALLOCATION MONTH")),
                sku          = _s(g(row, "SKU")),
                category     = _s(g(row, "CATEGORY")),
                product      = _s(g(row, "PRODUCT")),
                style_code   = style_code,
                gender       = gender,
                style_name   = _s(g(row, "STYLE NAME")),
                colour       = colour,
                retail_price = price,
                warehouse    = _s(g(row, "WAREHOUSE")),
                remarks      = _s(g(row, "REMARKS")),
            )

        qty = None
        try:
            q = g(row, "QTY")
            if q is not None:
                qty = int(float(str(q)))
        except (ValueError, TypeError):
            pass

        articles[key].sizes.append(SizeLine(
            ean     = ean,
            us_size = _s(g(row, "SIZE")),
            qty     = qty,
            row_num = i,
        ))

    result = list(articles.values())
    log.info("[Linesheet] %d variants | %d unique style codes | %d total EAN rows",
             len(result),
             len({a.style_code for a in result}),
             sum(len(a.sizes) for a in result))
    return result


# ══════════════════════════════════════════════════════════════════════════════
# XML BUILDER  (linelist-style: STEP header + Classifications + KEY_Article)
# ══════════════════════════════════════════════════════════════════════════════
def _indent(elem, level=0):
    pad = "\n" + "  " * level
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = pad + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = pad
        for child in elem:
            _indent(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = pad
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = pad


def _build_generic_product(
    style_code:   str,
    variants:     list[Article],
    mdd:          dict,
    cfg:          dict,
    season_id_cl: str,
) -> ET.Element:
    rep                = variants[0]
    brand_code         = cfg["brand_code"]
    brand_name         = cfg["brand_name"]
    sea_prefix         = cfg["season_prefix"]
    sea_year           = cfg["season_year"]
    color_desc_to_code = mdd["color_desc_to_code"]
    us_to_sap          = mdd["us_to_sap"]
    # Same KEY_InboundArticle transform as ONR linelist: omit 1st and 3rd characters.
    sap_style          = (style_code[1:2] + style_code[3:]) if len(style_code) > 3 else style_code
    generic_code       = f"{brand_code}{sap_style}"

    gender_code, gender_label, age_code = derive_gender(rep.gender, mdd)
    age_label = LOV_AGE.get(age_code, age_code)

    rep_color_3d = resolve_color_code(rep.colour, color_desc_to_code)

    gen_el = ET.Element(f"{{{STIBO_NS}}}Product")
    gen_el.set("UserTypeID", "PRD_GenericArticle")
    gen_el.set("ParentID",   "PPH_F-TempSubCat")
    _keyval(gen_el, "KEY_InboundArticle", generic_code)

    vals_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Values")
    ctx = {
        "brand_code":       brand_code,
        "brand_name":       brand_name,
        "sbu":              cfg["sbu"],
        "sbu_label":        mdd.get("sbu", {}).get(cfg["sbu"], ""),
        "comp_code":        cfg["comp_code"],
        "comp_label":       mdd.get("company", {}).get(cfg["comp_code"], ""),
        "country_code":     cfg["country_code"],
        "style_code":       style_code,
        "style_name":       rep.style_name,
        "colour":           rep.colour,
        "colour_id":        rep_color_3d if rep_color_3d != "000" else "",
        "generic_code":     generic_code,
        "gender_code":      gender_code,
        "gender_label":     gender_label,
        "gender_raw":       rep.gender,
        "age_code":         age_code,
        "age_label":        age_label,
        "age_raw":          age_label,
        "season_prefix":    sea_prefix,
        "season_year":      sea_year,
        "coo":              "VN",
        "article_type":     cfg["article_type"],
        "article_type_id":  cfg["article_type"],
        "retail_price":     str(int(rep.retail_price)) if rep.retail_price else "",
        "currency":         cfg.get("currency", "IDR"),
        "fob_currency":     "IDR",
        "division_label":   "Shoes",
        "division_letter":  "F",
        "prod_group_code":  (rep.category or "")[:2].upper(),
        "prod_group_label": rep.category or rep.product,
        "prod_cat_code":    (rep.product or "")[:2].upper(),
        "prod_cat_label":   rep.product or rep.category,
        "franchise":        derive_franchise(rep.style_name),
        "ecom_name_en":     derive_ecom_name(rep.style_name),
        "launch_date":      fmt_launch_date(rep.launch_date or ""),
        "sap_style_code":   sap_style,
        "generic_desc":     f"{brand_code} {rep.style_name} (A/{'M' if gender_code == 'M' else 'F'}) {rep.colour}"[:40],
        # CATEGORY = "Shoes Performance Running" → SAP map: (selection, vertical) → (L1, L2, _, sports_cat)
        "merch_l1":         _ONR_SAP_MAP.get(((rep.category or '').split()[0] if rep.category else '', ' '.join((rep.category or '').split()[1:])), ('Footwear','Sports Shoes','Running','Running'))[0],
        "merch_l2":         _ONR_SAP_MAP.get(((rep.category or '').split()[0] if rep.category else '', ' '.join((rep.category or '').split()[1:])), ('Footwear','Sports Shoes','Running','Running'))[1],
        "sports_category":  _ONR_SAP_MAP.get(((rep.category or '').split()[0] if rep.category else '', ' '.join((rep.category or '').split()[1:])), ('Footwear','Sports Shoes','Running','Running'))[3],
    }
    add_generic_values(vals_el, ctx)

    # ── Variants (PRD_VariantArticle — must live INSIDE the Generic) ──
    # One PRD_VariantArticle per (colour, size). Each variant holds its EAN
    # inside a DC_Barcode DataContainer. AT_MainEANIndicator flags it as the
    # main EAN for that variant.
    seen_variants: set[str] = set()
    for art in variants:
        color_3d = resolve_color_code(art.colour, color_desc_to_code)
        for sz in art.sizes:
            size_3d, size_from_mdd = resolve_size_code(sz.us_size, us_to_sap)
            size_lov_id = us_size_to_stibo_lov(sz.us_size)

            # Use H-notation from LOV ID for half sizes so KEY_InboundVariant
            # stays consistent with the size code used by Stibo.
            # e.g. size 6.5: size_lov_id="06H" (3 chars) → var_size="06H" ✓
            #      size 6  : size_lov_id="6"   (1 char)  → var_size=size_3d ✓
            if len(size_lov_id) == 3:
                var_size = size_lov_id  # Half sizes: "06H", "07H", etc.
            elif size_lov_id:
                var_size = size_lov_id.zfill(3)  # Whole sizes: "5" → "005", "6" → "006"
            else:
                var_size = size_3d 
            var_id   = f"{generic_code}{color_3d}{var_size}"

            if var_id in seen_variants:
                log.warning(
                    "[Variant] Duplicate skipped: %s  (EAN=%s, size=%s, row=%d)",
                    var_id, sz.ean, sz.us_size, sz.row_num,
                )
                continue
            seen_variants.add(var_id)

            var_el = ET.SubElement(gen_el, f"{{{STIBO_NS}}}Product")
            var_el.set("UserTypeID", "PRD_VariantArticle")
            _keyval(var_el, "KEY_InboundVariant", var_id)

            ET.SubElement(var_el, f"{{{STIBO_NS}}}Values")

            if sz.ean:
                dcs_el = ET.SubElement(var_el, f"{{{STIBO_NS}}}DataContainers")
                mdc_el = ET.SubElement(dcs_el, f"{{{STIBO_NS}}}MultiDataContainer")
                mdc_el.set("Type", "DC_Barcode")

                dc_el  = ET.SubElement(mdc_el, f"{{{STIBO_NS}}}DataContainer")
                dc_el.set("Analyzer", "true")
                dcv_el = ET.SubElement(dc_el,  f"{{{STIBO_NS}}}Values")
                _val(dcv_el, "AT_Barcode",          sz.ean)
                _val(dcv_el, "AT_BarcodeType",      id_val="P")
                _val(dcv_el, "AT_MainEANIndicator", id_val="Y")

    return gen_el


def build_xml(
    articles:    list[Article],
    mdd:         dict,
    output_path: Path,
    cfg:         dict,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    brand_code = cfg["brand_code"]
    brand_name = cfg["brand_name"]
    cfg["brand_name_for_clh"] = (cfg.get("brand_name_for_clh") or brand_name).replace(" ", "").upper()

    sea_prefix, sea_year = derive_season(cfg["season"])
    cfg["season_prefix"] = sea_prefix
    cfg["season_year"]   = sea_year
    season_id_cl         = season_id_full(brand_code, cfg["season"])

    by_style: dict[str, list[Article]] = defaultdict(list)
    for art in articles:
        by_style[art.style_code].append(art)

    log.info("[XML] Building — %d generics | %d variants | %d size variants",
             len(by_style), len(articles), sum(len(a.sizes) for a in articles))

    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    written_count = 0

    with open(output_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        open_step_xml(f, export_time)

        f.write("  <Products>\n")

        for style_code, variants in by_style.items():
            gen_el = _build_generic_product(style_code, variants, mdd, cfg, season_id_cl)
            _indent(gen_el, level=2)
            xml_str = strip_xmlns(ET.tostring(gen_el, encoding="unicode"))
            f.write(f"    {xml_str}\n")
            written_count += 1

        f.write("  </Products>\n")
        close_step_xml(f)

    log.info("[XML] Written → %s  (%.1f KB) — %d generics",
             output_path, output_path.stat().st_size / 1024, written_count)
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# LAMBDA ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def run(args, auditor=None):
    def first_xlsx(d: Path) -> Optional[Path]:
        files = list(d.glob("*.xlsx"))
        return files[0] if files else None

    mdd_f = first_xlsx(MDD_DIR)
    ll_files = list(LINELIST_DIR.glob("*.xlsx"))

    if not mdd_f:
        log.error("No MDD file found in %s — aborting.", MDD_DIR)
        sys.exit(1)
    if not ll_files:
        log.error("No Packaging List TAF linelist file found in %s — aborting.", LINELIST_DIR)
        sys.exit(1)

    mdd = load_mdd(mdd_f)

    brand_code   = getattr(args, "brand_code",   "ONR")    or "ONR"
    season       = getattr(args, "season",        "SS26")   or "SS26"
    article_type = getattr(args, "article_type_from_filename", "") or "Inline"

    sea_prefix = season[:2].upper()
    sea_year   = f"20{season[2:]}" if len(season) > 2 else season
    season_id  = f"{sea_prefix}{sea_year}"

    written: list[str] = []

    for ll_path in ll_files:
        log.info("─── Processing Packaging List TAF SS26 linesheet: %s ───", ll_path.name)
        articles = load_linesheet(ll_path)

        if not articles:
            log.warning("[Packaging List TAF SS26] No articles parsed — skipping %s", ll_path.name)
            continue

        # ── TEST MODE: limit to first N generics ────────────────────────────────
        # MAX_ARTICLES = 5
        # limited_style_codes = set(
        #     list(dict.fromkeys(article.style_code for article in articles))[:MAX_ARTICLES]
        # )
        # articles = [article for article in articles if article.style_code in limited_style_codes]

        # cfg: fallback to args, then embedded defaults
        fallback_comp    = getattr(args, "comp_code",    "0888") or "0888"
        fallback_sbu     = getattr(args, "sbu",          "SP")   or "SP"
        fallback_country = getattr(args, "country_code", "ID")   or "ID"

        cfg = {
            "brand_code":   brand_code,
            "brand_name":   mdd["brand"].get(brand_code, brand_code),
            "season":       season,
            "season_id":    season_id,
            "comp_code":    fallback_comp,
            "sbu":          fallback_sbu,
            "country_code": fallback_country,
            "article_type": article_type,
        }

        log.info(
            "[Packaging List TAF SS26] cfg → brand=%s comp=%s sbu=%s country=%s season=%s type=%s",
            cfg["brand_name"], cfg["comp_code"], cfg["sbu"],
            cfg["country_code"], cfg["season"], cfg["article_type"],
        )

        out_path = XML_OUT_DIR / (ll_path.stem + ".xml")
        build_xml(articles, mdd, out_path, cfg)
        written.append(str(out_path))

        total_eans = sum(len(a.sizes) for a in articles)
        print("═" * 52, flush=True)
        print("  ARTICLE SUMMARY — ON RUNNING PACKAGING LIST TAF", flush=True)
        print("═" * 52, flush=True)
        print(f"  Linesheet      : {ll_path.name}", flush=True)
        print(f"  Variants       : {len(articles)}", flush=True)
        print(f"  Unique styles  : {len({a.style_code for a in articles})}", flush=True)
        print(f"  Total EAN rows : {total_eans}", flush=True)
        print(f"  Season         : {cfg['season']}", flush=True)
        print(f"  Country        : {cfg['country_code']}  ({mdd['country'].get(cfg['country_code'], '?')})", flush=True)
        print(f"  SBU            : {cfg['sbu']}", flush=True)
        print(f"  Company Code   : {cfg['comp_code']}", flush=True)
        print(f"  XML output     : {out_path.name}  ({out_path.stat().st_size // 1024} KB)", flush=True)
        print("═" * 52, flush=True)

    if auditor and written:
        auditor.set_xml_uploads(written)


# ══════════════════════════════════════════════════════════════════════════════
# LOCAL CLI  (VS Code / direct python run)
# ══════════════════════════════════════════════════════════════════════════════
def main():
    import argparse
    p = argparse.ArgumentParser(
        description="Stibo Inbound XML Generator — ON Running Packaging List TAF SS26"
    )
    p.add_argument("--brand-code",   default="ONR")
    p.add_argument("--season",       default="SS26")
    p.add_argument("--comp-code",    default="0888",   help="Company code fallback")
    p.add_argument("--sbu",          default="SP",     help="SBU fallback")
    p.add_argument("--country-code", default="ID",     help="Country code fallback")
    p.add_argument("--article-type-from-filename", default="Inline")
    run(p.parse_args())


if __name__ == "__main__":
    main()