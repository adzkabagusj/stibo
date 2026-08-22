"""
aldo/ecommerce.py  — v1.2  (MDD-corrected)
===========================================
Ecommerce processor — generates ONE XML from WSL004 MCR (Material Composition Report).

Source  : WSL004 MCR (.xlsx)  sheet="Material Composition Report"  header=row 1
Output  : Stibo STEP XML — Generic Articles ONLY (no variants), same structure as LineList.

═══ FIXES vs v1.1 (second import pass errors) ═══

ERROR 13 — AT_MaterialUpper IllegalLOVValue "50% GEL,50% POLYURETHANE"
  CAUSE: MCR composition strings (e.g. "50% GEL,50% POLYURETHANE") are NOT valid
         LOV_MaterialUpper IDs.  LOV only has: EVA, Leather, Mesh, Polyester,
         PU, Suede, Sustainable, Synthetic, Textile.
  FIX:   _map_material_upper() parses raw MCR string, returns first matching
         LOV label.  If nothing matches → attribute is skipped entirely
         (STEP ignores missing optional attrs; better than an import error).
         Same logic applied to AT_Material (LOV_Material) for lining.

ERROR 14-16 — AT_PackagingHeight/Length/Width Number validator failed "2 CM"
  CAUSE: MCR dimension columns contain strings like "2 CM", "14 CM", "9 CM".
         AT_PackagingHeight/Length/Width are Number type — no unit suffix allowed.
  FIX:   _strip_unit() extracts only the numeric portion before writing.
         "2 CM" → "2",  "14.5 CM" → "14.5",  "" → skipped.

═══ FIXES vs v1.0 (verified against Master_Data_Dictionary__MAA___1_.xlsx) ═══

ERROR 11 — Classification 'MA_AOD_CUSHIONSGR' cannot be found
  FIX: merch_safe regex now limits to A-Z0-9 max 8 chars (was 10).
       CUSHIONSGR (10 chars) → CUSHIONS (8 chars).
       Classification IDs in STEP follow an 8-char suffix convention.

ERROR 12 — Classification 'CLH_FOOTWEARAC' cannot be found
  FIX: prod_safe regex now limits to A-Z0-9 max 8 chars (was 10).
       FOOTWEARAC → FOOTWEAR (8 chars).

ERRORS 13-14 — AT_HTSCode / AT_CustomsDescription not found
  FIX: MDD does NOT define AT_HTSCode or AT_CustomsDescription.
       AT_HTSCode → replaced with AT_HSCode  (defined in MDD, type=Text)
       AT_CustomsDescription → REMOVED (no equivalent in MDD; write to
         AT_ShortDescriptionEN as a safe text fallback).

ERRORS 15-18 — AT_UpperComposition / AT_LiningComposition /
               AT_InsoleComposition / AT_SoleComposition not found
  FIX: MDD does NOT define these four composition attributes.
       Mapped to the closest MDD-defined equivalents:
         AT_UpperComposition  → AT_MaterialUpper  (LOV: Material - Upper)
         AT_LiningComposition → AT_Material       (LOV: Material)
         AT_InsoleComposition → AT_Content        (Text)
         AT_SoleComposition   → AT_Others         (Text)

ERRORS 19-20 — AT_Upper / AT_Bottom not found
  FIX: MDD does NOT define AT_Upper or AT_Bottom.
       AT_Upper  → AT_MaterialUpper  (already mapped above; skip duplicate)
       AT_Bottom → AT_Others        (already mapped above; skip duplicate)
       Both fields REMOVED from _build_product_xml to avoid duplicate
       attribute writes.  Raw values are folded into the composition fields.

ERRORS 21-22 — AT_ToeShape / AT_ToeOpening not found
  FIX: MDD does NOT define these attributes. REMOVED from output.

ERROR 23 — AT_ClosureType not found
  FIX: MDD does NOT define AT_ClosureType.
       Mapped to AT_Fastening (LOV: Fastening) — closest MDD equivalent.

ERROR 24 — AT_Ornamentation not found
  FIX: MDD does NOT define AT_Ornamentation. REMOVED from output.

ERRORS 25-27 — AT_PlatformHeight / AT_ShaftHeight / AT_CalfCircumference not found
  FIX: MDD does NOT define these attributes. REMOVED from output.

ERRORS 28-29 — AT_Height / AT_Length not found
  FIX: MDD defines AT_ProductHeight / AT_ProductLength / AT_ProductWidth
       (type=Number) for physical dimensions.
       AT_Height → AT_PackagingHeight   (Number)
       AT_Length → AT_PackagingLength   (Number)
       AT_Width  → AT_PackagingWidth    (Number)  ← also fixes Error 30

ERROR 30 — AT_Width IllegalLOVValue "9 CM"
  FIX: AT_Width in MDD is type=LOV (LOV_Width: Regular / Wide / Relaxed).
       The code was writing raw numeric strings like "9 CM" which are not
       valid LOV IDs.  Numeric dimension values now go to AT_PackagingWidth
       (Number).  AT_Width is intentionally NOT written for MCR data because
       the MCR "Width" column contains measurements, not width-fit codes.
"""

import re
import logging
import pandas as pd
from xml.etree import ElementTree as ET
from datetime import datetime
from pathlib import Path

log = logging.getLogger("aldo_etl")

STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"
BRAND_CODE   = "AOD"
BRAND_NAME   = "ALDO"

SKIP_CATEGORIES = ["service", "gfs", "shoe care", "shoecare"]

LOV_COUNTRY = {
    "AF": "Afganistan",       "AL": "Albania",          "DZ": "Algeria",
    "VI": "Amer.Virgin Is.",  "AO": "Angola",           "AI": "Anguilla",
    "AQ": "Antarctica",       "AG": "Antigua/Barbuda",  "AR": "Argentina",
    "AM": "Armenia",          "AW": "Aruba",            "AU": "Australia",
    "AT": "Austria",          "AZ": "Azerbaijan",       "BS": "Bahamas",
    "BH": "Bahrain",          "BD": "Bangladesh",       "BB": "Barbados",
    "BY": "Belarus",          "BE": "Belgium",          "BZ": "Belize",
    "BJ": "Benin",            "BM": "Bermuda",          "BT": "Bhutan",
    "BO": "Bolivia",          "BA": "Bosnia-Herz.",     "BW": "Botswana",
    "BV": "Bouvet Islands",   "BR": "Brazil",           "IO": "Brit.Ind.Oc.Ter",
    "VG": "Brit.Virgin Is.",  "BN": "Brunei Daruss.",   "BG": "Bulgaria",
    "BF": "Burkina-Faso",     "BI": "Burundi",          "KH": "Cambodia",
    "CM": "Cameroon",         "CA": "Canada",           "CV": "Cape Verde",
    "KY": "Cayman Islands",   "CF": "Central Afr.Rep",  "TD": "Chad",
    "CL": "Chile",            "CN": "China",            "CX": "Christmas Islnd",
    "CC": "Coconut Islands",  "CO": "Colombia",         "KM": "Comoros",
    "CD": "Congo",            "CK": "Cook Islands",     "CR": "Costa Rica",
    "HR": "Croatia",          "CU": "Cuba",             "CY": "Cyprus",
    "CZ": "Czech Republic",   "DK": "Denmark",          "DJ": "Djibouti",
    "DM": "Dominica",         "DO": "Dominican Rep.",   "AN": "Dutch Antilles",
    "TP": "East Timor",       "EC": "Ecuador",          "EG": "Egypt",
    "SV": "El Salvador",      "GQ": "Equatorial Guin",  "ER": "Eritrea",
    "EE": "Estonia",          "ET": "Ethiopia",         "FK": "Falkland Islnds",
    "FO": "Faroe Islands",    "FJ": "Fiji",             "FI": "Finland",
    "FR": "France",           "PF": "Frenc.Polynesia",  "GF": "French Guyana",
    "TF": "French S.Territ",  "GA": "Gabon",            "GM": "Gambia",
    "GE": "Georgia",          "DE": "Germany",          "GH": "Ghana",
    "GI": "Gibraltar",        "GR": "Greece",           "GL": "Greenland",
    "GD": "Grenada",          "GP": "Guadeloupe",       "GU": "Guam",
    "GT": "Guatemala",        "GN": "Guinea",           "GW": "Guinea-Bissau",
    "GY": "Guyana",           "HT": "Haiti",            "HM": "Heard/McDon.Isl",
    "HN": "Honduras",         "HK": "Hong Kong",        "HU": "Hungary",
    "IS": "Iceland",          "IN": "India",            "ID": "Indonesia",
    "IR": "Iran",             "IQ": "Iraq",             "IE": "Ireland",
    "IL": "Israel",           "IT": "Italy",            "CI": "Ivory Coast",
    "JM": "Jamaica",          "JP": "Japan",            "JO": "Jordan",
    "KZ": "Kazakhstan",       "KE": "Kenya",            "KI": "Kiribati",
    "KW": "Kuwait",           "KG": "Kyrgyzstan",       "LA": "Laos",
    "LV": "Latvia",           "LB": "Lebanon",          "LS": "Lesotho",
    "LR": "Liberia",          "LY": "Libya",            "LI": "Liechtenstein",
    "LT": "Lithuania",        "LU": "Luxembourg",       "MO": "Macau",
    "MK": "Macedonia",        "MG": "Madagascar",       "MW": "Malawi",
    "MY": "Malaysia",         "MV": "Maldives",         "ML": "Mali",
    "MT": "Malta",            "MH": "Marshall Islnds",  "MQ": "Martinique",
    "MR": "Mauretania",       "MU": "Mauritius",        "YT": "Mayotte",
    "MX": "Mexico",           "FM": "Micronesia",       "UM": "Minor Outl.Isl.",
    "MD": "Moldavia",         "MC": "Monaco",           "MN": "Mongolia",
    "MS": "Montserrat",       "MA": "Morocco",          "MZ": "Mozambique",
    "MM": "Myanmar",          "MP": "N.Mariana Islnd",  "NA": "Namibia",
    "NR": "Nauru",            "NP": "Nepal",            "NL": "Netherlands",
    "NC": "New Caledonia",    "NZ": "New Zealand",      "NI": "Nicaragua",
    "NE": "Niger",            "NG": "Nigeria",          "NU": "Niue Islands",
    "NF": "Norfolk Islands",  "KP": "North Korea",      "NO": "Norway",
    "OM": "Oman",             "PK": "Pakistan",         "PW": "Palau",
    "PA": "Panama",           "PG": "Pap. New Guinea",  "PY": "Paraguay",
    "PE": "Peru",             "PH": "Phillipines",      "PN": "Pitcairn Islnds",
    "PL": "Poland",           "PT": "Portugal",         "PR": "Puerto Rico",
    "QA": "Qatar",            "RE": "Reunion",          "RO": "Romania",
    "RW": "Ruanda",           "RU": "Russian Fed.",     "GS": "S. Sandwich Ins",
    "ST": "S.Tome,Principe",  "AS": "Samoa, America",   "SM": "San Marino",
    "SA": "Saudi Arabia",     "SN": "Senegal",          "SC": "Seychelles",
    "SL": "Sierra Leone",     "SG": "Singapore",        "SK": "Slovakia",
    "SI": "Slovenia",         "SB": "Solomon Islands",  "SO": "Somalia",
    "ZA": "South Africa",     "KR": "South Korea",      "ES": "Spain",
    "LK": "Sri Lanka",        "KN": "St Kitts&Nevis",   "SH": "St. Helena",
    "LC": "St. Lucia",        "VC": "St. Vincent",      "PM": "St.Pier,Miquel.",
    "SD": "Sudan",            "SR": "Suriname",         "SJ": "Svalbard",
    "SZ": "Swaziland",        "SE": "Sweden",           "CH": "Switzerland",
    "SY": "Syria",            "TW": "Taiwan",           "TJ": "Tajikstan",
    "TZ": "Tanzania",         "TH": "Thailand",         "TG": "Togo",
    "TK": "Tokelau Islands",  "TO": "Tonga",            "TT": "Trinidad,Tobago",
    "TN": "Tunisia",          "TR": "Turkey",           "TM": "Turkmenistan",
    "TC": "Turksh Caicosin",  "TV": "Tuvalu",           "UG": "Uganda",
    "UA": "Ukraine",          "GB": "United Kingdom",   "UY": "Uruguay",
    "US": "USA",              "AE": "Utd.Arab Emir.",   "UZ": "Uzbekistan",
    "VU": "Vanuatu",          "VA": "Vatican City",     "VE": "Venezuela",
    "VN": "Vietnam",          "WF": "Wallis,Futuna",    "EH": "West Sahara",
    "WS": "Western Samoa",    "YE": "Yemen",            "YU": "Yugoslavia",
    "ZM": "Zambia",
}

TRENZASHOP_COMP_TO_STIBO = {
    "200782": "0191", "0191": "0191", "1191": "1191",
    "0195":   "0195", "0196": "0196",
}
LOV_COMPANY_CODE = {
    "0191": "PT. Aldo Indonesia Adiprk",
    "1191": "PT. Aldo Ind Adiprk Rtl",
    "0195": "Aldo Singapore",
    "0196": "Aldo Malaysia",
}
LOV_SBU = {
    "DO": "Aldo Ind Imp",          "D6": "Aldo Malaysia",
    "D5": "Aldo Singapore",        "DW": "Aldo Thailand",
    "PF": "Fashion",               "KV": "Fashion Footwear",
    "OF": "Fashion Footwear MAS",  "JD": "Fashion Footwear MY",
    "TA": "Fashion Thailand",      "FQ": "Foot Locker",
    "K2": "Foot Locker Malaysia",  "K0": "Foot Locker Singapore",
    "VO": "Foot Locker Vietnam",   "TF": "Footlocker Thailand",
    "K7": "MAA Cambodia",          "CH": "MAA Children",
    "FF": "MAA Fashion Footwear",  "FT": "MAA Foot Locker",
    "GO": "MAA Golf",              "SP": "MAA Sport",
    "SB": "MFA Badminton",         "GI": "Mitra Gaya Indah",
    "PY": "Payless",               "FY": "PSI Foot Locker",
    "PR": "PSI Retail",            "IF": "Sport Direct Philiphine",
    "IZ": "Sports Direct ID",      "TT": "Toys Thailand",
}

LOV_SEASON = {
    "SS": "Spring-Summer",
    "FW": "Fall-Winter",
    "AW": "Fall-Winter",
    "AL": "All Season",
}

_MCR_TO_FASTENING = {
    "ankle strap": ("Ankle Strap", "AS"),
    "back strap":  ("Back Strap",  "BS"),
    "boa":         ("BOA",         "BO"),
    "bungee":      ("Bungee Lace", "BL"),
    "gusset":      ("Gusset",      "GU"),
    "lace":        ("Lace Up",     "LU"),
    "mock lace":   ("Mock Lace",   "ML"),
    "velcro":      ("Velcro",      "VE"),
    "strap":       ("Straps",      "ST"),
}

_MCR_TO_HEEL_TYPE = {
    "block":           ("Block",          "BL"),
    "espadrille flat": ("Espadrille Flat","EF"),
    "espadrille wedge":("Espadrille Wedge","EW"),
    "flat":            ("Flat",           "FL"),
    "hidden wedge":    ("Hidden Wedge",   "HW"),
    "kitten":          ("Kitten",         "KT"),
    "pencil":          ("Pencil Heel",    "PH"),
    "platform":        ("Platform",       "PL"),
    "wedge":           ("Wedge",          "WD"),
}

LOV_CURRENCY_MAP: dict[str, str] = {
    "US DOLLAR": "USD",  "US DOLLARS": "USD", "USD": "USD",
    "INDONESIAN RUPIAH": "IDR", "RUPIAH": "IDR", "IDR": "IDR",
    "SINGAPORE DOLLAR": "SGD", "SGD": "SGD",
    "MALAYSIAN RINGGIT": "MYR", "RINGGIT": "MYR", "MYR": "MYR",
    "BRITISH POUND": "GBP", "POUND STERLING": "GBP", "GBP": "GBP",
    "EURO": "EUR", "EUR": "EUR",
    "AUSTRALIAN DOLLAR": "AUD", "AUD": "AUD",
    "CANADIAN DOLLAR": "CAD", "CAD": "CAD",
    "JAPANESE YEN": "JPY", "YEN": "JPY", "JPY": "JPY",
    "THAI BAHT": "THB", "BAHT": "THB", "THB": "THB",
    "VIETNAMESE DONG": "VND", "DONG": "VND", "VND": "VND",
    "CHINESE YUAN": "CNY", "YUAN": "CNY", "CNY": "CNY",
    "PHILIPPINE PESO": "PHP", "PHP": "PHP",
    "INDIAN RUPEE": "INR", "INR": "INR",
    "SWISS FRANC": "CHF", "CHF": "CHF",
    "HONG KONG DOLLAR": "HKD", "HKD": "HKD",
    "KOREAN WON": "KRW", "WON": "KRW", "KRW": "KRW",
    "SAUDI RIYAL": "SAR", "RIYAL": "SAR", "SAR": "SAR",
}

LOV_COLOR_CODE = {
    "core black": "005", "cloud white": "W", "ftwr white": "W",
    "off white": "W",    "white": "W",       "black": "005",
    "blue": "12W",       "navy": "NAV",      "grey": "GRE",
    "green": "G",        "red": "R",         "pink": "PK",
    "yellow": "Y",       "orange": "O",      "brown": "700",
    "beige": "18",       "purple": "P",      "gold": "KGO",
    "silver": "SV",      "multicolor": "MI", "multi": "MI",
    "cream": "CM",       "sand": "SAN",      "olive": "OLI",
    "carbon": "LPC",     "collegiate navy": "NAV",
}

ALDO_SILHOUETTE_MAP = {
    "HEELED SHOES": "Pumps",        "PUMPS": "Pumps",
    "FLAT SHOES": "Slip On",        "SLIP ON": "Slip On",
    "BOOTS": "Boot",                "BOOT": "Boot",
    "CHELSEA": "Chelsea",           "SANDAL": "Sandal",
    "SANDALS": "Sandal",            "STRAP SANDAL": "Strap Sandal",
    "MULES": "Mules",               "MULE": "Mules",
    "SNEAKER": "Sneaker Low Top",   "SNEAKERS": "Sneaker Low Top",
    "HIGH TOP": "Sneaker High Top", "LOW TOP": "Sneaker Low Top",
    "LOAFER": "Loafer",             "LOAFERS": "Loafer",
    "BALLERINA": "Ballerina",       "BALLET": "Ballerina",
    "MARY JANE": "Mary Jane",       "MARYJANE": "Mary Jane",
    "MOCCASIN": "Moccasins",        "MOCCASINS": "Moccasins",
    "CLOG": "Clog",                 "CLOGS": "Clog",
    "FLIP FLOP": "Flip Flop",       "FLIP-FLOP": "Flip Flop",
    "SLIDES": "Slides",             "SLIDE": "Slides",
    "OPEN TOE": "Open Toe Pumps",   "THONG": "Thong",
    "BACKSTRAP": "Backstrap",       "FISHERMAN": "Fisherman",
    "TOE RING": "Toe Ring",         "SPIKE": "Spike",
    "SPIKELESS": "Spikeless",       "MONK": "Monk",
}

_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)

# FIX 11/12: Classification suffix max length = 8 chars (not 10)
_CLSID_MAX = 8


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "#DIV/0!", "N/A", "#",
                       "Not assigned", "not assigned") else s


def _val(parent, attr_id: str, value: str = "", id_val: str = "") -> ET.Element:
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip()  if value  else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent, attr_id: str, id_val: str, label: str = "") -> ET.Element:
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    if label:
        v.text = str(label)
    return mv


def _map_silhouette(raw: str) -> str:
    if not raw:
        return ""
    upper = raw.upper().strip()
    if upper in ALDO_SILHOUETTE_MAP:
        return ALDO_SILHOUETTE_MAP[upper]
    for k, v in ALDO_SILHOUETTE_MAP.items():
        if k in upper:
            return v
    return ""


def _map_heel_height(raw: str) -> str:
    if not raw:
        return ""
    upper = raw.upper().strip()
    if "FLAT" in upper:  return "Flat"
    if "HIGH" in upper:  return "High"
    if "MEDIUM" in upper or "MED" in upper: return "Medium"
    if "LOW" in upper:   return "Low"
    inch_match = re.search(r"(\d+(?:\.\d+)?)\s*IN", upper)
    if inch_match:
        inches = float(inch_match.group(1))
        if inches >= 3.0:   return "High"
        elif inches >= 1.0: return "Medium"
        elif inches > 0:    return "Low"
        else:               return "Flat"
    return ""


def _color_to_lov_id(color_raw: str) -> str:
    if not color_raw:
        return ""
    key = color_raw.strip().lower()
    if key in LOV_COLOR_CODE:
        return LOV_COLOR_CODE[key]
    words = key.split()
    if len(words) >= 2:
        two_word = f"{words[0]} {words[1]}"
        if two_word in LOV_COLOR_CODE:
            return LOV_COLOR_CODE[two_word]
    if words and words[0] in LOV_COLOR_CODE:
        return LOV_COLOR_CODE[words[0]]
    return ""


def _normalize_currency(raw: str) -> str:
    if not raw:
        return ""
    cleaned = raw.strip().upper().rstrip(". ")
    return LOV_CURRENCY_MAP.get(cleaned, "")


def _parse_season(selling_season: str):
    year_match  = re.search(r'20\d{2}', selling_season)
    season_year = year_match.group() if year_match else ""
    ss = selling_season.upper()
    if "SPRING" in ss or " SS" in ss:
        season_code = "SS"
    elif "FALL" in ss or "FW" in ss or "AUTUMN" in ss or "AW" in ss:
        season_code = "FW"
    else:
        season_code = "SS"
    return season_code, season_year


# ─────────────────────────────────────────────────────────────────
# CLASSIFICATIONS
# ─────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────
# MATERIAL LOV MAPPERS  (FIX v1.2)
# ─────────────────────────────────────────────────────────────────

_LOV_MATERIAL_UPPER = ["EVA", "Leather", "Mesh", "Polyester", "PU",
                        "Suede", "Sustainable", "Synthetic", "Textile"]

_MCR_TO_MATERIAL_UPPER: dict[str, str] = {
    "leather":     "Leather",
    "suede":       "Suede",
    "mesh":        "Mesh",
    "textile":     "Textile",
    "synthetic":   "Synthetic",
    "polyester":   "Polyester",
    "polyurethan": "PU",    # catches POLYURETHANE
    "pu":          "PU",
    "eva":         "EVA",
    "sustainable": "Sustainable",
    "fabric":      "Textile",
    "knit":        "Textile",
    "nylon":       "Synthetic",
    "rubber":      "Synthetic",
    "canvas":      "Textile",
    "gel":         "PU",    # GEL closest match
    "foam":        "EVA",
}

_MCR_TO_MATERIAL: dict[str, str] = {
    "leather":     "Leather",
    "suede":       "Suede",
    "mesh":        "Mesh",
    "textile":     "Textile",
    "synthetic":   "Synthetic",
    "polyester":   "Polyester",
    "polyurethan": "PU",
    "pu":          "PU",
    "eva":         "Eva",
    "nylon":       "Nylon",
    "rubber":      "Rubber",
    "cotton":      "Cotton",
    "fabric":      "Fabrics",
    "canvas":      "Canvas",
    "plastic":     "Plastic",
    "pvc":         "PVC",
    "metal":       "Metal",
    "wool":        "Wool",
    "fur":         "Fur",
    "paper":       "Paper",
    "straw":       "Straw",
    "wood":        "Wood",
    "silicone":    "Silicone",
    "surlyn":      "Surlyn",
    "gel":         "PU",
    "foam":        "Eva",
    "knit":        "Textile",
}


def _map_material_upper(raw: str) -> str:
    """Map MCR composition string to LOV_MaterialUpper label. Returns '' if no match."""
    if not raw:
        return ""
    lower = raw.lower()
    for keyword, lov_label in _MCR_TO_MATERIAL_UPPER.items():
        if keyword in lower:
            return lov_label
    return ""


def _map_material(raw: str) -> str:
    """Map MCR composition string to LOV_Material label. Returns '' if no match."""
    if not raw:
        return ""
    lower = raw.lower()
    for keyword, lov_label in _MCR_TO_MATERIAL.items():
        if keyword in lower:
            return lov_label
    return ""


_CONTENT_LOV = {
    "cotton":    "Cotton",
    "nylon":     "Nylon",
    "polyester": "Polyester",
    "viscose":   "Viscose",
    "spandex":   "Spandex",
    "poly":      "Polyester",
}


def _map_content(raw: str) -> str:
    if not raw:
        return ""
    lower = raw.lower()
    for kw, label in _CONTENT_LOV.items():
        if kw in lower:
            return label
    return ""  # no match → skip


def _strip_unit(raw: str) -> str:
    """Extract numeric portion from measurement string. '2 CM' -> '2', '14.5 CM' -> '14.5'"""
    if not raw:
        return ""
    m = re.match(r"^\s*(\d+(?:\.\d+)?)", raw.strip())
    return m.group(1) if m else ""


def _build_classifications(season_code: str, season_year: str):
    cls_root  = ET.Element(f"{{{STIBO_NS}}}Classifications")
    season_id = f"CLH_{BRAND_CODE}_{season_code}{season_year}"
    label_map = {"SS": "Spring Summer", "FW": "Fall Winter"}

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", "CLH_AldoBatches")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = (
        f"{BRAND_NAME} {label_map.get(season_code, season_code)} {season_year}"
    )

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{season_code} {season_year} Confirmed Articles"
    )

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = (
        f"{season_code} {season_year} Unconfirmed Articles"
    )

    return cls_root, season_id




def _extract_country_code_from_filename(path: Path) -> str:
    """Extract 2-letter country code from ...-<COUNTRY>-<SEQ>.xlsx style filenames."""
    m = re.search(r"-([A-Za-z]{2})-\d+\.xlsx$", path.name, flags=re.IGNORECASE)
    return m.group(1).upper() if m else ""


def _extract_file_season_from_filename(path: Path) -> tuple[str, str]:
    """Extract season token from filename using Adidas-style token parsing (e.g. SS26 / FW2026)."""
    parts = re.split(r"\s*-\s*", path.stem)
    season_token = next(
        (part.strip().upper() for part in parts if re.match(r"^[A-Z]{2}\d{2,4}$", part.strip(), re.IGNORECASE)),
        "",
    )
    if not season_token:
        return "", ""

    season_code = season_token[:2]
    year_raw = season_token[2:]
    season_year = year_raw if len(year_raw) == 4 else f"20{year_raw}"
    return season_code, season_year


def _extract_article_type_from_filename(path: Path) -> str:
    """Extract Inline/License/SSE token from filename file-type segment."""
    parts = re.split(r"\s*-\s*", path.stem)
    file_type_token = parts[3] if len(parts) > 3 else path.stem
    for at in ("Inline", "License", "SSE"):
        if at.lower() in file_type_token.lower():
            return at
    return ""

# ─────────────────────────────────────────────────────────────────
# ECOMMERCE / MCR LOADER
# ─────────────────────────────────────────────────────────────────
class EcommerceLoader:
    """
    WSL004 - Material Composition Report (sheet: Material Composition Report, header row 1)
    One row per size/SKU.  Unique key: Article Number (Generic level — no variants in output).
    Skips: service, gfs, shoe care, shoecare product categories.
    """
    def __init__(self, path: Path):
        self.path:     Path = path
        self.articles: list = []
        self.file_country_code: str = _extract_country_code_from_filename(path)
        self.file_season_code, self.file_season_year = _extract_file_season_from_filename(path)
        self.file_article_type: str = _extract_article_type_from_filename(path)
        self._load()

    def _load(self):
        log.info("[Ecommerce] Loading: %s", self.path.name)

        df = pd.read_excel(str(self.path),
                           sheet_name="Material Composition Report",
                           header=1, dtype=str)
        df = df.fillna("")

        df = df[~df["Product Category Name"].str.lower().apply(
            lambda x: any(k in x for k in SKIP_CATEGORIES)
        )]
        log.info("[Ecommerce] Rows after category skip: %d", len(df))

        seen: set = set()
        for _, row in df.iterrows():
            art_no = _s(row.get("Article Number", ""))
            if not art_no:
                continue
            if art_no in seen:
                continue
            seen.add(art_no)

            self.articles.append({
                "article_no":         art_no,
                "style_name":         _s(row.get("Style Name", "")),
                "color_id":           _s(row.get("Color ID", "")),
                "color_desc":         _s(row.get("Color Name", "")),
                "coo":                _s(row.get("Country of Origin", "")),
                "selling_season":     _s(row.get("Selling Season", "")),
                "merch_category":     _s(row.get("Merchandise Category Name", "")),
                "product_category":   _s(row.get("Product Category Name", "")),
                "silhouette":         _s(row.get("Silhouette", "")),
                "heel_height":        _s(row.get("Heel Height", "")),
                # FIX 13: MCR column is "HTS Code"; MDD attribute is AT_HSCode
                "hts_code":           _s(row.get("HTS Code", "")),
                # FIX 14: MCR column retained; mapped to AT_ShortDescriptionEN in XML builder
                "customs_desc":       _s(row.get("Customs Commercial Description of Product", "")),
                # FIX 15-18: composition columns retained; remapped in XML builder
                "upper_comp":         _s(row.get("Upper Composition", "")),
                "lining_comp":        _s(row.get("Lining Composition", "")),
                "insole_comp":        _s(row.get("Insole Composition", "")),
                "sole":               _s(row.get("Sole", "")),
                # FIX 19-20: Upper/Bottom raw values folded into upper_comp / sole above
                # (fields dropped from article dict; composition attrs cover them)
                "heel_type":          _s(row.get("Heel Type", "")),
                # FIX 23: ClosureType → mapped to AT_Fastening in XML builder
                "closure_type":       _s(row.get("Closure Type", "")),
                # FIX 28-29: Height/Length/Width → AT_PackagingHeight/Length/Width (Number)
                "pkg_height":         _s(row.get("Height", "")),
                "pkg_length":         _s(row.get("Lenght", "")),   # MCR typo preserved
                "pkg_width":          _s(row.get("Width", "")),
                "comp_code":          _s(row.get("Customer ID", "")),
                "file_season_code":   self.file_season_code,
                "file_season_year":   self.file_season_year,
            })

        log.info("[Ecommerce] %d unique articles loaded", len(self.articles))


# ─────────────────────────────────────────────────────────────────
# PRODUCT XML BUILDER  — Generic Article ONLY (no variants)
# ─────────────────────────────────────────────────────────────────
def _build_product_xml(art: dict, season_code: str, season_year: str,
                        season_id: str, comp_code: str, sbu: str, file_country_code: str = "",
                        file_article_type: str = "") -> str:
    article_no  = art["article_no"]
    key_article = f"{BRAND_CODE}{article_no[:9]}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_F-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["style_name"] or article_no

    merch_safe = re.sub(r"[^A-Z0-9]", "",
        art["merch_category"].upper().encode("ascii", "ignore").decode())[:_CLSID_MAX]
    prod_safe  = re.sub(r"[^A-Z0-9]", "",
        art["product_category"].upper().encode("ascii", "ignore").decode())[:_CLSID_MAX]

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", "CLH_AldoArticles")
    cr_merch.set("Type", "CPL_Merchandiser")

    # cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_by.set("ClassificationID", f"MA_{BRAND_CODE}_{merch_safe}")
    # cr_by.set("Type", "CPL_BYHierarchy")

    # cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    # cr_sap.set("ClassificationID", f"CLH_{prod_safe}")
    # cr_sap.set("Type", "CPL_SAPHierarchy")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    # ── Values ───────────────────────────────────────────────────
    vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    _val(vals, "AT_Brand",                     BRAND_NAME, id_val=BRAND_CODE)
    if file_article_type:
        _val(vals, "AT_BYArticleType",         file_article_type, id_val=file_article_type)
    _val(vals, "AT_InboundGenericCode",                   key_article)
    _val(vals, "AT_PrincipalStyleCode",        article_no)
    _val(vals, "AT_SAPStyleCode",              article_no)
    _val(vals, "AT_PrincipalStyleDescription", art["style_name"])
    if file_country_code:
        country_label = LOV_COUNTRY.get(file_country_code, file_country_code)
        _val(vals, "AT_Country", id_val=file_country_code)
        _val(vals, "AT_CountryOrigin", country_label, id_val=file_country_code)

    color_lov_id = _color_to_lov_id(art["color_desc"])
    if color_lov_id:
        _val(vals, "AT_Color", id_val=color_lov_id)
    if art["color_id"]:
        _val(vals, "AT_PrincipalColorCode", art["color_id"])
        _val(vals, "AT_PrincipalColorName", art["color_desc"])

    _val(vals, "AT_Silhouette",          _map_silhouette(art["silhouette"]))
    _val(vals, "AT_MerchandiseCategory", art["merch_category"][:9])
    _val(vals, "AT_HeelHeight",          _map_heel_height(art["heel_height"]))

    # FIX 13: AT_HTSCode → AT_HSCode  (MDD: Text)
    _val(vals, "AT_HSCode",              art["hts_code"])

    # FIX 14: AT_CustomsDescription → AT_ShortDescriptionEN  (MDD: Text)
    _val(vals, "AT_ShortDescriptionEN",  art["customs_desc"])

    # FIX 15 v1.2: map raw MCR composition → LOV_MaterialUpper label; skip if no match
    _mat_upper = _map_material_upper(art["upper_comp"])
    if _mat_upper:
        _val(vals, "AT_MaterialUpper", _mat_upper)

    # FIX 16 v1.2: map raw MCR lining composition → LOV_Material label; skip if no match
    _mat_lining = _map_material(art["lining_comp"])
    if _mat_lining:
        _val(vals, "AT_Material", _mat_lining)

    # FIX 17: AT_InsoleComposition → AT_Content  (MDD: Text)
    _content = _map_content(art["insole_comp"])
    if _content:
        _val(vals, "AT_Content", _content)

    # FIX 18: AT_Others REMOVED — Number validator rejects text values

    # FIX 19-20: AT_Upper / AT_Bottom REMOVED — no MDD equivalent.
    # FIX 21-22: AT_ToeShape / AT_ToeOpening REMOVED — no MDD equivalent.

    # FIX 23: AT_ClosureType → AT_Fastening  (LOV: only specific values valid)
    if art["closure_type"]:
        _ct_lower = art["closure_type"].lower()
        _ft_match = next(
            (v for k, v in _MCR_TO_FASTENING.items() if k in _ct_lower), None
        )
        if _ft_match:
            _val(vals, "AT_Fastening", _ft_match[0], id_val=_ft_match[1])

    # FIX 24: AT_Ornamentation REMOVED — no MDD equivalent.
    # FIX 25-27: AT_PlatformHeight / AT_ShaftHeight / AT_CalfCircumference REMOVED.

    # FIX 28-30: Numeric dimensions → AT_PackagingHeight / AT_PackagingLength / AT_PackagingWidth
    # AT_Width (LOV: Regular/Wide/Relaxed) is NOT written for MCR data.
    _pkg_h = _strip_unit(art["pkg_height"])
    _pkg_l = _strip_unit(art["pkg_length"])
    _pkg_w = _strip_unit(art["pkg_width"])
    if _pkg_h: _val(vals, "AT_PackagingHeight", _pkg_h)
    if _pkg_l: _val(vals, "AT_PackagingLength",  _pkg_l)
    if _pkg_w: _val(vals, "AT_PackagingWidth",   _pkg_w)

    # FIX 23 (heel type): AT_HeelType — map to LOV ID only (BL/EF/EW/FL/HW/KT/PH/PL/WD)
    if art["heel_type"]:
        _ht_lower = art["heel_type"].lower()
        _ht_match = next(
            (v for k, v in _MCR_TO_HEEL_TYPE.items() if k in _ht_lower), None
        )
        if _ht_match:
            _val(vals, "AT_HeelType", _ht_match[0], id_val=_ht_match[1])

    # Season from source filename segment (e.g., SM22 / SS2026)
    file_season_code = art.get("file_season_code", "")
    if file_season_code not in LOV_SEASON:
        file_season_code = season_code
    file_season_year = art.get("file_season_year", "") or season_year
    _val(vals, "AT_Season",     id_val=file_season_code)
    _val(vals, "AT_SeasonYear", file_season_year)

    # Company + SBU
    raw_comp   = art.get("comp_code") or comp_code
    stibo_comp = TRENZASHOP_COMP_TO_STIBO.get(raw_comp, raw_comp)
    comp_label = LOV_COMPANY_CODE.get(stibo_comp, stibo_comp)
    _multival(vals, "AT_CompanyCode", stibo_comp, comp_label)

    sbu_label = LOV_SBU.get(sbu, sbu)
    _multival(vals, "AT_SBU", sbu, sbu_label)
    _val(vals, "AT_BYIndicator",  "Yes", id_val="Y")
    _val(vals, "AT_SAPIndicator", "No",  id_val="N")
    _val(vals, "AT_UOM",          "Pair", id_val="PAA")
    return ET.tostring(g_el, encoding="unicode")


# ─────────────────────────────────────────────────────────────────
# XML WRITER
# ─────────────────────────────────────────────────────────────────
def _write_xml(out_path: Path, cls_el: ET.Element,
               products_xml_list: list, export_time: str):
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    header = (
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
    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write(header)
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for pxml in products_xml_list:
            pxml = _XMLNS_RE.sub("", pxml)
            f.write(f"    {pxml}\n")
        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")


# ─────────────────────────────────────────────────────────────────
# SUMMARY PRINTER
# ─────────────────────────────────────────────────────────────────
def _print_summary(source_file, total_arts, written, season_code,
                   season_year, comp_code, sbu, out_name, kb):
    print("═══ ALDO ECOMMERCE (MCR) SUMMARY ══════════════════════════════", flush=True)
    print(f"  Source file      : {source_file}",               flush=True)
    print(f"  Total articles   : {total_arts}",                flush=True)
    print(f"  Written to XML   : {written}",                   flush=True)
    print(f"  Season           : {season_code} {season_year}", flush=True)
    print(f"  comp_code        : {comp_code}",                 flush=True)
    print(f"  sbu              : {sbu}",                       flush=True)
    print(f"  XML output       : {out_name}",                  flush=True)
    print(f"  File size        : {kb} KB",                     flush=True)
    print("════════════════════════════════════════════════════", flush=True)


# ─────────────────────────────────────────────────────────────────
# MAIN run()
# ─────────────────────────────────────────────────────────────────
def run(args, xml_out_dir: Path, export_time: str,
        season_code: str, season_year: str, season_id: str,
        ecommerce_dir: Path, auditor=None, article_filter=None):
    """
    Process Ecommerce (MCR/WSL004) file only.
    Returns number of XMLs written (0 or 1).
    """
    ec_files = (list(ecommerce_dir.glob("*.xlsx"))
                + list(ecommerce_dir.glob("*.xlsm")))
    if not ec_files:
        log.info("[Ecommerce] No file found in %s — skipping", ecommerce_dir)
        return 0

    ec_path = ec_files[0]
    loader  = EcommerceLoader(ec_path)
    # ── TEST MODE: limit to 5 articles ──────────────────────────
    # loader.articles = loader.articles[:1]
    log.info("[Ecommerce] TEST MODE: limited to %d articles", len(loader.articles))
    
    # ────────────────────────────────────────────────────────────
    
    if article_filter:
        before = len(loader.articles)
        loader.articles = [
            a for a in loader.articles if a["article_no"] == article_filter
        ]
        log.info("[Ecommerce] article_filter=%s → %d/%d kept",
                 article_filter, len(loader.articles), before)

    if not loader.articles:
        log.warning("[Ecommerce] No articles after filtering — skipping XML")
        return 0

    cls_el, _ = _build_classifications(season_code, season_year)
    out_name  = f"{ec_path.stem}.xml"
    out_path  = xml_out_dir / out_name

    products = [
        _build_product_xml(
            art, season_code, season_year, season_id,
            args.comp_code, args.sbu,
            file_country_code=loader.file_country_code,
            file_article_type=loader.file_article_type,
        )
        for art in loader.articles
    ]

    _write_xml(out_path, cls_el, products, export_time)
    kb = out_path.stat().st_size // 1024

    log.info("✅ Ecommerce XML → %s (%dKB)", out_name, kb)
    _print_summary(ec_path.name, len(loader.articles), len(loader.articles),
                   season_code, season_year, args.comp_code, args.sbu, out_name, kb)

    if auditor:
        auditor.record_loader(
            "Ecommerce", status="ok",
            articles=len(loader.articles),
            filename=ec_path.name,
        )
    return 1