"""
aldo/main.py
============
Aldo TrenzaShop line-list processor.

OOR report and OOR size-run processing lives in aldo/oor.py.
"""

import re
import os
import sys
import logging
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
import openpyxl

BASE_DIR = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

INPUT_DIR    = BASE_DIR / "input"
LINELIST_DIR = INPUT_DIR / "linelist"
OUTPUT_DIR   = BASE_DIR / "output"
XML_OUT_DIR  = OUTPUT_DIR / "xml"
LOG_DIR      = OUTPUT_DIR / "logs"

for d in [LINELIST_DIR, XML_OUT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

log_path = LOG_DIR / f"aldo_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("aldo_etl")

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────
STIBO_NS     = "http://www.stibosystems.com/step"
STIBO_XSI    = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

BRAND_CODE = "AOD"
BRAND_NAME = "ALDO"

HEADER_ROW_IDX   = 6
SIZE_COL_INDICES = list(range(59, 72))

# FIX 24: LOV_SEASON — maps sea_prefix → Stibo display text (same as TDD)
LOV_SEASON = {
    "SS": "Spring-Summer",
    "FW": "Fall-Winter",
    "FL": "Fall-Winter",
    "AW": "Fall-Winter",
    "SW": "Fall-Winter",
    "WN": "Fall-Winter",   # ← Winter = Fall-Winter bucket
    "SM": "Spring-Summer",
    "AL": "All Season",
    "SP": "Spring-Summer",
}

# FIX 28: Currency name/variant → valid Stibo Currency LOV ID
# Source sends "US DOLLAR", "USD", "IDR " etc — normalized to valid LOV IDs.
# Only IDs confirmed in Stibo Currency domain (image1) are included.
LOV_CURRENCY_MAP: dict[str, str] = {
    "US DOLLAR": "USD",  "US DOLLARS": "USD", "UNITED STATES DOLLAR": "USD", "USD": "USD",
    "INDONESIAN RUPIAH": "IDR", "RUPIAH": "IDR", "IDR": "IDR",
    "SINGAPORE DOLLAR": "SGD", "SGD": "SGD",
    "MALAYSIAN RINGGIT": "MYR", "RINGGIT": "MYR", "MYR": "MYR",
    "BRITISH POUND": "GBP", "POUND STERLING": "GBP", "STERLING": "GBP", "GBP": "GBP",
    "EURO": "EUR", "EUR": "EUR",
    "AUSTRALIAN DOLLAR": "AUD", "AUD": "AUD",
    "CANADIAN DOLLAR": "CAD", "CAD": "CAD",
    "JAPANESE YEN": "JPY", "YEN": "JPY", "JPY": "JPY",
    "THAI BAHT": "THB", "BAHT": "THB", "THB": "THB",
    "VIETNAMESE DONG": "VND", "DONG": "VND", "VND": "VND",
    "CHINESE YUAN": "CNY", "YUAN": "CNY", "CNY": "CNY", "RMB": "RMB",
    "PHILIPPINE PESO": "PHP", "PHP": "PHP",
    "CAMBODIAN RIEL": "KHR", "RIEL": "KHR", "KHR": "KHR",
    "INDIAN RUPEE": "INR", "INR": "INR",
    "PAKISTANI RUPEE": "PKR", "PKR": "PKR",
    "BANGLADESHI TAKA": "BDT", "TAKA": "BDT", "BDT": "BDT",
    "SWISS FRANC": "CHF", "CHF": "CHF",
    "NEW ZEALAND DOLLAR": "NZD", "NZD": "NZD",
    "HONG KONG DOLLAR": "HKD", "HKD": "HKD",
    "TAIWAN DOLLAR": "TWD", "NEW TAIWAN DOLLAR": "TWD", "TWD": "TWD",
    "KOREAN WON": "KRW", "WON": "KRW", "KRW": "KRW",
    "SAUDI RIYAL": "SAR", "RIYAL": "SAR", "SAR": "SAR",
    "SOUTH AFRICAN RAND": "ZAR", "RAND": "ZAR", "ZAR": "ZAR",
    "NORWEGIAN KRONE": "NOK", "NOK": "NOK",
    "DANISH KRONE": "DKK", "DKK": "DKK",
    "SWEDISH KRONA": "SEK", "SEK": "SEK",
    "TURKISH LIRA": "TRL", "TRL": "TRL",
    "BRAZILIAN REAL": "BRL", "REAL": "BRL", "BRL": "BRL",
    "MEXICAN PESO": "MXN", "MXN": "MXN",
    "QATARI RIYAL": "QAR", "QAR": "QAR",
    "KUWAITI DINAR": "KWD", "KWD": "KWD",
    "BAHRAINI DINAR": "BHD", "BHD": "BHD",
    "EGYPTIAN POUND": "EGP", "EGP": "EGP",
}


def _normalize_currency(raw: str) -> str:
    """
    FIX 28: Map raw currency string → valid Stibo Currency LOV ID.
    Returns "" if unmappable — caller must skip writing the attribute.
    """
    if not raw:
        return ""
    cleaned = raw.strip().upper()
    if cleaned in LOV_CURRENCY_MAP:
        return LOV_CURRENCY_MAP[cleaned]
    cleaned2 = cleaned.rstrip(". ")
    if cleaned2 in LOV_CURRENCY_MAP:
        return LOV_CURRENCY_MAP[cleaned2]
    log.warning("[Currency] Unknown currency '%s' — skipping AT_RetailPriceCurrency", raw)
    return ""


LOV_COUNTRY = {
    "VN": "Vietnam",    "CN": "China",       "ID": "Indonesia",
    "KH": "Cambodia",   "BD": "Bangladesh",  "IN": "India",
    "MY": "Malaysia",   "TH": "Thailand",    "PH": "Philippines",
    "PK": "Pakistan",   "TR": "Turkey",      "CA": "Canada",
    "IT": "Italy",      "PT": "Portugal",    "ES": "Spain",
    "CH": "Switzerland",
}

TRENZASHOP_COMP_TO_STIBO: dict[str, str] = {
    "200782": "0191",
    "0191":   "0191",
    "1191":   "1191",
    "0195":   "0195",
    "0196":   "0196",
}

LOV_COMPANY_CODE = {
    "0191": "PT. Aldo Indonesia Adiprk",
    "1191": "PT. Aldo Ind Adiprk Rtl",
    "0195": "Aldo Singapore",
    "0196": "Aldo Malaysia",
}

LOV_SBU: dict[str, str] = {
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
    "OM": "Sport MAA Malaysia (Reebok/ Wholesale)",
    "OU": "Sport MAA Singapore (Reebok/ Wholesale)",
    "QY": "Sport Malaysia (Converse)",
    "QS": "Sport Singapore (Converse)",
    "TW": "Sport Thailand",        "IZ": "Sports Direct ID",
    "TT": "Toys Thailand",
}

ALDO_SILHOUETTE_MAP: dict[str, str] = {
    "HEELED SHOES": "Pumps",       "PUMPS":       "Pumps",
    "FLAT SHOES":   "Slip On",     "SLIP ON":     "Slip On",
    "BOOTS":        "Boot",        "BOOT":        "Boot",
    "CHELSEA":      "Chelsea",     "SANDAL":      "Sandal",
    "SANDALS":      "Sandal",      "STRAP SANDAL":"Strap Sandal",
    "MULES":        "Mules",       "MULE":        "Mules",
    "SNEAKER":      "Sneaker Low Top", "SNEAKERS": "Sneaker Low Top",
    "HIGH TOP":     "Sneaker High Top","LOW TOP":  "Sneaker Low Top",
    "LOAFER":       "Loafer",      "LOAFERS":     "Loafer",
    "BALLERINA":    "Ballerina",   "BALLET":      "Ballerina",
    "MARY JANE":    "Mary Jane",   "MARYJANE":    "Mary Jane",
    "MOCCASIN":     "Moccasins",   "MOCCASINS":   "Moccasins",
    "MONK":         "Monk",        "CLOG":        "Clog",
    "CLOGS":        "Clog",        "FLIP FLOP":   "Flip Flop",
    "FLIP-FLOP":    "Flip Flop",   "SLIDES":      "Slides",
    "SLIDE":        "Slides",      "OPEN TOE":    "Open Toe Pumps",
    "THONG":        "Thong",       "BACKSTRAP":   "Backstrap",
    "FISHERMAN":    "Fisherman",   "TOE RING":    "Toe Ring",
    "SPIKE":        "Spike",       "SPIKELESS":   "Spikeless",
    "UNKNOWN":      "Unknown",
}

MDD_SIZE_LOV: dict[str, str] = {
    "1":  "001", "2":  "002", "3":  "003", "4":  "004",
    "5":  "005", "6":  "006", "7":  "007", "8":  "008",
    "9":  "009", "10": "010", "11": "011", "12": "012",
    "13": "013", "14": "014", "15": "015", "16": "016",
    "17": "017", "18": "018", "19": "019", "20": "020",
    "21": "021", "22": "022", "23": "023", "24": "024",
    "25": "025", "26": "026", "27": "027", "28": "028",
    "29": "029", "30": "030", "31": "031", "32": "032",
    "33": "033", "34": "034", "35": "035", "36": "036",
    "37": "037", "38": "038", "39": "039", "40": "040",
    "41": "041", "42": "042", "43": "043", "44": "044",
    "45": "045", "46": "046", "47": "047", "48": "048",
    "49": "049", "50": "050", "52": "052", "54": "054",
    "55": "055", "56": "056", "58": "058", "62": "062",
    "65": "065", "68": "068", "70": "070", "74": "074",
    "77": "077", "78": "078", "80": "080", "86": "086",
    "89": "089",
    "101": "101", "104": "104", "105": "105", "110": "110",
    "115": "115", "116": "116", "120": "120", "122": "122",
    "123": "123", "124": "124", "128": "128", "130": "130",
    "134": "134", "135": "135",
    "1.5":  "01H", "2.5":  "02H", "3.5":  "03H", "4.5":  "04H",
    "5.5":  "00S", "6.5":  "01-", "7.5":  "01X", "8.5":  "02-",
    "9.5":  "02X", "10.5": "01C", "11.5": "01K",
    "XS":  "00X", "S":   "00S", "M":   "00M", "L":   "00L",
    "XL":  "00X", "XXL": "00X", "XXXL":"00X",
    "OS":  "000", "ONE SIZE": "000",
    "S/M": "00S", "M/L": "00M", "L/XL": "00L",
    "10C": "10C", "C": "00C",   "F": "00F",   "H": "00H",
    "2T":  "002", "3T": "003",  "4T": "004",
    "0":   "000", "NO SIZE": "000",
}

MDD_SIZE_DESC: dict[str, str] = {
    "01H": "1.5",  "02H": "2.5",  "03H": "3.5",  "04H": "4.5",
    "01-": "6.5",  "01X": "7.5",  "02-": "8.5",  "02X": "9.5",
    "01C": "10.5", "01K": "11.5", "00S": "S",     "00X": "XS",
    "00M": "M",    "00L": "L",    "00C": "C",     "00F": "F",
    "00H": "H",    "000": "OS",   "10C": "10C",
}

def _mdd_size_code(size_raw: str) -> str:
    """Convert size string → MDD AT_Size LOV ID. Fallback = '000'."""
    s = size_raw.strip()
    if s in MDD_SIZE_LOV:
        return MDD_SIZE_LOV[s]
    upper = s.upper()
    for k, v in MDD_SIZE_LOV.items():
        if k.upper() == upper:
            return v
    try:
        f = float(s)
        if f == int(f):
            whole = str(int(f))
            if whole in MDD_SIZE_LOV:
                return MDD_SIZE_LOV[whole]
    except ValueError:
        pass
    log.warning("[SizeCode] Unknown size '%s' — using '000'", size_raw)
    return "000"


def _map_heel_height(raw: str) -> str:
    if not raw:
        return ""
    upper = raw.upper().strip()
    if "FLAT" in upper:
        return "Flat"
    if "HIGH" in upper:
        return "High"
    if "MEDIUM" in upper or "MED" in upper:
        return "Medium"
    if "LOW" in upper:
        return "Low"
    inch_match = re.search(r"(\d+(?:\.\d+)?)\s*IN", upper)
    if inch_match:
        inches = float(inch_match.group(1))
        if inches >= 3.0:    return "High"
        elif inches >= 1.0:  return "Medium"
        elif inches > 0:     return "Low"
        else:                return "Flat"
    return ""


def _map_silhouette(raw: str) -> str:
    if not raw:
        return ""
    upper = raw.upper().strip()
    if upper in ALDO_SILHOUETTE_MAP:
        return ALDO_SILHOUETTE_MAP[upper]
    for key, val in ALDO_SILHOUETTE_MAP.items():
        if key in upper:
            return val
    return ""


_XMLNS_RE = re.compile(r'\s+xmlns="[^"]*"')
ET.register_namespace('', STIBO_NS)


# ─────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────────────────────────────
def _s(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("None", "nan", "0", "#DIV/0!", "N/A", "#", "Not assigned") else s


def _val(parent, attr_id, value="", id_val=""):
    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    clean_id  = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip()  if value  else ""
    if clean_id and clean_id not in ("", "None", "nan"):
        el.set("ID", clean_id)
    if clean_val and clean_val not in ("None", "nan"):
        el.text = clean_val
    return el


def _multival(parent, attr_id, id_val, label=""):
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    if label:
        v.text = str(label)
    return mv


def _parse_season(selling_season: str):
    season_id_match = re.search(r'\[(\d+)\]', selling_season)
    season_id = season_id_match.group(1) if season_id_match else ""
    year_match = re.search(r'20\d{2}', selling_season)
    season_year = year_match.group() if year_match else ""
    ss = selling_season.upper()
    if "SPRING" in ss or " SS" in ss:
        season_code = "SS"
    elif "FALL" in ss or "FL" in ss or "FW" in ss or "AUTUMN" in ss or "AW" in ss:
        season_code = "FL"   # use FL if Aldo's internal season name says Fall
    else:
        season_code = "SS"
    return season_code, season_year, season_id


def _resolve_comp_code(raw_comp: str) -> tuple[str, str]:
    stibo_id = TRENZASHOP_COMP_TO_STIBO.get(raw_comp, raw_comp)
    label = LOV_COMPANY_CODE.get(stibo_id, stibo_id)
    return stibo_id, label


def _resolve_sbu(sbu_id: str) -> tuple[str, str]:
    label = LOV_SBU.get(sbu_id, sbu_id)
    return sbu_id, label


def _xml_header(export_time: str) -> str:
    return (
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


def build_classifications(season_code: str, season_year: str) -> ET.Element:
    cls_root     = ET.Element(f"{{{STIBO_NS}}}Classifications")
    full_season  = f"{season_code}{season_year}"
    season_id    = f"CLH_{BRAND_CODE}_{full_season}"
    batches_par  = "CLH_AldoBatches"
    label_map    = {"SS": "Spring Summer", "FW": "Fall Winter"}
    sea_name     = label_map.get(season_code, season_code)
    season_disp  = f"{BRAND_NAME} {sea_name} {season_year}"
    season_short = f"{season_code} {season_year}"

    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", batches_par)
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = season_disp

    confirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    confirmed.set("ID", f"{season_id}CA")
    confirmed.set("UserTypeID", "CLS_ConfirmedArticles")
    ET.SubElement(confirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Confirmed Articles"

    unconfirmed = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    unconfirmed.set("ID", f"{season_id}UA")
    unconfirmed.set("UserTypeID", "CLS_UnconfirmedArticles")
    ET.SubElement(unconfirmed, f"{{{STIBO_NS}}}Name").text = f"{season_short} Unconfirmed Articles"

    return cls_root


def _write_xml(out_path: Path, cls_el: ET.Element, products_xml_list: list,
               export_time: str):
    cls_str = _XMLNS_RE.sub("", ET.tostring(cls_el, encoding="unicode"))
    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write(_xml_header(export_time))
        f.write(f"  {cls_str}\n\n")
        f.write("  <Products>\n")
        for pxml in products_xml_list:
            pxml = _XMLNS_RE.sub("", pxml)
            f.write(f"    {pxml}\n")
        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")


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


# ─────────────────────────────────────────────────────────────────
# 1. TRENZASHOP LOADER + XML  ← UNTOUCHED — Line List intact
# ─────────────────────────────────────────────────────────────────
class TrenzaShopLoader:
    def __init__(self, path: Path):
        self.path     = path
        self.articles = []
        self.sizes    = {}
        self.file_country_code = _extract_country_code_from_filename(path)
        self.file_season_code, self.file_season_year = _extract_file_season_from_filename(path)
        self._load()

    def _load(self):
        log.info("[TrenzaShop] Loading: %s", self.path.name)
        wb   = openpyxl.load_workbook(str(self.path), read_only=True, data_only=True)
        ws   = wb["MainSheet"]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if len(rows) <= HEADER_ROW_IDX:
            log.error("[TrenzaShop] Too few rows")
            return

        header = rows[HEADER_ROW_IDX]
        col = {}
        for i, h in enumerate(header):
            if h and str(h).strip() and str(h).strip() != "HEADINGS":
                key = str(h).strip()
                if key not in col:
                    col[key] = i

        size_headers = []
        for idx in SIZE_COL_INDICES:
            if idx < len(header) and header[idx] is not None:
                label = str(header[idx]).strip()
                if label:
                    size_headers.append((idx, label))

        seen_articles = set()

        for row in rows:
            if not row or row[0] != "D":
                continue

            art_no = _s(row[col.get("Generic Article", 5)] if "Generic Article" in col else row[5])
            if not art_no:
                continue
            try:
                int(art_no)
            except ValueError:
                continue

            sizes_for_row = []
            for idx, size_label in size_headers:
                qty = row[idx] if idx < len(row) else None
                if qty is not None and str(qty).strip() not in ("", "None", "nan", "0"):
                    try:
                        qty_val = float(str(qty).strip())
                        if qty_val > 0:
                            sizes_for_row.append({"size": size_label, "qty": qty_val})
                    except ValueError:
                        pass

            def g(col_name):
                idx = col.get(col_name)
                if idx is None or idx >= len(row):
                    return ""
                return _s(row[idx])

            if art_no not in seen_articles:
                seen_articles.add(art_no)
                self.articles.append({
                    "article_no":         art_no,
                    "style_name":         g("Style Name"),
                    "color_id":           g("Color ID"),
                    "color_desc":         g("Color Description"),
                    "coo":                g("COO"),
                    "main_material":      g("Main Material"),
                    "product_category":   g("Product Category"),
                    "merch_category":     g("Merchandise Category"),
                    "silhouette":         g("Silhouette"),
                    "article_collection": g("Article Collection"),
                    "heel_height":        g("Heel Height"),
                    "sugg_retail":        g("International Sugg. Retail"),
                    "currency":           g("Currency"),
                    "wholesale_price":    g("Wholesale Price"),
                    "selling_season":     g("Selling Season"),
                })
                self.sizes[art_no] = sizes_for_row
            else:
                existing       = self.sizes.get(art_no, [])
                existing_sizes = {s["size"] for s in existing}
                for s in sizes_for_row:
                    if s["size"] not in existing_sizes:
                        existing.append(s)
                        existing_sizes.add(s["size"])
                self.sizes[art_no] = existing

        log.info("[TrenzaShop] %d unique articles loaded", len(self.articles))


def _build_trenzashop_product_xml(art, sizes, season_code, season_year,
                                   season_id, comp_code, sbu, article_type="", country_code="") -> str:
    # ── UNTOUCHED — TrenzaShop / Line List logic preserved exactly ──
    article_no  = art["article_no"]
    key_article = f"{BRAND_CODE}{article_no}"

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID",   "PPH_F-TempSubCat")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    ET.SubElement(g_el, f"{{{STIBO_NS}}}Name").text = art["style_name"] or article_no

    cr_merch = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_merch.set("ClassificationID", "CLH_AldoArticles")
    cr_merch.set("Type", "CPL_Merchandiser")

#     merch_cat_safe = re.sub(r"[^A-Z0-9]", "", art["merch_category"].upper())[:8]
#     cr_by = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
#     cr_by.set("ClassificationID", f"MA_{BRAND_CODE}_{merch_cat_safe}")
#     cr_by.set("Type", "CPL_BYHierarchy")

#     prod_cat_safe = re.sub(r"[^A-Z0-9]", "", art["product_category"].upper())[:8]
#     cr_sap = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
#     cr_sap.set("ClassificationID", f"CLH_{prod_cat_safe}")
#     cr_sap.set("Type", "CPL_SAPHierarchy")

    cr_unconf = ET.SubElement(g_el, f"{{{STIBO_NS}}}ClassificationReference")
    cr_unconf.set("ClassificationID", f"{season_id}UA")
    cr_unconf.set("Type", "CPL_UnConfirmedForSeason")

    vals = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")

    _val(vals, "AT_Brand", BRAND_NAME, id_val=BRAND_CODE)
    _val(vals, "AT_InboundGenericCode", key_article)
    _val(vals, "AT_PrincipalStyleCode",        art["article_no"])
    _val(vals, "AT_SAPStyleCode",              art["article_no"])
    _val(vals, "AT_PrincipalStyleDescription", art["style_name"])

    color_id   = _s(art["color_id"])
    color_desc = _s(art["color_desc"])
    _val(vals, "AT_Color")
    if color_id:
        _val(vals, "AT_PrincipalColorCode", color_id)
        _val(vals, "AT_PrincipalColorName", color_desc)

    coo_code  = _s(art["coo"])
    coo_label = LOV_COUNTRY.get(coo_code, coo_code)
    if country_code:
        country_label = LOV_COUNTRY.get(country_code, country_code)
        _val(vals, "AT_Country", id_val=country_code)

    sil_mapped  = _map_silhouette(art["silhouette"])
    heel_mapped = _map_heel_height(art["heel_height"])
    _val(vals, "AT_Silhouette",          sil_mapped)
    _val(vals, "AT_Collection1",         art["article_collection"])
    _val(vals, "AT_HeelHeight",          heel_mapped)
    _val(vals, "AT_MerchandiseCategory", art["merch_category"][:9])
    _val(vals, "AT_PrincipalMerchandiseHierarchyL2", art["product_category"])
    _val(vals, "AT_PrincipalMerchandiseHierarchyL3", art["merch_category"])
    _val(vals, "AT_PrincipalMerchandiseHierarchyL4", art["silhouette"])
    _val(vals, "AT_OriginalPrice",       art["sugg_retail"])
    _val(vals, "AT_CurrentPrice",        art["sugg_retail"])

    curr = _normalize_currency(art.get("currency") or "")
    if curr:
        _val(vals, "AT_RetailPriceCurrency",id_val=curr)
    sea_label = LOV_SEASON.get(season_code, season_code)  # unknown codes fall back to raw code string
    _val(vals, "AT_Season", value=sea_label, id_val=season_code)
    full_year = f"20{season_year}" if len(season_year) == 2 else season_year
    _val(vals, "AT_SeasonYear", full_year)
    if article_type:
        _val(vals, "AT_BYArticleType", article_type, id_val=article_type)
    _val(vals, "AT_SAPProductFlag", id_val="A")
        

    stibo_comp, comp_label = _resolve_comp_code(comp_code)
    _multival(vals, "AT_CompanyCode", stibo_comp, comp_label)
    sbu_id, sbu_label = _resolve_sbu(sbu)
    _multival(vals, "AT_SBU", sbu_id, sbu_label)
    _val(vals, "AT_BYIndicator",        "Yes",     id_val="Y")
    _val(vals, "AT_SAPIndicator",       "No",      id_val="N")
    _val(vals, "AT_SAPArticleCategory", "Generic", id_val="1")
    _val(vals, "AT_UOM",                "Pair",    id_val="PAA")
    _val(vals, "AT_SportsCategoryEN",   "Other")
    # AT_CountrySize — V6 mapping: Footwear → US, Accessories → No Size
    prod_cat_upper = (art.get("product_category") or "").upper()
    if "FOOTWEAR" in prod_cat_upper or "SHOE" in prod_cat_upper or "BOOT" in prod_cat_upper or "SANDAL" in prod_cat_upper:
        _val(vals, "AT_CountrySize", "US")
    elif "ACCESSOR" in prod_cat_upper or "BAG" in prod_cat_upper or "WALLET" in prod_cat_upper or "BELT" in prod_cat_upper:
        _val(vals, "AT_CountrySize", "No Size")
    # Skip AT_CountrySize for unknown categories


    # TrenzaShop Variants — AT_Size text only (no id_val), AT_Variant only
    for size_rec in sizes:
        size_val  = str(size_rec["size"]).strip()
        size_code = _mdd_size_code(size_val)

        color_3 = re.sub(r"[^0-9]", "", color_id).zfill(3)[:3] if color_id else "000"
        variant_key = f"{BRAND_CODE}{article_no}{color_3}{size_code}"
        variant_key = variant_key[:18]

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")
        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID", "KEY_Variant")
        v_kv.text = variant_key
        ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = f"Size {size_val}"

        v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
        # FIX 26: AT_Size — text only, no id_val (same as TDD)
        _val(v_vals, "AT_Size", size_val)
        _val(v_vals, "AT_Variant", variant_key)

    return ET.tostring(g_el, encoding="unicode")



def _extract_country_code_from_filename(path: Path) -> str:
    """Extract 2-letter country code from ...-<COUNTRY>-<SEQ>.xlsx style filenames."""
    m = re.search(r"(?:^|[-_])([A-Za-z]{2})(?:[-_]\d+)?$", path.stem, flags=re.IGNORECASE)
    return m.group(1).upper() if m else ""


def _extract_file_season_from_filename(path: Path) -> tuple[str, str]:
    """
    Extract season code and year directly from filename — no whitelist.
    Matches any 2-letter code immediately followed by a 4-digit or 2-digit year.
    e.g. WN2027, SS26, FW2028, FL29 — all work without maintaining a list.
    """
    upper = path.stem.upper()
    # Match any 2-letter uppercase token directly followed by a year (no separator)
    m = re.search(r"(?<![A-Z])([A-Z]{2})(20\d{2}|\d{2})(?![0-9A-Z])", upper)
    if not m:
        return "", ""

    raw_code = m.group(1)
    year = m.group(2)
    if len(year) == 2:
        year = f"20{year}"

    return raw_code, year


def _extract_article_type_from_filename(path: Path) -> str:
    """
    Extract Inline/License article type from filename.
    Defaults to empty string when not present.
    """
    upper = path.stem.upper()
    if re.search(r"(?:^|[-_ ])LICEN[SC]E(?:D)?(?:$|[-_ ])", upper):
        return "License"
    if re.search(r"(?:^|[-_ ])INLINE(?:$|[-_ ])", upper):
        return "Inline"
    return ""


# MAIN ORCHESTRATOR
# ─────────────────────────────────────────────────────────────────
def run(args, auditor=None, article_filter=None, process_types=None):
    if process_types is None:
        process_types = {"linelist"}
    log.info("=== ALDO ETL START ===")
    log.info("brand=%s  comp_code=%s  sbu=%s  season=%s  seq=%s",
             args.brand, args.comp_code, args.sbu, args.season, args.seq)

    if article_filter:
        log.info("article_filter=%s  (single-article mode)", article_filter)

    export_time   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_written = 0

    season_code = args.season[:2].upper() if len(args.season) >= 2 else "SS"
    year_part   = args.season[2:]
    season_year = f"20{year_part}" if len(year_part) == 2 else year_part
    season_id   = f"CLH_{BRAND_CODE}_{season_code}{season_year}"

    log.info("Season: %s %s  →  %s", season_code, season_year, season_id)

    def _apply_filter(loader, source_label: str):
        if not article_filter:
            return
        before = len(loader.articles)
        loader.articles = [
            a for a in loader.articles if a["article_no"] == article_filter
        ]
        after = len(loader.articles)
        log.info("[%s] article_filter=%s → %d/%d articles kept",
                 source_label, article_filter, after, before)

    # ── 1. TrenzaShop ────────────────────────────────────────────
    if "linelist" in process_types:
        ll_files = list(LINELIST_DIR.glob("*.xlsm")) + list(LINELIST_DIR.glob("*.xlsx"))
        if ll_files:
            ll_path = ll_files[0]
            loader  = TrenzaShopLoader(ll_path)

            if loader.articles and (not args.season or args.season == "SS26"):
                first_season = loader.articles[0].get("selling_season", "")
                if first_season:
                    sc, sy, _ = _parse_season(first_season)
                    season_code = sc
                    season_year = sy
                    season_id   = f"CLH_{BRAND_CODE}_{season_code}{season_year}"

            _apply_filter(loader, "TrenzaShop")

            # loader.articles = loader.articles[:5]

            if loader.articles:
                file_sc = getattr(loader, "file_season_code", "") or season_code
                file_sy = getattr(loader, "file_season_year", "") or season_year
                file_sid = f"CLH_{BRAND_CODE}_{file_sc}{file_sy}"
                file_article_type = _extract_article_type_from_filename(ll_path) or getattr(args, "article_type_from_filename", "")

                cls_el   = build_classifications(file_sc, file_sy)
                out_name = f"{ll_path.stem}.xml"
                out_path = XML_OUT_DIR / out_name
                products = [
                    _build_trenzashop_product_xml(
                        art, loader.sizes.get(art["article_no"], []),
                        file_sc, file_sy, file_sid,
                        args.comp_code, args.sbu,
                        file_article_type,
                        getattr(loader, "file_country_code", "") or getattr(args, "country_code", ""),
                    )
                    for art in loader.articles
                ]
                _write_xml(out_path, cls_el, products, export_time)
                kb = out_path.stat().st_size // 1024
                log.info("✅ TrenzaShop XML → %s (%dKB)", out_name, kb)
                total_written += 1
                _print_summary("TrenzaShop", ll_path.name, len(loader.articles),
                               len(loader.articles), file_sc, file_sy,
                               args.comp_code, args.sbu, out_name, kb)
                if auditor:
                    auditor.record_loader("TrenzaShop", status="ok",
                                          articles=len(loader.articles), filename=ll_path.name)

    if total_written == 0:
        log.error("No input files found — no XML generated.")
        if auditor:
            auditor.set_etl_status("error", error="No input files found")
        return

    if auditor:
        auditor.set_etl_status("ok")

    log.info("=== ALDO ETL DONE — %d XML(s) generated ===", total_written)


def _print_summary(file_type, source_file, total_arts, written,
                   season_code, season_year, comp_code, sbu, out_name, kb):
    print(f"═══ ALDO {file_type.upper()} SUMMARY ══════════════════════════════════", flush=True)
    print(f"  Source file      : {source_file}",               flush=True)
    print(f"  Total articles   : {total_arts}",                flush=True)
    print(f"  Written to XML   : {written}",                   flush=True)
    print(f"  Season           : {season_code} {season_year}", flush=True)
    print(f"  comp_code        : {comp_code}",                 flush=True)
    print(f"  sbu              : {sbu}",                       flush=True)
    print(f"  XML output       : {out_name}",                  flush=True)
    print(f"  File size        : {kb} KB",                     flush=True)
    print("════════════════════════════════════════════════════", flush=True)

