"""
backlog_main.py

Backlog to STEP XML mapper.

Key behavior:
- Variant barcode values are always written on the variant product.
- DC_Barcode is written for new variant/EAN IDs and skipped for known existing
  IDs, so reruns do not recreate Stibo DataContainer objects.

Barcode DC modes:
- BACKLOG_BARCODE_DC_MODE=auto        default; write DC only when ID is not known locally
- BACKLOG_BARCODE_DC_MODE=values_only write only variant barcode attributes
- BACKLOG_BARCODE_DC_MODE=always      always write new DC IDs in this XML

Optional existing barcode DC list:
  input/barcode_dc_existing.txt

You can also drop Stibo import/error logs in input/ or output/logs/.
The loader extracts lines like:
    Data Container Object with ID 'ADIB75807520520_4062061443048' already exists.

One ID per line, for example:
    ADIB75807520520_4062061443048
"""

import io
import logging
import os
import re
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import boto3
import openpyxl
import pandas as pd

log = logging.getLogger("backlog_main")

STIBO_NS = "http://www.stibosystems.com/step"
STIBO_XSI = "http://www.w3.org/2001/XMLSchema-instance"
STIBO_SCHEMA = "http://www.stibosystems.com/step PIM.xsd"

ET.register_namespace("", STIBO_NS)
_XMLNS_RE = re.compile(r"\s+xmlns=\"[^\"]*\"")


# ==================================================================
# SECTION 1 -- DYNAMIC MDD LOADER
# ==================================================================

_SAP_SIZE_BASE = {
    # Garment / children sizes.
    "075": "07J",
    "085": "07P",
    "090": "09H",
    "100": "10H",
    "110": "11H",
    "120": "12H",
    "130": "13H",
    "140": "140",
    "150": "15H",
    "160": "16H",
    "170": "17H",
    "180": "18H",
    "185": "18H",
    "200": "200",
    # Footwear EU sizes.
    "370": "37",
    "380": "38",
    "390": "39",
    "400": "40",
    "410": "41",
    "420": "42",
    "430": "43",
    "440": "44",
    "450": "45",
    "460": "46",
    "470": "47",
    "480": "48",
    "490": "49",
    "500": "50",
    "520": "52",
    "530": "53",
    "540": "54",
    "550": "55",
    "560": "56",
    "570": "57",
    "580": "58",
    "590": "59",
    "600": "60",
    "610": "61",
    "620": "62",
    "630": "63",
    "640": "64",
    "650": "65",
    "660": "66",
    "670": "67",
    "680": "68",
    "690": "69",
    "700": "70",
    "710": "71",
    "720": "72",
    "730": "73",
    "740": "74",
    "750": "75",
    "760": "76",
    "770": "77",
    "780": "78",
    "790": "79",
    "800": "80",
    # Apparel text sizes.
    "010": "0XS",
    "020": "0XS",
    "030": "0ST",
    "040": "0SM",
    "050": "0LL",
    "060": "0XL",
    "070": "2XL",
    "072": "07J",
    "080": "3XL",
}


def _csv_set(env_name, default_values):
    raw = os.environ.get(env_name)
    if raw is None:
        return set(default_values)
    return {x.strip() for x in raw.split(",") if x.strip()}


STIBO_REJECTED_SIZE_IDS = _csv_set(
    "STIBO_REJECTED_SIZE_IDS",
    {
        # IDs confirmed rejected by LOV_SizeCode in STEP imports.
        "175", "195", "205", "230", "265", "270", "275",
        "280", "285", "290", "415",
        "53", "57", "59", "60", "61", "63", "64",
        "66", "67", "69", "71", "72", "73",
        "75", "76", "79",
    },
)

_BARCODE_DC_IDS_IN_XML = set()
_EXISTING_BARCODE_DC_IDS = set()
_GENERATED_BARCODE_DC_IDS = set()
_BARCODE_DC_ID_RE = re.compile(
    r"(?:Data Container Object with ID|DataContainer ID=)\s*['\"]([^'\"]+_[^'\"]+)['\"]"
)


def _size_lookup_key(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def _size_alias_keys(value):
    key = _size_lookup_key(value)
    if not key:
        return []
    keys = [key]
    compact = re.sub(r"[\s_]+", "", key)
    aliases = [compact]
    # Only add alphanumeric-only alias when key has no punctuation (e.g. '#', '-', '.').
    # Prevents '#01' → '01' and '3-' → '3' collisions.
    if re.fullmatch(r"[A-Z0-9\s_]+", key):
        aliases.append(re.sub(r"[^A-Z0-9]+", "", key))
    for alias in aliases:
        if alias and alias not in keys:
            keys.append(alias)
    return keys


def _is_size_header(value):
    return _size_lookup_key(value) in {
        "SIZE CODE", "SIZE CODE DESCRIPTION", "DESCRIPTION",
        "TECH SIZE", "TECH.SIZE", "TECHNICAL SIZE",
        "SAP SIZE", "SAP SIZE CODE", "LOV ID",
    }


class MDDLoader:
    DEFAULT_BUCKET = os.environ.get("RAW_BUCKET", "map-stibo-inbound-raw-dev")
    DEFAULT_KEY = "raw/metadata/Master Data Dictionary (MAA).xlsx"

    def __init__(self):
        self.lovs: dict[str, dict[str, str]] = {}
        self.LOV_BRAND: dict[str, str] = {}
        self.LOV_BRAND_GROUP: dict[str, str] = {}
        self.LOV_GENDER: dict[str, str] = {}
        self.LOV_AGE: dict[str, str] = {}
        self.LOV_BY_AGE_ID: dict[str, str] = {}
        self.LOV_SEASON: dict[str, str] = {}
        self.LOV_COUNTRY_ORIGIN: dict[str, str] = {}
        self.COUNTRY_NAME_TO_CODE: dict[str, str] = {}
        self.LOV_COMPANY_CODE: dict[str, str] = {}
        self.LOV_SBU: dict[str, str] = {}
        self.LOV_COLOR_CODE: dict[str, str] = {}
        self.COLOR_DESC_TO_CODE: dict[str, str] = {}
        self.LOV_SAP_ARTICLE_CATEGORY: dict[str, str] = {}
        self.LOV_BY_ARTICLE_TYPE: dict[str, str] = {}
        self.LOV_UOM: dict[str, str] = {}
        self.LOV_NATURE_OF_ARTICLE: dict[str, str] = {}
        self.LOV_COUNTRY: dict[str, str] = {}
        self.LOV_SIZE_CODE: set[str] = set()
        self.SIZE_DESC_TO_LOV: dict[str, str] = {}
        self.SAP_SIZE_TO_LOV: dict[str, str] = {}
        self._loaded = False

    def _register_size_alias(self, alias, lov_id):
        lov = str(lov_id or "").strip()
        # Skip supplier-internal source codes like #01, #02 — not valid Stibo LOV IDs.
        if not lov or _is_size_header(lov) or lov.startswith("#"):
            return
        for key in _size_alias_keys(alias):
            if not _is_size_header(key):
                self.SIZE_DESC_TO_LOV.setdefault(key, lov)

    def _load_gender_lov(self, wb):
        """
        Load Gender LOV: Col A = display value (e.g. 'Male'),
                         Col C = LOV ID (e.g. 'M').
        Also registers common aliases (W, M, U, MEN, WOMEN, FEMALE, etc.).
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "GENDER" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Gender LOV sheet not found")
            return
        gender_lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or not row[0]:
                continue
            display_value = str(row[0]).strip()           # Col A = display value
            lov_id        = str(row[2]).strip() if len(row) > 2 and row[2] else ""  # Col C = LOV ID
            if display_value and lov_id:
                gender_lov[display_value.upper()] = lov_id
        # Register common aliases so short codes / alternate spellings resolve too
        for alias, canonical in [
            ("W",     "FEMALE"),
            ("F",     "FEMALE"),
            ("M",     "MALE"),
            ("MEN",   "MALE"),
            ("WOMEN", "FEMALE"),
            ("U",     "UNISEX"),
        ]:
            if canonical in gender_lov:
                gender_lov.setdefault(alias, gender_lov[canonical])
        self.lovs["GenderLOV"] = gender_lov
        log.info("[MDD] Gender LOV loaded: %d entries", len(gender_lov))

    def _load_age_lov(self, wb):
        """
        Load Age LOV: Col A = display value (e.g. 'Children'),
                      Col B = LOV ID (e.g. 'K').
        """
        sheet_name = next(
            (s for s in wb.sheetnames if "AGE" in s.upper() and "LOV" in s.upper()),
            None,
        )
        if not sheet_name:
            log.warning("[MDD] Age LOV sheet not found")
            return
        age_lov: dict[str, str] = {}
        for row in list(wb[sheet_name].iter_rows(values_only=True))[1:]:
            if not row or not row[0]:
                continue
            display_value = str(row[0]).strip()           # Col A = display value
            lov_id        = str(row[1]).strip() if len(row) > 1 and row[1] else ""  # Col B = LOV ID
            if display_value and lov_id:
                age_lov[display_value.upper()] = lov_id
        self.lovs["AgeLOV"] = age_lov
        log.info("[MDD] Age LOV loaded: %d entries", len(age_lov))

    @classmethod
    def from_s3(cls, bucket=None, key=None):
        bucket = bucket or cls.DEFAULT_BUCKET
        key = key or cls.DEFAULT_KEY
        loader = cls()
        try:
            log.info("[MDD] Downloading from s3://%s/%s", bucket, key)
            s3_client = boto3.client("s3")
            resp = s3_client.get_object(Bucket=bucket, Key=key)
            data = resp["Body"].read()
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            loader._parse(wb)
            wb.close()
            log.info("[MDD] Loaded from S3 successfully.")
        except Exception as e:
            log.warning("[MDD] S3 load failed (%s) -- using minimal defaults.", e)
            loader._apply_defaults()
        return loader

    @classmethod
    def from_local(cls, path):
        loader = cls()
        try:
            log.info("[MDD] Loading from local: %s", path)
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            loader._parse(wb)
            wb.close()
            log.info("[MDD] Loaded from local successfully.")
        except Exception as e:
            log.warning("[MDD] Local load failed (%s) -- using minimal defaults.", e)
            loader._apply_defaults()
        return loader

    def _parse(self, wb):
        sn = set(wb.sheetnames)

        if "Brand LOV" in sn:
            for r in self._rows(wb["Brand LOV"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_BRAND[code.upper()] = label.upper()

        if "Brand Group" in sn:
            for r in self._rows(wb["Brand Group"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_BRAND_GROUP[code.upper()] = label.upper()

        # Load Gender LOV dynamically (Col A → Col C mapping)
        self._load_gender_lov(wb)

        # Load Age LOV dynamically (Col A → Col B mapping)
        self._load_age_lov(wb)

        if "Age LOV" in sn:
            for r in self._rows(wb["Age LOV"], skip=1):
                label, code = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_AGE[code.upper()] = label
            self.LOV_BY_AGE_ID.update({
                "ADULT": "ADULT", "ADULTS": "ADULT", "AD": "ADULT",
                "JUNIOR": "KIDS", "JR": "KIDS", "YOUTH": "KIDS",
                "KIDS": "KIDS", "KD": "KIDS", "KID": "KIDS",
                "CHILDREN": "KIDS", "CHILD": "KIDS", "CH": "KIDS",
                "INFANT": "INFANT", "IN": "INFANT",
                "ALL AGES": "ALL AGES", "ALL": "ALL AGES", "AA": "ALL AGES",
                "GRADE SCHOOL": "GRADE SCHOOL", "PRESCHOOL": "PRESCHOOL",
            })

        if "Season LOV" in sn:
            for r in self._rows(wb["Season LOV"], skip=1):
                code, name = self._v(r, 0), self._v(r, 1)
                if code and name:
                    self.LOV_SEASON[code.strip().upper()] = name
            for k, v in {
                "SS26": "Spring-Summer", "SS27": "Spring-Summer",
                "SS25": "Spring-Summer", "FW26": "Fall-Winter",
                "AW": "Fall-Winter", "SM": "Summer",
            }.items():
                self.LOV_SEASON.setdefault(k, v)

        if "Country Origin LOV" in sn:
            for r in self._rows(wb["Country Origin LOV"], skip=1):
                code, name = self._v(r, 0), self._v(r, 1)
                if code and name:
                    c = code.upper()
                    self.LOV_COUNTRY_ORIGIN.setdefault(c, name)
                    self.COUNTRY_NAME_TO_CODE[name.upper()] = c

        if "Company Code LOV" in sn:
            for r in self._rows(wb["Company Code LOV"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_COMPANY_CODE[str(code).strip()] = label

        if "SBU LOV" in sn:
            for r in self._rows(wb["SBU LOV"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_SBU[code.upper()] = label

        if "Color Code LOV" in sn:
            for r in self._rows(wb["Color Code LOV"], skip=1):
                raw_code = r[0] if r else None
                desc = self._v(r, 1)
                if raw_code is None or not desc:
                    continue
                code = str(raw_code).strip()
                self.LOV_COLOR_CODE[code.upper()] = desc.upper()
                self.COLOR_DESC_TO_CODE[desc.lower()] = code

        if "Article Category LOV" in sn:
            for r in self._rows(wb["Article Category LOV"], skip=1):
                raw_code = r[0] if r else None
                label = self._v(r, 1)
                if raw_code is None or not label:
                    continue
                self.LOV_SAP_ARTICLE_CATEGORY[str(raw_code).strip()] = label.strip()

        if "BY Article Type" in sn:
            for r in self._rows(wb["BY Article Type"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_BY_ARTICLE_TYPE[code] = label
                    self.LOV_BY_ARTICLE_TYPE[code.capitalize()] = label
                    self.LOV_BY_ARTICLE_TYPE[code.upper()] = label

        if "UOM LOV" in sn:
            for r in self._rows(wb["UOM LOV"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_UOM[code.upper()] = label

        if "Nature of Article LOV" in sn:
            for r in self._rows(wb["Nature of Article LOV"], skip=1):
                code, label = self._v(r, 0), self._v(r, 1)
                if code and label:
                    self.LOV_NATURE_OF_ARTICLE[code.upper()] = label

        if "Country LOV" in sn:
            for r in self._rows(wb["Country LOV"], skip=1):
                code, name = self._v(r, 0), self._v(r, 1)
                if code and name:
                    self.LOV_COUNTRY[code.upper()] = name

        if "Size Code LOV" in sn:
            for r in self._rows(wb["Size Code LOV"], skip=1):
                for col_idx in [0, 2, 5]:
                    if col_idx < len(r) and r[col_idx] is not None:
                        code = str(r[col_idx]).strip()
                        if code and code not in ("None", "nan") and not _is_size_header(code) and not code.startswith("#"):
                            self.LOV_SIZE_CODE.add(code)
                            self._register_size_alias(code, code)
                            desc = self._v(r, col_idx + 1)
                            self._register_size_alias(desc, code)

                # Some MDD versions store the source display/raw size in column D
                # and the valid STEP LOV ID in column F.
                raw_alias = self._v(r, 3)
                lov_id = self._v(r, 5)
                if raw_alias and lov_id:
                    self._register_size_alias(raw_alias, lov_id)
            log.info("[MDD] Size Code LOV: %d valid codes loaded", len(self.LOV_SIZE_CODE))
        else:
            log.warning("[MDD] 'Size Code LOV' sheet not found -- size validation disabled")

        if "SAP Size Mapping" in sn:
            for r in self._rows(wb["SAP Size Mapping"], skip=1):
                sap_code = self._v(r, 0)
                lov_id = self._v(r, 1)
                if sap_code and lov_id:
                    self.SAP_SIZE_TO_LOV[sap_code.strip()] = lov_id.strip()
            self._validate_sap_size_mapping()
            log.info("[MDD] SAP Size Mapping: %d entries loaded", len(self.SAP_SIZE_TO_LOV))
        else:
            self._build_sap_size_mapping()

        self._loaded = True
        log.info(
            "[MDD] Parsed -- brands=%d gender=%d age=%d season=%d "
            "country_origin=%d company=%d sbu=%d color=%d size_codes=%d sap_size_map=%d",
            len(self.LOV_BRAND), len(self.LOV_GENDER), len(self.LOV_AGE),
            len(self.LOV_SEASON), len(self.LOV_COUNTRY_ORIGIN),
            len(self.LOV_COMPANY_CODE), len(self.LOV_SBU), len(self.LOV_COLOR_CODE),
            len(self.LOV_SIZE_CODE), len(self.SAP_SIZE_TO_LOV),
        )

    def _build_sap_size_mapping(self):
        if not self.LOV_SIZE_CODE:
            log.warning("[MDD] LOV_SIZE_CODE empty -- SAP size mapping unvalidated")
            self.SAP_SIZE_TO_LOV = dict(_SAP_SIZE_BASE)
            return

        numeric_lov = sorted(
            [c for c in self.LOV_SIZE_CODE if re.match(r"^\d+$", c) and self.is_valid_size_lov(c)],
            key=lambda x: int(x),
        )

        invalid_count = 0
        for sap_code, candidate in _SAP_SIZE_BASE.items():
            if self.is_valid_size_lov(candidate):
                self.SAP_SIZE_TO_LOV[sap_code] = candidate
            else:
                nearest = self._nearest_numeric_lov(candidate, numeric_lov)
                if nearest:
                    log.warning(
                        "[MDD] SAP size '%s' -> '%s' invalid -- using nearest '%s'",
                        sap_code, candidate, nearest,
                    )
                    self.SAP_SIZE_TO_LOV[sap_code] = nearest
                else:
                    log.warning(
                        "[MDD] SAP size '%s' -> '%s' invalid and no fallback -- mapping skipped",
                        sap_code, candidate,
                    )
                invalid_count += 1

        if invalid_count:
            log.warning("[MDD] %d SAP size mappings had invalid LOV targets", invalid_count)

    def _validate_sap_size_mapping(self):
        if not self.SAP_SIZE_TO_LOV:
            return

        numeric_lov = sorted(
            [c for c in self.LOV_SIZE_CODE if re.match(r"^\d+$", c) and self.is_valid_size_lov(c)],
            key=lambda x: int(x),
        )

        fixed = {}
        for sap_code, lov_id in self.SAP_SIZE_TO_LOV.items():
            if self.is_valid_size_lov(lov_id):
                fixed[sap_code] = lov_id
                continue
            nearest = self._nearest_numeric_lov(lov_id, numeric_lov)
            if nearest:
                log.warning(
                    "[MDD] SAP mapping '%s' -> '%s' invalid -- using nearest '%s'",
                    sap_code, lov_id, nearest,
                )
                fixed[sap_code] = nearest
            else:
                log.warning("[MDD] SAP mapping '%s' -> '%s' invalid -- skipped", sap_code, lov_id)
        self.SAP_SIZE_TO_LOV = fixed

    @staticmethod
    def _nearest_numeric_lov(candidate, numeric_lov_sorted):
        if not numeric_lov_sorted:
            return None
        try:
            target = int(candidate)
        except (ValueError, TypeError):
            return None
        return min(numeric_lov_sorted, key=lambda x: abs(int(x) - target))

    @staticmethod
    def _rows(ws, skip=0):
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i < skip:
                continue
            if any(c is not None for c in row):
                yield row

    @staticmethod
    def _v(row, idx):
        if row is None or idx >= len(row) or row[idx] is None:
            return ""
        v = row[idx]
        try:
            if pd.isnull(v):
                return ""
        except (TypeError, ValueError):
            pass
        return str(v).strip()

    def _apply_defaults(self):
        self.LOV_BRAND = {
            "ADI": "ADIDAS", "NIK": "NIKE", "NEW": "NEW BALANCE",
            "SMI": "SMIGGLE", "ALD": "ALDO", "CRO": "CROCS",
            "LOT": "LOTTO", "BCK": "BIRKENSTOCK",
        }
        self.LOV_GENDER = {"M": "Male", "F": "Female", "W": "Female", "U": "Unisex"}
        self.LOV_AGE = {"AD": "Adults", "CH": "Kids", "AA": "All Ages"}
        self.LOV_BY_AGE_ID = {
            "ADULT": "ADULT", "ADULTS": "ADULT", "AD": "ADULT",
            "INFANT": "INFANT",
            "JUNIOR": "KIDS", "YOUTH": "KIDS", "KIDS": "KIDS", "KID": "KIDS",
            "CHILDREN": "KIDS", "CHILD": "KIDS", "CH": "KIDS",
            "ALL AGES": "ALL AGES", "ALL": "ALL AGES", "AA": "ALL AGES",
            "GRADE SCHOOL": "GRADE SCHOOL",
            "PRESCHOOL": "PRESCHOOL",
        }
        self.LOV_SEASON = {
            "SS": "Spring-Summer", "FW": "Fall-Winter", "AL": "All Season",
            "SM": "Summer", "SP": "Spring", "FL": "Fall", "WN": "Winter", "CO": "Core",
        }
        self.LOV_COUNTRY_ORIGIN = {
            "CN": "China", "VN": "Vietnam", "BD": "Bangladesh", "IN": "India",
            "ID": "Indonesia", "KH": "Cambodia", "MY": "Malaysia",
            "TH": "Thailand", "PK": "Pakistan",
        }
        self.COUNTRY_NAME_TO_CODE = {v.upper(): k for k, v in self.LOV_COUNTRY_ORIGIN.items()}
        self.LOV_COMPANY_CODE = {
            "0888": "PT. Map Aktif Adiperkasa",
            "0886": "PT. MAP FTL Adiperkasa",
            "0882": "Magna Management Asia",
        }
        self.LOV_SBU = {"SP": "MAA Sport", "GO": "MAA Golf", "CH": "MAA Children"}
        self.LOV_SAP_ARTICLE_CATEGORY = {"0": "Single", "1": "Generic", "10": "Sell set (Hampers)"}
        self.LOV_BY_ARTICLE_TYPE = {"Inline": "Inline", "License": "License", "SSE": "SSE"}
        self.LOV_UOM = {"EA": "Each", "PR": "Pair", "SET": "Set"}
        self.LOV_NATURE_OF_ARTICLE = {"WH1": "Wholesale", "REG": "REGULAR"}
        self.LOV_COUNTRY = {
            "ID": "Indonesia", "MY": "Malaysia", "SG": "Singapore",
            "TH": "Thailand", "VN": "Vietnam", "PH": "Philippines",
        }
        self.LOV_SIZE_CODE = set()
        self.SIZE_DESC_TO_LOV = {}
        self.SAP_SIZE_TO_LOV = dict(_SAP_SIZE_BASE)
        self._loaded = True
        log.warning(
            "[MDD] Using defaults -- SAP size mapping loaded from built-in base (%d entries).",
            len(self.SAP_SIZE_TO_LOV),
        )

    def brand(self, code):
        return self.LOV_BRAND.get((code or "").upper(), code or "")

    def brand_group(self, name):
        return self.LOV_BRAND_GROUP.get((name or "").upper(), name or "")

    def gender_label(self, code):
        return self.LOV_GENDER.get((code or "").upper(), code or "")

    def gender_sap(self, raw):
        _map = {"M": "M", "MALE": "M", "F": "F", "FEMALE": "F", "W": "F", "U": "U", "UNISEX": "U"}
        return _map.get((raw or "").strip().upper(), raw or "U")

    def age_label(self, code):
        return self.LOV_AGE.get((code or "").upper(), code or "")

    def age_sap(self, raw):
        _map = {
            "ADULT": "AD", "ADULTS": "AD", "AD": "AD",
            "JUNIOR": "CH", "JR": "CH", "YOUTH": "CH",
            "KIDS": "CH", "KD": "CH", "KID": "CH",
            "CHILDREN": "CH", "CHILD": "CH", "CH": "CH", "INFANT": "CH",
            "ALL AGES": "AA", "ALL": "AA", "AA": "AA",
        }
        return _map.get((raw or "").strip().upper(), "AD")

    def by_age_id(self, raw):
        normalized = (raw or "").strip().upper()
        return self.LOV_BY_AGE_ID.get(normalized, normalized)

    def season_label(self, code):
        clean = (code or "").strip().upper()
        return self.LOV_SEASON.get(clean, self.LOV_SEASON.get(clean[:2], code or ""))

    def country_origin_name(self, code):
        return self.LOV_COUNTRY_ORIGIN.get((code or "").upper(), code or "")

    def country_origin_code(self, raw):
        raw = (raw or "").strip()
        upper = raw.upper()
        if upper in self.LOV_COUNTRY_ORIGIN:
            return upper, self.LOV_COUNTRY_ORIGIN[upper]
        code = self.COUNTRY_NAME_TO_CODE.get(upper, raw[:2].upper())
        label = self.LOV_COUNTRY_ORIGIN.get(code, raw)
        return code, label

    def company_code(self, code):
        return self.LOV_COMPANY_CODE.get(str(code or "").strip(), code or "")

    def sbu(self, code):
        return self.LOV_SBU.get((code or "").upper(), code or "")

    def color_code_from_desc(self, desc):
        if not desc:
            return ""
        parts = [p.strip().lower() for p in re.split(r"[/,]", desc)]
        for part in parts:
            if part in self.COLOR_DESC_TO_CODE:
                return self.COLOR_DESC_TO_CODE[part]
            first = part.split()[0] if part.split() else ""
            if first in self.COLOR_DESC_TO_CODE:
                return self.COLOR_DESC_TO_CODE[first]
            words = part.split()
            if len(words) >= 2:
                two = f"{words[0]} {words[1]}"
                if two in self.COLOR_DESC_TO_CODE:
                    return self.COLOR_DESC_TO_CODE[two]
        return ""

    def color_label(self, code):
        return self.LOV_COLOR_CODE.get((code or "").upper(), code or "")

    def article_category(self, code):
        return self.LOV_SAP_ARTICLE_CATEGORY.get(str(code or "").strip(), code or "")

    def by_article_type(self, code):
        return self.LOV_BY_ARTICLE_TYPE.get(code or "", code or "")

    def uom(self, code):
        return self.LOV_UOM.get((code or "").upper(), code or "")

    def nature_of_article(self, code):
        return self.LOV_NATURE_OF_ARTICLE.get((code or "").upper(), code or "")

    def country(self, code):
        return self.LOV_COUNTRY.get((code or "").upper(), code or "")

    def is_valid_size_lov(self, size_id: str) -> bool:
        if not size_id:
            return False
        size_id = str(size_id).strip()
        if size_id in STIBO_REJECTED_SIZE_IDS:
            return False
        if not self.LOV_SIZE_CODE:
            return True
        return size_id in self.LOV_SIZE_CODE

    def size_code_from_description(self, raw: str) -> str:
        for key in _size_alias_keys(raw):
            lov = self.SIZE_DESC_TO_LOV.get(key)
            if lov:
                return lov
        return ""

    def resolve_sap_size(self, sap_code: str) -> str:
        return self.SAP_SIZE_TO_LOV.get((sap_code or "").strip(), "")


_MDD_CACHE = None


def get_mdd(mdd_path=None):
    global _MDD_CACHE
    if _MDD_CACHE is not None:
        return _MDD_CACHE
    if mdd_path and Path(mdd_path).exists():
        _MDD_CACHE = MDDLoader.from_local(mdd_path)
    else:
        _MDD_CACHE = MDDLoader.from_s3()
    return _MDD_CACHE


# ==================================================================
# SECTION 2 -- CONSTANTS & HELPERS
# ==================================================================

DIVISION_PARENT_MAP = {
    "ACCESSORIES": "E", "FOOTWEAR": "F", "APPAREL": "A",
    "EQUIPMENT": "Q", "HARDWARE": "Q", "TOYS": "T",
}

_COLOR_3DIGIT = {
    "005": "BLK", "W": "WHT", "NAV": "NVY", "GRE": "GRY",
    "12W": "BLU", "PK": "PNK", "Y": "YLW", "O": "ORG",
    "700": "BRN", "18": "BGE", "P": "PRP", "KGO": "GLD",
    "SV": "SLV", "MI": "MUL", "G": "GRN", "R": "RED",
    "LPC": "CAR", "OLI": "OLV", "CM": "CRM", "SAN": "SND", "000": "NOC",
}

COL_ALIASES = {
    "article_no": [
        "material", "so article", "article no", "article number", "article",
        "style no", "style number", "material no", "material - jun",
    ],
    "material_desc": [
        "material description", "iam material description",
        "model name", "description", "article description",
    ],
    "tech_size": ["technical size", "tech.size", "tech size"],
    "ean": ["ean/upc", "ean", "upc", "barcode", "ean code"],
    "season": ["season - jun ( based on po )", "season", "seasonal indicator"],
    "backlog_qty": ["purchase order qt.", "backlog qty", "qty", "quantity"],
    "ns_value": ["ns value\n( new disc. ) - jun", "ns value ( new disc. ) - jun", "ns value"],
    "rrp": ["rrp (lcy)", "rrp", "retail price", "price"],
    "currency": ["doc. currency", "document currency", "currency"],
    "channel": ["channel - jun", "channel"],
    "frs_whs": ["frs / whs", "frs/whs", "banner"],
    "comm_support": ["comm\nsupport", "comm support", "comm ( new disc. ) - jun"],
    "sold_to": ["customer name", "sold-to name", "sold to name"],
    "rdd": ["requested delivery date", "rdd"],
    "order_no": ["ib purchase order no", "order no", "customer po#"],
    "division": ["product division", "division", "prod.div."],
    "age_group": ["age group", "agr group"],
    "gender": ["gender"],
    "biz_segment": ["business segment", "segment"],
    "product_group": ["product group"],
    "product_type": ["product type"],
    "key_cat": ["key category cluster", "key cat cluster"],
    "collection": ["collection"],
    "intro_date": ["rid from mdc s26", "retail intro date", "intro date"],
    "carry_over": ["carry over flag", "carry over"],
    "color_desc": [
        "color/ description 1", "colour/ description 1", "color description",
        "colour description", "color desc", "colour", "color",
        "colorway name", "colour name",
    ],
}

BACKLOG_SHEET_CANDIDATES = ["Backlog", "BACKLOG", "backlog", "Sheet1", "Data"]
_EXCEL_EPOCH = datetime(1899, 12, 30)
_HEADER_KEYWORDS = {
    "material", "material description", "technical size",
    "ean", "size", "gender", "season", "division", "product division",
    "article", "style", "qty", "quantity", "business segment",
    "requested delivery date", "channel", "frs", "rrp", "ns value",
    "so article", "seasonal indicator", "material - jun",
    "product group", "product type", "doc. currency",
}
_JUNK_PATTERNS = {"add by jun", "add by"}

_SEA_CANON = {
    "FW": "FW", "AW": "FW", "FL": "FW", "WN": "FW",
    "SS": "SS", "SP": "SS",
    "SM": "SM",
    "CO": "CO",
    "AL": "AL",
}


def _barcode_dc_mode():
    mode = os.environ.get("BACKLOG_BARCODE_DC_MODE")
    legacy = os.environ.get("BACKLOG_INCLUDE_BARCODE_DC")

    if mode:
        return mode.strip().lower()
    if legacy is not None and legacy.strip().lower() in ("false", "0", "no", "n", "off"):
        return "values_only"
    return "auto"


def _excel_date_to_int(v):
    if isinstance(v, datetime):
        return max(0, (v - _EXCEL_EPOCH).days)
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return 0


def _s(v):
    if v is None:
        return ""
    try:
        if pd.isnull(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    s = str(v).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _get_division_code(division):
    div_upper = (division or "").strip().upper()
    for key, code in DIVISION_PARENT_MAP.items():
        if key in div_upper:
            return code
    return "F"


def _parse_season(season_raw):
    s = (season_raw or "").strip()
    s = re.sub(r"^\d+\.\s*", "", s).strip().upper()
    if not s or s in ("BULK / PIPELINE", "BULK/PIPELINE", "TDD ( ROS )", "UNIFORM"):
        return "SS", "26", "2026"
    m0 = re.match(r"^(\d{2})([SF])$", s)
    if m0:
        yr_short = m0.group(1)
        sea_prefix = "SS" if m0.group(2) == "S" else "FW"
        return sea_prefix, yr_short, f"20{yr_short}"
    m = re.match(r"^([SFA])(W?)(\d{2,4})$", s)
    if m:
        digits = m.group(3)
        sea_prefix = "FW" if m.group(1) == "F" else "SS"
        full_year, short_year = (digits, digits[2:]) if len(digits) == 4 else (f"20{digits}", digits)
        return sea_prefix, short_year, full_year
    m2 = re.match(r"^(SS|FW|AW|AL|SP|SM|FL|WN|CO)(\d{2,4})$", s)
    if m2:
        prefix_raw = m2.group(1)
        sea_prefix = _SEA_CANON.get(prefix_raw, "SS")
        digits = m2.group(2)
        full_year, short_year = (digits, digits[2:]) if len(digits) == 4 else (f"20{digits}", digits)
        return sea_prefix, short_year, full_year
    digits = re.sub(r"[^0-9]", "", s)
    sea_prefix = "FW" if "F" in s else "SS"
    if len(digits) == 4:
        return sea_prefix, digits[2:], digits
    if len(digits) == 2:
        return sea_prefix, digits, f"20{digits}"
    return "SS", "26", "2026"


def _parse_season_year_from_filename(filename):
    m = re.search(
        r"[-_](SS|FW|SM|SP|FL|WN|AL|CO)(\d{2,4})(?:[-_]|$)",
        filename,
        re.IGNORECASE,
    )
    if m:
        digits = m.group(2)
        return digits if len(digits) == 4 else f"20{digits}"
    return None


def _parse_season_prefix_from_filename(filename):
    m = re.search(
        r"[-_](SS|FW|SM|SP|FL|WN|AL|CO)\d{2,4}(?:[-_]|$)",
        filename,
        re.IGNORECASE,
    )
    if not m:
        return None
    return _SEA_CANON.get(m.group(1).upper(), m.group(1).upper())


def _resolve_col(col_map, field):
    for alias in COL_ALIASES.get(field, []):
        normalised = re.sub(r"[\n\r\t]+", " ", alias).strip()
        idx = col_map.get(normalised)
        if idx is not None:
            return idx
        idx = col_map.get(alias)
        if idx is not None:
            return idx
    return None


def _get_ean_category(ean, brand_code="ADI"):
    ean_clean = re.sub(r"[^0-9]", "", ean) if ean else ""
    if brand_code.upper() == "ADI":
        return "ZV" if ean_clean.startswith("70") else "ZQ"
    return ""


def _color_lov_to_3digit(color_code):
    if not color_code:
        return ""
    mapped = _COLOR_3DIGIT.get(color_code.upper(), "")
    # Skip placeholder-like values so variant keys don't carry synthetic color codes.
    if mapped in {"NOC", "000"}:
        return ""
    return mapped


def _size_to_3digit(tech_size, display_size=""):
    raw = (display_size or tech_size or "").strip()
    if not raw or raw == "*":
        return "000"
    half_lov = _half_size_lov_candidates(raw)
    if half_lov:
        clean_half = re.sub(r"[^A-Z0-9]", "", half_lov[0].upper())
        if clean_half.endswith("H") and clean_half[:-1].isdigit():
            return f"{int(clean_half[:-1]):02d}H"
        return clean_half[-3:].ljust(3, "0") if clean_half else "000"
    clean = re.sub(r"[^A-Z0-9]", "", raw.upper())
    return clean[-3:].ljust(3, "0") if clean else "000"


def _size_for_variant_key(tech_size, display_size=""):
    raw = (display_size or tech_size or "").strip()
    if not raw or raw == "*":
        return ""

    half_lov = _half_size_lov_candidates(raw)
    if half_lov:
        return half_lov[0]

    if re.match(r"^0*\d+(?:\.0+)?$", raw):
        return str(int(float(raw)))

    return re.sub(r"\s+", "", raw)


def _half_size_lov_candidates(raw_size):
    raw = (raw_size or "").strip().upper()
    if not raw:
        return []

    hyphen_match = re.match(r"^0*(\d+)\s*-$", raw)
    decimal_half_match = re.match(r"^0*(\d+)(?:[.,]5)$", raw)
    explicit_h_match = re.match(r"^0*(\d+)H$", raw)

    match = hyphen_match or decimal_half_match or explicit_h_match
    if not match:
        return []

    return [f"{int(match.group(1))}H"]


def _sap_size_keys(raw_size):
    raw = (raw_size or "").strip().upper()
    if not raw:
        return []
    clean = re.sub(r"[^A-Z0-9]", "", raw.split(".")[0])
    if not clean:
        return []
    keys = [clean]
    if clean.isdigit() and len(clean) < 3:
        padded = clean.zfill(3)
        if padded not in keys:
            keys.insert(0, padded)
    return keys


def _resolve_size_lov_id(tech_sz, display_sz, mdd=None):
    text_to_lov = {
        "XS": "0XS", "S": "0ST", "M": "0SM", "L": "0LL", "XL": "0XL",
        "2XL": "2XL", "XXL": "2XL", "3XL": "3XL", "XXXL": "3XL",
        "4XL": "4XL", "5XL": "5XL", "2XS": "2XS", "XXS": "2XS", "OS": "0OS",
    }

    def _mdd_size_lov(raw):
        if not mdd or not raw:
            return ""

        lov_codes = getattr(mdd, "LOV_SIZE_CODE", set())

        # Half-size patterns ("3-", "7.5", "7H") take priority over description lookup.
        # Prevents "3-" from mapping to a MDD LOV "3-" entry instead of the correct "3H".
        half = _half_size_lov_candidates(raw)
        if half:
            for h in half:
                if h in lov_codes:
                    return h
            return ""

        desc_lov = mdd.size_code_from_description(raw)
        if desc_lov:
            return desc_lov

        exact_key = _size_lookup_key(raw)
        if exact_key in lov_codes:
            return exact_key

        for candidate in _size_alias_keys(raw):
            if candidate in lov_codes:
                return candidate

        mapped = text_to_lov.get(_size_lookup_key(raw), "")
        if mapped and mapped in lov_codes:
            return mapped

        return ""

    display_raw = (display_sz or "").strip()
    tech_raw = (tech_sz or "").strip()

    if display_raw:
        mapped = _mdd_size_lov(display_raw)
        if mapped:
            return mapped, True

    for sap_key in _sap_size_keys(tech_raw):
        if mdd:
            sap_lov = mdd.resolve_sap_size(sap_key)
            if sap_lov:
                return sap_lov, True
        elif sap_key in _SAP_SIZE_BASE:
            return _SAP_SIZE_BASE[sap_key], True

    if tech_raw and not display_raw:
        mapped = _mdd_size_lov(tech_raw)
        if mapped:
            return mapped, True

    raw = display_raw or tech_raw
    if not raw:
        return "", False

    half_size_candidates = _half_size_lov_candidates(raw)
    if half_size_candidates:
        return half_size_candidates[0], True

    return raw, True


def _build_variant_key(b_code, article_no, color_code, tech_size, display_size=""):
    color_3 = _color_lov_to_3digit(color_code).strip().upper()
    size_3 = _size_to_3digit(tech_size, display_size)
    size_key = _size_for_variant_key(tech_size, display_size)
    raw = f"{b_code}{article_no}{color_3}{size_key or size_3}"
    return raw[:18]


def _collect_barcode_dc_ids_from_xml(path):
    ids = set()
    try:
        for _, elem in ET.iterparse(path, events=("end",)):
            if elem.tag.endswith("DataContainer"):
                dc_id = elem.get("ID")
                if dc_id and "_" in dc_id:
                    ids.add(dc_id)
            elem.clear()
    except ET.ParseError as e:
        log.warning("[BarcodeDC] Could not parse existing XML %s (%s)", path, e)
    except OSError as e:
        log.warning("[BarcodeDC] Could not read existing XML %s (%s)", path, e)
    return ids


def _looks_like_barcode_dc_id(text):
    text = (text or "").strip()
    return bool(re.match(r"^[A-Z0-9]{6,30}_[0-9]{8,20}$", text))


def _collect_barcode_dc_ids_from_text(path):
    ids = set()
    try:
        raw = path.read_text(encoding="utf-8", errors="ignore")
    except OSError as e:
        log.warning("[BarcodeDC] Could not read existing ID text %s (%s)", path, e)
        return ids

    for match in _BARCODE_DC_ID_RE.finditer(raw):
        dc_id = match.group(1).strip()
        if dc_id:
            ids.add(dc_id)

    for line in raw.splitlines():
        dc_id = line.strip()
        if _looks_like_barcode_dc_id(dc_id):
            ids.add(dc_id)

    return ids


def _load_existing_barcode_dc_ids(base):
    global _EXISTING_BARCODE_DC_IDS

    base = Path(base)
    ids = set()
    text_paths = [
        base / "input" / "barcode_dc_existing.txt",
        base / "output" / "logs" / "barcode_dc_written.txt",
    ]

    for path in text_paths:
        if not path.exists():
            continue
        ids.update(_collect_barcode_dc_ids_from_text(path))

    for scan_root in [base / "input", base / "output" / "logs"]:
        if not scan_root.exists():
            continue
        for pattern in ["*.txt", "*.log", "*.err"]:
            for path in scan_root.glob(pattern):
                ids.update(_collect_barcode_dc_ids_from_text(path))

    xml_dir = base / "output" / "xml"
    if xml_dir.exists():
        for xml_path in xml_dir.glob("*.xml"):
            ids.update(_collect_barcode_dc_ids_from_xml(xml_path))

    _EXISTING_BARCODE_DC_IDS = ids
    log.info("[BarcodeDC] Loaded %d existing barcode DC IDs", len(ids))


def _save_written_barcode_dc_ids(base):
    all_ids = set(_EXISTING_BARCODE_DC_IDS) | set(_GENERATED_BARCODE_DC_IDS)
    if not all_ids:
        return

    out_path = Path(base) / "output" / "logs" / "barcode_dc_written.txt"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(sorted(all_ids)) + "\n", encoding="utf-8")
    log.info("[BarcodeDC] Saved %d barcode DC IDs -> %s", len(all_ids), out_path)


def _reset_barcode_tracking(base):
    _BARCODE_DC_IDS_IN_XML.clear()
    _GENERATED_BARCODE_DC_IDS.clear()
    _load_existing_barcode_dc_ids(base)


# ==================================================================
# SECTION 3 -- LOADERS
# ==================================================================

class EANListLoader:
    def __init__(self, wb):
        self.by_article = {}
        self.article_meta = {}
        self._load(wb)

    def _load(self, wb):
        ean_sheet = None
        for sn in wb.sheetnames:
            if "EAN" in sn.upper():
                ean_sheet = wb[sn]
                log.info("[EAN LIST] Found sheet: '%s'", sn)
                break
        if not ean_sheet:
            log.warning("[EAN LIST] Sheet not found")
            return

        all_rows = list(ean_sheet.iter_rows(values_only=True))
        if not all_rows:
            return

        header_idx = 1 if len(all_rows) > 1 else 0
        for i, row in enumerate(all_rows[:10]):
            vals = [str(c).strip().lower() for c in row if c]
            if any("article" in v or "tech" in v for v in vals):
                header_idx = i
                break

        hdr = [
            re.sub(r"[\n\r\t]+", " ", str(c)).strip().lower() if c else ""
            for c in all_rows[header_idx]
        ]
        log.info("[EAN LIST] Headers: %s", hdr)

        def ci(name):
            for i, h in enumerate(hdr):
                if name.lower() in h:
                    return i
            return None

        def exact_ci(*names):
            wanted = {name.lower() for name in names}
            for i, h in enumerate(hdr):
                if h in wanted:
                    return i
            return None

        def size_ci():
            exact = exact_ci("size")
            if exact is not None:
                return exact
            for i, h in enumerate(hdr):
                if "size" in h and "tech" not in h and "technical" not in h:
                    return i
            return None

        art_col = ci("article") if ci("article") is not None else 0
        tech_col = ci("tech") if ci("tech") is not None else 1
        resolved_size_col = size_ci()
        size_col = resolved_size_col if resolved_size_col is not None else 2
        ean_col = ci("ean") if ci("ean") is not None else (ci("upc") if ci("upc") is not None else 3)
        rrp_col = next((i for i, h in enumerate(hdr) if "rrp" in h or "mdc" in h), 8)
        color_sap_col = next((i for i, h in enumerate(hdr) if "color - sap" in h or "colour - sap" in h), 9)
        color_col = next((i for i, h in enumerate(hdr) if h in ("color", "colour")), 11)
        mat_desc_col = ci("material description") if ci("material description") is not None else 10
        div_col = next((i for i, h in enumerate(hdr) if "product division" in h or "division" in h), 12)
        grp_col = next((i for i, h in enumerate(hdr) if "product group" in h), 13)
        type_col = next((i for i, h in enumerate(hdr) if "product type" in h), 14)
        keycat_col = next((i for i, h in enumerate(hdr) if "key" in h and "cat" in h), 15)
        age_col = next((i for i, h in enumerate(hdr) if "age" in h), 16)
        gender_col = next((i for i, h in enumerate(hdr) if "gender" in h), 17)
        seg_col = next((i for i, h in enumerate(hdr) if "business" in h and "segment" in h), 18)

        s26_col = next((i for i, h in enumerate(hdr) if h.strip() == "s26"), None)
        if s26_col is None:
            s26_col = 6
        log.info("[EAN LIST] S26 flag column index: %d", s26_col)

        def g(row, idx):
            if idx is None or idx >= len(row):
                return ""
            return _s(row[idx])

        count = 0
        skipped_season = 0
        for row in all_rows[header_idx + 1:]:
            if not row or all(c is None for c in row):
                continue
            s26_flag = str(row[s26_col]).strip().upper() if s26_col < len(row) and row[s26_col] else ""
            if s26_flag != "Y":
                skipped_season += 1
                continue

            article = g(row, art_col)
            if not article:
                continue

            size_rec = {
                "tech_size": g(row, tech_col),
                "size": g(row, size_col),
                "ean": g(row, ean_col),
                "rrp": g(row, rrp_col),
                "color_sap": g(row, color_sap_col),
                "color_desc": g(row, color_col),
                "division": g(row, div_col),
                "product_group": g(row, grp_col),
                "product_type": g(row, type_col),
                "key_cat": g(row, keycat_col),
                "age_group": g(row, age_col),
                "gender": g(row, gender_col),
                "biz_segment": g(row, seg_col),
            }
            self.by_article.setdefault(article, []).append(size_rec)
            if article not in self.article_meta:
                self.article_meta[article] = {
                    "material_desc": g(row, mat_desc_col),
                    "color_desc": g(row, color_sap_col) or g(row, color_col),
                    "division": g(row, div_col),
                    "product_group": g(row, grp_col),
                    "product_type": g(row, type_col),
                    "key_cat": g(row, keycat_col),
                    "age_group": g(row, age_col),
                    "gender": g(row, gender_col),
                    "biz_segment": g(row, seg_col),
                    "rrp": g(row, rrp_col),
                }
            count += 1

        log.info(
            "[EAN LIST] Loaded %d S26 size records for %d articles | Skipped %d non-S26 rows",
            count, len(self.by_article), skipped_season,
        )

    def get_sizes(self, article_no):
        return self.by_article.get(article_no, [])

    def get_meta(self, article_no):
        return self.article_meta.get(article_no, {})

    @property
    def all_articles(self):
        return set(self.by_article.keys())


class BacklogSheetLoader:
    def __init__(self, wb):
        self.article_meta = {}
        self._load(wb)

    def _find_sheet(self, wb):
        for name in BACKLOG_SHEET_CANDIDATES:
            if name in wb.sheetnames:
                return name
        for name in wb.sheetnames:
            if any(kw in name.lower() for kw in ["backlog", "data", "order"]):
                return name
        return wb.sheetnames[0]

    def _find_header_row(self, rows):
        best_row, best_score = 0, 0
        for i, row in enumerate(rows[:20]):
            if not row:
                continue
            vals = [str(c).strip().lower() for c in row if c is not None]
            junk = sum(1 for v in vals if any(j in v for j in _JUNK_PATTERNS))
            if vals and junk / len(vals) > 0.4:
                continue
            score = sum(1 for v in vals if any(kw in v for kw in _HEADER_KEYWORDS))
            if score > best_score:
                best_score = score
                best_row = i
                if score >= 5:
                    break
        return best_row

    def _load(self, wb):
        sheet_name = self._find_sheet(wb)
        log.info("[Backlog] Using sheet: '%s'", sheet_name)
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return

        header_idx = self._find_header_row(rows)
        hdr_row = rows[header_idx]
        col_map = {}
        for i, h in enumerate(hdr_row):
            if h is not None:
                key = re.sub(r"[\n\r\t]+", " ", str(h)).strip().lower()
                col_map[key] = i

        def ci(field):
            return _resolve_col(col_map, field)

        def g(row, fld):
            idx = ci(fld)
            if idx is None or idx >= len(row):
                return ""
            return _s(row[idx])

        for row in rows[header_idx + 1:]:
            if not row or all(c is None for c in row):
                continue
            article_no = g(row, "article_no")
            if not article_no or article_no in self.article_meta:
                continue

            qty_idx = ci("backlog_qty")
            qty_raw = row[qty_idx] if (qty_idx is not None and qty_idx < len(row)) else None
            if isinstance(qty_raw, datetime):
                backlog_qty = _excel_date_to_int(qty_raw)
            elif qty_raw is not None:
                try:
                    backlog_qty = int(float(str(qty_raw).strip()))
                except (ValueError, TypeError):
                    backlog_qty = 0
            else:
                backlog_qty = 0

            self.article_meta[article_no] = {
                "article_no": article_no,
                "material_desc": g(row, "material_desc"),
                "color_description": g(row, "color_desc"),
                "product_division": g(row, "division"),
                "product_group": g(row, "product_group"),
                "product_type": g(row, "product_type"),
                "business_segment": g(row, "biz_segment"),
                "key_cat_cluster": g(row, "key_cat"),
                "collection": g(row, "collection"),
                "age_group": g(row, "age_group"),
                "gender": g(row, "gender"),
                "rrp": g(row, "rrp"),
                "currency": g(row, "currency"),
                "channel": g(row, "channel"),
                "frs_whs": g(row, "frs_whs"),
                "season": g(row, "season"),
                "retail_intro_date": g(row, "intro_date"),
                "comm_support": g(row, "comm_support"),
                "carry_over_flag": g(row, "carry_over"),
                "sold_to_name": g(row, "sold_to"),
                "backlog_qty": backlog_qty,
                "ns_value": g(row, "ns_value"),
                "order_no": g(row, "order_no"),
            }

        log.info("[Backlog] Loaded metadata for %d unique articles", len(self.article_meta))

    def get_meta(self, article_no):
        return self.article_meta.get(article_no, {})

    @property
    def all_articles(self):
        return set(self.article_meta.keys())


# ==================================================================
# SECTION 4 -- MAPPER
# ==================================================================

def _map_article(bl_meta, ean_meta, sizes, brand_code, mdd):
    def pick(bl_key, ean_key=None):
        bl_val = str(bl_meta.get(bl_key, "") or "").strip()
        ean_val = str(ean_meta.get(ean_key or bl_key, "") or "").strip()
        return ean_val if (ean_val and not bl_val) else (bl_val or ean_val)

    gender_raw = pick("gender")
    age_raw = pick("age_group")
    colour_raw = pick("color_description", "color_desc") or ean_meta.get("color_sap", "")
    sap_gender = mdd.gender_sap(gender_raw)
    sap_age = mdd.age_sap(age_raw)
    rrp = pick("rrp")
    if not rrp and ean_meta.get("rrp"):
        rrp = ean_meta["rrp"]
    if not colour_raw:
        colour_raw = ean_meta.get("color_desc", "") or ean_meta.get("color_sap", "")

    return {
        "article_no": bl_meta["article_no"],
        "model_name": pick("material_desc") or ean_meta.get("material_desc", ""),
        "model_no": bl_meta["article_no"],
        "brand_code": brand_code,
        "gender_code": sap_gender,
        "gender_raw": gender_raw,
        "age_code": sap_age,
        "age_raw": age_raw,
        "colour": colour_raw,
        "division": pick("product_division", "division"),
        "biz_segment": pick("business_segment", "biz_segment"),
        "key_cat_cluster": pick("key_cat_cluster", "key_cat"),
        "product_type": pick("product_type"),
        "product_group": pick("product_group"),
        "collection": bl_meta.get("collection", ""),
        "article_type": "Generic",
        "art_category": "1",
        "season": bl_meta.get("season", ""),
        "intro_date": bl_meta.get("retail_intro_date", ""),
        "rrp": rrp,
        "currency": bl_meta.get("currency", "IDR") or "IDR",
        "channel": bl_meta.get("channel", ""),
        "frs_whs": bl_meta.get("frs_whs", ""),
        "comms": bl_meta.get("comm_support", ""),
        "carry_over": bl_meta.get("carry_over_flag", ""),
        "backlog_qty": bl_meta.get("backlog_qty", 0),
        "ns_value": bl_meta.get("ns_value", ""),
        "_sizes": sizes,
    }


# ==================================================================
# SECTION 5 -- XML VALUE HELPERS
# ==================================================================

def _val(parent, attr_id, value="", id_val=""):
    clean_id = str(id_val).strip() if id_val else ""
    clean_val = str(value).strip() if value else ""
    if clean_id in ("None", "nan", "0", ""):
        clean_id = ""
    if clean_val in ("None", "nan", ""):
        clean_val = ""
    if not clean_id and not clean_val:
        return None

    el = ET.SubElement(parent, f"{{{STIBO_NS}}}Value")
    el.set("AttributeID", attr_id)
    if clean_id:
        el.set("ID", clean_id)
    elif clean_val:
        el.text = clean_val
    return el


def _multival(parent, attr_id, id_val):
    if not id_val:
        return None
    mv = ET.SubElement(parent, f"{{{STIBO_NS}}}MultiValue")
    mv.set("AttributeID", attr_id)
    v = ET.SubElement(mv, f"{{{STIBO_NS}}}Value")
    v.set("ID", str(id_val))
    return mv


def _xml_child(parent, tag):
    return parent.find(f"{{{STIBO_NS}}}{tag}")


def _find_mdc(dc_root, dc_type):
    for mdc in dc_root.findall(f"{{{STIBO_NS}}}MultiDataContainer"):
        if mdc.get("Type") == dc_type:
            return mdc
    return None


def _find_dc(mdc, dc_id):
    for dc in mdc.findall(f"{{{STIBO_NS}}}DataContainer"):
        if dc.get("ID") == dc_id:
            return dc
    return None


def _set_dc_value(vals, attr_id, text=None, id_val=None):
    existing = None
    for v in vals.findall(f"{{{STIBO_NS}}}Value"):
        if v.get("AttributeID") == attr_id:
            existing = v
            break

    v = existing or ET.SubElement(vals, f"{{{STIBO_NS}}}Value")
    v.set("AttributeID", attr_id)
    if id_val:
        v.set("ID", str(id_val))
        v.text = None
    else:
        v.attrib.pop("ID", None)
        v.text = str(text or "")
    return v


def _add_barcode_dc(parent_el, ean, variant_key, brand_code="ADI"):
    if not ean:
        return False

    ean_clean = re.sub(r"[^0-9]", "", str(ean).strip())
    if not ean_clean:
        return False

    mode = _barcode_dc_mode()
    if mode in ("values_only", "value_only", "false", "off", "none", "skip"):
        log.info("[BarcodeDC] Mode=%s; skipped DC_Barcode for %s_%s", mode, variant_key, ean_clean)
        return False

    dc_id = ean_clean

    if dc_id in _BARCODE_DC_IDS_IN_XML:
        log.warning("[BarcodeDC] Duplicate in same XML skipped: %s", dc_id)
        return False

    if mode not in ("always", "force", "dc") and dc_id in _EXISTING_BARCODE_DC_IDS:
        log.info("[BarcodeDC] Existing barcode DC skipped: %s", dc_id)
        return False

    _BARCODE_DC_IDS_IN_XML.add(dc_id)
    _GENERATED_BARCODE_DC_IDS.add(dc_id)

    dc_root = _xml_child(parent_el, "DataContainers")
    if dc_root is None:
        dc_root = ET.SubElement(parent_el, f"{{{STIBO_NS}}}DataContainers")

    mdc = _find_mdc(dc_root, "DC_Barcode")
    if mdc is None:
        mdc = ET.SubElement(dc_root, f"{{{STIBO_NS}}}MultiDataContainer")
        mdc.set("Type", "DC_Barcode")
    mdc.set("update", "true")

    dc = ET.SubElement(mdc, f"{{{STIBO_NS}}}DataContainer")
    # ID field removed — Stibo will auto-generate the DataContainer ID

    dc.set("update", "true")
    dc.set("Analyzer", "true")
    dc.attrib.pop("Type", None)

    vals = _xml_child(dc, "Values")
    if vals is None:
        vals = ET.SubElement(dc, f"{{{STIBO_NS}}}Values")

    _set_dc_value(vals, "AT_Barcode", text=ean_clean)
    _set_dc_value(vals, "AT_BarcodeType", id_val="P")
    _set_dc_value(vals, "AT_MainEANIndicator", id_val="Y")

    ean_cat = _get_ean_category(ean_clean, brand_code)
    if ean_cat:
        _set_dc_value(vals, "AT_EANCategory", id_val=ean_cat)

    return True

def _add_generic_values(vals_el, art, comp_code, sbu, mdd):
    # _multival(vals_el, "AT_SBU", sbu)
    # _multival(vals_el, "AT_CompanyCode", comp_code)

    b_code = art["brand_code"]
    b_label = mdd.brand(b_code)
    _val(vals_el, "AT_Brand", b_label, id_val=b_code)
    _val(vals_el, "AT_BrandGroup", b_label, id_val=b_label)
    _val(vals_el, "AT_PrincipalStyleCode", art["article_no"])
    _val(vals_el, "AT_SAPStyleCode", art["article_no"])
    _val(vals_el, "AT_PrincipalStyleDescription", art["model_name"])
    # _val(vals_el, "AT_PrincipalColorCode", art["colour"])

    colour_code = mdd.color_code_from_desc(art.get("colour", ""))
    if colour_code:
        _val(vals_el, "AT_Color", id_val=colour_code)

    # ── Gender ── (dynamic lookup from MDD Gender LOV: Col A → Col C) ──────────
    gender_lov    = (mdd.lovs.get("GenderLOV") or {}) if mdd else {}
    gender_raw_up = (art.get("gender_raw") or "").strip().upper()
    g_lov_id      = (
        gender_lov.get(gender_raw_up)
        or gender_lov.get((art.get("gender_code") or "").upper())
        or art["gender_code"]
    )
    _val(vals_el, "AT_Gender",   id_val=g_lov_id)
    _val(vals_el, "AT_BYGender", id_val=g_lov_id)

    # ── Age ── (dynamic lookup from MDD Age LOV: Col A → Col B) ──────────────
    age_lov       = (mdd.lovs.get("AgeLOV") or {}) if mdd else {}
    age_raw_up    = (art.get("age_raw") or "").strip().upper()
    sap_age_id    = (
        age_lov.get(age_raw_up)
        or age_lov.get((art.get("age_code") or "").upper())
        or art["age_code"]
    )
    _val(vals_el, "AT_SAPAge", id_val=sap_age_id)

    age_raw_upper = (art.get("age_raw", "") or "").strip().upper()
    by_age_id = mdd.by_age_id(age_raw_upper or sap_age_id)
    _val(vals_el, "AT_BYAge", by_age_id, id_val=by_age_id)

    sea_prefix, _, full_year = _parse_season(art.get("season", ""))
    filename_prefix = art.get("_filename_season_prefix")
    filename_year = art.get("_filename_season_year")
    if filename_prefix:
        sea_prefix = filename_prefix
    if filename_year:
        full_year = filename_year

    sea_label = mdd.season_label(sea_prefix)
    # _val(vals_el, "AT_Season", sea_label, id_val=sea_prefix)
    # _val(vals_el, "AT_SeasonYear", full_year)

    cat_code = art["art_category"]
    cat_label = mdd.article_category(cat_code)
    _val(vals_el, "AT_SAPArticleCategory", cat_label, id_val=cat_code)

    at_label = mdd.by_article_type("Inline")
    _val(vals_el, "AT_BYArticleType", at_label, id_val="Inline")

    _val(vals_el, "AT_UOM", mdd.uom("EA"), id_val="EA")
    _val(vals_el, "AT_OriginalPrice", art["rrp"])

    curr = art.get("currency") or "IDR"
    _val(vals_el, "AT_RetailPriceCurrency", curr, id_val=curr)

    frs = (art.get("frs_whs", "") or "").upper()
    if "WHS" in frs:
        noa_label = mdd.nature_of_article("WH1")
        _val(vals_el, "AT_NatureOfArticle", noa_label, id_val="WH1")

    _val(vals_el, "AT_MainVendorIdentification", "1")
    _val(vals_el, "AT_InboundGenericCode", f"{b_code}{art['article_no']}")


def _add_variant_values(vals_el, art, size_rec, variant_key, mdd):
    _val(vals_el, "AT_InboundVariantCode", variant_key)

    tech_sz = size_rec.get("tech_size", "").strip()
    display_sz = size_rec.get("size", "").strip()
    size_lov, is_valid = _resolve_size_lov_id(tech_sz, display_sz, mdd)
    label = display_sz or tech_sz

    if size_lov and is_valid:
        _val(vals_el, "AT_Size", id_val=size_lov)
        log.debug(
            "[Variant] AT_Size resolved -- article=%s source=%s tech=%s resolved=%s",
            art["article_no"], label, tech_sz, size_lov,
        )
    else:
        log.warning(
            "[Variant] Skipping AT_Size invalid LOV -- article=%s tech=%s display=%s resolved=%s",
            art["article_no"], tech_sz, display_sz, size_lov,
        )


# ==================================================================
# SECTION 6 -- XML BUILDERS
# ==================================================================

def _build_classifications(brand, brand_code, season_raw, mdd, filename_season_prefix=None):
    sea_prefix, _, full_year = _parse_season(season_raw)

    if filename_season_prefix:
        sea_prefix = filename_season_prefix.upper()

    season_id = f"CLH_{brand_code}_{sea_prefix}{full_year}"
    batches_parent = f"CLH_{brand.capitalize()}Batches"
    sea_name = mdd.season_label(sea_prefix)

    cls_root = ET.Element(f"{{{STIBO_NS}}}Classifications")
    season_cls = ET.SubElement(cls_root, f"{{{STIBO_NS}}}Classification")
    season_cls.set("ID", season_id)
    season_cls.set("UserTypeID", "CLS_Season")
    season_cls.set("ParentID", batches_parent)
    season_cls.set("update", "true")
    ET.SubElement(season_cls, f"{{{STIBO_NS}}}Name").text = f"{brand} {sea_name} {full_year}"

    ca = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    ca.set("ID", f"{season_id}CA")
    ca.set("update", "true")
    ET.SubElement(ca, f"{{{STIBO_NS}}}Name").text = f"{sea_prefix} {full_year} Confirmed Articles"

    ua = ET.SubElement(season_cls, f"{{{STIBO_NS}}}Classification")
    ua.set("ID", f"{season_id}UA")
    ua.set("update", "true")
    ET.SubElement(ua, f"{{{STIBO_NS}}}Name").text = f"{sea_prefix} {full_year} Unconfirmed Articles"

    return cls_root, season_id


def _build_product_xml(art, brand, brand_code, comp_code, sbu, season_id, mdd):
    article_no = art["article_no"]
    if not article_no:
        return ""

    sizes = art.get("_sizes", [])
    div_letter = _get_division_code(art.get("division", ""))
    parent_id = f"PPH_{div_letter}-TempSubCat"
    b_code = art.get("brand_code", brand_code) or brand_code
    key_article = f"{b_code}{article_no}"
    colour_code = mdd.color_code_from_desc(art.get("colour", ""))

    g_el = ET.Element(f"{{{STIBO_NS}}}Product")
    g_el.set("UserTypeID", "PRD_GenericArticle")
    g_el.set("ParentID", parent_id)
    g_el.set("update", "true")

    kv = ET.SubElement(g_el, f"{{{STIBO_NS}}}KeyValue")
    kv.set("KeyID", "KEY_InboundArticle")
    kv.text = key_article

    vals_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Values")
    _add_generic_values(vals_el, art, comp_code, sbu, mdd)

    seen_keys = set()
    var_counter = 0
    for size_rec in sizes:
        tech_sz = size_rec.get("tech_size", "").strip()
        display_sz = size_rec.get("size", "").strip()
        ean = size_rec.get("ean", "").strip()

        variant_key = _build_variant_key(b_code, article_no, colour_code, tech_sz, display_sz)
        if variant_key in seen_keys:
            ean_sfx = re.sub(r"[^0-9]", "", ean)[-3:] if ean else f"{var_counter:03d}"
            variant_key = (variant_key[:15] + ean_sfx)[:18]
        if variant_key in seen_keys:
            log.warning("[Backlog] Skipping duplicate KEY_InboundVariant: %s", variant_key)
            continue

        seen_keys.add(variant_key)
        var_counter += 1

        v_el = ET.SubElement(g_el, f"{{{STIBO_NS}}}Product")
        v_el.set("UserTypeID", "PRD_VariantArticle")
        v_el.set("update", "true")

        v_kv = ET.SubElement(v_el, f"{{{STIBO_NS}}}KeyValue")
        v_kv.set("KeyID", "KEY_InboundVariant")
        v_kv.text = variant_key

        # ET.SubElement(v_el, f"{{{STIBO_NS}}}Name").text = f"Size {var_counter}"

        v_vals = ET.SubElement(v_el, f"{{{STIBO_NS}}}Values")
        _add_variant_values(v_vals, art, size_rec, variant_key, mdd)

        if ean:
            _add_barcode_dc(v_el, ean, variant_key, b_code)
        else:
            log.warning("[Variant] EAN missing -- article=%s tech_size=%s", article_no, tech_sz)

    return ET.tostring(g_el, encoding="unicode")


# ==================================================================
# SECTION 7 -- MAIN ENTRY POINT
# ==================================================================

def run_backlog(args, auditor=None, tmp_dir=None):
    if tmp_dir:
        base = Path(tmp_dir)
    else:
        base = Path(os.environ.get("LAMBDA_TMP_DIR", str(Path(__file__).parent)))

    input_dir = base / "input"
    backlog_dir = input_dir / "backlog"
    mdd_dir = input_dir / "mdd"
    xml_out_dir = base / "output" / "xml"
    log_dir = base / "output" / "logs"

    for _d in [backlog_dir, xml_out_dir, log_dir]:
        _d.mkdir(parents=True, exist_ok=True)

    local_mdd = next(mdd_dir.glob("*.xlsx"), None) if mdd_dir.exists() else None
    mdd = get_mdd(local_mdd)
    _reset_barcode_tracking(base)

    backlog_files = list(backlog_dir.glob("*.xlsx"))
    if not backlog_files:
        log.warning("[Backlog] No Backlog file found in %s -- skipping.", backlog_dir)
        return

    backlog_path = backlog_files[0]
    log.info("--- Backlog Pipeline: %s ---", backlog_path.name)

    wb = openpyxl.load_workbook(str(backlog_path), read_only=True, data_only=True)
    ean_loader = EANListLoader(wb)
    bl_loader = BacklogSheetLoader(wb)
    wb.close()

    if not bl_loader.all_articles:
        log.warning("[Backlog] No articles loaded -- skipping.")
        return

    filename_year = _parse_season_year_from_filename(backlog_path.stem)
    filename_sea_prefix = _parse_season_prefix_from_filename(backlog_path.stem)
    log.info("[Backlog] Season year from filename: %s", filename_year or "not detected")
    log.info("[Backlog] Season prefix from filename: %s", filename_sea_prefix or "not detected")

    matched_articles = bl_loader.all_articles & ean_loader.all_articles
    skipped_no_ean = bl_loader.all_articles - ean_loader.all_articles

    log.info(
        "[Backlog] Backlog articles=%d | EAN articles=%d | Matched=%d | Skipped(no EAN)=%d",
        len(bl_loader.all_articles), len(ean_loader.all_articles),
        len(matched_articles), len(skipped_no_ean),
    )
    if skipped_no_ean:
        log.info("[Backlog] Skipped articles (no EAN match) -- sample: %s", sorted(skipped_no_ean)[:10])

    articles = []
    missing_ean = 0
    for article_no in sorted(matched_articles):
        bl_meta = bl_loader.get_meta(article_no)
        ean_meta = ean_loader.get_meta(article_no)
        sizes = ean_loader.get_sizes(article_no)
        if not sizes:
            missing_ean += 1
            log.warning("[Backlog] No EAN sizes found for article: %s", article_no)
            continue

        mapped = _map_article(bl_meta, ean_meta, sizes, args.brand_code, mdd)
        if filename_year:
            mapped["_filename_season_year"] = filename_year
        if filename_sea_prefix:
            mapped["_filename_season_prefix"] = filename_sea_prefix
        articles.append(mapped)

    log.info("[Backlog] Mapped %d articles | %d missing EAN (skipped)", len(articles), missing_ean)

    if not articles:
        log.warning("[Backlog] No articles to write -- check EAN LIST S26=Y coverage.")
        return

    # ══════════════════════════════════════════════════════════════════
    # TEST LIMITER — Set to None for production, or a number to limit articles
    # ══════════════════════════════════════════════════════════════════
    # TEST_ARTICLE_LIMIT = 5  # Set to None to process all matched articles
    # if TEST_ARTICLE_LIMIT is not None:
    #     articles = articles[:TEST_ARTICLE_LIMIT]
    #     log.info("[Backlog] TEST MODE: Limited to first %d articles", TEST_ARTICLE_LIMIT)

    first_season = articles[0].get("season", args.season) if articles else args.season
    sea_prefix, _, full_year = _parse_season(first_season)
    if filename_year:
        full_year = filename_year
    effective_prefix = filename_sea_prefix or sea_prefix
    season_code = f"{effective_prefix}{full_year[2:]}"

    _, season_id = _build_classifications(
        args.brand, args.brand_code, season_code, mdd,
        filename_season_prefix=filename_sea_prefix,
    )

    out_name = backlog_path.stem + ".xml"
    out_path = xml_out_dir / out_name
    export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log.info("[Backlog] Writing XML -> %s", out_name)
    written_count = 0
    variant_count = 0
    barcode_value_count = 0

    with open(out_path, "w", encoding="utf-8", buffering=1 << 20) as f:
        f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        f.write(
            f"<STEP-ProductInformation"
            f" xmlns=\"{STIBO_NS}\""
            f" xmlns:xsi=\"{STIBO_XSI}\""
            f" xsi:schemaLocation=\"{STIBO_SCHEMA}\""
            f" ExportTime=\"{export_time}\""
            f" ExportContext=\"Context1\""
            f" ContextID=\"Context1\""
            f" WorkspaceID=\"Main\""
            f" UseContextLocale=\"false\">\n\n"
        )
        f.write("  <Products>\n")
        for art in articles:
            if not art.get("article_no"):
                continue
            product_xml = _build_product_xml(
                art, args.brand, args.brand_code,
                args.comp_code, args.sbu, season_id, mdd,
            )
            product_xml = _XMLNS_RE.sub("", product_xml)
            f.write(f"    {product_xml}\n")
            del product_xml
            written_count += 1
            variant_count += len(art.get("_sizes") or [])
            barcode_value_count += sum(1 for s in (art.get("_sizes") or []) if s.get("ean"))
        f.write("  </Products>\n")
        f.write("</STEP-ProductInformation>\n")

    _save_written_barcode_dc_ids(base)

    file_kb = out_path.stat().st_size // 1024
    log.info("OK Backlog XML written -> %s  (%dKB)", out_path, file_kb)

    print("=== BACKLOG SUMMARY =====================================", flush=True)
    print(f"  Backlog file         : {backlog_path.name}", flush=True)
    print(f"  MDD source           : {'local' if local_mdd else 'S3'}", flush=True)
    print(f"  Size codes in LOV    : {len(mdd.LOV_SIZE_CODE)}", flush=True)
    print(f"  SAP size mappings    : {len(mdd.SAP_SIZE_TO_LOV)}", flush=True)
    print(f"  Rejected size IDs    : {','.join(sorted(STIBO_REJECTED_SIZE_IDS))}", flush=True)
    print(f"  Season prefix (file) : {filename_sea_prefix or 'not detected'}", flush=True)
    print(f"  Season year (file)   : {filename_year or 'not detected'}", flush=True)
    print(f"  Season ID used       : {season_id}", flush=True)
    print(f"  Backlog articles     : {len(bl_loader.all_articles)}", flush=True)
    print(f"  EAN articles (S26=Y) : {len(ean_loader.all_articles)}", flush=True)
    print(f"  Matched articles     : {len(matched_articles)}", flush=True)
    print(f"  Skipped (no EAN)     : {len(skipped_no_ean)}", flush=True)
    print(f"  Total variant rows   : {variant_count}", flush=True)
    print(f"  Barcode values       : {barcode_value_count}", flush=True)
    print(f"  Barcode DC mode      : {_barcode_dc_mode()}", flush=True)
    print(f"  Barcode DC existing  : {len(_EXISTING_BARCODE_DC_IDS)}", flush=True)
    print(f"  Barcode DC written   : {len(_GENERATED_BARCODE_DC_IDS)}", flush=True)
    print(f"  XML articles written : {written_count}", flush=True)
    print(f"  XML file size        : {file_kb}KB", flush=True)
    print(f"  Output               : {out_name}", flush=True)
    print("=========================================================", flush=True)

    if auditor:
        try:
            auditor.set_backlog_result({
                "status": "ok",
                "backlog_file": backlog_path.name,
                "backlog_total": len(bl_loader.all_articles),
                "ean_total": len(ean_loader.all_articles),
                "matched": len(matched_articles),
                "skipped_no_ean": len(skipped_no_ean),
                "articles": written_count,
                "variants": variant_count,
                "barcodes": barcode_value_count,
                "barcode_dc_mode": _barcode_dc_mode(),
                "barcode_dc_existing": len(_EXISTING_BARCODE_DC_IDS),
                "barcode_dc_written": len(_GENERATED_BARCODE_DC_IDS),
                "xml_output": out_name,
                "filename_year": filename_year,
                "filename_season_prefix": filename_sea_prefix,
                "season_id": season_id,
                "size_lov_count": len(mdd.LOV_SIZE_CODE),
                "sap_size_map_count": len(mdd.SAP_SIZE_TO_LOV),
            })
        except AttributeError:
            pass


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate STEP XML from Backlog workbook.")
    parser.add_argument("--brand", default=os.environ.get("BACKLOG_BRAND", "ADIDAS"))
    parser.add_argument("--brand-code", default=os.environ.get("BACKLOG_BRAND_CODE", "ADI"))
    parser.add_argument("--comp-code", default=os.environ.get("BACKLOG_COMP_CODE", "0888"))
    parser.add_argument("--sbu", default=os.environ.get("BACKLOG_SBU", "SP"))
    parser.add_argument("--season", default=os.environ.get("BACKLOG_SEASON", "SS26"))
    parser.add_argument("--tmp-dir", default=os.environ.get("LAMBDA_TMP_DIR"))
    parser.add_argument("--log-level", default=os.environ.get("LOG_LEVEL", "INFO"))
    cli_args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, cli_args.log_level.upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )
    run_backlog(cli_args, tmp_dir=cli_args.tmp_dir)








