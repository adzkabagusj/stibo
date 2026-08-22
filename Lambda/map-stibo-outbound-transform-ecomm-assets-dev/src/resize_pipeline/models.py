"""
Data classes shared across the entire pipeline.

These are plain dataclasses — no logic, no I/O.
Every other module imports from here so this file has zero internal dependencies.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import List


# ── Asset code constants ───────────────────────────────────────────────────────
# Last two characters after the final underscore in an asset name.
# e.g. "ADIHP3837_2_NW"  →  "NW"
VALID_ASSET_CODES: frozenset[str] = frozenset({"NW", "AW", "NM", "AM"})


@dataclass
class AssetInfo:
    """
    One product image asset parsed from the STIBO XML.

    Fields
    ------
    asset_id     : STIBO asset ID          e.g. "PI-1304806"
    asset_name   : STIBO asset name        e.g. "ADIHP3837_2_NW"
    relative_url : REST API path           e.g. "/restapi/assets/PI-1304806/content"
    asset_code   : resize margin selector  e.g. "NW"
    """
    asset_id:     str
    asset_name:   str
    relative_url: str
    asset_code:   str   # must be one of VALID_ASSET_CODES


@dataclass
class ProductInfo:
    """
    Top-level product metadata parsed from the STIBO XML.

    Fields
    ------
    product_id   : STIBO product ID        e.g. "1304550"
    product_name : human-readable name     e.g. "ADILETTE COMFORT 2.0"
    asset_ids    : ordered list of asset IDs linked to this product
    channel_ids  : sales channel Value IDs e.g. ["AT_CH_Digital_MY"]
    """
    product_id:   str
    product_name: str
    asset_ids:    List[str]
    channel_ids:  List[str]


@dataclass
class ResizeJob:
    """
    One unit of resize work: a single asset × a single spec column.

    Created by pipeline.py from the cross-product of assets and resolved specs.

    Fields
    ------
    asset_info : the source asset to resize
    spec_name  : key into RESIZE_SPECS     e.g. "digital_flagship"
    spec       : the resolved spec dict    e.g. { canvas_width: 800, ... }
    """
    asset_info: AssetInfo
    spec_name:  str
    spec:       dict


@dataclass
class OutputFile:
    """
    A successfully written JPEG file, ready to be added to a ZIP.

    Created by image_processor.py, consumed by zipper.py.

    Fields
    ------
    filepath      : absolute path to the JPEG on /tmp
    canvas_width  : pixel width of the canvas  (used for ZIP grouping)
    canvas_height : pixel height of the canvas (used for ZIP grouping)
    spec_name     : which spec produced this file
    """
    filepath:      Path
    canvas_width:  int
    canvas_height: int
    spec_name:     str