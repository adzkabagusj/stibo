"""
Parse a STIBO STEP product XML file.

Expected XML structure (simplified)
-------------------------------------
    <STEP-ProductInformation>
      <Products>
        <Product ID="1304550">
          <Name xml:lang="en-US">ADILETTE COMFORT 2.0</Name>
          <AssetCrossReference Type="REF_ProductImage" AssetID="PI-1304782"/>
          <Values>
            <MultiValue AttributeID="AT_EcommSalesChannel">
              <Value ID="AT_CH_Digital_MY">Digital MY</Value>
            </MultiValue>
          </Values>
        </Product>
      </Products>
      <Assets>
        <Asset ID="PI-1304782">
          <Name xml:lang="en-US">ADIHP3837_1_NW</Name>
          <AssetContentSpecifications>
            <AssetContentSpecification>
              <RelativeURL>/restapi/assets/PI-1304782/content</RelativeURL>
            </AssetContentSpecification>
          </AssetContentSpecifications>
        </Asset>
      </Assets>
    </STEP-ProductInformation>

RelativeURL can appear in two forms depending on the STIBO export version:

  Element form (standard):
    <AssetContentSpecification>
      <RelativeURL>/restapi/assets/PI-1304782/content</RelativeURL>
    </AssetContentSpecification>

  Attribute form (some exports):
    <AssetContentSpecification RelativeURL="/restapi/assets/PI-1304782/content"/>

Both forms are handled; element form is tried first.

Asset code extraction
----------------------
The last two characters after the final underscore in the asset name
determine the resize margin selector:

    "ADIHP3837_2_NW"  ->  "NW"
    "ADIHP3837_AM"    ->  "AM"

Valid codes: NW, AW, NM, AM
Assets with unrecognised codes are skipped with a warning.
"""

from pathlib import Path
from typing import Dict, List, Optional

from lxml import etree

from .models import AssetInfo, ProductInfo, VALID_ASSET_CODES


def _extract_asset_code(asset_name: str) -> Optional[str]:
    """
    Extract the two-character asset code from an asset name.

    Parameters
    ----------
    asset_name : str
        Raw asset name from the XML e.g. "ADIHP3837_2_NW".

    Returns
    -------
    str or None
        Uppercased two-character code if recognised, else None.

    Examples
    --------
    >>> _extract_asset_code("ADIHP3837_2_NW")
    'NW'
    >>> _extract_asset_code("ADIHP3837_AM")
    'AM'
    >>> _extract_asset_code("ADIHP3837_XX")
    None
    """
    parts = asset_name.strip().upper().split("_")
    if not parts:
        return None

    # Take the last segment and use its first two characters as the code
    code = parts[-1][:2]
    return code if code in VALID_ASSET_CODES else None


def _extract_relative_url(asset_node) -> Optional[str]:
    """
    Extract the RelativeURL from an <Asset> node.

    Tries two forms in order:

    1. Element form (standard STIBO export):
         <AssetContentSpecification>
           <RelativeURL>/restapi/assets/.../content</RelativeURL>
         </AssetContentSpecification>

    2. Attribute form (some STIBO export variants):
         <AssetContentSpecification RelativeURL="/restapi/assets/.../content"/>

    Parameters
    ----------
    asset_node : lxml.etree._Element
        The <Asset> XML element to search within.

    Returns
    -------
    str or None
        The URL string if found, else None.
    """
    # ── Form 1: <RelativeURL> as a child element ──────────────────────────────
    url_node = asset_node.find(".//RelativeURL")
    if url_node is not None and url_node.text and url_node.text.strip():
        return url_node.text.strip()

    # ── Form 2: RelativeURL as an attribute on <AssetContentSpecification> ────
    spec_node = asset_node.find(".//AssetContentSpecification")
    if spec_node is not None:
        attr_val = spec_node.get("RelativeURL")
        if attr_val and attr_val.strip():
            return attr_val.strip()

    return None


def parse_product_xml(
    xml_path: Path,
) -> tuple[List[ProductInfo], Dict[str, AssetInfo]]:
    """
    Parse a STIBO STEP product XML file.

    Parameters
    ----------
    xml_path : Path
        Absolute path to the XML file on disk (written to /tmp by the handler).

    Returns
    -------
    products : list[ProductInfo]
        List of product metadata for ALL products in the XML.
    assets : dict[str, AssetInfo]
        Mapping of Asset ID -> AssetInfo for every valid asset linked
        to any product. Assets with unrecognised codes are excluded.

    Raises
    ------
    ValueError
        If the XML contains no <Product> node.
    lxml.etree.XMLSyntaxError
        If the file cannot be parsed as XML.
    """
    tree = etree.parse(str(xml_path))
    root = tree.getroot()

    # ── Locate ALL <Product> nodes ────────────────────────────────────────────
    product_nodes = root.findall(".//Product")
    if not product_nodes:
        raise ValueError(
            f"No <Product> node found in {xml_path.name}. "
            "Verify the XML is a valid STIBO STEP product export."
        )

    print(f"  Found {len(product_nodes)} product(s) in XML")

    products: List[ProductInfo] = []
    all_asset_ids: set = set()

    # ── Parse each product ────────────────────────────────────────────────────
    for idx, product_node in enumerate(product_nodes, start=1):
        product_id = product_node.get("ID", "UNKNOWN")

        # Skip products without an ID attribute (e.g. template nodes)
        if product_id == "UNKNOWN" or not product_id:
            print(f"\n  [WARN] Product #{idx} has no ID attribute - skipped.")
            continue

        # ── Product name ──────────────────────────────────────────────────────
        name_node    = product_node.find("Name")
        product_name = (
            name_node.text.strip()
            if name_node is not None and name_node.text
            else product_id
        )

        # ── Asset cross-references (<AssetCrossReference Type="REF_ProductImage">) ─
        asset_ids: List[str] = [
            ref.get("AssetID")
            for ref in product_node.findall(
                ".//AssetCrossReference[@Type='REF_ProductImage']"
            )
            if ref.get("AssetID")
        ]

        # ── Sales channel Value IDs ───────────────────────────────────────────
        channel_ids: List[str] = [
            val.get("ID")
            for val in product_node.findall(
                ".//MultiValue[@AttributeID='AT_EcommSalesChannel']/Value"
            )
            if val.get("ID")
        ]

        product_info = ProductInfo(
            product_id=product_id,
            product_name=product_name,
            asset_ids=asset_ids,
            channel_ids=channel_ids,
        )

        print(f"\n  Product #{idx}")
        print(f"    ID         : {product_id}")
        print(f"    Name       : {product_name}")
        print(f"    Asset IDs  : {asset_ids or 'none'}")
        print(f"    Channel IDs: {channel_ids or 'none'}")

        if not asset_ids:
            print(
                "    [WARN] No REF_ProductImage asset cross-references found. "
                "This product will produce no output."
            )

        if not channel_ids:
            print(
                "    [WARN] No AT_EcommSalesChannel values found. "
                "This product will produce no output."
            )

        # Only add products that have both assets and channels
        if asset_ids and channel_ids:
            products.append(product_info)
            all_asset_ids.update(asset_ids)
        else:
            print(f"    [SKIP] Product {product_id} excluded (missing assets or channels).")

    if not products:
        print("\n  [WARN] No valid products found with both assets and channels.")
        return [], {}

    print(f"\n  Valid products to process: {len(products)}")

    # ── Parse <Asset> nodes ───────────────────────────────────────────────────
    assets: Dict[str, AssetInfo] = {}
    _first_asset_logged = False

    for asset_node in root.findall(".//Asset"):
        asset_id = asset_node.get("ID")

        # Only process assets that are cross-referenced by any product
        if not asset_id or asset_id not in all_asset_ids:
            continue

        # ── Debug: print raw XML of first matched asset to verify structure ───
        if not _first_asset_logged:
            raw_xml = etree.tostring(asset_node, pretty_print=True).decode("utf-8")
            print(f"  [DEBUG] First matched asset node raw XML:\n{raw_xml}")
            _first_asset_logged = True

        # ── Asset name ────────────────────────────────────────────────────────
        name_node  = asset_node.find("Name")
        asset_name = (
            name_node.text.strip()
            if name_node is not None and name_node.text
            else asset_id
        )

        # ── Relative URL for download ─────────────────────────────────────────
        relative_url = _extract_relative_url(asset_node)
        if not relative_url:
            print(
                f"  [WARN] Asset {asset_id} ('{asset_name}') has no RelativeURL "
                "in either element or attribute form - skipped."
            )
            continue

        # ── Asset code (NW / AW / NM / AM) ───────────────────────────────────
        asset_code = _extract_asset_code(asset_name)
        if asset_code is None:
            print(
                f"  [WARN] Asset {asset_id}: cannot determine asset code from "
                f"name '{asset_name}'. "
                f"Expected one of {sorted(VALID_ASSET_CODES)} as the last "
                "underscore-separated segment - skipped."
            )
            continue

        assets[asset_id] = AssetInfo(
            asset_id=asset_id,
            asset_name=asset_name,
            relative_url=relative_url,
            asset_code=asset_code,
        )
        print(f"  Asset {asset_id}: '{asset_name}' (code={asset_code})")

    return products, assets