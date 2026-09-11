from pathlib import Path
from typing import Optional

from botocore.exceptions import ClientError


ROOT_METADATA_PATTERNS = {
    "mdd": ["master data dictionary", "mdd"],
    "attributes": [
        "brand mapping", "brand_mapping",
        "mapping files template", "mapping template", "brand template",
        "attributes list", "attribute list", "attributes_list",
    ],
}

# Brands that must use the legacy Attributes List v6 file.
# All other brands will use the latest NEW Brand Mapping file.
# "onr" and "2xu" are always-V6 (whole brand); "nike" is passed as the
# principal only for non-360 Nike files — Nike 360 keeps using the NEW
# Brand Mapping file (NIK 360 tab) — see nike/lambda_function.py.
# NOTE: Lotto was NOT added here — every Lotto ETL module that actually
# reads the attributes file (orderform_main.py / pricelist_main.py's RNA
# lookup, recap_main.py's AttributesListLoader) hardcodes the NEW mapping
# file's name/sheet ("NEW - Brand mapping files Template.xlsx" /
# "Lotto(Inline + Licensed)"), so switching Lotto to V6 here would silently
# empty out AT_BrandType/AT_BrandCategory (and recap's whole attribute map)
# rather than fail loudly. Flagged back to the user — see chat.
V6_BRANDS: set[str] = {
    "adidas", "new_balance", "smiggle", "aldo", "diadora",
    "onr", "2xu", "nike",
}

# Attribute patterns for v6 brands (matches Attributes List_complete_v6_*.xlsx)
_ATTR_PATTERNS_V6: list[str] = [
    "attributes list", "attribute list", "attributes_list", "attributes_complete",
]

# Attribute patterns for all other brands (matches NEW - Brand mapping files Template.xlsx)
_ATTR_PATTERNS_NEW_MAPPING: list[str] = [
    "brand mapping", "brand_mapping",
    "mapping files template", "mapping template", "brand template",
]

# Backward compatibility for router.py
ROOT_METADATA_KEYS = {
    "mdd": "Master Data Dictionary (MAA).xlsx",
    "attributes": "Attributes List_complete_v6_25032026.xlsx",
}


def _safe_ts_to_epoch(ts) -> int:
    """Convert timestamp to epoch for comparison."""
    try:
        if hasattr(ts, "timestamp"):
            return int(ts.timestamp())
        elif isinstance(ts, (int, float)):
            return int(float(ts))
        elif isinstance(ts, str):
            from datetime import datetime as _dt
            d = _dt.fromisoformat(ts.replace("Z", "+00:00"))
            return int(d.timestamp())
    except Exception:
        pass
    return 0


def add_root_metadata_files(
    *,
    s3_client,
    bucket: str,
    found: dict,
    include_types: Optional[set[str]] = None,
    exclude_types: Optional[set[str]] = None,
    principal: Optional[str] = None,
    log=None,
) -> None:
    """
    Add shared support workbooks stored at the bucket root.

    Brand-specific files already present in found keep priority. This lets a
    brand override a support file while allowing root-level shared MDD and
    Attributes List files to replace the old raw/metadata/ location.

    Scans bucket root for files matching patterns and picks the latest based on LastModified.

    For brands in V6_BRANDS (adidas, new_balance, smiggle, aldo, diadora) the legacy
    Attributes List v6 file is selected. All other brands use the latest NEW Brand Mapping file.
    """
    include_types = include_types or set(ROOT_METADATA_PATTERNS.keys())
    exclude_types = exclude_types or set()

    # Determine which attribute file pattern to use for this brand
    _principal_norm = (principal or "").lower().replace(" ", "_").replace("-", "_")
    _attr_patterns = (
        _ATTR_PATTERNS_V6 if _principal_norm in V6_BRANDS else _ATTR_PATTERNS_NEW_MAPPING
    )

    # Build effective patterns per type, overriding attributes based on brand group
    _effective_patterns = {
        ftype: (_attr_patterns if ftype == "attributes" else patterns)
        for ftype, patterns in ROOT_METADATA_PATTERNS.items()
    }

    if log:
        log.info(
            "  [AttrFile] brand='%s' → using %s attribute patterns",
            principal or "unknown",
            "v6" if _principal_norm in V6_BRANDS else "new-brand-mapping",
        )

    # List all objects at bucket root (non-recursive)
    try:
        paginator = s3_client.get_paginator("list_objects_v2")
        root_files = {}  # ftype -> {key, last_modified, filename}
        
        for page in paginator.paginate(Bucket=bucket, Delimiter="/"):
            for obj in page.get("Contents", []):
                key = obj["Key"]
                filename = Path(key).name.lower()
                
                # Skip if not an xlsx/xlsb/csv file
                if not filename.endswith((".xlsx", ".xlsb", ".csv")):
                    continue
                
                # Match against effective patterns
                for ftype, patterns in _effective_patterns.items():
                    if ftype not in include_types or ftype in exclude_types:
                        continue
                    
                    if any(pattern.lower() in filename for pattern in patterns):
                        last_modified = obj["LastModified"]
                        
                        # Keep the latest file for this type
                        if ftype not in root_files or (
                            _safe_ts_to_epoch(last_modified) > _safe_ts_to_epoch(root_files[ftype]["last_modified"])
                        ):
                            root_files[ftype] = {
                                "key": key,
                                "filename": Path(key).name,
                                "last_modified": last_modified,
                            }
                            if log:
                                log.info("  Found root %-12s <- %s  (modified: %s)", 
                                        ftype, Path(key).name, last_modified)
        
        # Add the latest files to found
        for ftype, file_info in root_files.items():
            found[ftype] = file_info
            if log:
                log.info("  Classified  %-22s <- %s  [bucket-root, LATEST]", ftype, file_info["filename"])
    
    except ClientError as exc:
        if log:
            log.warning("  Could not list root files in s3://%s/: %s", bucket, exc)
