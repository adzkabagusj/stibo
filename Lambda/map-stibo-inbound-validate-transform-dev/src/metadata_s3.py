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
    log=None,
) -> None:
    """
    Add shared support workbooks stored at the bucket root.

    Brand-specific files already present in found keep priority. This lets a
    brand override a support file while allowing root-level shared MDD and
    Attributes List files to replace the old raw/metadata/ location.
    
    Scans bucket root for files matching patterns and picks the latest based on LastModified.
    """
    include_types = include_types or set(ROOT_METADATA_PATTERNS.keys())
    exclude_types = exclude_types or set()

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
                
                # Match against patterns
                for ftype, patterns in ROOT_METADATA_PATTERNS.items():
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
