"""
Full pipeline orchestration.

Flow
----
  1. Parse product XML       -> ProductInfo + dict[AssetInfo]
  2. Resolve channel IDs     -> set of spec names
  3. Build resize jobs       -> list[ResizeJob]
  4. Download each asset     -> dict[asset_id, PIL.Image]
  5. Process each job        -> list[OutputFile]
  6. Zip by canvas size      -> list[Path]
"""

from pathlib import Path
from typing import Dict, List, Set

from . import config
from .downloader import download_image
from .image_processor import process_image
from .models import AssetInfo, OutputFile, ResizeJob
from .xml_parser import parse_product_xml
from .zipper import zip_outputs


def _resolve_specs(channel_ids: List[str]) -> Set[str]:
    """
    Map sales-channel Value IDs to the set of unique spec column names.

    Channels with no mapping in CHANNEL_TO_SPEC are skipped with a warning.

    Parameters
    ----------
    channel_ids : list[str]
        Sales channel Value IDs from the product XML
        e.g. ["AT_CH_Digital_MY", "AT_CH_Zalora"].

    Returns
    -------
    set[str]
        Unique spec names e.g. {"digital_flagship", "zalora"}.
    """
    specs: Set[str] = set()

    for ch_id in channel_ids:
        normalised = ch_id.strip().upper()
        spec = config.CHANNEL_TO_SPEC.get(normalised)
        if spec:
            specs.add(spec)
        else:
            print(
                f"  [WARN] Channel ID '{ch_id}' (normalised: '{normalised}') "
                "has no mapping in CHANNEL_TO_SPEC "
                "- skipped. Check the 'Channels' sheet in Image Requirements.xlsx."
            )

    return specs


def _build_jobs(
    assets:         Dict[str, AssetInfo],
    asset_to_specs: Dict[str, Set[str]],
) -> List[ResizeJob]:
    """
    Build ResizeJob objects for each asset using only the specs that
    correspond to the products referencing that asset.

    Deduplication is implicit: asset_to_specs already holds a *set* of
    unique spec names per asset, so each (asset, spec) pair appears once.

    Parameters
    ----------
    assets : dict[str, AssetInfo]
        Parsed assets from xml_parser.parse_product_xml().
    asset_to_specs : dict[str, set[str]]
        Per-asset set of resolved spec column names, keyed by asset_id.

    Returns
    -------
    list[ResizeJob]
        One job per unique (asset, spec) combination.
    """
    jobs: List[ResizeJob] = []

    for asset_id, spec_names in asset_to_specs.items():
        asset_info = assets.get(asset_id)
        if asset_info is None:
            continue

        for spec_name in sorted(spec_names):          # sorted for predictable log order

            # ── Guard: spec must exist in RESIZE_SPECS ────────────────────────
            spec = config.RESIZE_SPECS.get(spec_name)
            if spec is None:
                print(
                    f"  [WARN] Spec '{spec_name}' not found in RESIZE_SPECS "
                    "- skipped. Check the 'Specs' sheet in Image Requirements.xlsx."
                )
                continue

            # ── Guard: asset code must have margins in this spec ──────────────
            if asset_info.asset_code not in spec.get("margins", {}):
                print(
                    f"  [WARN] Asset code '{asset_info.asset_code}' has no margin "
                    f"entry in spec '{spec_name}' - {asset_info.asset_id} skipped."
                )
                continue

            jobs.append(
                ResizeJob(
                    asset_info=asset_info,
                    spec_name=spec_name,
                    spec=spec,
                )
            )

    return jobs


def run_pipeline(
    xml_path:   Path,
    output_dir: Path,
) -> List[Path]:
    """
    Execute the full resizing pipeline for one product XML file.

    Parameters
    ----------
    xml_path : Path
        Path to the product XML file on disk (written to /tmp by the handler).
    output_dir : Path
        Root directory for all output - JPEGs go in output_dir/images/,
        ZIPs go directly in output_dir/.

    Returns
    -------
    list[Path]
        Paths of all created ZIP files. Empty list if no output was produced.

    Raises
    ------
    RuntimeError
        If config dicts are empty (Excel not loaded), env vars are missing,
        or the OIDC token request fails.
    ValueError
        Propagated from xml_parser if the XML is malformed.
    """
    # ── Guard: config must be loaded ──────────────────────────────────────────
    if not config.RESIZE_SPECS:
        raise RuntimeError(
            "RESIZE_SPECS is empty - Image Requirements.xlsx was not loaded. "
            "Check that S3_BUCKET is set correctly."
        )
    if not config.CHANNEL_TO_SPEC:
        raise RuntimeError(
            "CHANNEL_TO_SPEC is empty - Image Requirements.xlsx was not loaded. "
            "Check that S3_BUCKET is set correctly."
        )

    images_dir = output_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    # ── Step 1: Parse XML ─────────────────────────────────────────────────────
    print(f"\n{'─' * 60}")
    print(f"[1/4] Parsing  {xml_path.name}")
    print(f"{'─' * 60}")

    products, assets = parse_product_xml(xml_path)

    if not products:
        print("\n  [ERROR] No valid products found - aborting pipeline.")
        return []

    if not assets:
        print("\n  [ERROR] No valid assets found - aborting pipeline.")
        return []

    # ── Build per-asset spec mapping from each product's channel IDs ──────────
    # Each asset only receives specs from the product(s) that reference it.
    # set.update() deduplicates channel IDs that resolve to the same spec.
    asset_to_specs: Dict[str, Set[str]] = {}
    for product_info in products:
        product_specs = _resolve_specs(product_info.channel_ids)
        print(f"\n  Product {product_info.product_id} specs: {sorted(product_specs)}")
        for asset_id in product_info.asset_ids:
            asset_to_specs.setdefault(asset_id, set()).update(product_specs)

    all_spec_names: Set[str] = set().union(*asset_to_specs.values()) if asset_to_specs else set()

    if not all_spec_names:
        print("\n  [ERROR] No recognised sales channels across all products - aborting pipeline.")
        return []

    print(f"\n  Combined specs : {sorted(all_spec_names)}")

    # ── Build jobs - each asset only gets specs from its own products ──────────
    filtered_assets = {k: v for k, v in assets.items() if k in asset_to_specs}

    jobs = _build_jobs(filtered_assets, asset_to_specs)
    if not jobs:
        print("\n  [ERROR] No valid resize jobs - aborting pipeline.")
        return []

    print(f"  Total products : {len(products)}")
    print(f"  Total assets   : {len(filtered_assets)}")
    print(f"  Jobs built     : {len(jobs)}")

    # ── Step 2: Download each unique asset once ───────────────────────────────
    print(f"\n{'─' * 60}")
    print(f"[2/4] Downloading {len(filtered_assets)} asset(s)")
    print(f"{'─' * 60}")

    image_cache: Dict[str, object] = {}

    for asset_id, asset_info in filtered_assets.items():
        print(f"\n  Asset {asset_id}  '{asset_info.asset_name}'")
        try:
            image_cache[asset_id] = download_image(
                asset_info.relative_url,
            )
        except Exception as exc:
            print(f"  [ERROR] Download failed: {exc}")

    if not image_cache:
        print("\n  [ERROR] All downloads failed - aborting pipeline.")
        return []

    # ── Step 3: Process each job ──────────────────────────────────────────────
    print(f"\n{'─' * 60}")
    print(f"[3/4] Processing {len(jobs)} resize job(s)")
    print(f"{'─' * 60}")

    output_files: List[OutputFile] = []

    for job in jobs:
        source = image_cache.get(job.asset_info.asset_id)
        if source is None:
            print(
                f"\n  [SKIP] {job.asset_info.asset_id} / {job.spec_name} "
                "- download failed earlier."
            )
            continue

        print(f"\n  {job.asset_info.asset_name}  x  {job.spec_name}")
        try:
            out_file = process_image(source, job, images_dir)
            output_files.append(out_file)
        except Exception as exc:
            print(
                f"  [ERROR] {job.asset_info.asset_id} / {job.spec_name}: {exc}"
            )

    if not output_files:
        print("\n  [ERROR] No images processed - no ZIPs will be created.")
        return []

    # ── Step 4: Zip by canvas size ────────────────────────────────────────────
    print(f"\n{'─' * 60}")
    print(f"[4/4] Creating ZIP archives")
    print(f"{'─' * 60}")

    zip_paths = zip_outputs(output_files, output_dir)

    print(f"\n  Pipeline complete - {len(zip_paths)} ZIP(s) written to {output_dir}")

    return zip_paths