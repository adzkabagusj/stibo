"""
Group processed images by canvas size and create timestamped ZIP archives.

Grouping rule
─────────────
All output files with the same canvas_width × canvas_height go into
the same ZIP — even if they came from different specs.
This means if two specs both produce 800×800 images, they share one ZIP.

Output filename format
───────────────────────
  {width}x{height}-{YYYY-MM-DD_HH.MM.SS}.zip
  e.g. 800x800-2026-06-09_10.30.00.zip
"""

import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import List

from .models import OutputFile


def zip_outputs(
    output_files: List[OutputFile],
    output_dir:   Path,
) -> List[Path]:
    """
    Create one ZIP archive per unique canvas size from *output_files*.

    Parameters
    ----------
    output_files : list[OutputFile]
        All successfully processed JPEG files from image_processor.py.
    output_dir : Path
        Directory where the ZIP files will be written.

    Returns
    -------
    list[Path]
        Sorted list of created ZIP file paths (sorted by canvas size).

    Raises
    ------
    RuntimeError
        If output_files is empty — nothing to zip.
    """
    if not output_files:
        raise RuntimeError(
            "zip_outputs called with an empty output_files list. "
            "No ZIP archives created."
        )

    # ── Shared timestamp for all ZIPs in this invocation ─────────────────────
    timestamp = datetime.now().strftime("%Y-%m-%d_%H.%M.%S")

    # ── Group files by canvas size ────────────────────────────────────────────
    # key: (canvas_width, canvas_height)
    # value: list of OutputFile
    groups: defaultdict[tuple[int, int], List[OutputFile]] = defaultdict(list)

    for f in output_files:
        groups[(f.canvas_width, f.canvas_height)].append(f)

    print(
        f"\n  {len(groups)} canvas size group(s) → "
        f"{len(output_files)} file(s) total"
    )

    zip_paths: List[Path] = []

    # ── Create one ZIP per group, sorted by canvas size for predictable output ─
    for (w, h), files in sorted(groups.items()):

        zip_name = f"{w}x{h}-{timestamp}.zip"
        zip_path = output_dir / zip_name

        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for out_file in files:
                # Store file inside a folder named by canvas size e.g. 800x800/filename.jpg
                zf.write(out_file.filepath, arcname=f"{w}x{h}-{timestamp}/{out_file.filepath.name}")

        # ── Size reporting ────────────────────────────────────────────────────
        zip_size_kb        = zip_path.stat().st_size // 1024
        uncompressed_kb    = sum(f.filepath.stat().st_size for f in files) // 1024
        file_names_preview = ", ".join(f.filepath.name for f in files[:3])
        if len(files) > 3:
            file_names_preview += f" … +{len(files) - 3} more"

        print(
            f"  📦  {zip_name}\n"
            f"      {len(files)} file(s)  |  "
            f"{uncompressed_kb} KB uncompressed  →  {zip_size_kb} KB zipped\n"
            f"      {file_names_preview}"
        )

        zip_paths.append(zip_path)

    return zip_paths