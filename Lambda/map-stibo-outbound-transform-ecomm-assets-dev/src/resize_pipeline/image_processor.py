"""
Transform a source image according to a ResizeJob:

  1. Create an RGB canvas at the target dimensions with the spec background color.
  2. Subtract margins to get the drawable area.
  3. Scale the source image to fill the drawable area (aspect-ratio preserved).
  4. Centre the scaled image within the drawable area.
  5. Export as JPEG, stepping quality down until the file is ≤ 500 KB.
"""

import io
from pathlib import Path

from PIL import Image

from . import config
from .models import OutputFile, ResizeJob


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """
    Convert a hex color string to an RGB tuple.

    Parameters
    ----------
    hex_color : str
        Color string with or without leading '#' e.g. "#FFFFFF" or "FFFFFF".

    Returns
    -------
    tuple[int, int, int]
        (R, G, B) each in range 0–255.

    Examples
    --------
    >>> _hex_to_rgb("#FFFFFF")
    (255, 255, 255)
    >>> _hex_to_rgb("F5F5F5")
    (245, 245, 245)
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(
            f"Invalid background_color '{hex_color}'. "
            "Expected a 6-digit hex string e.g. '#FFFFFF'."
        )
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _fit_to_box(image: Image.Image, box_w: int, box_h: int) -> Image.Image:
    """
    Scale *image* to fit inside a box_w × box_h bounding box while
    preserving the original aspect ratio.

    Parameters
    ----------
    image : PIL.Image.Image
        Source image (any mode).
    box_w : int
        Maximum width of the target bounding box in pixels.
    box_h : int
        Maximum height of the target bounding box in pixels.

    Returns
    -------
    PIL.Image.Image
        Resized image — never larger than the original in either dimension.
    """
    img_w, img_h = image.size
    scale        = min(box_w / img_w, box_h / img_h)
    new_w        = max(1, round(img_w * scale))
    new_h        = max(1, round(img_h * scale))
    return image.resize((new_w, new_h), Image.LANCZOS)


def _compress_jpeg(image_rgb: Image.Image, dpi: int) -> bytes:
    """
    Export *image_rgb* as JPEG bytes.

    Quality is stepped down from INITIAL_JPEG_QUALITY to MIN_JPEG_QUALITY
    in steps of JPEG_QUALITY_STEP until the result fits within
    MAX_FILE_SIZE_BYTES (500 KB).

    Parameters
    ----------
    image_rgb : PIL.Image.Image
        RGB image to export — must be in 'RGB' mode (no alpha).
    dpi : int
        DPI value embedded in the JPEG header.

    Returns
    -------
    bytes
        Compressed JPEG bytes ≤ MAX_FILE_SIZE_BYTES.

    Raises
    ------
    RuntimeError
        If the image cannot be compressed below MAX_FILE_SIZE_BYTES
        even at the minimum quality.
    """
    quality = config.INITIAL_JPEG_QUALITY

    while quality >= config.MIN_JPEG_QUALITY:
        buf = io.BytesIO()
        image_rgb.save(
            buf,
            format="JPEG",
            quality=quality,
            dpi=(dpi, dpi),
            optimize=True,
        )
        data = buf.getvalue()

        if len(data) <= config.MAX_FILE_SIZE_BYTES:
            print(
                f"    JPEG quality={quality} → "
                f"{len(data) // 1024} KB ✓"
            )
            return data

        print(
            f"    JPEG quality={quality} → "
            f"{len(data) // 1024} KB  (too large, stepping down …)"
        )
        quality -= config.JPEG_QUALITY_STEP

    raise RuntimeError(
        f"Cannot compress image to ≤ {config.MAX_FILE_SIZE_BYTES // 1024} KB "
        f"even at JPEG quality={config.MIN_JPEG_QUALITY}. "
        "The image may be too large or too detailed."
    )


def process_image(
    source:     Image.Image,
    job:        ResizeJob,
    output_dir: Path,
) -> OutputFile:
    """
    Resize and transform *source* according to *job* and write the
    result as a JPEG file inside *output_dir*.

    Steps
    -----
    1. Create an RGB canvas at canvas_width × canvas_height filled
       with the spec background color.
    2. Subtract the asset-code margins to get the drawable area.
    3. Scale *source* to fit the drawable area (aspect-ratio preserved).
    4. Paste the scaled image centred inside the drawable area.
    5. Compress to JPEG ≤ 500 KB and write to disk.

    Parameters
    ----------
    source : PIL.Image.Image
        RGBA source image returned by downloader.download_image().
    job : ResizeJob
        Describes which spec and asset code to use.
    output_dir : Path
        Directory where the output JPEG will be written.

    Returns
    -------
    OutputFile
        Metadata about the written file — consumed by zipper.py.

    Raises
    ------
    ValueError
        If the margins leave no drawable area on the canvas.
    RuntimeError
        If JPEG compression cannot meet the 500 KB limit.
    """
    spec       = job.spec
    asset_code = job.asset_info.asset_code

    canvas_w:   int   = spec["canvas_width"]
    canvas_h:   int   = spec["canvas_height"]
    bg_color:   tuple = _hex_to_rgb(spec["background_color"])
    dpi:        int   = spec["dpi"]
    margins:    dict  = spec["margins"][asset_code]

    # ── 1. Create canvas ──────────────────────────────────────────────────────
    canvas = Image.new("RGB", (canvas_w, canvas_h), bg_color)

    # ── 2. Compute drawable area after margins ────────────────────────────────
    draw_x = margins["left"]
    draw_y = margins["top"]
    draw_w = canvas_w - margins["left"] - margins["right"]
    draw_h = canvas_h - margins["top"]  - margins["bottom"]

    if draw_w <= 0 or draw_h <= 0:
        raise ValueError(
            f"Spec '{job.spec_name}' / asset code '{asset_code}': "
            f"margins {margins} leave no drawable area on a "
            f"{canvas_w}×{canvas_h} canvas."
        )

    print(
        f"    Canvas {canvas_w}×{canvas_h}  "
        f"Drawable area {draw_w}×{draw_h} at ({draw_x},{draw_y})"
    )

    # ── 3. Scale source image to fit drawable area ────────────────────────────
    scaled     = _fit_to_box(source, draw_w, draw_h)
    sc_w, sc_h = scaled.size

    # ── 4. Centre scaled image within the drawable area ───────────────────────
    paste_x = draw_x + (draw_w - sc_w) // 2
    paste_y = draw_y + (draw_h - sc_h) // 2

    # Use the alpha channel as a mask if present (RGBA source)
    if scaled.mode == "RGBA":
        canvas.paste(scaled, (paste_x, paste_y), mask=scaled.split()[3])
    else:
        canvas.paste(scaled.convert("RGB"), (paste_x, paste_y))

    # ── 5. Compress and write JPEG ────────────────────────────────────────────
    jpeg_bytes = _compress_jpeg(canvas, dpi)

    # Output filename: {asset_name}_{spec_name}.jpg
    filename = f"{job.asset_info.asset_name}.jpg"
    out_path  = output_dir / filename
    out_path.write_bytes(jpeg_bytes)

    size_kb = len(jpeg_bytes) // 1024
    print(f"    ✓  {filename}  ({canvas_w}×{canvas_h}, {size_kb} KB)")

    return OutputFile(
        filepath=out_path,
        canvas_width=canvas_w,
        canvas_height=canvas_h,
        spec_name=job.spec_name,
    )