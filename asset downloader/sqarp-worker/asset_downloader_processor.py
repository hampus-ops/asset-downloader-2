#!/usr/bin/env python3
"""Download assets listed in an asset_downloader.xlsx workbook.

Output folders are always created beside this script:

    image-.../    Isolated images become cropped transparent RGBA PNGs;
                  environmental images retain their original dimensions
                  and become opaque RGB PNGs
    documents/    Downloaded document assets

Run with an explicit workbook:

    python asset_downloader_processor.py "C:\\path\\to\\asset_downloader.xlsx"

If no path is supplied, the script looks for asset_downloader.xlsx beside itself.
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import gc
import hashlib
import mimetypes
import os
import re
import sys
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook
import numpy as np
from PIL import Image, ImageOps, ImageSequence
from scipy.ndimage import binary_erosion, binary_propagation, label as label_regions


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

PPI = 300
ORIGINALS_ONLY = False
REMOVE_BACKGROUND = True
AI_BACKGROUND_FALLBACK = True
CROP_ENABLED = True
DOCUMENTS_TO_PDF = True
REQUEST_TIMEOUT_SECONDS = 60
# Balanced fast mode: downloads overlap while background removal remains
# deliberately limited below so Windows stays responsive.
DOWNLOAD_WORKERS = 6
DOWNLOAD_BATCH_SIZE = 18

# PNG compression is lossless at every level. Level 1 is much faster to save
# than Pillow's default, at the cost of somewhat larger output files.
PNG_COMPRESS_LEVEL = 1

# u2netp is rembg's lightweight model. It is substantially smaller and faster
# than BiRefNet on a CPU. Limiting ONNX to two threads keeps Windows responsive.
BACKGROUND_MODEL = "u2netp"
ONNX_CPU_THREADS = 2

# Ignore almost-invisible mask noise when finding the product bounds, then keep
# a small transparent breathing space around the product.
ALPHA_CROP_THRESHOLD = 10
CROP_PADDING_PIXELS = 5

# Product photos on a clean white studio background are safer and much faster
# to cut out by removing only white pixels connected to the outer edge. This
# preserves thin products, labels and separate product parts that an AI model
# can otherwise mistake for background.
WHITE_BORDER_MIN_VALUE = 220
WHITE_BORDER_MIN_COVERAGE = 0.08
WHITE_BORDER_MAX_COLOR_SPREAD = 22
WHITE_BACKGROUND_TOLERANCE = 55
WHITE_ALPHA_SOFT_START = 7
WHITE_ALPHA_SOFT_END = 42
EXTREME_IMAGE_ASPECT_RATIO = 0.15
FULL_FRAME_MIN_WIDTH_COVERAGE = 0.85
FULL_FRAME_MIN_HEIGHT_COVERAGE = 0.95
PACKAGE_MIN_LIGHT_NEUTRAL_COVERAGE = 0.30
INNER_HOLE_STRICT_TOLERANCE = 12
INNER_HOLE_MIN_AREA_FRACTION = 0.0003
INNER_HOLE_MAX_AREA_FRACTION = 0.08
INNER_HOLE_MIN_ASPECT_RATIO = 0.65
INNER_HOLE_MAX_ASPECT_RATIO = 1.55
INNER_HOLE_MIN_CIRCULARITY = 0.88

# A white backdrop is not proof that pale pixels belong to the background.
# Compare the color cutout against an independent foreground prediction before
# accepting it. Exclude only the compact openings identified above.
WHITE_GUARD_FOREGROUND_THRESHOLD = 200
WHITE_GUARD_MIN_LOST_PIXELS = 32
WHITE_GUARD_MAX_LOST_FRACTION = 0.0005
_background_session = None

SCRIPT_DIR = Path(__file__).resolve().parent
RUN_TIMESTAMP = (
    os.environ.get("ASSET_RUN_TIMESTAMP", "").strip()
    or datetime.now().strftime("%Y%m%d-%H%M")
)
RESUME_AFTER = max(0, int(os.environ.get("ASSET_RESUME_AFTER", "0")))
REPAIR_ISOLATED_THROUGH = max(
    0,
    int(os.environ.get("ASSET_REPAIR_ISOLATED_THROUGH", "0")),
)
REPAIR_EXTREME_ONLY = os.environ.get("ASSET_REPAIR_EXTREME_ONLY", "0") == "1"
IMAGES_DIR = SCRIPT_DIR / f"image-{RUN_TIMESTAMP}"
DOCUMENTS_DIR = SCRIPT_DIR / f"documents-{RUN_TIMESTAMP}"
NUMBA_CACHE_DIR = SCRIPT_DIR / ".numba_cache"

HEADERS = {
    "User-Agent": "Mozilla/5.0 asset-downloader-processor/1.0",
    "Accept": "image/jpeg,image/png,image/webp,application/pdf,*/*;q=0.8",
}

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

CONTENT_TYPE_EXTENSIONS = {
    "application/pdf": ".pdf",
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "image/tiff": ".tif",
}


@dataclass(frozen=True)
class Asset:
    row_number: int
    asset_file_name: str
    cdn_filename: str
    asset_type: str
    asset_url: str
    image_type: str


@dataclass(frozen=True)
class Download:
    path: Path
    content: bytes
    content_type: str


def clean_cell(value: Any) -> str:
    return "" if value is None else str(value).strip()


def read_assets(workbook_path: Path) -> list[Asset]:
    """Read non-empty asset rows from the workbook's File-list sheet."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if "File-list" not in workbook.sheetnames:
            raise ValueError("The workbook does not contain a 'File-list' sheet.")

        sheet = workbook["File-list"]
        rows = sheet.iter_rows(values_only=True)
        try:
            raw_headers = next(rows)
        except StopIteration as exc:
            raise ValueError("The 'File-list' sheet is empty.") from exc

        headers = [clean_cell(value).lower() for value in raw_headers]
        required = {"asset_file_name", "cdn_filename", "asset_type", "asset_url"}
        missing = sorted(required - set(headers))
        if missing:
            raise ValueError(f"Missing required column(s): {', '.join(missing)}")

        column = {name: headers.index(name) for name in required}
        image_type_column = (
            headers.index("image type") if "image type" in headers else None
        )
        assets: list[Asset] = []

        for row_number, row in enumerate(rows, start=2):
            values = list(row)

            def value(name: str) -> str:
                index = column[name]
                return clean_cell(values[index]) if index < len(values) else ""

            def image_type_value() -> str:
                if image_type_column is None or image_type_column >= len(values):
                    return ""
                return clean_cell(values[image_type_column]).casefold()

            url = value("asset_url")
            asset_type = value("asset_type").lower()
            if not url and not asset_type:
                continue
            if not url:
                print(f"Row {row_number}: skipped because asset_url is blank.")
                continue
            if asset_type not in {"image", "document"}:
                print(
                    f"Row {row_number}: skipped because asset_type is "
                    f"{asset_type!r}, not 'image' or 'document'."
                )
                continue
            if asset_type == "image" and image_type_column is None:
                raise ValueError(
                    "Missing required column 'image type' for image assets."
                )

            assets.append(
                Asset(
                    row_number=row_number,
                    asset_file_name=value("asset_file_name"),
                    cdn_filename=value("cdn_filename"),
                    asset_type=asset_type,
                    asset_url=url,
                    image_type=image_type_value(),
                )
            )

        return assets
    finally:
        workbook.close()


def safe_filename(name: str, fallback: str) -> str:
    """Make a single Windows-safe filename while preventing path traversal."""
    name = Path(unquote(name or "")).name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).rstrip(" .")
    if not name:
        name = fallback

    path = Path(name)
    if path.stem.upper() in WINDOWS_RESERVED_NAMES:
        name = f"_{name}"
    return name[:240]


def initial_filename(asset: Asset) -> str:
    url_name = Path(urlparse(asset.asset_url).path).name
    return safe_filename(
        asset.asset_file_name or asset.cdn_filename or url_name,
        f"asset_row_{asset.row_number}",
    )


def extension_for_content_type(content_type: str) -> str:
    media_type = content_type.partition(";")[0].strip().lower()
    return CONTENT_TYPE_EXTENSIONS.get(media_type) or mimetypes.guess_extension(media_type) or ""


def match_extension_to_content(filename: str, content_type: str) -> str:
    """Avoid files such as PNG data being saved with a misleading .pdf suffix."""
    expected = extension_for_content_type(content_type)
    if not expected:
        return filename

    aliases = {
        ".jpeg": ".jpg",
        ".jpe": ".jpg",
        ".tiff": ".tif",
    }
    current = aliases.get(Path(filename).suffix.lower(), Path(filename).suffix.lower())
    expected = aliases.get(expected.lower(), expected.lower())
    if current == expected:
        return filename

    path = Path(filename)
    if path.suffix:
        print(
            f"  Note: changed extension from {path.suffix} to {expected} "
            f"to match {content_type.partition(';')[0]}."
        )
        return f"{path.stem}{expected}"
    return f"{filename}{expected}"


def build_output_path(folder: Path, filename: str) -> Path:
    """Return a deterministic path so reruns overwrite instead of duplicating."""
    return folder / filename


def fetch_asset(asset: Asset) -> Download:
    request = Request(asset.asset_url, headers=HEADERS)
    with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        content = response.read()
        content_type = response.headers.get("Content-Type", "").strip()

    filename = match_extension_to_content(initial_filename(asset), content_type)
    downloaded_path = (
        IMAGES_DIR if asset.asset_type == "image" else DOCUMENTS_DIR
    ) / filename
    return Download(downloaded_path, content, content_type)


def decode_image(content: bytes) -> Image.Image:
    with Image.open(BytesIO(content)) as source:
        source = ImageOps.exif_transpose(source)
        if source.mode in {"RGBA", "LA"} or "transparency" in source.info:
            return source.convert("RGBA")
        return source.convert("RGB")


def has_source_transparency(image: Image.Image) -> bool:
    """Return true when the downloaded source contains visible alpha data."""
    if "A" not in image.getbands():
        return False
    minimum_alpha, _ = image.getchannel("A").getextrema()
    return minimum_alpha < 255


def crop_to_visible_product(image: Image.Image) -> Image.Image:
    """Crop an RGBA image to visible alpha, retaining a small safe padding."""
    if not CROP_ENABLED:
        return image.copy()
    alpha = image.getchannel("A")
    thresholded = alpha.point(
        lambda value: 255 if value > ALPHA_CROP_THRESHOLD else 0
    )
    bounds = thresholded.getbbox()
    if bounds is None:
        raise ValueError("Background removal produced an empty image.")

    left, top, right, bottom = bounds
    left = max(0, left - CROP_PADDING_PIXELS)
    top = max(0, top - CROP_PADDING_PIXELS)
    right = min(image.width, right + CROP_PADDING_PIXELS)
    bottom = min(image.height, bottom + CROP_PADDING_PIXELS)
    return image.crop((left, top, right, bottom))


def white_background_color(image: Image.Image) -> tuple[int, int, int] | None:
    """Return the border's white color when it is clearly a studio backdrop."""
    rgb = np.asarray(image.convert("RGB"), dtype=np.int16)
    border = np.concatenate(
        (rgb[0, :, :], rgb[-1, :, :], rgb[:, 0, :], rgb[:, -1, :]),
        axis=0,
    )
    channel_spread = border.max(axis=1) - border.min(axis=1)
    white_neutral = (
        (border.min(axis=1) >= WHITE_BORDER_MIN_VALUE)
        & (channel_spread <= WHITE_BORDER_MAX_COLOR_SPREAD)
    )
    coverage = float(white_neutral.mean())
    corner_pixels = np.asarray(
        [rgb[0, 0], rgb[0, -1], rgb[-1, 0], rgb[-1, -1]],
        dtype=np.int16,
    )
    corner_spread = corner_pixels.max(axis=1) - corner_pixels.min(axis=1)
    white_corners = (
        (corner_pixels.min(axis=1) >= WHITE_BORDER_MIN_VALUE)
        & (corner_spread <= WHITE_BORDER_MAX_COLOR_SPREAD)
    )
    # Some long, narrow product packages occupy nearly every border pixel.
    # Accept those when at least three corners still clearly show the studio
    # background, while requiring a useful amount of white edge evidence.
    if coverage < WHITE_BORDER_MIN_COVERAGE or int(white_corners.sum()) < 3:
        return None
    background = np.median(border[white_neutral], axis=0)
    return tuple(int(value) for value in background)


def save_white_background_isolated_png(
    image: Image.Image,
    source_name: str,
) -> Path | None:
    """Remove only a clean, edge-connected white studio background."""
    background = white_background_color(image)
    if background is None:
        return None

    rgb_image = image.convert("RGB")
    rgb = np.asarray(rgb_image, dtype=np.int16)
    background_array = np.asarray(background, dtype=np.int16)
    difference = np.max(np.abs(rgb - background_array), axis=2)

    # Candidate background is marked black. Flood-filling from border seeds
    # ensures white labels and highlights inside the product remain opaque.
    background_candidate = difference <= WHITE_BACKGROUND_TOLERANCE
    height, width = background_candidate.shape
    step = max(1, min(width, height) // 32)
    seed_mask = np.zeros_like(background_candidate, dtype=bool)
    for x in range(0, width, step):
        seed_mask[0, x] = background_candidate[0, x]
        seed_mask[height - 1, x] = background_candidate[height - 1, x]
    for y in range(0, height, step):
        seed_mask[y, 0] = background_candidate[y, 0]
        seed_mask[y, width - 1] = background_candidate[y, width - 1]
    seed_mask[0, 0] = background_candidate[0, 0]
    seed_mask[0, width - 1] = background_candidate[0, width - 1]
    seed_mask[height - 1, 0] = background_candidate[height - 1, 0]
    seed_mask[height - 1, width - 1] = background_candidate[height - 1, width - 1]
    connected_background = binary_propagation(
        seed_mask,
        mask=background_candidate,
    )

    # A product can contain a true opening whose pixels show the same studio
    # background but are enclosed by the product. Detect only substantial,
    # compact, rounded regions so white labels and printed details stay opaque.
    strict_white = difference <= INNER_HOLE_STRICT_TOLERANCE
    enclosed_white = strict_white & ~connected_background
    labeled, region_count = label_regions(enclosed_white)
    enclosed_holes = np.zeros_like(connected_background, dtype=bool)
    image_area = width * height
    minimum_hole_area = max(16, int(image_area * INNER_HOLE_MIN_AREA_FRACTION))
    maximum_hole_area = max(minimum_hole_area, int(image_area * INNER_HOLE_MAX_AREA_FRACTION))

    for region_id in range(1, region_count + 1):
        region = labeled == region_id
        area = int(region.sum())
        if area < minimum_hole_area or area > maximum_hole_area:
            continue
        ys, xs = np.nonzero(region)
        region_width = int(xs.max()) - int(xs.min()) + 1
        region_height = int(ys.max()) - int(ys.min()) + 1
        aspect_ratio = region_width / region_height
        if not (
            INNER_HOLE_MIN_ASPECT_RATIO
            <= aspect_ratio
            <= INNER_HOLE_MAX_ASPECT_RATIO
        ):
            continue
        perimeter = int((region & ~binary_erosion(region)).sum())
        if perimeter <= 0:
            continue
        circularity = 4.0 * np.pi * area / (perimeter * perimeter)
        if circularity < INNER_HOLE_MIN_CIRCULARITY:
            continue

        expanded_region = binary_propagation(
            region,
            mask=background_candidate & ~connected_background,
        )
        enclosed_holes |= expanded_region

    removable_background = connected_background | enclosed_holes
    soft_alpha = np.clip(
        (difference - WHITE_ALPHA_SOFT_START)
        * 255
        / (WHITE_ALPHA_SOFT_END - WHITE_ALPHA_SOFT_START),
        0,
        255,
    ).astype(np.uint8)
    alpha = np.where(removable_background, soft_alpha, 255).astype(np.uint8)

    # White packaging can be connected to the backdrop in color space. A
    # color flood-fill then erases or fades the product, even with valid edges.
    # Check every white-background cutout, including fully erased pale areas.
    if not AI_BACKGROUND_FALLBACK:
        # Without an independent foreground check, keep pale products intact.
        rgb_image.close()
        return None
    mask = foreground_mask(image, load_background_session())
    try:
        predicted_alpha = np.asarray(mask)
        product_interior = binary_erosion(
            predicted_alpha >= WHITE_GUARD_FOREGROUND_THRESHOLD,
            iterations=2,
        )
        lost_product = product_interior & (alpha < 245) & ~enclosed_holes
        lost_count = int(lost_product.sum())
        allowed = max(
            WHITE_GUARD_MIN_LOST_PIXELS,
            int(product_interior.sum() * WHITE_GUARD_MAX_LOST_FRACTION),
        )
        if lost_count > allowed:
            print(
                f"  White-product safeguard: color removal would damage "
                f"{lost_count} product pixels; using foreground mask."
            )
            rgb_image.close()
            return save_isolated_png(image, source_name, None, mask=mask)
    finally:
        mask.close()

    rgba = rgb_image.convert("RGBA")
    rgba.putalpha(Image.fromarray(alpha, mode="L"))
    output = crop_to_visible_product(rgba)
    rgba.close()
    rgb_image.close()

    target = build_output_path(IMAGES_DIR, f"{Path(source_name).stem}.png")
    output.save(
        target,
        "PNG",
        dpi=(PPI, PPI),
        compress_level=PNG_COMPRESS_LEVEL,
    )
    output.close()
    return target


def should_keep_extreme_package_background(image: Image.Image) -> bool:
    """Keep white behind unusually narrow, full-frame package photography."""
    width, height = image.size
    if width <= 0 or height <= 0:
        return False
    if min(width, height) / max(width, height) >= EXTREME_IMAGE_ASPECT_RATIO:
        return False

    background = white_background_color(image)
    if background is None:
        return False

    rgb = np.asarray(image.convert("RGB"), dtype=np.int16)
    background_array = np.asarray(background, dtype=np.int16)
    difference = np.max(np.abs(rgb - background_array), axis=2)
    foreground = difference > WHITE_BACKGROUND_TOLERANCE
    ys, xs = np.nonzero(foreground)
    if not len(xs):
        return False

    occupied_width = (int(xs.max()) - int(xs.min()) + 1) / width
    occupied_height = (int(ys.max()) - int(ys.min()) + 1) / height
    channel_spread = rgb.max(axis=2) - rgb.min(axis=2)
    light_neutral_package = (
        foreground
        & (rgb.min(axis=2) >= 130)
        & (channel_spread <= 35)
    )
    light_neutral_coverage = float(light_neutral_package.mean())
    return (
        occupied_width >= FULL_FRAME_MIN_WIDTH_COVERAGE
        and occupied_height >= FULL_FRAME_MIN_HEIGHT_COVERAGE
        and light_neutral_coverage >= PACKAGE_MIN_LIGHT_NEUTRAL_COVERAGE
    )


def foreground_mask(image: Image.Image, rembg_session: Any) -> Image.Image:
    """Predict at working resolution, return alpha at original pixel size."""
    from rembg import remove

    mask = remove(
        image, session=rembg_session, only_mask=True, post_process_mask=False,
    ).convert("L")
    if mask.size != image.size:
        resized = mask.resize(image.size, Image.Resampling.LANCZOS)
        mask.close()
        mask = resized
    if mask.getextrema()[1] < WHITE_GUARD_FOREGROUND_THRESHOLD:
        mask.close()
        raise ValueError("Foreground prediction is empty or uncertain; output not saved.")
    # The model can leave whole solid surfaces at alpha 254 and faint haze
    # outside the object. Snap only near-certain extremes; retain soft edges.
    normalized = mask.point(lambda value: 255 if value >= 245 else (0 if value <= 3 else value))
    mask.close()
    return normalized


def save_isolated_png(
    image: Image.Image,
    source_name: str,
    rembg_session: Any,
    *,
    mask: Image.Image | None = None,
) -> Path:
    """Remove the background, crop the product and save a transparent PNG."""
    owns_mask = mask is None
    if mask is None:
        mask = foreground_mask(image, rembg_session)
    rgba = image.convert("RGBA")
    rgba.putalpha(mask)
    if owns_mask:
        mask.close()
    output = crop_to_visible_product(rgba)
    rgba.close()
    target = build_output_path(IMAGES_DIR, f"{Path(source_name).stem}.png")
    output.save(
        target,
        "PNG",
        dpi=(PPI, PPI),
        compress_level=PNG_COMPRESS_LEVEL,
    )
    output.close()
    return target


def save_existing_transparent_png(image: Image.Image, source_name: str) -> Path:
    """Keep a source alpha channel instead of exposing hidden RGB pixels."""
    rgba = image.convert("RGBA")
    output = crop_to_visible_product(rgba)
    rgba.close()
    target = build_output_path(IMAGES_DIR, f"{Path(source_name).stem}.png")
    output.save(
        target,
        "PNG",
        dpi=(PPI, PPI),
        compress_level=PNG_COMPRESS_LEVEL,
    )
    output.close()
    return target


def save_reviewed_source_png(image: Image.Image, source_name: str) -> Path | None:
    """Reuse an approved full-resolution alpha only for exactly matching RGB.

    A filename alone is not sufficient: variants and replaced source images
    must never inherit a different product's manually reviewed contour.
    """
    rgb = image.convert("RGB")
    key = hashlib.sha256(
        f"{rgb.width}x{rgb.height}:RGB:".encode() + rgb.tobytes()
    ).hexdigest()
    rgb.close()
    mask_path = SCRIPT_DIR / "reviewed_masks" / f"{key}.png"
    if not mask_path.is_file():
        return None
    with Image.open(mask_path) as source_mask:
        if source_mask.mode != "L" or source_mask.size != image.size:
            raise ValueError("Reviewed source mask has an invalid size or mode.")
        mask = source_mask.copy()
    try:
        if mask.getextrema() != (0, 255):
            raise ValueError("Reviewed source mask must contain foreground and background.")
        return save_isolated_png(image, source_name, None, mask=mask)
    finally:
        mask.close()


def save_environmental_png(image: Image.Image, source_name: str) -> Path:
    """Keep the background and original pixel dimensions in an RGB PNG."""
    target = build_output_path(IMAGES_DIR, f"{Path(source_name).stem}.png")
    output = image.convert("RGB")
    output.save(
        target,
        "PNG",
        dpi=(PPI, PPI),
        compress_level=PNG_COMPRESS_LEVEL,
    )
    output.close()
    return target


def image_frame_to_rgb(frame: Image.Image) -> Image.Image:
    """Flatten transparency onto white because PDF pages cannot use RGBA."""
    frame = ImageOps.exif_transpose(frame)
    if frame.mode in {"RGBA", "LA"} or "transparency" in frame.info:
        rgba = frame.convert("RGBA")
        background = Image.new("RGBA", rgba.size, "white")
        return Image.alpha_composite(background, rgba).convert("RGB")
    return frame.convert("RGB")


def save_document_as_pdf(content: bytes, source_name: str) -> tuple[Path, bool]:
    """Save/convert a PDF, or retain an existing file during a safe rerun.

    Returns (path, was_already_present).
    """
    target = build_output_path(DOCUMENTS_DIR, f"{Path(source_name).stem}.pdf")

    # A PDF may be open in Explorer's preview pane or Adobe Reader. Do not try
    # to overwrite a valid existing file on a rerun; Windows would reject it.
    if target.exists() and target.stat().st_size > 0:
        return target, True

    if content.lstrip().startswith(b"%PDF-"):
        target.write_bytes(content)
        return target, False

    with Image.open(BytesIO(content)) as image:
        pages = [image_frame_to_rgb(frame.copy()) for frame in ImageSequence.Iterator(image)]

    if not pages:
        raise ValueError("Document is neither a PDF nor a supported image.")

    pages[0].save(
        target,
        "PDF",
        resolution=PPI,
        save_all=True,
        append_images=pages[1:],
    )
    return target, False


def create_output_folders() -> None:
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)


def load_background_session() -> Any:
    global _background_session
    if _background_session is not None:
        return _background_session
    # Microsoft Store Python installations can make Numba try to cache compiled
    # PyMatting files inside a protected/fragile site-packages path. Redirect the
    # cache to a normal writable folder before importing rembg (and thus Numba).
    NUMBA_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault("NUMBA_CACHE_DIR", str(NUMBA_CACHE_DIR))
    os.environ["OMP_NUM_THREADS"] = str(ONNX_CPU_THREADS)

    try:
        import onnxruntime as ort
        from rembg import new_session
    except ImportError as exc:
        raise RuntimeError(
            "rembg is required. Install the packages from requirements.txt."
        ) from exc

    session_options = ort.SessionOptions()
    session_options.intra_op_num_threads = ONNX_CPU_THREADS
    session_options.inter_op_num_threads = 1
    session_options.execution_mode = ort.ExecutionMode.ORT_SEQUENTIAL
    # Do not busy-spin while inference threads wait for work.
    session_options.add_session_config_entry("session.intra_op.allow_spinning", "0")
    session_options.add_session_config_entry("session.inter_op.allow_spinning", "0")

    print(
        f"Loading {BACKGROUND_MODEL} "
        "(the first run may download model weights)..."
    )
    session = new_session(BACKGROUND_MODEL, sess_opts=session_options)
    _background_session = session
    print("Background-removal model loaded.\n")
    return session


def unique_assets(assets: Iterable[Asset]) -> list[Asset]:
    """Reserve names before parallel downloads and ZIP flushing can remove files."""
    used = set()
    result = []
    for asset in assets:
        name = Path(initial_filename(asset))
        stem = name.stem[:200]
        candidate = stem
        number = 1
        key = (asset.asset_type, candidate.casefold())
        while key in used:
            number += 1
            candidate = f"{stem}-row-{asset.row_number}-{number}"
            key = (asset.asset_type, candidate.casefold())
        used.add(key)
        result.append(replace(asset, asset_file_name=candidate + name.suffix))
    return result


def process_assets(assets: Iterable[Asset], report=None, before_save=None) -> tuple[int, int]:
    assets = unique_assets(assets)
    total_count = len(assets)
    completed = min(RESUME_AFTER, total_count)
    assets = assets[completed:]
    successes = completed
    failures = 0

    if completed:
        print(
            f"Resuming after {completed} already completed assets; "
            f"{len(assets)} remain."
        )

    if not assets:
        return successes, failures

    worker_count = min(DOWNLOAD_WORKERS, len(assets))
    print(
        f"Downloading with {worker_count} parallel connections "
        f"in batches of {DOWNLOAD_BATCH_SIZE}..."
    )

    # Only one bounded batch is submitted at a time. This prevents completed
    # futures from retaining the bytes for an entire large workbook in RAM.
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        rembg_session = None
        for batch_start in range(0, len(assets), DOWNLOAD_BATCH_SIZE):
            batch = assets[batch_start : batch_start + DOWNLOAD_BATCH_SIZE]
            future_to_asset = {
                executor.submit(fetch_asset, asset): asset for asset in batch
            }

            for future in as_completed(future_to_asset):
                completed += 1
                asset = future_to_asset[future]
                label = asset.asset_file_name or asset.cdn_filename or asset.asset_url
                print(f"[{completed}/{total_count}] Row {asset.row_number}: {label}")
                image = None
                downloaded = None
                error = None
                try:
                    downloaded = future.result()
                    if before_save is not None:
                        before_save(len(downloaded.content))
                    if ORIGINALS_ONLY or (asset.asset_type == "document" and not DOCUMENTS_TO_PDF):
                        downloaded.path.write_bytes(downloaded.content)
                        successes += 1
                        continue
                    if asset.asset_type == "image":
                        image = decode_image(downloaded.content)
                        if before_save is not None:
                            before_save(max(len(downloaded.content), image.width * image.height * 4 + 65536))
                        if REPAIR_EXTREME_ONLY:
                            if (
                                asset.image_type == "isolated"
                                and not has_source_transparency(image)
                                and should_keep_extreme_package_background(image)
                            ):
                                saved_path = save_environmental_png(
                                    image,
                                    initial_filename(asset),
                                )
                                print(
                                    "  extreme package repair: original white "
                                    f"background retained: {saved_path}"
                                )
                            else:
                                print(
                                    "  not an extreme full-frame package; "
                                    "existing result retained."
                                )
                            successes += 1
                            continue
                        if asset.image_type == "isolated" and REMOVE_BACKGROUND:
                            reviewed_path = (
                                None if has_source_transparency(image)
                                else save_reviewed_source_png(image, initial_filename(asset))
                            )
                            if reviewed_path is not None:
                                print(f"  isolated (reviewed source contour): {reviewed_path}")
                                successes += 1
                                continue
                            if has_source_transparency(image):
                                saved_path = save_existing_transparent_png(
                                    image,
                                    initial_filename(asset),
                                )
                                print(
                                    "  isolated (from Excel; source transparency "
                                    f"preserved): cropped transparent PNG: {saved_path}"
                                )
                            else:
                                if should_keep_extreme_package_background(image):
                                    saved_path = save_environmental_png(
                                        image,
                                        initial_filename(asset),
                                    )
                                    print(
                                        "  isolated exception (extreme full-frame "
                                        "package): original white background retained: "
                                        f"{saved_path}"
                                    )
                                else:
                                    saved_path = save_white_background_isolated_png(
                                        image,
                                        initial_filename(asset),
                                    )
                                    if saved_path is not None:
                                        print(
                                            "  isolated (from Excel; clean white "
                                            "background, foreground checked): transparent PNG: "
                                            f"{saved_path}"
                                        )
                                    else:
                                        if not AI_BACKGROUND_FALLBACK:
                                            save_environmental_png(image, initial_filename(asset))
                                            successes += 1
                                            continue
                                        if rembg_session is None:
                                            rembg_session = load_background_session()
                                        saved_path = save_isolated_png(
                                            image,
                                            initial_filename(asset),
                                            rembg_session,
                                        )
                                        print(
                                            "  isolated (from Excel; complex "
                                            "background): AI transparent PNG: "
                                            f"{saved_path}"
                                        )
                        else:
                            saved_path = save_environmental_png(
                                image,
                                initial_filename(asset),
                            )
                            source_type = asset.image_type or "blank"
                            print(
                                f"  environmental (Image type: {source_type}): "
                                f"background retained: {saved_path}"
                            )
                    else:
                        saved_path, already_present = save_document_as_pdf(
                            downloaded.content,
                            initial_filename(asset),
                        )
                        if already_present:
                            print(f"  PDF already exists; retained: {saved_path}")
                        else:
                            print(f"  PDF document:    {saved_path}")

                    successes += 1
                except Exception as exc:
                    error = exc
                    failures += 1
                    print(f"  ERROR ({type(exc).__name__}): {exc}")
                finally:
                    if image is not None:
                        image.close()
                    downloaded = None
                    if report is not None:
                        report(completed, total_count, asset, error)

            future_to_asset.clear()
            gc.collect()

    return successes, failures


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download and process assets from asset_downloader.xlsx."
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        type=Path,
        default=SCRIPT_DIR / "asset_downloader.xlsx",
        help="Path to the input .xlsx file (default: beside this script).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook_path = args.workbook.expanduser().resolve()

    if not workbook_path.is_file():
        print(f"ERROR: Workbook not found: {workbook_path}", file=sys.stderr)
        return 2
    if workbook_path.suffix.lower() != ".xlsx":
        print("ERROR: The input file must be an .xlsx workbook.", file=sys.stderr)
        return 2

    try:
        assets = read_assets(workbook_path)
        if not assets:
            print("No downloadable image or document rows were found.")
            return 0

        if REPAIR_ISOLATED_THROUGH:
            repair_limit = min(REPAIR_ISOLATED_THROUGH, len(assets))
            assets = [
                asset
                for asset in assets[:repair_limit]
                if asset.asset_type == "image" and asset.image_type == "isolated"
            ]
            print(
                "Repair mode: reprocessing "
                f"{len(assets)} isolated images among the first "
                f"{repair_limit} assets."
            )

        create_output_folders()
        print(f"Workbook:  {workbook_path}")
        print(f"Images:    {IMAGES_DIR}")
        print(f"Documents: {DOCUMENTS_DIR}")
        print(f"Assets:    {len(assets)}\n")

        successes, failures = process_assets(assets)
        print(f"\nDone: {successes} succeeded, {failures} failed.")
        return 1 if failures else 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
