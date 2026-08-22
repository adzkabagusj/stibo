"""
Download product images from the STIBO REST API.

Auth
-----
Uses OIDC client-credentials flow via config.get_bearer_token().
The token is cached at module level in config.py and only refreshed
when within STIBO_TOKEN_LEEWAY seconds of expiry.

Download URL
-------------
Built from:
  config.ENVIRONMENT_BASE_URL  +  relative_url  +  context/workspace params

e.g.
  https://your-stibo-host.com/restapi/assets/PI-1304806/content
    ?context=Context1&workspace=Approved

The downloaded image is converted to RGBA so image_processor.py can
safely composite it onto any background color.
"""

import io
import urllib.error
import urllib.request

from PIL import Image, UnidentifiedImageError

from . import config


def build_download_url(relative_url: str) -> str:
    """
    Build the full STIBO asset download URL from a relative URL.

    Parameters
    ----------
    relative_url : str
        Relative URL from the product XML
        e.g. "/restapi/assets/PI-1304806/content"

    Returns
    -------
    str
        Full URL with context and workspace query parameters appended.

    Examples
    --------
    >>> build_download_url("/restapi/assets/PI-1304806/content")
    'https://host/restapi/assets/PI-1304806/content?context=Context1&workspace=Approved'
    """
    base = config.ENVIRONMENT_BASE_URL.rstrip("/")
    path = relative_url.lstrip("/")
    return (
        f"{base}/{path}"
        f"?context={config.STIBO_CONTEXT}"
        f"&workspace={config.STIBO_WORKSPACE}"
    )


def download_image(relative_url: str) -> Image.Image:
    """
    Download a single asset from the STIBO REST API and return it as
    an RGBA Pillow Image.

    The bearer token is fetched automatically via config.get_bearer_token()
    which handles caching and refresh transparently.

    Parameters
    ----------
    relative_url : str
        Relative URL from the product XML
        e.g. "/restapi/assets/PI-1304806/content"

    Returns
    -------
    PIL.Image.Image
        Downloaded image converted to RGBA mode.

    Raises
    ------
    RuntimeError
        If STIBO_BASE_URL is not set, the OIDC token cannot be obtained,
        or the HTTP request fails.
    ValueError
        If the response body cannot be decoded as an image.
    """
    # ── Guard: must have a base URL ───────────────────────────────────────────
    if not config.ENVIRONMENT_BASE_URL:
        raise RuntimeError(
            "STIBO_BASE_URL environment variable is not set. "
            "Cannot build asset download URL."
        )

    # ── Obtain token (cached or fresh) ────────────────────────────────────────
    # Raises RuntimeError if OIDC env vars are missing or token request fails
    token = config.get_bearer_token()

    url = build_download_url(relative_url)
    print(f"    ->  GET {url}")

    # ── Build request ─────────────────────────────────────────────────────────
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept":        "image/*",
        },
        method="GET",
    )

    # ── HTTP request ──────────────────────────────────────────────────────────
    # timeout is passed as a plain integer to urlopen - urllib.request has no
    # Timeout or ConnectionError attributes; all errors surface as URLError
    # (which wraps socket.timeout for timeouts) or HTTPError for non-2xx.
    try:
        with urllib.request.urlopen(req, timeout=config.DOWNLOAD_TIMEOUT_SECONDS) as resp:
            raw_bytes = resp.read()
    except urllib.error.HTTPError as exc:
        raise RuntimeError(
            f"STIBO returned HTTP {exc.code} for {url}. "
            f"Reason: {exc.reason}"
        ) from exc
    except urllib.error.URLError as exc:
        # Covers both connection errors and socket timeouts
        raise RuntimeError(
            f"Request failed for {url}: {exc.reason}"
        ) from exc

    content_length_kb = len(raw_bytes) // 1024
    print(f"    ok  {content_length_kb} KB received")

    # ── Decode image bytes ────────────────────────────────────────────────────
    try:
        image = Image.open(io.BytesIO(raw_bytes))
        image.load()   # force full decode - catches truncated files early
    except UnidentifiedImageError:
        raise ValueError(
            f"Response from {url} could not be identified as an image."
        )
    except Exception as exc:
        raise ValueError(
            f"Failed to decode image from {url}: {exc}"
        ) from exc

    # ── Convert to RGBA for safe alpha compositing in image_processor ─────────
    return image.convert("RGBA")