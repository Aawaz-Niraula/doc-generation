"""assets.py — remote asset acquisition (embedded fonts, topic images).

Everything here is time-boxed and fails soft: a slow CDN can never block or
break a generation, it only means the renderer falls back (system font stacks,
geometric SVG panels instead of photos).
"""

from __future__ import annotations

import asyncio
import base64
import re
import urllib.parse
from typing import Dict, Optional

import httpx

from content_model import DocumentSpec
from theme import Theme

# ─── Real font embedding (WeasyPrint-safe) ────────────────────────────────────
#
# WeasyPrint can mis-parse the Google Fonts CSS endpoint, so we fetch actual
# static TTFs from the Fontsource CDN and inline them as base64 @font-face
# rules. This makes the AI-chosen typography *actually render* in the PDF.

_TTF_MAGIC = (b"\x00\x01\x00\x00", b"true", b"OTTO", b"ttcf")
_FONTSOURCE = "https://cdn.jsdelivr.net/fontsource/fonts/{slug}@latest/latin-{w}-normal.ttf"
_FONT_WEIGHTS = {
    "display": (700, 800, 900),
    "heading": (500, 700),
    "body":    (400, 600, 700),
    "data":    (500, 700),
}
# Families whose CDN slug differs from a naive lower-hyphen transform.
_FONT_SLUG_FIXUPS = {
    "source sans 3": "source-sans-3", "source sans pro": "source-sans-pro",
    "ibm plex sans": "ibm-plex-sans", "ibm plex mono": "ibm-plex-mono",
    "pt sans": "pt-sans", "pt serif": "pt-serif",
}


def _font_slug(name: str) -> str:
    key = " ".join(str(name).strip().lower().split())
    if key in _FONT_SLUG_FIXUPS:
        return _FONT_SLUG_FIXUPS[key]
    return re.sub(r"[^a-z0-9]+", "-", key).strip("-")


async def _fetch_one_face(client: httpx.AsyncClient, family: str, weight: int) -> Optional[str]:
    try:
        r = await client.get(_FONTSOURCE.format(slug=_font_slug(family), w=weight))
        if r.status_code == 200 and r.content[:4] in _TTF_MAGIC and len(r.content) < 900_000:
            b64 = base64.b64encode(r.content).decode("ascii")
            return (f"@font-face{{font-family:'{family}';font-style:normal;"
                    f"font-weight:{weight};font-display:swap;"
                    f"src:url(data:font/ttf;base64,{b64}) format('truetype');}}")
    except Exception:
        pass
    return None


async def build_embedded_font_css(theme: Theme, timeout: float = 6.0) -> str:
    """Fetch the theme's PDF typography as base64 @font-face rules.

    Returns a <style> block, or '' on any failure so PDF rendering never
    depends on the network.
    """
    families: Dict[str, set] = {}
    for key, weights in _FONT_WEIGHTS.items():
        fam = str(theme.pdf_fonts.get(key, "")).strip()
        if fam and re.search(r"[A-Za-z]", fam):
            families.setdefault(fam, set()).update(weights)
    if not families:
        return ""
    jobs = [(fam, w) for fam, ws in families.items() for w in sorted(ws)]
    try:
        async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
            results = await asyncio.gather(
                *[_fetch_one_face(client, fam, w) for fam, w in jobs],
                return_exceptions=True,
            )
    except Exception:
        return ""
    faces = [r for r in results if isinstance(r, str)]
    if not faces:
        return ""
    return "<style>" + "".join(faces) + "</style>"


# ─── Topic image prefetch ─────────────────────────────────────────────────────
#
# Pollinations generates the image on request, so the timeout is generous but
# the fetches run concurrently: worst case adds ~25s once, not per image.

_IMAGE_STYLE = (" documentary editorial photograph, natural lighting, no text, "
                "no logo, high quality")


async def _fetch_single_image(client: httpx.AsyncClient, url: str, section) -> None:
    try:
        resp = await client.get(url)
        if resp.status_code == 200 and resp.content[:2] in (b"\xff\xd8", b"\x89P"):
            b64 = base64.b64encode(resp.content).decode("utf-8")
            mime = "image/png" if resp.content[:2] == b"\x89P" else "image/jpeg"
            section.image_data = f"data:{mime};base64,{b64}"
    except Exception:
        pass  # renderer falls back to a geometric panel


async def prefetch_images(spec: DocumentSpec, timeout: float = 25.0,
                          width: int = 900, height: int = 560) -> None:
    """Fill `section.image_data` for every section that requested an image."""
    tasks = []
    async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
        for section in spec.sections:
            kw = (section.image_keyword or "").strip()
            if kw:
                encoded = urllib.parse.quote(kw + _IMAGE_STYLE)
                url = (f"https://image.pollinations.ai/prompt/{encoded}"
                       f"?width={width}&height={height}&nologo=true")
                tasks.append(_fetch_single_image(client, url, section))
        if tasks:
            await asyncio.gather(*tasks, return_exceptions=True)


def image_bytes(section) -> Optional[bytes]:
    """Decode a section's data-URI image to raw bytes (for python-pptx)."""
    data = str(section.image_data or "")
    if not data.startswith("data:image/"):
        return None
    try:
        return base64.b64decode(data.split(",", 1)[1])
    except Exception:
        return None
