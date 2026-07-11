"""assets.py — remote asset acquisition (embedded fonts, topic images).

Everything here is time-boxed and fails soft: a slow CDN can never block or
break a generation, it only means the renderer falls back (system font stacks,
geometric SVG panels instead of photos).
"""

from __future__ import annotations

import asyncio
import base64
import os
import re
import urllib.parse
from typing import Dict, List, Optional

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
# Real photography first: a provider chain of stock/CC photo search APIs is
# tried per section — Pexels and Unsplash when API keys are configured, then
# the keyless Openverse and Wikimedia Commons — and only if every real-photo
# source fails does the AI image generator (Pollinations) run as last resort.
# Everything stays concurrent per section and fails soft.

_PEXELS_KEY = os.getenv("PEXELS_API_KEY")
_UNSPLASH_KEY = os.getenv("UNSPLASH_ACCESS_KEY")
_UA = {"User-Agent": "BotBot-DocStudio/1.0 (document generation service)"}

_IMAGE_STYLE = (" dramatic documentary editorial photograph, cinematic natural "
                "lighting, rich color depth, no text, no logo, high detail")

_JPEG_MAGIC, _PNG_MAGIC = b"\xff\xd8", b"\x89P"
_MIN_IMAGE_BYTES = 12_000          # skip icons / broken thumbs
_MAX_IMAGE_BYTES = 7_000_000       # keep the deck size sane


async def _pexels_urls(client: httpx.AsyncClient, kw: str) -> List[str]:
    if not _PEXELS_KEY:
        return []
    r = await client.get(
        "https://api.pexels.com/v1/search",
        params={"query": kw, "per_page": 4, "orientation": "landscape", "size": "large"},
        headers={"Authorization": _PEXELS_KEY, **_UA},
    )
    if r.status_code != 200:
        return []
    return [p["src"]["large2x"] for p in r.json().get("photos", [])
            if isinstance(p.get("src"), dict) and p["src"].get("large2x")]


async def _unsplash_urls(client: httpx.AsyncClient, kw: str) -> List[str]:
    if not _UNSPLASH_KEY:
        return []
    r = await client.get(
        "https://api.unsplash.com/search/photos",
        params={"query": kw, "per_page": 4, "orientation": "landscape"},
        headers={"Authorization": f"Client-ID {_UNSPLASH_KEY}", **_UA},
    )
    if r.status_code != 200:
        return []
    return [p["urls"]["regular"] for p in r.json().get("results", [])
            if isinstance(p.get("urls"), dict) and p["urls"].get("regular")]


async def _openverse_urls(client: httpx.AsyncClient, kw: str) -> List[str]:
    """Openverse aggregates CC-licensed photography; no API key required."""
    r = await client.get(
        "https://api.openverse.org/v1/images/",
        params={"q": kw, "page_size": 8, "license_type": "commercial",
                "category": "photograph", "mature": "false"},
        headers=_UA,
    )
    if r.status_code != 200:
        return []
    urls, thumbs = [], []
    for item in r.json().get("results", []):
        w, h = item.get("width") or 0, item.get("height") or 0
        if item.get("url") and w >= 900 and h and 0.9 <= w / h <= 2.4:
            urls.append(item["url"])
            if item.get("thumbnail"):
                thumbs.append(item["thumbnail"])
    # origin links can be stale (dead Flickr etc.); Openverse's own thumbnail
    # proxy is smaller but reliably up — cap origin tries so a thumbnail is
    # always reachable within the caller's candidate budget
    return urls[:2] + thumbs[:2]


async def _wikimedia_urls(client: httpx.AsyncClient, kw: str) -> List[str]:
    """Wikimedia Commons file search — keyless, reliable for real-world topics."""
    r = await client.get(
        "https://commons.wikimedia.org/w/api.php",
        params={
            "action": "query", "format": "json", "generator": "search",
            "gsrsearch": f"filetype:bitmap {kw}", "gsrnamespace": 6, "gsrlimit": 6,
            "prop": "imageinfo", "iiprop": "url|size|mime", "iiurlwidth": 1600,
        },
        headers=_UA,
    )
    if r.status_code != 200:
        return []
    pages = (r.json().get("query") or {}).get("pages") or {}
    ranked = sorted(pages.values(), key=lambda p: p.get("index", 99))
    urls = []
    for page in ranked:
        for info in page.get("imageinfo") or []:
            if info.get("mime") in ("image/jpeg", "image/png") and \
                    (info.get("width") or 0) >= 1000 and info.get("thumburl"):
                urls.append(info["thumburl"])
    return urls


def _pollinations_url(kw: str, width: int, height: int) -> str:
    encoded = urllib.parse.quote(kw + _IMAGE_STYLE)
    return (f"https://image.pollinations.ai/prompt/{encoded}"
            f"?width={width}&height={height}&nologo=true&enhance=true")


async def _download_image(client: httpx.AsyncClient, url: str,
                          timeout: Optional[float] = None) -> Optional[str]:
    """Fetch one candidate URL → data URI, or None if it isn't a usable photo."""
    try:
        resp = await client.get(url, headers=_UA,
                                timeout=timeout if timeout else client.timeout)
        content = resp.content
        if resp.status_code == 200 and content[:2] in (_JPEG_MAGIC, _PNG_MAGIC) \
                and _MIN_IMAGE_BYTES <= len(content) <= _MAX_IMAGE_BYTES:
            mime = "image/png" if content[:2] == _PNG_MAGIC else "image/jpeg"
            return f"data:{mime};base64,{base64.b64encode(content).decode('utf-8')}"
    except Exception:
        pass
    return None


async def _fetch_single_image(client: httpx.AsyncClient, section, kw: str,
                              width: int, height: int,
                              gen_sem: asyncio.Semaphore) -> None:
    """Curated stock first (keyed APIs return genuinely relevant photography),
    then AI generation (always on-topic and cinematic), and the keyless CC
    searches last — their relevance is a lottery (vintage drawings for
    industrial terms) and a wrong image is worse than a generated one."""
    for provider in (_pexels_urls, _unsplash_urls):
        try:
            candidates = await provider(client, kw)
        except Exception:
            continue
        for url in candidates[:3]:
            data = await _download_image(client, url)
            if data:
                section.image_data = data
                return
    # AI generation renders on request — long read timeout, and gated to a few
    # in flight at once: ten concurrent anonymous requests trip rate limiting
    async with gen_sem:
        data = await _download_image(client, _pollinations_url(kw, width, height),
                                     timeout=45)
    if data:
        section.image_data = data
        return
    for provider in (_openverse_urls, _wikimedia_urls):
        try:
            candidates = await provider(client, kw)
        except Exception:
            continue
        for url in candidates[:3]:
            data = await _download_image(client, url)
            if data:
                section.image_data = data
                return
    # else: renderer falls back to a geometric panel


async def prefetch_images(spec: DocumentSpec, timeout: float = 20.0,
                          width: int = 1600, height: int = 900) -> None:
    """Fill `section.image_data` for every section that requested an image.
    Per-section chains run concurrently and are individually time-boxed so a
    slow provider can never eat the generation budget."""
    gen_sem = asyncio.Semaphore(3)

    async def one(section, kw: str) -> None:
        try:
            await asyncio.wait_for(
                _fetch_single_image(client, section, kw, width, height, gen_sem),
                timeout=75)
        except Exception:
            pass

    async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
        tasks = [one(sec, sec.image_keyword.strip())
                 for sec in spec.sections if (sec.image_keyword or "").strip()]
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
