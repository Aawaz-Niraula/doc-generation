"""theme.py — the design system.

Everything visual that is a *decision* lives here, not in render code:

- Named palettes (dominant color carries 60–70% of visual weight, one accent).
- A typography scale (pt sizes) and metric-safe PPTX font pairs.
- WCAG contrast enforcement — colors that would produce unreadable text are
  corrected before any renderer sees them.
- Layout registries + the no-consecutive-repeat rule.
- Font-metric estimation used by the PPTX overflow guard and by QA.

Renderers receive a resolved `Theme` and never invent colors or sizes ad hoc.
"""

from __future__ import annotations

import math
import random
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from content_model import DesignSpec

# ─── Color math ───────────────────────────────────────────────────────────────

_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def is_hex(value: str) -> bool:
    return bool(_HEX_RE.match(str(value or "").strip()))


def hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = str(h).lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, int(round(v)))) for v in rgb))


def mix(a: str, b: str, t: float) -> str:
    """Linear blend from color a to color b, t in [0,1]."""
    ra, ga, ba = hex_to_rgb(a)
    rb, gb, bb = hex_to_rgb(b)
    return rgb_to_hex((ra + (rb - ra) * t, ga + (gb - ga) * t, ba + (bb - ba) * t))


def shade(h: str, f: float) -> str:
    """Darken (f<1) or lighten toward white-less scale (f>1)."""
    return rgb_to_hex(tuple(v * f for v in hex_to_rgb(h)))


def relative_luminance(h: str) -> float:
    def channel(c: float) -> float:
        c = c / 255.0
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = hex_to_rgb(h)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(fg: str, bg: str) -> float:
    l1, l2 = relative_luminance(fg), relative_luminance(bg)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def ensure_contrast(fg: str, bg: str, minimum: float = 4.5) -> str:
    """Nudge fg darker/lighter until it clears `minimum` against bg (WCAG AA)."""
    if contrast_ratio(fg, bg) >= minimum:
        return fg
    target = "#FFFFFF" if relative_luminance(bg) < 0.5 else "#0B1220"
    out = fg
    for _ in range(20):
        out = mix(out, target, 0.18)
        if contrast_ratio(out, bg) >= minimum:
            return out
    return target


def readable_on(bg: str) -> str:
    """White or near-black ink, whichever reads better on bg."""
    return "#FFFFFF" if contrast_ratio("#FFFFFF", bg) >= contrast_ratio("#0B1220", bg) else "#0B1220"


# ─── Typography scale (encoded once, consumed everywhere) ─────────────────────

@dataclass(frozen=True)
class TypeScale:
    cover_title: float = 44.0     # deck/report cover headline
    slide_title: float = 36.0     # per-slide titles (36–44 bold per spec)
    section_header: float = 22.0  # card headers / sub-heads (20–24 bold)
    body: float = 15.0            # body text (14–16)
    caption: float = 10.5         # captions / footers (10–12, muted)
    stat_hero: float = 64.0       # big-stat callouts (60–72)
    stat_card: float = 44.0       # stats inside a 3-up card row
    quote: float = 26.0


SCALE = TypeScale()

# Metric-safe fonts that ship with Microsoft Office and render consistently in
# LibreOffice / older installs. Serif display + sans body, mono for data.
# (Never Aptos — no reliable substitute on older installs.)
PPTX_FONT_PAIRS: List[Dict[str, str]] = [
    {"display": "Cambria", "heading": "Calibri", "body": "Calibri", "data": "Consolas"},
    {"display": "Cambria", "heading": "Arial",   "body": "Arial",   "data": "Consolas"},
    {"display": "Times New Roman", "heading": "Arial", "body": "Arial", "data": "Consolas"},
]

# ─── Palettes ─────────────────────────────────────────────────────────────────
#
# Each palette: ONE dominant color (60–70% of visual weight), a secondary tone
# for cards/panels, a supporting mid tone, one sharp accent, and derived light
# tints. `keywords` steer topic→theme matching so a finance deck and a wellness
# deck never look interchangeable.

@dataclass(frozen=True)
class Palette:
    name: str
    dominant: str       # deep base — covers, closing, dark panels
    secondary: str      # card / panel tone on dark
    support: str        # mid tone
    accent: str         # sharp accent, used sparingly
    keywords: Tuple[str, ...] = ()
    mood: str = ""

    @property
    def tint(self) -> str:      # light fill derived from dominant
        return mix(self.dominant, "#FFFFFF", 0.86)

    @property
    def wash(self) -> str:      # page background (committed white-ish, never beige)
        return mix(self.dominant, "#FFFFFF", 0.955)

    @property
    def ink(self) -> str:       # headline ink on light backgrounds
        return ensure_contrast(shade(self.dominant, 0.72), self.wash, 7.0)


PALETTES: Dict[str, Palette] = {p.name: p for p in [
    Palette("boardroom", "#0F2D5E", "#1A3F7A", "#5BA8D4", "#E8A33D",
            ("finance", "invest", "bank", "market", "revenue", "economy", "business", "strategy", "quarterly"),
            "authoritative navy with brass accent"),
    Palette("vitality",  "#1F5F4E", "#2A7A65", "#7BC4A8", "#F2762E",
            ("wellness", "health", "fitness", "nutrition", "mindful", "medical", "care", "sleep", "yoga"),
            "grounded green with citrus energy"),
    Palette("circuit",   "#161B33", "#232B52", "#6C7BD9", "#3EE0CF",
            ("ai", "software", "tech", "digital", "cyber", "data", "cloud", "robot", "quantum", "startup"),
            "midnight indigo with electric teal"),
    Palette("terra",     "#4A3428", "#6B4A36", "#C4906B", "#D96C2C",
            ("food", "coffee", "agriculture", "farm", "craft", "culinary", "restaurant", "wine", "travel"),
            "earthen umber with roasted orange"),
    Palette("scholar",   "#3B2A52", "#50396E", "#9B7FC0", "#E0B341",
            ("education", "learning", "university", "research", "history", "culture", "philosophy", "book"),
            "academic violet with gilded accent"),
    Palette("pulse",     "#5C1030", "#7A1B42", "#D4658C", "#FF4D4D",
            ("music", "fashion", "brand", "creative", "art", "media", "entertainment", "sport", "gaming"),
            "bold wine with signal red"),
    Palette("tide",      "#0C4A5E", "#136178", "#59B8CE", "#F4B942",
            ("ocean", "climate", "energy", "sustain", "environment", "water", "solar", "renewable", "green"),
            "deep sea teal with solar gold"),
    Palette("carbon",    "#1C2126", "#2A3138", "#8A97A3", "#E4572E",
            ("engineering", "manufactur", "automotive", "construction", "industrial", "logistics", "security"),
            "graphite with furnace ember"),
    Palette("dawn",      "#7A3B2E", "#96503F", "#E0A38A", "#2E86AB",
            ("community", "social", "nonprofit", "people", "hr", "team", "leadership", "psychology"),
            "warm clay with trusted blue"),
    Palette("noir",      "#101014", "#1D1D24", "#6E6E7A", "#C9A84C",
            ("luxury", "premium", "executive", "architecture", "real estate", "legal", "law"),
            "near-black with muted gold"),
]}

DEFAULT_PALETTE = "circuit"


def match_palette(topic: str, rng: random.Random) -> Palette:
    """Score palettes by topic keyword hits; tie-break deterministically random."""
    text = (topic or "").lower()
    best, best_score = None, 0
    for pal in PALETTES.values():
        score = sum(1 for kw in pal.keywords if kw in text)
        if score > best_score:
            best, best_score = pal, score
    if best is not None:
        return best
    names = sorted(PALETTES)
    return PALETTES[names[rng.randrange(len(names))]]


# ─── Resolved theme ───────────────────────────────────────────────────────────

@dataclass
class Theme:
    name: str
    # committed colors — every renderer uses only these
    dominant: str
    secondary: str
    support: str
    accent: str
    tint: str
    wash: str
    ink: str                     # headline ink on light
    body_ink: str                # body text on light
    on_dark: str = "#FFFFFF"
    on_dark_muted: str = "rgba(255,255,255,.72)"
    # typography
    scale: TypeScale = field(default_factory=lambda: SCALE)
    pdf_fonts: Dict[str, str] = field(default_factory=dict)    # embedded Google fonts
    pptx_fonts: Dict[str, str] = field(default_factory=dict)   # metric-safe Office fonts
    # identity
    aesthetic_label: str = "Editorial Intelligence"
    signature_element: str = "numbered accent dots"            # the ONE repeated motif
    dark_deck: bool = False       # True → fully dark premium deck; False → sandwich
    seed: int = 0

    def accent_readable_on_dark(self) -> str:
        return ensure_contrast(self.accent, self.dominant, 3.0)

    def accent_readable_on_light(self) -> str:
        return ensure_contrast(self.accent, self.wash, 3.0)


def _clean_font(name: str, fallback: str) -> str:
    name = str(name or "").strip()
    return name if re.search(r"[A-Za-z]", name) else fallback


def resolve_theme(design: DesignSpec, topic: str = "", requested: Optional[str] = None,
                  seed: Optional[int] = None) -> Theme:
    """Build the final Theme: requested palette > valid AI colors > topic match.

    AI-proposed colors are trusted for *identity* but every derived pairing is
    pushed through the contrast guard, so an unreadable suggestion can never
    reach a renderer.
    """
    seed = random.randrange(1 << 30) if seed is None else seed
    rng = random.Random(seed)

    base = None
    if requested and requested.lower() in PALETTES:
        base = PALETTES[requested.lower()]
    elif design.palette_name in PALETTES:
        base = PALETTES[design.palette_name]

    colors = design.colors or {}
    ai_dominant = colors.get("dominant", "")
    use_ai = base is None and is_hex(ai_dominant)

    if use_ai:
        dominant = ai_dominant.upper()
        # AI "dominant" must behave as a deep base; if it's light, darken it.
        if relative_luminance(dominant) > 0.45:
            dominant = shade(dominant, 0.42)
        secondary = colors.get("secondary", "")
        secondary = secondary.upper() if is_hex(secondary) else mix(dominant, "#FFFFFF", 0.12)
        if contrast_ratio(secondary, dominant) < 1.08:   # too close to read as a card
            secondary = mix(dominant, "#FFFFFF", 0.14)
        support = colors.get("supporting_1", "")
        support = support.upper() if is_hex(support) else mix(dominant, "#FFFFFF", 0.45)
        accent = colors.get("accent", "")
        accent = accent.upper() if is_hex(accent) else "#E8A33D"
        name = design.aesthetic_name or "ai-directed"
        tint = mix(dominant, "#FFFFFF", 0.86)
        wash = mix(dominant, "#FFFFFF", 0.955)
    else:
        pal = base or match_palette(topic, rng)
        dominant, secondary, support, accent = pal.dominant, pal.secondary, pal.support, pal.accent
        name, tint, wash = pal.name, pal.tint, pal.wash

    # accent must read against BOTH the dark base and the light wash
    accent = ensure_contrast(accent, dominant, 2.6)
    ink = ensure_contrast(shade(dominant, 0.68), wash, 7.0)
    body_ink = ensure_contrast("#374151", wash, 6.0)

    pdf_typo = design.typography or {}
    pdf_fonts = {
        "display": _clean_font(pdf_typo.get("h1_font") or pdf_typo.get("display_font"), "Playfair Display"),
        "heading": _clean_font(pdf_typo.get("h2_font") or pdf_typo.get("h1_font"), "DM Sans"),
        "body":    _clean_font(pdf_typo.get("body_font"), "Source Sans 3"),
        "data":    _clean_font(pdf_typo.get("data_font") or pdf_typo.get("body_font"), "JetBrains Mono"),
    }
    pptx_fonts = PPTX_FONT_PAIRS[rng.randrange(len(PPTX_FONT_PAIRS))]

    return Theme(
        name=name,
        dominant=dominant, secondary=secondary, support=support, accent=accent,
        tint=tint, wash=wash, ink=ink, body_ink=body_ink,
        pdf_fonts=pdf_fonts, pptx_fonts=pptx_fonts,
        aesthetic_label=design.aesthetic_label or "Editorial Intelligence",
        signature_element=design.signature_element or "numbered accent dots",
        dark_deck=rng.random() < 0.30,
        seed=seed,
    )


# ─── Layout registry ──────────────────────────────────────────────────────────
#
# Names every renderer must draw from. The planner in each renderer picks from
# these based on what content a section actually has; `no_consecutive_repeat`
# is the shared invariant.

PDF_LAYOUTS = ("cover", "feature", "split", "statgrid", "timeline", "cards", "quote", "closing")
PPTX_LAYOUTS = ("title", "stat", "two_col", "chart", "quote", "image_text", "cards", "closing")

# Slide geometry (inches) — 16:9
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.6            # ≥ 0.5" from every edge
GAP = 0.35              # consistent gap between blocks


def no_consecutive_repeat(assigned: List[str], options_per_item: List[List[str]],
                          rng: random.Random) -> List[str]:
    """Repair a layout assignment so no two consecutive items match, choosing
    replacements only from each item's own compatible options."""
    out = list(assigned)
    for i in range(1, len(out)):
        if out[i] == out[i - 1]:
            nxt = out[i + 1] if i + 1 < len(out) else None
            alts = [o for o in options_per_item[i] if o != out[i - 1] and o != nxt]
            if alts:
                out[i] = alts[rng.randrange(len(alts))]
    return out


# ─── Font-metric estimation (overflow guard + QA share this) ─────────────────
#
# Conservative average glyph widths (em) for the metric-safe families above.
# Estimation errs slightly wide so "fits by estimate" ⇒ fits on screen.

_AVG_CHAR_EM = 0.52
_AVG_CHAR_EM_BOLD = 0.56
_LINE_HEIGHT_EM = 1.28


def estimate_text_lines(text: str, width_in: float, size_pt: float, bold: bool = False) -> int:
    """Estimated wrapped line count for `text` in a box `width_in` wide."""
    if not text:
        return 0
    char_w_pt = size_pt * (_AVG_CHAR_EM_BOLD if bold else _AVG_CHAR_EM)
    chars_per_line = max(4, int((width_in * 72.0) / char_w_pt))
    lines = 0
    for paragraph in str(text).split("\n"):
        n = len(paragraph)
        lines += max(1, math.ceil(n / chars_per_line)) if n else 1
    return lines


def estimate_text_height_in(text: str, width_in: float, size_pt: float, bold: bool = False,
                            spacing_pt: float = 0.0) -> float:
    """Estimated rendered height (inches) of wrapped text."""
    lines = estimate_text_lines(text, width_in, size_pt, bold)
    return (lines * size_pt * _LINE_HEIGHT_EM + max(0, lines - 1) * spacing_pt) / 72.0


def fit_text_size(text: str, width_in: float, height_in: float, start_pt: float,
                  min_pt: Optional[float] = None, bold: bool = False) -> float:
    """The auto-fit guard: largest size ≤ start_pt whose estimated height fits.

    Never returns less than `min_pt` (default 55% of start); at the floor the
    caller must clip/split content instead of shipping overflow.
    """
    floor = min_pt if min_pt is not None else max(8.0, start_pt * 0.55)
    size = start_pt
    while size > floor and estimate_text_height_in(text, width_in, size, bold) > height_in:
        size -= 1.0
    return size


def clip_words(value: str, limit: int) -> str:
    words = str(value or "").split()
    if len(words) <= limit:
        return " ".join(words)
    return " ".join(words[:limit]).rstrip(".,;:") + "."


def clip_chars(value: str, limit: int) -> str:
    text = " ".join(str(value or "").split())
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0].rstrip(".,;:") + "."
