"""render_pdf.py — premium editorial PDF renderer (WeasyPrint, HTML → PDF).

Consumes the same DocumentSpec as render_pptx / render_docx. WeasyPrint is the
engine (not raw canvas) so text flows, wraps, and paginates; real fonts arrive
pre-fetched as base64 @font-face CSS (assets.build_embedded_font_css) so the
AI-chosen typography actually renders.

Design-system rules enforced here:
- no accent rules under titles, no edge stripes, no single-edge card borders —
  hierarchy comes from whitespace, background tints, and soft shadows;
- one repeated motif: the numbered accent circle;
- every page carries at least one visual (chart, table, stat, photo, or
  geometric panel);
- clickable outline via bookmark-level CSS; PDF metadata via <title>/<meta>.

WeasyPrint-safe CSS only: block/float/table layout, no CSS Grid.
"""

from __future__ import annotations

import math
import random
from html import escape
from typing import Optional

from content_model import ChartBlock, DocumentSpec, Section, Stat, TableBlock
from theme import Theme, clip_chars, clip_words, mix, no_consecutive_repeat, readable_on

PAGE_SIZES = {"a4": ("A4", "210mm", "297mm"), "letter": ("Letter", "215.9mm", "279.4mm")}


# ─── Visual blocks ────────────────────────────────────────────────────────────

def _chart_html(chart: Optional[ChartBlock], theme: Theme, title: str = "", big: bool = False) -> str:
    if not chart or len(chart.points) < 2:
        return ""
    data = chart.points
    max_val = max((v for _, v in data), default=1) or 1
    rows = []
    for label, val in data:
        width = max(7, min(100, (val / max_val) * 100))
        peak = " peak" if val == max_val else ""
        rows.append(f"""
<div class="chart-row">
  <div class="chart-label">{escape(clip_chars(label, 22))}</div>
  <div class="chart-track"><div class="chart-bar{peak}" style="width:{width:.1f}%"></div></div>
  <div class="chart-value">{escape(f'{val:g}')}</div>
</div>""")
    delta = ""
    if len(data) >= 2 and data[0][1]:
        try:
            chg = (data[-1][1] - data[0][1]) / abs(data[0][1]) * 100
            arrow = "▲" if chg >= 0 else "▼"
            delta = f'<span class="chart-delta">{arrow} {abs(chg):.0f}%</span>'
        except Exception:
            delta = ""
    cls = "premium-chart big" if big else "premium-chart"
    return f"""
<div class="{cls}">
  <div class="chart-head"><span class="chart-kicker">{escape(clip_chars(chart.title or title or 'Measured signal', 40))}</span>{delta}</div>
  <div class="chart-rows">{"".join(rows)}</div>
</div>"""


def _hero_stat_html(stat: Optional[Stat]) -> str:
    if not stat or not stat.value:
        return ""
    return f"""
<div class="premium-stat">
  <div class="premium-stat-value">{escape(stat.value)}</div>
  <div class="premium-stat-label">{escape(clip_chars(stat.label or "Key metric", 42))}</div>
  <p>{escape(clip_words(stat.context, 26))}</p>
</div>"""


def _table_html(table: Optional[TableBlock]) -> str:
    if not table or table.is_empty():
        return ""
    hdr = "".join(f"<th>{escape(clip_chars(h, 22))}</th>" for h in table.headers[:5])
    body = ""
    for row in table.rows[:7]:
        cells = "".join(
            f"<td>{escape(clip_chars(c, 34))}</td>"
            for c in (list(row) + [""] * len(table.headers))[:len(table.headers[:5])]
        )
        body += f"<tr>{cells}</tr>"
    return (f'<div class="hero-table"><table class="premium-table">'
            f'<thead><tr>{hdr}</tr></thead><tbody>{body}</tbody></table></div>')


def _big_stat_card(stat: Stat, index: int, compact: bool = False) -> str:
    """Tall stat tile with an oversized figure — for stat-grid pages."""
    return f"""
<div class="bigstat{' compact' if compact else ''}">
  <span class="n-dot">{index + 1:02d}</span>
  <div class="bigstat-val">{escape(stat.value or "—")}</div>
  <div class="bigstat-lbl">{escape(clip_chars(stat.label or "Metric", 24))}</div>
  <div class="bigstat-ctx">{escape(clip_words(stat.context, 18))}</div>
</div>"""


def _insight_cards(items: list, limit: int = 3, hero: bool = False) -> str:
    """Equal-width insight cards (table-cell layout, never wraps). The numbered
    accent circle is the document's repeated motif."""
    items = [str(h) for h in (items or []) if str(h).strip()][:limit]
    if not items:
        return ""
    cells = "".join(
        f'<div class="cell"><div class="ins-card"><span class="n-dot">{i + 1:02d}</span>'
        f'<span class="ins-txt">{escape(clip_words(str(h), 26 if hero else 20))}</span></div></div>'
        for i, h in enumerate(items)
    )
    cls = "trow ins-row cards-hero" if hero else "trow ins-row"
    return f'<div class="{cls}">{cells}</div>'


def _duo_cards(section: Section) -> str:
    """Side-by-side comparison columns (left/right points) as two cards."""
    if not (section.left and section.right and section.left.points and section.right.points):
        return ""
    cols = []
    for i, col in enumerate((section.left, section.right)):
        title = escape(clip_chars(col.title or ("Perspective A" if i == 0 else "Perspective B"), 34))
        pts = "".join(f'<div class="insight">{escape(clip_words(p, 20))}</div>'
                      for p in col.points[:3])
        cols.append(f'<div class="cell"><div class="duo-card{" alt" if i else ""}">'
                    f'<div class="duo-title">{title.upper()}</div>{pts}</div></div>')
    return f'<div class="trow duo-row">{"".join(cols)}</div>'


def _geo_panel(theme: Theme, idx: int, label: str, tall: bool = False,
               variant: Optional[int] = None) -> str:
    """Clean, art-directed geometric data-visual for pages without real data
    visuals. Tall columns use crop-safe centered compositions; wide areas cycle
    through four distinct data-shaped motifs."""
    uid = f"gp{idx}"
    acc, mid, prim, card2 = theme.accent, theme.support, theme.dominant, theme.secondary
    ink = "rgba(255,255,255,.9)"
    faint = "rgba(255,255,255,.14)"
    rng = random.Random(idx * 7 + 3)

    if tall:
        W, H = 380, 750
        variant = (idx if variant is None else variant) % 2
        cx, cy = W / 2, H / 2
        if variant == 0:  # concentric arcs
            rings = "".join(
                f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{r}" fill="none" stroke="{acc}" stroke-width="{w}" opacity="{o}" stroke-dasharray="{d}" stroke-linecap="round"/>'
                for r, w, o, d in [(160, "2", ".16", "0"), (128, "16", ".26", f"{rng.randint(220, 420)} 1200"),
                                   (96, "12", ".42", f"{rng.randint(150, 300)} 1200"), (64, "9", ".62", f"{rng.randint(90, 220)} 1200")])
            core = (f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="30" fill="{acc}" opacity=".9"/>'
                    f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="15" fill="#fff" opacity=".9"/>')
            ticks = "".join(
                f'<circle cx="{cx + 200 * math.cos(math.radians(a)):.0f}" cy="{cy + 200 * math.sin(math.radians(a)):.0f}" r="3" fill="{mid}" opacity=".5"/>'
                for a in range(0, 360, 30))
            art = ticks + rings + core
        else:  # stacked horizontal bars
            n = 7
            gap = (H - 160) / n
            bars = ""
            for i in range(n):
                bw = rng.randint(120, W - 90)
                y = 100 + i * gap
                c = acc if i % 3 == 0 else mid
                bars += f'<rect x="60" y="{y:.0f}" width="{bw}" height="{gap * 0.5:.0f}" rx="6" fill="{c}" opacity="{0.45 + i * 0.06:.2f}"/>'
            art = bars
    else:
        W, H = 600, 335
        variant = ((idx - 1) if variant is None else variant) % 4

    if not tall and variant == 0:  # ascending area + line
        n = 8
        ys = []
        y = H - 90
        for i in range(n):
            y = max(70, min(H - 70, y - rng.randint(-8, 46)))
            ys.append(y)
        xs = [60 + i * ((W - 120) / (n - 1)) for i in range(n)]
        line = "M " + " L ".join(f"{x:.0f},{y:.0f}" for x, y in zip(xs, ys))
        area = f"M {xs[0]:.0f},{H - 60} L " + " L ".join(f"{x:.0f},{y:.0f}" for x, y in zip(xs, ys)) + f" L {xs[-1]:.0f},{H - 60} Z"
        grid = "".join(f'<line x1="60" y1="{H - 60 - i * ((H - 130) / 4):.0f}" x2="{W - 40}" y2="{H - 60 - i * ((H - 130) / 4):.0f}" stroke="{faint}" stroke-width="1"/>' for i in range(5))
        dots = "".join(f'<circle cx="{x:.0f}" cy="{y:.0f}" r="5" fill="{acc}" stroke="#fff" stroke-width="2"/>' for x, y in zip(xs, ys))
        art = (f'<defs><linearGradient id="{uid}a" x1="0" y1="0" x2="0" y2="1">'
               f'<stop offset="0%" stop-color="{acc}" stop-opacity=".38"/>'
               f'<stop offset="100%" stop-color="{acc}" stop-opacity="0"/></linearGradient></defs>'
               f'{grid}<path d="{area}" fill="url(#{uid}a)"/>'
               f'<path d="{line}" fill="none" stroke="{acc}" stroke-width="3.5" stroke-linejoin="round" stroke-linecap="round"/>{dots}')
    elif not tall and variant == 1:  # column bars
        n = 7
        bw = (W - 130) / n
        bars = ""
        for i in range(n):
            bh = rng.randint(70, H - 130)
            x = 65 + i * bw
            y = H - 60 - bh
            c = acc if i % 3 == 0 else mid
            bars += f'<rect x="{x:.0f}" y="{y:.0f}" width="{bw * 0.62:.0f}" height="{bh}" rx="5" fill="{c}" opacity="{0.55 + i * 0.05:.2f}"/>'
        grid = "".join(f'<line x1="55" y1="{H - 60 - i * ((H - 130) / 4):.0f}" x2="{W - 40}" y2="{H - 60 - i * ((H - 130) / 4):.0f}" stroke="{faint}" stroke-width="1"/>' for i in range(5))
        art = f'{grid}{bars}<line x1="55" y1="{H - 60}" x2="{W - 40}" y2="{H - 60}" stroke="{faint}" stroke-width="1.5"/>'
    elif not tall and variant == 2:  # concentric arcs / radial
        cx, cy = W / 2, H / 2 + 6
        rings = "".join(
            f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{r}" fill="none" stroke="{acc}" stroke-width="{w}" opacity="{o}" stroke-dasharray="{d}" stroke-linecap="round"/>'
            for r, w, o, d in [(150, "2", ".18", "0"), (120, "14", ".28", f"{rng.randint(180, 320)} 900"),
                               (92, "11", ".42", f"{rng.randint(120, 240)} 900"), (64, "8", ".6", f"{rng.randint(80, 180)} 900")])
        core = (f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="30" fill="{acc}" opacity=".9"/>'
                f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="15" fill="#fff" opacity=".85"/>')
        art = rings + core
    elif not tall:  # node network
        nodes = [(120, 110, 15), (260, 70, 11), (420, 120, 17), (510, 210, 12),
                 (360, 250, 14), (180, 250, 10), (300, 165, 20), (470, 60, 9)]
        edges = [(0, 6), (1, 6), (2, 6), (2, 3), (3, 4), (4, 5), (5, 0), (6, 4), (1, 2), (2, 7)]
        lines = "".join(f'<line x1="{nodes[a][0]}" y1="{nodes[a][1]}" x2="{nodes[b][0]}" y2="{nodes[b][1]}" stroke="{faint}" stroke-width="1.4"/>' for a, b in edges)
        dots = "".join(
            f'<circle cx="{x}" cy="{y}" r="{r}" fill="{acc if i % 2 == 0 else mid}" opacity="{0.5 + i * 0.05:.2f}"/>'
            f'<circle cx="{x}" cy="{y}" r="{max(3, r - 6)}" fill="#fff" opacity=".55"/>'
            for i, (x, y, r) in enumerate(nodes))
        art = lines + dots

    bg = (f'<defs><linearGradient id="{uid}bg" x1="0" y1="0" x2="1" y2="1">'
          f'<stop offset="0%" stop-color="{prim}"/><stop offset="100%" stop-color="{card2}"/></linearGradient></defs>'
          f'<rect width="{W}" height="{H}" rx="20" fill="url(#{uid}bg)"/>')
    lab = (f'<circle cx="49" cy="44" r="15" fill="{acc}"/>'
           f'<text x="49" y="49" fill="{readable_on(acc)}" font-family="monospace" font-size="13" font-weight="700" text-anchor="middle">{idx:02d}</text>'
           f'<text x="34" y="{H - 34}" fill="{ink}" font-family="monospace" font-size="16" font-weight="700" letter-spacing="2">{escape(clip_chars(label, 34)).upper()}</text>')
    return (f'<div class="geo-panel"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" '
            f'preserveAspectRatio="xMidYMid slice">{bg}{art}{lab}</svg></div>')


def _photo_html(section: Section, caption: str, big: bool = False) -> str:
    if not section.image_data:
        return ""
    cls = "photo-plate big" if big else "photo-plate"
    return (f'<div class="{cls}"><img src="{section.image_data}" alt="" />'
            f'<div class="photo-caption">{escape(clip_chars(caption, 70 if big else 60))}</div></div>')


def _hero_visual(section: Section, theme: Theme, idx: int, heading: str) -> str:
    """The page-filling centerpiece: chart > table > hero-stat > photo > geometric."""
    return (_chart_html(section.chart, theme, heading, big=True)
            or _table_html(section.table)
            or _hero_stat_html(section.hero_stat)
            or _photo_html(section, heading, big=True)
            or _geo_panel(theme, idx, heading, variant=section.geo_seed))


def _side_visual(section: Section, theme: Theme, idx: int, heading: str) -> str:
    """A visual sized for the narrower side column."""
    return (_hero_stat_html(section.hero_stat)
            or _table_html(section.table)
            or _chart_html(section.chart, theme, heading)
            or _photo_html(section, heading)
            or _geo_panel(theme, idx, heading, tall=True, variant=section.geo_seed))


def _backdrop(theme: Theme, idx: int, dark: bool = False) -> str:
    ink = "rgba(255,255,255,.42)" if dark else theme.dominant
    accent = theme.accent
    return f"""
<svg class="premium-backdrop" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" preserveAspectRatio="none">
  <defs>
    <linearGradient id="premiumGrad{idx}" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="{theme.dominant}" stop-opacity="{'.97' if dark else '.06'}"/>
      <stop offset="62%" stop-color="{theme.secondary}" stop-opacity="{'.85' if dark else '.03'}"/>
      <stop offset="100%" stop-color="{accent}" stop-opacity="{'.5' if dark else '.09'}"/>
    </linearGradient>
  </defs>
  <rect width="595" height="842" fill="url(#premiumGrad{idx})"/>
  <path d="M-40 720 C130 590 220 735 380 592 C492 492 548 552 650 414" fill="none" stroke="{accent}" stroke-width="2.5" opacity=".32"/>
  <path d="M-50 112 C120 168 206 52 360 118 C485 172 535 112 650 42" fill="none" stroke="{ink}" stroke-width=".8" opacity=".16"/>
  <circle cx="512" cy="118" r="168" fill="none" stroke="{accent}" stroke-width="1" opacity=".16"/>
  <circle cx="78" cy="752" r="236" fill="{accent}" opacity="{'.10' if dark else '.05'}"/>
</svg>"""


# ─── Layout planning ──────────────────────────────────────────────────────────

def plan_pdf_layouts(sections: list, seed: int) -> None:
    """Assign each interior page a *content-faithful* archetype (nothing the
    model produced is dropped), with per-run variety from mirror flips, geometry
    seeds, and feature↔split alternation — then repair consecutive repeats."""
    rng = random.Random(seed)
    n = len(sections)

    def compatible(sec: Section) -> list:
        opts = []
        if sec.steps:
            opts.append("timeline")
        if len(sec.stats) >= 2:
            opts.append("statgrid")
        if sec.kind == "case_study" or sec.attribution or (sec.callout and not sec.body):
            opts.append("quote")
        if sec.kind in ("future",) and (sec.callout or sec.statements):
            opts.append("quote")
        if len(sec.tiles) >= 3 or (len(sec.highlights) >= 3 and not sec.body):
            opts.append("cards")
        if sec.has_data_visual() or sec.image_data or sec.body:
            opts.extend(["feature", "split"])
        return opts or ["split"]

    options, assigned = [], []
    for i, sec in enumerate(sections):
        sec.geo_seed = rng.randrange(4)
        sec.mirror = rng.random() < 0.5
        if i == 0 or sec.kind == "cover":
            options.append(["cover"]); assigned.append("cover"); continue
        if i == n - 1 or sec.kind == "closing":
            options.append(["closing"]); assigned.append("closing"); continue
        opts = compatible(sec)
        choice = opts[0]
        # feature and split both fully present a data+narrative page — alternate.
        if choice in ("feature", "split") and sec.body and sec.has_data_visual():
            prev = assigned[-1] if assigned else ""
            if prev in ("feature", "split"):
                choice = "split" if prev == "feature" else "feature"
            elif rng.random() < 0.42:
                choice = "split" if choice == "feature" else "feature"
        options.append(opts)
        assigned.append(choice)

    for sec, layout in zip(sections, no_consecutive_repeat(assigned, options, rng)):
        sec.layout = layout


# ─── Stylesheet ───────────────────────────────────────────────────────────────

def _stylesheet(theme: Theme, page_w: str, page_h: str, page_kw: str) -> str:
    T = theme
    on_acc = readable_on(T.accent)
    quote_tint = mix(T.dominant, "#FFFFFF", 0.92)
    return f"""
@page {{ size: {page_kw}; margin: 0; }}
* {{ box-sizing: border-box; }}
html, body {{ margin:0; padding:0; font-family:var(--font-body); color:var(--text); }}
:root {{
  --primary:{T.dominant}; --accent:{T.accent}; --mid:{T.support}; --light:{T.tint};
  --wash:{T.wash}; --text:{T.ink}; --dark:{T.dominant}; --body:{T.body_ink};
  --card2:{T.secondary}; --muted:{T.support};
  --font-display:'{T.pdf_fonts["display"]}', Georgia, 'Times New Roman', serif;
  --font-heading:'{T.pdf_fonts["heading"]}', 'Helvetica Neue', Arial, sans-serif;
  --font-body:'{T.pdf_fonts["body"]}', 'Helvetica Neue', Arial, sans-serif;
  --font-data:'{T.pdf_fonts["data"]}', 'SF Mono', 'Courier New', monospace;
}}
.page {{ width:{page_w}; height:{page_h}; page-break-after:always; break-after:page; position:relative; overflow:hidden; background:var(--wash); display:block; }}
.page:last-child {{ page-break-after:auto; break-after:auto; }}
.premium-backdrop {{ position:absolute; inset:0; width:{page_w}; height:{page_h}; z-index:0; display:block; }}
.content {{ position:relative; z-index:1; display:block; }}
.kicker {{ font-family:var(--font-data); font-size:7.5pt; letter-spacing:.28em; text-transform:uppercase; font-weight:700; color:var(--accent); display:block; }}
h1,h2,h3 {{ margin:0; font-family:var(--font-display); letter-spacing:-.02em; line-height:1.03; display:block; font-weight:800; }}
.cover h1 {{ bookmark-level:1; bookmark-label:content(text); }}
.doc-head h2, .dark-page h2 {{ bookmark-level:1; bookmark-label:content(text); }}
p {{ margin:0; color:var(--body); font-size:10pt; line-height:1.62; display:block; }}

/* ── the repeated motif: numbered accent circle ── */
.n-dot {{ display:block; width:8mm; height:8mm; border-radius:50%; background:var(--accent); color:{on_acc}; text-align:center; font:800 8.5pt/8mm var(--font-data); letter-spacing:0; }}

/* ── running foot (whitespace-separated, no rule) ── */
.folio {{ position:absolute; left:20mm; right:20mm; bottom:11mm; z-index:2; display:block; overflow:hidden; font:7pt var(--font-data); color:rgba(0,0,0,.40); letter-spacing:.14em; text-transform:uppercase; }}
.folio.dark {{ color:rgba(255,255,255,.5); }}
.folio span {{ float:left; display:block; max-width:72%; overflow:hidden; }}
.folio b {{ float:right; display:block; color:var(--accent); font-size:8pt; letter-spacing:.1em; }}

/* ── shared header block (hierarchy by whitespace, no accent rule) ── */
.doc {{ padding:19mm 20mm 20mm; display:block; }}
.doc-head {{ display:block; margin-bottom:9mm; }}
.doc-head .num {{ display:inline-block; font:800 12pt var(--font-data); color:var(--accent); letter-spacing:-.02em; margin-right:4mm; vertical-align:middle; }}
.doc-head h2 {{ font-size:32pt; color:var(--primary); max-width:172mm; margin:4.5mm 0 0; }}
.lede {{ max-width:158mm; margin-top:6mm; font-size:11pt; line-height:1.62; color:var(--body); }}

/* ── utility table rows (equal cols, never wrap) ── */
.trow {{ display:table; width:100%; table-layout:fixed; border-collapse:separate; }}
.trow > .cell {{ display:table-cell; vertical-align:top; padding-right:5mm; }}
.trow > .cell:last-child {{ padding-right:0; }}

/* ── cover ── */
.cover {{ background:var(--dark); color:#fff; }}
.cover .content {{ padding:24mm 24mm; height:{page_h}; position:relative; display:block; }}
.cover-top {{ margin-bottom:12mm; display:block; }}
.cover-pill {{ display:inline-block; border:1px solid rgba(255,255,255,.24); border-radius:999px; padding:2.4mm 7mm; background:rgba(255,255,255,.07); margin-right:4mm; font:7pt var(--font-data); letter-spacing:.24em; text-transform:uppercase; color:rgba(255,255,255,.78); vertical-align:middle; }}
.cover-body {{ display:block; min-height:118mm; }}
.cover h1 {{ max-width:168mm; font-size:54pt; color:#fff; line-height:1.0; }}
.cover-sub {{ max-width:140mm; margin-top:10mm; color:rgba(255,255,255,.8); font-size:14pt; line-height:1.5; font-family:var(--font-heading); }}
.cover-callout {{ max-width:130mm; color:#fff; font:italic 15pt var(--font-display); line-height:1.4; }}
.cover-meta {{ position:absolute; left:24mm; right:24mm; bottom:22mm; display:block; overflow:hidden; }}
.cover-meta > div:first-child {{ float:left; max-width:calc(100% - 56mm); display:block; }}
.cover-num {{ float:right; display:block; font:800 78pt var(--font-data); color:rgba(255,255,255,.12); letter-spacing:-.08em; line-height:.8; }}
.cover-author {{ margin-top:8mm; color:rgba(255,255,255,.5); font:7pt var(--font-data); letter-spacing:.2em; text-transform:uppercase; }}

/* ── narrative / body ── */
.body-col p + p {{ margin-top:4mm; }}
.pullquote {{ margin-top:8mm; padding:6mm 7mm; background:{quote_tint}; border-radius:3mm; font:italic 14pt var(--font-display); line-height:1.42; color:var(--primary); display:block; box-shadow:0 8px 22px rgba(15,23,42,.06); }}
.dark-page .pullquote {{ background:rgba(255,255,255,.07); color:#fff; box-shadow:none; }}
.insights {{ margin-top:7mm; display:block; }}
.insight {{ display:block; position:relative; padding:3.6mm 0 3.6mm 7mm; font-size:9.4pt; line-height:1.42; color:var(--body); }}
.insight:before {{ content:""; position:absolute; left:0; top:4.6mm; width:3mm; height:3mm; border-radius:50%; background:var(--accent); }}
.dark-page .insight {{ color:rgba(255,255,255,.78); }}

/* ── insight cards (tint + shadow, no edge stripes) ── */
.ins-row {{ margin-top:8mm; }}
.ins-card {{ display:block; background:#fff; border-radius:3mm; padding:5.5mm 5mm; min-height:33mm; box-shadow:0 10px 26px rgba(15,23,42,.09); }}
.ins-card .n-dot {{ margin-bottom:4mm; }}
.ins-txt {{ display:block; color:var(--body); font:9.6pt/1.5 var(--font-body); }}
.dark-page .ins-card {{ background:rgba(255,255,255,.06); box-shadow:none; }}
.dark-page .ins-txt {{ color:rgba(255,255,255,.82); }}
.cards-hero {{ margin-top:11mm; }}
.cards-hero .ins-card {{ min-height:130mm; padding:9mm 7mm; }}
.cards-hero .n-dot {{ width:11mm; height:11mm; font-size:11pt; line-height:11mm; margin-bottom:8mm; }}
.cards-hero .ins-txt {{ font-size:11pt; line-height:1.55; }}

/* ── two-column split ── */
.split {{ margin-top:2mm; display:table; width:100%; table-layout:fixed; }}
.split .main-col {{ display:table-cell; vertical-align:top; }}
.split .side-col {{ display:table-cell; vertical-align:top; width:78mm; }}
.split .gut {{ padding-right:9mm; }}

/* ── hero stat card ── */
.premium-stat {{ display:block; background:linear-gradient(140deg,var(--primary),var(--card2) 60%,var(--accent)); color:#fff; padding:11mm; border-radius:3.5mm; box-shadow:0 16px 40px rgba(15,23,42,.24); min-height:98mm; position:relative; overflow:hidden; }}
.premium-stat:after {{ content:""; position:absolute; right:-30mm; top:-30mm; width:100mm; height:100mm; border-radius:50%; border:1.2mm solid rgba(255,255,255,.12); }}
.premium-stat-value {{ display:block; margin-top:20mm; font:800 68pt var(--font-data); letter-spacing:-.05em; line-height:.9; }}
.premium-stat-label {{ display:block; margin-top:6mm; font:800 9pt var(--font-data); letter-spacing:.2em; text-transform:uppercase; color:rgba(255,255,255,.82); }}
.premium-stat p {{ margin-top:5mm; max-width:120mm; color:rgba(255,255,255,.76); font-size:10pt; line-height:1.55; }}
.side-col .premium-stat {{ min-height:150mm; padding:10mm; }}
.side-col .premium-stat-value {{ margin-top:58mm; font-size:52pt; }}

/* ── chart ── */
.premium-chart {{ display:block; background:linear-gradient(160deg,#14161c,#0c0d11); color:#fff; padding:9mm 11mm; border-radius:3.5mm; box-shadow:0 16px 40px rgba(15,23,42,.22); }}
.premium-chart.big {{ min-height:90mm; }}
.side-col .premium-chart {{ min-height:150mm; }}
.chart-head {{ display:block; overflow:hidden; margin-bottom:7mm; padding-bottom:4mm; }}
.chart-kicker {{ float:left; font:800 8.5pt var(--font-data); letter-spacing:.16em; text-transform:uppercase; color:rgba(255,255,255,.82); max-width:70%; }}
.chart-delta {{ float:right; font:800 8.5pt var(--font-data); color:var(--accent); letter-spacing:.06em; }}
.chart-rows {{ display:block; }}
.chart-row {{ display:table; width:100%; table-layout:fixed; margin:5.2mm 0; border-collapse:separate; border-spacing:0; }}
.chart-label {{ display:table-cell; width:30mm; vertical-align:middle; padding-right:3mm; font:8pt var(--font-data); color:rgba(255,255,255,.72); }}
.chart-track {{ display:table-cell; vertical-align:middle; height:6mm; background:rgba(255,255,255,.08); border-radius:999px; overflow:hidden; }}
.chart-value {{ display:table-cell; width:16mm; vertical-align:middle; text-align:right; font:9pt var(--font-data); color:#fff; font-weight:800; }}
.chart-bar {{ display:block; height:6mm; background:linear-gradient(90deg,var(--mid),var(--accent)); border-radius:999px; }}
.chart-bar.peak {{ background:linear-gradient(90deg,var(--accent),#fff); }}

/* ── table ── */
.hero-table {{ display:block; border-radius:3.5mm; overflow:hidden; box-shadow:0 14px 34px rgba(15,23,42,.12); }}
.premium-table {{ display:table; width:100%; border-collapse:collapse; overflow:hidden; font-size:9pt; background:#fff; }}
.premium-table thead {{ display:table-header-group; }}
.premium-table tbody {{ display:table-row-group; }}
.premium-table tr {{ display:table-row; }}
.premium-table th {{ display:table-cell; background:var(--primary); color:#fff; text-align:left; padding:4.5mm 4mm; font:800 7.5pt var(--font-data); letter-spacing:.1em; text-transform:uppercase; }}
.premium-table td {{ display:table-cell; padding:4.2mm 4mm; color:{mix(theme.ink, "#FFFFFF", 0.12)}; }}
.premium-table tbody tr:first-child td {{ font-weight:700; }}
.premium-table tr:nth-child(even) td {{ background:var(--wash); }}

/* ── photo ── */
.photo-plate {{ display:block; margin:0; border-radius:3.5mm; overflow:hidden; position:relative; box-shadow:0 16px 40px rgba(15,23,42,.22); background:#111; }}
.photo-plate img {{ width:100%; height:96mm; object-fit:cover; display:block; }}
.side-col .photo-plate img {{ height:150mm; }}
.photo-plate:after {{ content:""; position:absolute; inset:0; background:linear-gradient(180deg,transparent 50%,rgba(0,0,0,.5)); }}
.photo-caption {{ position:absolute; left:6mm; right:6mm; bottom:5mm; z-index:2; color:#fff; font:7.5pt var(--font-data); letter-spacing:.14em; text-transform:uppercase; display:block; }}

/* ── geometric data panel ── */
.geo-panel {{ display:block; border-radius:3.5mm; overflow:hidden; box-shadow:0 16px 40px rgba(15,23,42,.2); line-height:0; }}
.geo-panel svg {{ display:block; width:100%; height:96mm; }}
.side-col .geo-panel svg {{ height:150mm; }}

/* ── big stat grid ── */
.stat-row {{ margin-top:9mm; }}
.bigstat {{ display:block; background:#fff; border-radius:3.5mm; padding:9mm 7mm; min-height:150mm; position:relative; overflow:hidden; box-shadow:0 12px 30px rgba(15,23,42,.10); }}
.bigstat-val {{ display:block; margin-top:44mm; font:800 44pt var(--font-display); color:var(--primary); letter-spacing:-.03em; line-height:.92; }}
.bigstat-lbl {{ display:block; margin-top:6mm; font:800 8pt var(--font-data); letter-spacing:.16em; text-transform:uppercase; color:var(--accent); }}
.bigstat-ctx {{ display:block; margin-top:5mm; font:9.6pt/1.55 var(--font-body); color:var(--body); }}

.bigstat.compact {{ min-height:82mm; }}
.bigstat.compact .bigstat-val {{ margin-top:10mm; font-size:36pt; }}

/* ── comparison duo cards ── */
.duo-row {{ margin-top:8mm; }}
.duo-card {{ display:block; background:#fff; border-radius:3.5mm; padding:7mm 6mm; min-height:52mm; box-shadow:0 12px 30px rgba(15,23,42,.09); }}
.duo-card.alt {{ background:linear-gradient(150deg,var(--primary),var(--card2)); }}
.duo-title {{ display:block; font:800 9pt var(--font-data); letter-spacing:.16em; text-transform:uppercase; color:var(--accent); margin-bottom:3mm; }}
.duo-card.alt .insight {{ color:rgba(255,255,255,.84); }}

/* ── timeline ── */
.tl-row {{ margin-top:9mm; }}
.tl-card {{ display:block; background:#fff; border-radius:3.5mm; padding:8mm 6mm; min-height:148mm; position:relative; box-shadow:0 12px 30px rgba(15,23,42,.09); }}
.tl-card .n-dot {{ margin-bottom:6mm; }}
.tl-date {{ display:block; font:800 12pt var(--font-data); color:var(--primary); letter-spacing:-.02em; }}
.tl-step {{ display:block; margin-top:2mm; font:800 8pt var(--font-data); letter-spacing:.14em; text-transform:uppercase; color:var(--accent); }}
.tl-desc {{ display:block; margin-top:4mm; font:9.6pt/1.55 var(--font-body); color:var(--body); }}

/* ── dark feature (quote / manifesto) ── */
.dark-page {{ background:var(--dark); color:#fff; }}
.dark-page .content {{ padding:22mm 22mm 20mm; display:block; }}
.dark-page h2 {{ color:#fff; }}
.dark-page p {{ color:rgba(255,255,255,.76); }}
.dark-page .kicker {{ color:{mix(theme.support, "#FFFFFF", 0.4)}; }}
.dark-page .doc-head h2 {{ color:#fff; }}
.giant-quote {{ display:block; margin:14mm 0 0; font:italic 33pt/1.22 var(--font-display); color:#fff; max-width:170mm; letter-spacing:-.01em; }}
.giant-quote .mark {{ font-size:52pt; color:var(--accent); line-height:0; }}
.quote-attr {{ display:block; margin-top:10mm; font:800 10pt var(--font-data); letter-spacing:.12em; text-transform:uppercase; color:var(--accent); }}

/* ── closing (dark, as impactful as the cover) ── */
.closing h2 {{ font-size:42pt; }}
.cta-pill {{ display:inline-block; margin-top:9mm; padding:4.5mm 10mm; border-radius:999px; background:var(--accent); color:{on_acc}; font:800 11.5pt var(--font-heading); letter-spacing:.02em; }}
.closing-cards {{ margin-top:12mm; }}
.closing-cards .ins-card {{ min-height:58mm; padding:7mm 6mm; }}
"""


# ─── Page renderers ───────────────────────────────────────────────────────────

def _render_page(section: Section, spec: DocumentSpec, theme: Theme,
                 idx: int, total: int) -> str:
    layout = section.layout or "split"
    title = escape(clip_chars(spec.title, 90))
    eyebrow = escape(clip_chars(section.eyebrow or f"Section {idx}", 44))
    heading = clip_words(section.heading or spec.title, 13)
    heading_html = escape(heading)
    body = escape(clip_words(section.body, 96 if idx in (1, total) else 82))
    callout = escape(clip_words(section.callout, 30))
    highlights_raw = section.fallback_highlights()[:4]
    highlights = [escape(clip_words(h, 18)) for h in highlights_raw]

    dark = layout in ("cover", "quote")
    folio_cls = "folio dark" if dark else "folio"
    folio = f'<div class="{folio_cls}"><span>{title}</span><b>{idx:02d} / {total:02d}</b></div>'
    head_html = (f'<div class="doc-head"><div class="kicker"><span class="num">{idx:02d}</span>{eyebrow}</div>'
                 f'<h2>{heading_html}</h2></div>')
    pullquote = f'<div class="pullquote">{callout}</div>' if callout else ""
    insights_list = "".join(f'<div class="insight">{h}</div>' for h in highlights)

    if layout == "cover":
        subtitle = escape(clip_chars(section.subtitle or spec.subtitle, 155))
        author = escape(clip_chars(spec.author or "AI Document Studio", 90))
        aesthetic_label = escape(clip_chars(theme.aesthetic_label, 42))
        signature = escape(clip_chars(theme.signature_element, 70))
        return f"""
<section class="page cover">
  {_backdrop(theme, idx, dark=True)}
  <div class="content">
    <div class="cover-top"><span class="cover-pill">{aesthetic_label}</span><span class="cover-pill">{eyebrow}</span></div>
    <div class="cover-body"><h1>{heading_html}</h1><p class="cover-sub">{subtitle}</p></div>
    <div class="cover-meta"><div><p class="cover-callout">{callout or subtitle}</p><p class="cover-author">{author} &nbsp;·&nbsp; motif: {signature}</p></div><div class="cover-num">01</div></div>
  </div>
</section>"""

    if layout == "quote":
        chips = _insight_cards(section.statements or highlights_raw, limit=3)
        attr = escape(clip_chars(section.attribution, 48))
        quote_text = callout or body
        return f"""
<section class="page dark-page">
  {_backdrop(theme, idx, dark=True)}
  <div class="content">
    <div class="doc-head"><div class="kicker"><span class="num" style="color:var(--accent)">{idx:02d}</span>{eyebrow}</div><h2>{heading_html}</h2></div>
    <div class="giant-quote"><span class="mark">“</span>{quote_text}</div>
    {f'<div class="quote-attr">{attr}</div>' if attr else ''}
    {chips}
  </div>
  {folio}
</section>"""

    if layout == "closing":
        takeaways = section.takeaways or section.statements or highlights_raw
        cards = _insight_cards(takeaways, limit=3)
        closing_body = body or escape(clip_words(spec.conclusion, 96))
        cta = escape(clip_words(section.cta or section.callout, 20))
        return f"""
<section class="page dark-page closing">
  {_backdrop(theme, idx, dark=True)}
  <div class="content doc">
    {head_html}
    <p class="lede">{closing_body}</p>
    {f'<span class="cta-pill">{cta}</span>' if cta else ''}
    <div class="closing-cards">{cards}</div>
  </div>
  <div class="folio dark"><span>{title}</span><b>{idx:02d} / {total:02d}</b></div>
</section>"""

    if layout == "statgrid":
        # If the section also carries a chart, use compact stat tiles and keep
        # the chart on the page — content-faithful, and no two stat pages match.
        chart = _chart_html(section.chart, theme, heading)
        cells = "".join(
            f'<div class="cell">{_big_stat_card(s, i, compact=bool(chart))}</div>'
            for i, s in enumerate(section.stats[:3]))
        chart_block = f'<div style="margin-top:8mm">{chart}</div>' if chart else ''
        return f"""
<section class="page">
  {_backdrop(theme, idx)}
  <div class="content doc">
    {head_html}
    <p class="lede">{body}</p>
    <div class="trow stat-row">{cells}</div>
    {chart_block}
  </div>
  {folio}
</section>"""

    if layout == "timeline":
        cells = ""
        for i, s in enumerate(section.steps[:4] or []):
            date = escape(clip_chars(s.date or f"{i + 1:02d}", 16))
            step = escape(clip_chars(s.title or f"Phase {i + 1}", 22))
            desc = escape(clip_words(s.description or s.title, 24))
            cells += (f'<div class="cell"><div class="tl-card"><span class="n-dot">{i + 1:02d}</span>'
                      f'<span class="tl-date">{date}</span><span class="tl-step">{step}</span>'
                      f'<span class="tl-desc">{desc}</span></div></div>')
        if not cells:
            for i, h in enumerate(highlights_raw[:4]):
                cells += (f'<div class="cell"><div class="tl-card"><span class="n-dot">{i + 1:02d}</span>'
                          f'<span class="tl-step">Phase {i + 1}</span><span class="tl-desc">{h}</span></div></div>')
        return f"""
<section class="page">
  {_backdrop(theme, idx)}
  <div class="content doc">
    {head_html}
    <p class="lede">{body}</p>
    <div class="trow tl-row">{cells}</div>
  </div>
  {folio}
</section>"""

    if layout == "cards":
        if section.tiles:
            items = [t.value or t.label for t in section.tiles[:4]]
        else:
            items = highlights_raw[:4]
        cards = _insight_cards(items, limit=4, hero=True)
        return f"""
<section class="page">
  {_backdrop(theme, idx)}
  <div class="content doc">
    {head_html}
    <p class="lede">{body}</p>
    {cards}
  </div>
  {folio}
</section>"""

    if layout == "feature":
        hero = _hero_visual(section, theme, idx, heading)
        # comparison columns beat generic insight cards when the section has them
        cards = _duo_cards(section) or _insight_cards(highlights_raw, limit=3)
        return f"""
<section class="page">
  {_backdrop(theme, idx)}
  <div class="content doc">
    {head_html}
    <p class="lede">{body}</p>
    <div style="margin-top:8mm">{hero}</div>
    {cards}
  </div>
  {folio}
</section>"""

    # split — narrative + tall side visual (default)
    visual = _side_visual(section, theme, idx, heading)
    narrative = f'<p>{body}</p>{pullquote}<div class="insights">{insights_list}</div>'
    if section.mirror:
        cols = (f'<div class="side-col gut">{visual}</div>'
                f'<div class="main-col body-col">{narrative}</div>')
    else:
        cols = (f'<div class="main-col body-col gut">{narrative}</div>'
                f'<div class="side-col">{visual}</div>')
    return f"""
<section class="page">
  {_backdrop(theme, idx)}
  <div class="content doc">
    {head_html}
    <div class="split">{cols}</div>
  </div>
  {folio}
</section>"""


# ─── Entry point ──────────────────────────────────────────────────────────────

def render_pdf_html(spec: DocumentSpec, theme: Theme, page_size: str = "a4",
                    font_face_css: str = "", seed: Optional[int] = None) -> str:
    """Build the full HTML document (also used by tests without WeasyPrint)."""
    page_kw, page_w, page_h = PAGE_SIZES.get(page_size.lower(), PAGE_SIZES["a4"])
    plan_pdf_layouts(spec.sections, theme.seed if seed is None else seed)
    total = len(spec.sections)
    pages = [_render_page(sec, spec, theme, i, total)
             for i, sec in enumerate(spec.sections, 1)]
    css = _stylesheet(theme, page_w, page_h, page_kw)
    meta = (f"<title>{escape(spec.title)}</title>"
            f'<meta name="author" content="{escape(spec.author or "AI Document Studio")}"/>'
            f'<meta name="description" content="{escape(clip_chars(spec.subtitle or spec.title, 150))}"/>'
            f'<meta name="generator" content="doc-generation engine"/>')
    return (f'<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>{meta}'
            f'{font_face_css}<style>{css}</style></head><body>{"".join(pages)}</body></html>')


def render_pdf(spec: DocumentSpec, theme: Theme, page_size: str = "a4",
               font_face_css: str = "", seed: Optional[int] = None) -> bytes:
    """Render the spec to PDF bytes. WeasyPrint is imported lazily so PPTX/DOCX
    generation works on machines without the pango/cairo native libraries."""
    from weasyprint import CSS, HTML
    import io

    html_doc = render_pdf_html(spec, theme, page_size, font_face_css, seed)
    page_kw, _, _ = PAGE_SIZES.get(page_size.lower(), PAGE_SIZES["a4"])
    pagination_css = CSS(string=f"""
        @page {{ size: {page_kw}; margin: 0; }}
        .page {{ page-break-after: always; break-after: page; }}
        .page:last-child {{ page-break-after: auto; break-after: auto; }}
        .premium-stat, .premium-chart, .premium-table, .photo-plate, .geo-panel,
        .bigstat, .tl-card, .ins-card, .pullquote {{ break-inside: avoid; page-break-inside: avoid; }}
        h1, h2, h3 {{ break-after: avoid; page-break-after: avoid; }}
        p {{ orphans: 3; widows: 3; }}
    """)
    buffer = io.BytesIO()
    HTML(string=html_doc).write_pdf(buffer, stylesheets=[pagination_css])
    return buffer.getvalue()
