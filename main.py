from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Pt as PPTXPt, Inches as PPTXInches, Emu
from pptx.dml.color import RGBColor as PPTXRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt as PPTXPt
from weasyprint import HTML, CSS
import httpx, io, json, os, re, uuid, math
from html import escape
from vercel.blob import AsyncBlobClient

app = FastAPI()

DEEPINFRA_KEY = os.getenv("DEEPINFRA_KEY")
VERCEL_BLOB_TOKEN = os.getenv("VERCEL_BLOB_TOKEN") or os.getenv("BLOB_READ_WRITE_TOKEN")


# ─── Health check ────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok"}


# ─── Utilities ───────────────────────────────────────────────────────────────

NUMBER_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15,
}

def requested_count(prompt: str, unit: str, default: int, minimum: int, maximum: int) -> int:
    normalized = prompt.lower()
    match = re.search(r"\b(\d{1,2})\s*(?:-|\s)?"+re.escape(unit)+r"s?\b", normalized)
    if match:
        return max(minimum, min(maximum, int(match.group(1))))
    for word, value in NUMBER_WORDS.items():
        if re.search(r"\b"+word+r"\s+(?:-|\s)?"+re.escape(unit)+r"s?\b", normalized):
            return max(minimum, min(maximum, value))
    return default


async def _deepinfra_call(messages: list, temperature: float = 0.4) -> str:
    async with httpx.AsyncClient(timeout=120) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={
                "model": "deepseek-ai/DeepSeek-V4-Flash",
                "temperature": temperature,
                "messages": messages,
            },
        )
    raw = res.json()["choices"][0]["message"]["content"]
    return re.sub(r"```json|```", "", raw).strip()


# ─── Vercel Blob upload ───────────────────────────────────────────────────────

async def upload_to_vercel_blob(buffer: io.BytesIO, filename: str, content_type: str) -> dict:
    if not VERCEL_BLOB_TOKEN:
        raise RuntimeError("VERCEL_BLOB_TOKEN or BLOB_READ_WRITE_TOKEN is not configured")
    buffer.seek(0)
    pathname = f"generated/{uuid.uuid4().hex}-{filename}"
    async with AsyncBlobClient(token=VERCEL_BLOB_TOKEN) as client:
        blob = await client.put(
            pathname, buffer.read(), access="public",
            content_type=content_type, add_random_suffix=False, overwrite=True,
        )
    return {
        "url": blob.url,
        "downloadUrl": getattr(blob, "download_url", blob.url),
        "pathname": blob.pathname,
        "contentType": getattr(blob, "content_type", content_type),
    }


# ══════════════════════════════════════════════════════════════════════════════
# PDF GENERATION
# ══════════════════════════════════════════════════════════════════════════════

PALETTES = {
    "indigo":  {
        "primary":"#312E81","accent":"#6366F1","mid":"#818CF8","light":"#E0E7FF",
        "wash":"#F5F3FF","text":"#1E1B4B","dark":"#1E1B4B","soft":"#C7D2FE",
        "white":"#FFFFFF","body":"#374151","card2":"#4338CA","muted":"#A5B4FC",
    },
    "teal":    {
        "primary":"#134E4A","accent":"#0D9488","mid":"#2DD4BF","light":"#CCFBF1",
        "wash":"#F0FDFA","text":"#0F2F2E","dark":"#0F2F2E","soft":"#99F6E4",
        "white":"#FFFFFF","body":"#374151","card2":"#0F766E","muted":"#5EEAD4",
    },
    "crimson": {
        "primary":"#7F1D1D","accent":"#DC2626","mid":"#F87171","light":"#FEE2E2",
        "wash":"#FFF5F5","text":"#450A0A","dark":"#450A0A","soft":"#FCA5A5",
        "white":"#FFFFFF","body":"#374151","card2":"#991B1B","muted":"#FCA5A5",
    },
    "emerald": {
        "primary":"#064E3B","accent":"#059669","mid":"#34D399","light":"#D1FAE5",
        "wash":"#F0FDF4","text":"#022C22","dark":"#022C22","soft":"#6EE7B7",
        "white":"#FFFFFF","body":"#374151","card2":"#065F46","muted":"#6EE7B7",
    },
    "slate":   {
        "primary":"#1E293B","accent":"#475569","mid":"#94A3B8","light":"#E2E8F0",
        "wash":"#F8FAFC","text":"#0F172A","dark":"#0F172A","soft":"#CBD5E1",
        "white":"#FFFFFF","body":"#374151","card2":"#334155","muted":"#94A3B8",
    },
    "violet":  {
        "primary":"#4C1D95","accent":"#7C3AED","mid":"#A78BFA","light":"#EDE9FE",
        "wash":"#FAF5FF","text":"#2E1065","dark":"#2E1065","soft":"#C4B5FD",
        "white":"#FFFFFF","body":"#374151","card2":"#5B21B6","muted":"#C4B5FD",
    },
    "amber":   {
        "primary":"#78350F","accent":"#D97706","mid":"#FBBF24","light":"#FDE68A",
        "wash":"#FFFBEB","text":"#451A03","dark":"#451A03","soft":"#FCD34D",
        "white":"#FFFFFF","body":"#374151","card2":"#92400E","muted":"#FCD34D",
    },
    "rose":    {
        "primary":"#881337","accent":"#E11D48","mid":"#FB7185","light":"#FFE4E6",
        "wash":"#FFF1F2","text":"#4C0519","dark":"#4C0519","soft":"#FDA4AF",
        "white":"#FFFFFF","body":"#374151","card2":"#9F1239","muted":"#FDA4AF",
    },
    "navy":    {
        "primary":"#0F2D5E","accent":"#1D6FA4","mid":"#5BA8D4","light":"#D6EAF8",
        "wash":"#EBF5FB","text":"#0A1F3D","dark":"#0A1F3D","soft":"#AED6F1",
        "white":"#FFFFFF","body":"#374151","card2":"#154F7A","muted":"#7EC8E3",
    },
}

INTERIOR_LAYOUTS = ["split","feature","grid","quote","timeline","stats","editorial","manifesto"]


async def call_ai_pdf(prompt: str, page_count: int) -> dict:
    system = f"""You are a premium editorial content strategist. Return ONLY valid JSON, no markdown fences.
Generate content for exactly {page_count} PDF pages. Each page maps 1:1 to one physical printed page.

PALETTE — pick one that emotionally fits the topic:
indigo=tech/professional, teal=health/sustainability, crimson=energy/urgency, emerald=growth/nature,
slate=enterprise/legal, violet=creative/luxury, amber=warmth/food/culture, rose=beauty/lifestyle, navy=corporate/finance

PAGE TYPES (use variety, never same type twice in a row):
- "cover"      → page 1 only: big hero title, subtitle, author tagline
- "split"      → 60/40 two-column: editorial body + highlights sidebar
- "feature"    → full-width hero band heading + body + 3-card row
- "grid"       → heading + body + 2×2 fact/stat tiles
- "quote"      → dark full-bleed page, giant pull-quote + body + pill badges
- "timeline"   → heading + 3-4 connected numbered steps + callout
- "stats"      → heading + 3 big-number hero stat cards + body
- "editorial"  → magazine-style: large heading left, body right, bottom band callout
- "manifesto"  → bold typographic page: stacked large statements + body
- "closing"    → last page: takeaways + CTA gradient band

For "grid": provide "tiles": [{{"label":"label","value":"fact or number"}}] (3-4 tiles)
For "timeline": provide "steps": [{{"step":"label","desc":"one sentence"}}] (3-4 steps)
For "stats": provide "stats": [{{"number":"value","label":"what it measures","sub":"one sentence context"}}] (exactly 3)
For "closing": provide "takeaways": ["short takeaway"] (3-4 items)
For "manifesto": provide "statements": ["Bold declarative sentence."] (3 statements)

JSON schema:
{{
  "title": "Document title",
  "subtitle": "Compelling subtitle",
  "palette": "indigo|teal|crimson|emerald|slate|violet|amber|rose|navy",
  "author": "Prepared for [topic] — 2025",
  "pages": [
    {{
      "type": "cover|split|feature|grid|quote|timeline|stats|editorial|manifesto|closing",
      "eyebrow": "SECTION LABEL",
      "heading": "Compelling page headline",
      "body": "Rich paragraph, minimum 60 words, specific and insightful.",
      "highlights": ["Punchy point one", "Punchy point two", "Punchy point three"],
      "callout": "One memorable quotable sentence.",
      "tiles": [], "steps": [], "stats": [], "takeaways": [], "statements": []
    }}
  ]
}}

RULES:
- pages[0].type MUST be "cover"
- pages[{page_count-1}].type MUST be "closing" if page_count > 2
- Interior pages: vary types — mix editorial/manifesto/stats/timeline/grid/quote/split/feature
- body: minimum 60 words of rich specific content
- highlights: 3-4 items always
- pages array MUST have exactly {page_count} items"""

    raw = await _deepinfra_call([
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ], temperature=0.38)
    data = json.loads(raw)
    pages = data.get("pages", [])
    while len(pages) < page_count:
        i = len(pages)
        lt = INTERIOR_LAYOUTS[i % len(INTERIOR_LAYOUTS)]
        pages.append({
            "type": lt, "eyebrow": f"SECTION {i+1}",
            "heading": f"Key Insight {i+1}",
            "body": "This section presents essential analysis and actionable perspective on the topic, offering structured insight to guide informed decision-making and strategic planning across all relevant dimensions.",
            "highlights": ["Structured insight", "Clear implication", "Practical takeaway"],
            "callout": "The strongest documents stay precise, structured, and visually deliberate.",
            "tiles": [], "steps": [], "stats": [], "takeaways": [], "statements": [],
        })
    data["pages"] = pages[:page_count]
    return data


# ─── SVG decorative helpers (WeasyPrint renders inline SVG natively) ─────────

def _svg_circles(c1: str, c2: str, c3: str) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;overflow:hidden;pointer-events:none;" viewBox="0 0 595 842" preserveAspectRatio="none">'
        f'<circle cx="520" cy="80" r="200" fill="{c1}" opacity="0.10"/>'
        f'<circle cx="50" cy="780" r="160" fill="{c2}" opacity="0.09"/>'
        f'<circle cx="480" cy="680" r="110" fill="{c3}" opacity="0.07"/>'
        f'<circle cx="280" cy="400" r="350" fill="{c1}" opacity="0.025"/>'
        f'</svg>'
    )

def _svg_diagonal_band(color: str, x1=0, y1=580, x2=595, y2=500, x3=595, y3=842, x4=0, y4=842) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">'
        f'<polygon points="{x1},{y1} {x2},{y2} {x3},{y3} {x4},{y4}" fill="{color}" opacity="0.9"/>'
        f'</svg>'
    )

def _svg_grid_dots(color: str, spacing: int = 24) -> str:
    dots = "".join(
        f'<circle cx="{x}" cy="{y}" r="1.2" fill="{color}" opacity="0.30"/>'
        for x in range(0, 600, spacing) for y in range(0, 850, spacing)
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{dots}</svg>'
    )

def _svg_half_circle(color: str, side: str = "right") -> str:
    if side == "right":
        return (
            f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;'
            f'top:0;right:0;width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">'
            f'<circle cx="595" cy="421" r="280" fill="{color}" opacity="0.12"/>'
            f'<circle cx="595" cy="421" r="180" fill="{color}" opacity="0.08"/>'
            f'</svg>'
        )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;'
        f'top:0;left:0;width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">'
        f'<circle cx="0" cy="421" r="280" fill="{color}" opacity="0.12"/>'
        f'<circle cx="0" cy="421" r="180" fill="{color}" opacity="0.08"/>'
        f'</svg>'
    )

def _svg_cross_hatch(color: str) -> str:
    lines = "".join(
        f'<line x1="{x}" y1="0" x2="{x}" y2="842" stroke="{color}" stroke-width="0.4" opacity="0.20"/>'
        for x in range(0, 600, 36)
    ) + "".join(
        f'<line x1="0" y1="{y}" x2="595" y2="{y}" stroke="{color}" stroke-width="0.4" opacity="0.20"/>'
        for y in range(0, 850, 36)
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{lines}</svg>'
    )

def _svg_arc_decoration(color: str) -> str:
    arcs = "".join(
        f'<circle cx="595" cy="0" r="{r}" fill="none" stroke="{color}" stroke-width="1.2" opacity="{round(0.14 - i*0.025, 3)}"/>'
        for i, r in enumerate([100, 160, 220, 290, 360])
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{arcs}</svg>'
    )

def _svg_stripe_band(color: str, accent: str) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">'
        f'<polygon points="0,620 220,842 0,842" fill="{color}" opacity="0.16"/>'
        f'<polygon points="0,700 130,842 0,842" fill="{accent}" opacity="0.20"/>'
        f'</svg>'
    )

def _svg_diagonal_lines(color: str) -> str:
    """Subtle diagonal line pattern."""
    lines = "".join(
        f'<line x1="{x-200}" y1="0" x2="{x}" y2="842" stroke="{color}" stroke-width="0.6" opacity="0.12"/>'
        for x in range(-100, 900, 32)
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{lines}</svg>'
    )

def _svg_hexagons(color: str) -> str:
    """Scattered hexagon outlines for texture."""
    hexes = ""
    positions = [(80,120,28),(500,80,20),(520,720,36),(60,700,22),(290,200,16),(450,450,24),(150,500,18)]
    for cx,cy,r in positions:
        pts = " ".join(
            f"{round(cx + r * math.cos(math.radians(60*i-30)),1)},{round(cy + r * math.sin(math.radians(60*i-30)),1)}"
            for i in range(6)
        )
        hexes += f'<polygon points="{pts}" fill="none" stroke="{color}" stroke-width="0.8" opacity="0.18"/>'
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{hexes}</svg>'
    )

def _svg_noise_overlay(color: str) -> str:
    """Fine noise texture via tiny dots."""
    import random
    rng = random.Random(42)
    dots = "".join(
        f'<circle cx="{rng.randint(0,595)}" cy="{rng.randint(0,842)}" r="0.7" fill="{color}" opacity="0.08"/>'
        for _ in range(280)
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;pointer-events:none;" viewBox="0 0 595 842">{dots}</svg>'
    )

def _svg_cover_geometric(primary: str, accent: str, mid: str) -> str:
    """Bold geometric shapes for cover pages."""
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;overflow:hidden;pointer-events:none;" viewBox="0 0 595 842">'
        # Large rotated rectangle top-right
        f'<rect x="320" y="-80" width="380" height="380" rx="18" fill="{accent}" opacity="0.13" transform="rotate(18 510 110)"/>'
        # Medium circle bottom-left
        f'<circle cx="40" cy="800" r="160" fill="{mid}" opacity="0.10"/>'
        # Small accent circle mid-right
        f'<circle cx="560" cy="500" r="70" fill="{accent}" opacity="0.09"/>'
        # Thin ring decoration
        f'<circle cx="510" cy="110" r="145" fill="none" stroke="{accent}" stroke-width="1.5" opacity="0.18"/>'
        f'<circle cx="510" cy="110" r="200" fill="none" stroke="{mid}" stroke-width="0.8" opacity="0.12"/>'
        f'</svg>'
    )

def _svg_manifesto_bg(dark: str, accent: str) -> str:
    """Dramatic background for manifesto pages."""
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;inset:0;'
        f'width:210mm;height:297mm;overflow:hidden;pointer-events:none;" viewBox="0 0 595 842">'
        f'<rect x="0" y="0" width="595" height="842" fill="{dark}"/>'
        f'<rect x="-60" y="200" width="200" height="600" fill="{accent}" opacity="0.07" transform="rotate(-12 40 500)"/>'
        f'<circle cx="595" cy="0" r="320" fill="{accent}" opacity="0.06"/>'
        f'<circle cx="0" cy="842" r="240" fill="{accent}" opacity="0.05"/>'
        f'</svg>'
    )


# ─── Footer / running elements ────────────────────────────────────────────────

def _footer(title: str, idx: int, total: int, accent: str, mid: str, dark_bg: bool = False) -> str:
    title_color = "rgba(255,255,255,0.5)" if dark_bg else mid
    rule_color  = "rgba(255,255,255,0.15)" if dark_bg else "var(--light)"
    return f"""<div class="footer" style="border-top-color:{rule_color};">
  <span class="footer-title" style="color:{title_color};">{title}</span>
  <span class="footer-pager" style="background:{accent};color:white;">{idx}&thinsp;/&thinsp;{total}</span>
</div>"""

def _eyebrow(text: str, accent: str, light: bool = False) -> str:
    color = "rgba(255,255,255,0.65)" if light else accent
    return f'<div class="eyebrow" style="color:{color};">{text}</div>'


# ─── Main renderer ────────────────────────────────────────────────────────────

def render_pdf_html(structure: dict, page_count: int) -> str:
    pname = str(structure.get("palette","indigo")).lower()
    P = PALETTES.get(pname, PALETTES["indigo"])

    title    = escape(str(structure.get("title")    or "Report"))
    subtitle = escape(str(structure.get("subtitle") or ""))
    author   = escape(str(structure.get("author")   or ""))

    css_vars = f"""
:root {{
  --primary: {P['primary']};
  --accent:  {P['accent']};
  --mid:     {P['mid']};
  --light:   {P['light']};
  --wash:    {P['wash']};
  --text:    {P['text']};
  --dark:    {P['dark']};
  --soft:    {P['soft']};
  --body:    {P['body']};
  --white:   {P['white']};
  --card2:   {P['card2']};
  --muted:   {P['muted']};
}}"""

    global_css = f"""
@page {{
  size: A4;
  margin: 0;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{
  font-family: Georgia, 'Times New Roman', serif;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}}
{css_vars}

/* ═══════════════════════════════════════════
   PAGE BASE — WeasyPrint pagination control
   ═══════════════════════════════════════════ */
.page {{
  width: 210mm;
  height: 297mm;
  page-break-after: always;
  page-break-inside: avoid;
  break-inside: avoid;
  position: relative;
  overflow: hidden;
}}
.page:last-child {{
  page-break-after: auto;
}}

/* ═══════════════════════════════════════════
   GLOBAL ORPHAN / WIDOW / BREAK CONTROL
   Prevents single lines stranded at top/bottom
   ═══════════════════════════════════════════ */
p, li, .body-text, .callout-italic, .timeline-step-desc, .stats-sub,
.split-card p, .feature-card p, .grid-tile-val, .quote-body,
.manifesto-body, .closing-takeaway p, .editorial-highlight-list li {{
  orphans: 3;
  widows: 3;
}}

/* Cards, tiles, steps — never break internally */
.split-card, .feature-card, .grid-tile, .stats-card, .timeline-step,
.closing-takeaway, .quote-box, .timeline-callout, .feature-callout-bar,
.grid-callout-bar, .closing-cta, .manifesto-statement, .stats-pill,
.quote-pill, .editorial-callout-band, .cover-bottom, .cover-inner {{
  page-break-inside: avoid;
  break-inside: avoid;
}}

/* Lists — avoid orphaned first/last items */
ul, ol, .feature-cards, .stats-cards, .grid-tiles, .quote-pills,
.stats-pills, .closing-takeaways, .manifesto-statements {{
  page-break-inside: avoid;
  break-inside: avoid;
}}

/* Headings stay with following content */
h1, h2, h3, .eyebrow {{
  page-break-after: avoid;
  break-after: avoid;
}}

/* ═══════════════════════════════════════════
   TYPOGRAPHY
   ═══════════════════════════════════════════ */
h1 {{ font-family: Georgia, serif; font-weight: 700; letter-spacing: -0.03em; line-height: 1.05; }}
h2 {{ font-family: Georgia, serif; font-weight: 700; letter-spacing: -0.02em; line-height: 1.10; }}
h3 {{ font-family: Georgia, serif; font-weight: 700; line-height: 1.18; }}
p  {{ font-family: 'Helvetica Neue', Arial, sans-serif; font-weight: 400; }}

.eyebrow {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; font-weight: 700; letter-spacing: 0.28em;
  text-transform: uppercase; margin-bottom: 5mm;
}}

.footer {{
  position: absolute; bottom: 10mm; left: 18mm; right: 18mm;
  border-top: 0.25mm solid var(--light); padding-top: 3.5mm;
  display: flex; justify-content: space-between; align-items: center;
}}
.footer-title {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; letter-spacing: 0.10em;
}}
.footer-pager {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  padding: 1.2mm 4.5mm; border-radius: 10mm; font-size: 7pt; font-weight: 700;
}}
.body-text {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10.5pt; line-height: 1.78; color: var(--body);
}}
.callout-italic {{
  font-family: Georgia, serif; font-style: italic;
  font-size: 11pt; line-height: 1.58; color: var(--primary);
}}

/* ═══════════════════════════════════════════
   COVER
   ═══════════════════════════════════════════ */
.cover {{
  background: linear-gradient(148deg, var(--dark) 0%, var(--primary) 55%, var(--card2) 100%);
  color: white;
  display: flex; flex-direction: column; justify-content: space-between;
}}
.cover-inner {{
  position: relative; z-index: 2;
  padding: 24mm 28mm 0;
}}
.cover-label {{
  display: inline-block;
  background: rgba(255,255,255,0.13);
  border: 1px solid rgba(255,255,255,0.20);
  border-radius: 20mm; padding: 2mm 9mm;
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; letter-spacing: 0.26em; text-transform: uppercase;
  margin-bottom: 13mm; color: rgba(255,255,255,0.85);
}}
.cover-title {{
  font-size: 48pt; max-width: 150mm; margin-bottom: 9mm;
  line-height: 1.02; letter-spacing: -0.04em;
}}
.cover-subtitle {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 13.5pt; line-height: 1.55; opacity: 0.82; max-width: 138mm;
  font-weight: 300;
}}
.cover-bottom {{
  position: relative; z-index: 2;
  padding: 0 28mm 22mm;
}}
.cover-rule {{
  height: 0.35mm; background: rgba(255,255,255,0.20); margin-bottom: 9mm;
}}
.cover-meta {{
  display: flex; justify-content: space-between; align-items: flex-end;
}}
.cover-author {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 9pt; opacity: 0.58; margin-bottom: 2.5mm; letter-spacing: 0.04em;
}}
.cover-callout {{
  font-family: Georgia, serif; font-size: 11pt; font-style: italic;
  max-width: 108mm; line-height: 1.55; opacity: 0.88;
}}
.cover-num {{
  font-size: 64pt; font-weight: 200; opacity: 0.15; line-height: 1;
  letter-spacing: -0.06em;
}}

/* ═══════════════════════════════════════════
   SPLIT
   ═══════════════════════════════════════════ */
.split {{
  display: grid; grid-template-columns: 60% 40%;
  background: white;
}}
.split-left {{
  padding: 20mm 13mm 24mm 20mm;
  display: flex; flex-direction: column;
}}
.split-accent-bar {{
  width: 20mm; height: 1mm; background: var(--accent);
  border-radius: 1mm; margin-bottom: 9mm;
}}
.split-right {{
  padding: 20mm 15mm 20mm 13mm;
  display: flex; flex-direction: column;
  background: linear-gradient(170deg, var(--primary) 0%, var(--dark) 100%);
  color: white;
}}
.split-hl-label {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7.5pt; font-weight: 700; letter-spacing: 0.22em;
  text-transform: uppercase; color: var(--muted); margin-bottom: 9mm;
}}
.split-card {{
  background: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.12);
  border-left: 3mm solid var(--muted);
  border-radius: 0 3mm 3mm 0;
  padding: 5.5mm 7mm; margin-bottom: 4mm;
}}
.split-card p {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 9.8pt; line-height: 1.55; color: rgba(255,255,255,0.90);
}}
.split-num {{
  margin-top: auto; font-size: 48pt; font-weight: 100;
  color: rgba(255,255,255,0.10); text-align: right; line-height: 1;
  letter-spacing: -0.06em;
}}
.split-blockquote {{
  border-left: 3.5mm solid var(--accent); padding: 5.5mm 9mm;
  background: var(--wash); border-radius: 0 4mm 4mm 0; margin-top: auto;
}}

/* ═══════════════════════════════════════════
   FEATURE
   ═══════════════════════════════════════════ */
.feature {{
  background: linear-gradient(145deg, var(--wash) 0%, white 65%);
  padding: 20mm 24mm 22mm;
}}
.feature-ghost-num {{
  position: absolute; right: -8mm; bottom: -10mm;
  font-size: 220pt; font-weight: 900; color: var(--light);
  line-height: 1; z-index: 0; opacity: 0.55;
  font-family: Georgia, serif; letter-spacing: -0.08em;
}}
.feature-inner {{ position: relative; z-index: 1; }}
.feature-band {{
  background: linear-gradient(138deg, var(--primary), var(--card2), var(--accent));
  border-radius: 7mm; padding: 12mm 16mm; margin-bottom: 12mm;
  box-shadow: 0 4px 20px rgba(0,0,0,0.18);
}}
.feature-band h2 {{ font-size: 28pt; color: white; letter-spacing: -0.025em; }}
.feature-band-sub {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10pt; color: rgba(255,255,255,0.72); margin-top: 4mm;
  line-height: 1.5; font-weight: 300;
}}
.feature-cards {{
  display: grid; gap: 5mm; margin-bottom: 11mm;
}}
.feature-card {{
  background: white; border: 1.2px solid var(--light);
  border-radius: 5mm; padding: 7mm 9mm;
  border-top: 4mm solid var(--accent);
  box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}}
.feature-card-num {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; font-weight: 700; letter-spacing: 0.18em;
  color: var(--accent); text-transform: uppercase; margin-bottom: 2.5mm;
}}
.feature-card p {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 9.5pt; line-height: 1.58; color: var(--text);
}}
.feature-callout-bar {{
  background: var(--light); border-radius: 4mm;
  padding: 6.5mm 11mm; display: flex; align-items: center; gap: 6mm;
}}
.feature-callout-pip {{
  width: 3mm; min-height: 10mm; background: var(--accent);
  border-radius: 2mm; flex-shrink: 0;
}}

/* ═══════════════════════════════════════════
   GRID
   ═══════════════════════════════════════════ */
.grid-page {{
  padding: 20mm 22mm 22mm; background: var(--wash);
}}
.grid-tiles {{
  display: grid; grid-template-columns: 1fr 1fr; gap: 5.5mm; margin-bottom: 9mm;
}}
.grid-tile {{
  background: white; border-radius: 5mm; padding: 9mm 10mm;
  border: 1.2px solid var(--light);
  border-top: 4.5mm solid var(--accent);
  box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}}
.grid-tile-label {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; color: var(--mid); text-transform: uppercase;
  letter-spacing: 0.16em; margin-bottom: 4mm; font-weight: 600;
}}
.grid-tile-val {{
  font-family: Georgia, serif;
  font-size: 16pt; font-weight: 700; color: var(--primary); line-height: 1.25;
}}
.grid-callout-bar {{
  background: linear-gradient(135deg, var(--primary), var(--accent));
  border-radius: 5mm; padding: 6.5mm 13mm; color: white;
}}

/* ═══════════════════════════════════════════
   QUOTE
   ═══════════════════════════════════════════ */
.quote-page {{
  background: linear-gradient(165deg, var(--dark) 0%, var(--primary) 50%, var(--dark) 100%);
  padding: 26mm 28mm; color: white;
  display: flex; flex-direction: column; justify-content: space-between;
}}
.quote-deco-open {{
  position: absolute; top: 4mm; left: 10mm;
  font-size: 150pt; font-weight: 900; color: rgba(255,255,255,0.045);
  line-height: 1; font-family: Georgia, serif; z-index: 0;
}}
.quote-deco-close {{
  position: absolute; bottom: 18mm; right: 10mm;
  font-size: 150pt; font-weight: 900; color: rgba(255,255,255,0.045);
  line-height: 1; font-family: Georgia, serif; z-index: 0;
}}
.quote-box {{
  background: rgba(255,255,255,0.07);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 7mm; padding: 14mm 16mm; margin-bottom: 10mm;
  border-left: 5mm solid var(--mid);
  position: relative; z-index: 1;
}}
.quote-box p {{
  font-family: Georgia, serif; font-size: 21pt; font-weight: 600;
  line-height: 1.35; color: white; letter-spacing: -0.01em;
}}
.quote-body {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10.5pt; line-height: 1.75; color: rgba(255,255,255,0.78);
  max-width: 145mm; margin-bottom: 10mm; position: relative; z-index: 1;
}}
.quote-pills {{ display: flex; flex-wrap: wrap; gap: 3.5mm; margin-bottom: 9mm; position: relative; z-index: 1; }}
.quote-pill {{
  background: rgba(255,255,255,0.11);
  border: 1px solid rgba(255,255,255,0.16);
  border-radius: 20mm;
  padding: 2mm 8mm; font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 8.5pt; color: white;
}}
.quote-footer {{
  border-top: 1px solid rgba(255,255,255,0.15); padding-top: 5mm;
  display: flex; justify-content: space-between;
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; color: rgba(255,255,255,0.40);
  position: relative; z-index: 1;
}}

/* ═══════════════════════════════════════════
   TIMELINE
   ═══════════════════════════════════════════ */
.timeline-page {{
  background: white; padding: 20mm 22mm 22mm;
}}
.timeline-sidebar {{
  position: absolute; right: 0; top: 0; bottom: 0; width: 36mm;
  background: linear-gradient(180deg, var(--wash) 0%, var(--light) 100%);
  z-index: 0;
}}
.timeline-sidebar-accent {{
  position: absolute; right: 0; top: 0; bottom: 0; width: 4mm;
  background: linear-gradient(180deg, var(--primary) 0%, var(--accent) 100%);
  z-index: 0;
}}
.timeline-inner {{ position: relative; z-index: 1; max-width: 148mm; }}
.timeline-step {{
  display: flex; align-items: flex-start; gap: 7mm; margin-bottom: 0;
}}
.timeline-node {{
  display: flex; flex-direction: column; align-items: center; flex-shrink: 0;
}}
.timeline-circle {{
  width: 17mm; height: 17mm; border-radius: 50%;
  background: linear-gradient(135deg, var(--primary), var(--accent));
  display: flex; align-items: center; justify-content: center;
  color: white; font-weight: 700; font-size: 10pt;
  font-family: 'Helvetica Neue', Arial, sans-serif;
  box-shadow: 0 2px 8px rgba(0,0,0,0.20);
}}
.timeline-connector {{
  width: 0.6mm; height: 8mm; background: var(--light);
  margin: 1mm 0 1mm 8.2mm;
}}
.timeline-content {{ padding-top: 2mm; flex: 1; }}
.timeline-step-label {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-weight: 700; font-size: 11.5pt; color: var(--primary); margin-bottom: 2mm;
}}
.timeline-step-desc {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10pt; color: #4B5563; line-height: 1.60;
}}
.timeline-callout {{
  background: var(--wash); border-radius: 5mm; padding: 7mm 11mm;
  border-left: 3.5mm solid var(--accent); margin-top: 9mm;
}}

/* ═══════════════════════════════════════════
   STATS
   ═══════════════════════════════════════════ */
.stats-page {{
  background: linear-gradient(145deg, var(--wash) 0%, white 100%);
  padding: 20mm 22mm 22mm;
}}
.stats-cards {{
  display: grid; grid-template-columns: repeat(3,1fr); gap: 5.5mm; margin-bottom: 11mm;
}}
.stats-card {{
  background: linear-gradient(148deg, var(--primary), var(--accent));
  border-radius: 7mm; padding: 11mm 9mm; text-align: center; color: white;
  box-shadow: 0 4px 16px rgba(0,0,0,0.20);
}}
.stats-num {{
  font-family: Georgia, serif; font-size: 40pt; font-weight: 700;
  line-height: 1; margin-bottom: 4mm; letter-spacing: -0.03em;
}}
.stats-label {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 8.5pt; opacity: 0.88; line-height: 1.45; margin-bottom: 3.5mm;
  font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em;
}}
.stats-sub {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7.5pt; opacity: 0.65; line-height: 1.45;
}}
.stats-pills {{ display: flex; flex-wrap: wrap; gap: 3.5mm; margin-top: 9mm; }}
.stats-pill {{
  background: var(--light); border-radius: 20mm; padding: 2.5mm 8mm;
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 9pt; color: var(--primary); font-weight: 500;
}}

/* ═══════════════════════════════════════════
   EDITORIAL
   ═══════════════════════════════════════════ */
.editorial-page {{
  background: white; padding: 0; display: grid;
  grid-template-rows: auto 1fr auto;
}}
.editorial-topband {{
  height: 3.5mm;
  background: linear-gradient(90deg, var(--primary) 0%, var(--accent) 55%, var(--mid) 100%);
}}
.editorial-body-grid {{
  display: grid; grid-template-columns: 52% 48%;
  padding: 17mm 20mm 0; gap: 14mm; align-items: start;
}}
.editorial-left h2 {{
  font-size: 32pt; color: var(--primary); line-height: 1.08;
  margin-bottom: 9mm; letter-spacing: -0.025em;
}}
.editorial-rule {{
  height: 0.8mm; width: 22mm; background: var(--accent);
  border-radius: 1mm; margin-bottom: 8mm;
}}
.editorial-highlight-list {{
  list-style: none; padding: 0; margin-top: 8mm;
}}
.editorial-highlight-list li {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10pt; color: var(--body); padding: 4.5mm 0;
  border-bottom: 0.3mm solid var(--light); line-height: 1.55;
}}
.editorial-highlight-list li::before {{
  content: "→ "; color: var(--accent); font-weight: 700;
}}
.editorial-right .body-text {{ margin-bottom: 9mm; }}
.editorial-callout-band {{
  background: var(--primary); padding: 10mm 20mm;
  margin-top: auto;
}}
.editorial-callout-band p {{
  font-family: Georgia, serif; font-size: 13pt; font-style: italic;
  color: white; line-height: 1.55;
}}

/* ═══════════════════════════════════════════
   MANIFESTO
   ═══════════════════════════════════════════ */
.manifesto-page {{
  color: white; padding: 22mm 26mm;
  display: flex; flex-direction: column; justify-content: center;
}}
.manifesto-statements {{ margin-bottom: 13mm; }}
.manifesto-statement {{
  font-family: Georgia, serif; font-size: 21pt; font-weight: 700;
  line-height: 1.20; color: white; margin-bottom: 8mm;
  padding-bottom: 8mm; border-bottom: 0.3mm solid rgba(255,255,255,0.14);
}}
.manifesto-statement:last-child {{ border-bottom: none; margin-bottom: 0; }}
.manifesto-body {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 10.5pt; line-height: 1.72; color: rgba(255,255,255,0.70);
  max-width: 145mm; margin-bottom: 11mm;
}}
.manifesto-footer {{
  border-top: 0.3mm solid rgba(255,255,255,0.18); padding-top: 5mm;
  display: flex; justify-content: space-between;
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7pt; color: rgba(255,255,255,0.35);
}}

/* ═══════════════════════════════════════════
   CLOSING
   ═══════════════════════════════════════════ */
.closing-page {{
  background: linear-gradient(158deg, var(--wash) 0%, white 50%, var(--light) 100%);
  padding: 0;
}}
.closing-topbar {{
  height: 4.5mm;
  background: linear-gradient(90deg, var(--primary) 0%, var(--accent) 60%, var(--mid) 100%);
}}
.closing-inner {{ padding: 15mm 24mm 24mm; }}
.closing-takeaways {{ margin-bottom: 10mm; }}
.closing-takeaway {{
  display: flex; align-items: flex-start; margin-bottom: 5.5mm;
}}
.closing-dot {{
  min-width: 7mm; height: 7mm; border-radius: 50%;
  background: linear-gradient(135deg, var(--primary), var(--accent));
  margin-right: 6mm; margin-top: 1.5mm; flex-shrink: 0;
  box-shadow: 0 2px 6px rgba(0,0,0,0.18);
}}
.closing-takeaway p {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 11pt; line-height: 1.60; color: var(--text);
}}
.closing-cta {{
  background: linear-gradient(138deg, var(--primary), var(--accent));
  border-radius: 6mm; padding: 9mm 14mm; color: white;
  box-shadow: 0 4px 18px rgba(0,0,0,0.18);
}}
.closing-cta p {{
  font-family: Georgia, serif; font-size: 13pt; font-style: italic;
  line-height: 1.55;
}}
.closing-tk-header {{
  font-family: 'Helvetica Neue', Arial, sans-serif;
  font-size: 7.5pt; font-weight: 700; letter-spacing: 0.20em;
  text-transform: uppercase; color: var(--accent); margin-bottom: 7mm;
}}
"""

    pages_html = []

    for idx, page in enumerate(structure["pages"][:page_count], start=1):
        ptype      = str(page.get("type","split")).lower()
        eyebrow    = escape(str(page.get("eyebrow") or f"SECTION {idx}"))
        heading    = escape(str(page.get("heading") or title))
        body       = escape(str(page.get("body") or ""))
        callout    = escape(str(page.get("callout") or ""))
        highlights = [escape(str(h)) for h in (page.get("highlights") or [])[:5]]
        tiles      = page.get("tiles") or []
        steps      = page.get("steps") or []
        stats      = page.get("stats") or []
        takeaways  = page.get("takeaways") or []
        statements = page.get("statements") or []

        # ── COVER ──────────────────────────────────────────────────────────
        if ptype == "cover":
            pages_html.append(f"""
<section class="page cover">
  {_svg_cover_geometric(P['primary'], P['accent'], P['mid'])}
  {_svg_noise_overlay('rgba(255,255,255,0.8)')}
  <div class="cover-inner">
    <div class="cover-label">{eyebrow}</div>
    <h1 class="cover-title">{heading}</h1>
    <p class="cover-subtitle">{subtitle}</p>
  </div>
  <div class="cover-bottom">
    <div class="cover-rule"></div>
    <div class="cover-meta">
      <div>
        <p class="cover-author">{author}</p>
        <p class="cover-callout">{callout}</p>
      </div>
      <div class="cover-num">01</div>
    </div>
  </div>
</section>""")

        # ── SPLIT ──────────────────────────────────────────────────────────
        elif ptype == "split":
            cards = "".join(f'<div class="split-card"><p>{h}</p></div>' for h in highlights)
            pages_html.append(f"""
<section class="page split">
  {_svg_grid_dots(P['soft'])}
  <div class="split-left">
    {_eyebrow(eyebrow, P['accent'])}
    <h2 style="font-size:26pt;color:{P['primary']};margin-bottom:6mm;">{heading}</h2>
    <div class="split-accent-bar"></div>
    <p class="body-text" style="flex:1;margin-bottom:9mm;">{body}</p>
    <div class="split-blockquote">
      <p class="callout-italic">{callout}</p>
    </div>
    {_footer(title, idx, page_count, P['accent'], P['mid'])}
  </div>
  <div class="split-right">
    <div class="split-hl-label">Key Highlights</div>
    {cards}
    <div class="split-num">{idx:02d}</div>
  </div>
</section>""")

        # ── FEATURE ────────────────────────────────────────────────────────
        elif ptype == "feature":
            n = min(len(highlights), 3) or 3
            cards = "".join(
                f'<div class="feature-card"><div class="feature-card-num">0{i+1}</div><p>{h}</p></div>'
                for i, h in enumerate(highlights[:3])
            )
            pages_html.append(f"""
<section class="page feature">
  {_svg_half_circle(P['accent'], 'right')}
  {_svg_hexagons(P['light'])}
  <div class="feature-ghost-num">{idx:02d}</div>
  <div class="feature-inner">
    {_eyebrow(eyebrow, P['accent'])}
    <div class="feature-band">
      <h2>{heading}</h2>
      <p class="feature-band-sub">{callout}</p>
    </div>
    <p class="body-text" style="margin-bottom:10mm;max-width:155mm;">{body}</p>
    <div class="feature-cards" style="grid-template-columns:repeat({n},1fr);">
      {cards}
    </div>
    <div class="feature-callout-bar">
      <div class="feature-callout-pip"></div>
      <p style="font-family:Georgia,serif;font-size:11pt;font-weight:600;
        color:{P['primary']};font-style:italic;">{callout}</p>
    </div>
  </div>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

        # ── GRID ───────────────────────────────────────────────────────────
        elif ptype == "grid":
            tile_items = tiles[:4] or [{"label": "Key Point", "value": h} for h in highlights[:4]]
            tile_html = "".join(f"""
<div class="grid-tile">
  <div class="grid-tile-label">{escape(str(t.get('label','Key Fact')))}</div>
  <div class="grid-tile-val">{escape(str(t.get('value','—')))}</div>
</div>""" for t in tile_items)
            pages_html.append(f"""
<section class="page grid-page">
  {_svg_circles(P['accent'], P['light'], P['soft'])}
  {_svg_diagonal_lines(P['mid'])}
  <div style="position:relative;z-index:1;">
    {_eyebrow(eyebrow, P['accent'])}
    <h2 style="font-size:27pt;color:{P['primary']};margin-bottom:5mm;">{heading}</h2>
    <div style="width:22mm;height:1mm;background:{P['accent']};border-radius:1mm;margin-bottom:7mm;"></div>
    <p class="body-text" style="margin-bottom:9mm;max-width:155mm;">{body}</p>
    <div class="grid-tiles">{tile_html}</div>
    <div class="grid-callout-bar">
      <p style="font-family:Georgia,serif;font-size:11.5pt;font-style:italic;color:white;">{callout}</p>
    </div>
  </div>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

        # ── QUOTE ──────────────────────────────────────────────────────────
        elif ptype == "quote":
            pills = "".join(f'<span class="quote-pill">{h}</span>' for h in highlights)
            pages_html.append(f"""
<section class="page quote-page">
  {_svg_arc_decoration(P['soft'])}
  {_svg_noise_overlay('rgba(255,255,255,0.5)')}
  <div class="quote-deco-open">\u201C</div>
  <div class="quote-deco-close">\u201D</div>
  <div style="position:relative;z-index:1;">
    {_eyebrow(eyebrow, P['muted'], light=True)}
    <div class="quote-box">
      <p>{callout}</p>
    </div>
    <p class="quote-body">{body}</p>
  </div>
  <div style="position:relative;z-index:1;">
    <div class="quote-pills">{pills}</div>
    <div class="quote-footer">
      <span>{title}</span><span>{idx}&thinsp;/&thinsp;{page_count}</span>
    </div>
  </div>
</section>""")

        # ── TIMELINE ───────────────────────────────────────────────────────
        elif ptype == "timeline":
            step_list = steps[:4] or [
                {"step": f"Phase {i+1}", "desc": highlights[i] if i < len(highlights) else "Key phase"}
                for i in range(3)
            ]
            steps_html = ""
            for si, s in enumerate(step_list):
                connector = '<div class="timeline-connector"></div>' if si < len(step_list)-1 else ""
                steps_html += f"""
<div class="timeline-step">
  <div class="timeline-node">
    <div class="timeline-circle">{si+1}</div>
    {connector}
  </div>
  <div class="timeline-content">
    <div class="timeline-step-label">{escape(str(s.get('step','')))}</div>
    <p class="timeline-step-desc">{escape(str(s.get('desc','')))}</p>
  </div>
</div>"""
            pages_html.append(f"""
<section class="page timeline-page">
  <div class="timeline-sidebar"></div>
  <div class="timeline-sidebar-accent"></div>
  {_svg_cross_hatch(P['light'])}
  <div class="timeline-inner">
    {_eyebrow(eyebrow, P['accent'])}
    <h2 style="font-size:26pt;color:{P['primary']};margin-bottom:6mm;">{heading}</h2>
    <p class="body-text" style="margin-bottom:9mm;max-width:130mm;">{body}</p>
    {steps_html}
    <div class="timeline-callout">
      <p class="callout-italic">{callout}</p>
    </div>
  </div>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

        # ── STATS ──────────────────────────────────────────────────────────
        elif ptype == "stats":
            stat_list = stats[:3] or [
                {"number": "—", "label": h, "sub": ""} for h in highlights[:3]
            ]
            stat_cards = "".join(f"""
<div class="stats-card">
  <div class="stats-num">{escape(str(s.get('number','—')))}</div>
  <div class="stats-label">{escape(str(s.get('label','Metric')))}</div>
  <div class="stats-sub">{escape(str(s.get('sub','')))}</div>
</div>""" for s in stat_list)
            pills = "".join(f'<span class="stats-pill">{h}</span>' for h in highlights)
            pages_html.append(f"""
<section class="page stats-page">
  {_svg_stripe_band(P['primary'], P['accent'])}
  {_svg_arc_decoration(P['soft'])}
  <div style="position:relative;z-index:1;">
    {_eyebrow(eyebrow, P['accent'])}
    <h2 style="font-size:27pt;color:{P['primary']};margin-bottom:11mm;">{heading}</h2>
    <div class="stats-cards">{stat_cards}</div>
    <p class="body-text" style="max-width:152mm;margin-bottom:0;">{body}</p>
    <div class="stats-pills">{pills}</div>
  </div>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

        # ── EDITORIAL ──────────────────────────────────────────────────────
        elif ptype == "editorial":
            hi_li = "".join(f'<li>{h}</li>' for h in highlights)
            pages_html.append(f"""
<section class="page editorial-page">
  {_svg_half_circle(P['accent'], 'left')}
  <div class="editorial-topband"></div>
  <div class="editorial-body-grid" style="position:relative;z-index:1;">
    <div class="editorial-left">
      {_eyebrow(eyebrow, P['accent'])}
      <div class="editorial-rule"></div>
      <h2>{heading}</h2>
      <ul class="editorial-highlight-list">{hi_li}</ul>
    </div>
    <div class="editorial-right">
      <p class="body-text">{body}</p>
      <div style="background:{P['wash']};border-radius:5mm;padding:7mm 10mm;
        border-left:3.5mm solid {P['accent']};">
        <p class="callout-italic">{callout}</p>
      </div>
    </div>
  </div>
  <div class="editorial-callout-band" style="position:relative;z-index:1;">
    <p>{title}&nbsp;&nbsp;·&nbsp;&nbsp;{eyebrow}&nbsp;&nbsp;·&nbsp;&nbsp;{idx}&thinsp;/&thinsp;{page_count}</p>
  </div>
</section>""")

        # ── MANIFESTO ──────────────────────────────────────────────────────
        elif ptype == "manifesto":
            stmt_list = statements[:3] or [callout] + highlights[:2]
            stmts_html = "".join(
                f'<div class="manifesto-statement">{escape(str(s))}</div>'
                for s in stmt_list
            )
            pages_html.append(f"""
<section class="page manifesto-page">
  {_svg_manifesto_bg(P['dark'], P['accent'])}
  {_svg_grid_dots('rgba(255,255,255,0.22)')}
  <div style="position:relative;z-index:1;">
    {_eyebrow(eyebrow, P['muted'], light=True)}
    <div class="manifesto-statements">{stmts_html}</div>
    <p class="manifesto-body">{body}</p>
    <div class="manifesto-footer">
      <span>{title}</span><span>{idx}&thinsp;/&thinsp;{page_count}</span>
    </div>
  </div>
</section>""")

        # ── CLOSING ────────────────────────────────────────────────────────
        elif ptype == "closing":
            tk_list = takeaways[:4] or highlights[:4]
            tk_html = "".join(f"""
<div class="closing-takeaway">
  <div class="closing-dot"></div>
  <p>{escape(str(t))}</p>
</div>""" for t in tk_list)
            pages_html.append(f"""
<section class="page closing-page">
  {_svg_diagonal_band(P['light'], 0, 530, 595, 470, 595, 842, 0, 842)}
  {_svg_arc_decoration(P['soft'])}
  <div class="closing-topbar"></div>
  <div class="closing-inner" style="position:relative;z-index:1;">
    {_eyebrow(eyebrow, P['accent'])}
    <h2 style="font-size:30pt;color:{P['primary']};margin-bottom:9mm;">{heading}</h2>
    <p class="body-text" style="margin-bottom:10mm;max-width:146mm;">{body}</p>
    <div style="background:white;border-radius:7mm;padding:11mm;
      border:1.5px solid {P['light']};margin-bottom:10mm;
      box-shadow:0 2px 12px rgba(0,0,0,0.06);">
      <div class="closing-tk-header">Key Takeaways</div>
      <div class="closing-takeaways">{tk_html}</div>
    </div>
    <div class="closing-cta">
      <p>{callout}</p>
    </div>
  </div>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

        # ── FALLBACK ───────────────────────────────────────────────────────
        else:
            hi_li = "".join(
                f'<li style="padding:4mm 0;border-bottom:0.3mm solid {P["light"]};">→ {h}</li>'
                for h in highlights
            )
            pages_html.append(f"""
<section class="page" style="background:white;padding:20mm 24mm;">
  {_eyebrow(eyebrow, P['accent'])}
  <h2 style="font-size:27pt;color:{P['primary']};margin-bottom:9mm;">{heading}</h2>
  <div style="width:20mm;height:1mm;background:{P['accent']};border-radius:1mm;margin-bottom:8mm;"></div>
  <p class="body-text" style="margin-bottom:10mm;">{body}</p>
  <ul style="list-style:none;padding:0;margin-bottom:10mm;
    font-family:'Helvetica Neue',Arial,sans-serif;font-size:10.5pt;color:{P['text']};">{hi_li}</ul>
  <blockquote style="border-left:4mm solid {P['accent']};padding:6mm 10mm;
    background:{P['wash']};border-radius:0 5mm 5mm 0;">
    <p class="callout-italic">{callout}</p>
  </blockquote>
  {_footer(title, idx, page_count, P['accent'], P['mid'])}
</section>""")

    # WeasyPrint: pass unbreakable CSS via the CSS object for robust pagination
    pagination_css = CSS(string="""
        @page { size: A4; margin: 0; }
        .page { page-break-after: always; page-break-inside: avoid; }
        p, li { orphans: 3; widows: 3; }
        .split-card, .feature-card, .grid-tile, .stats-card, .timeline-step,
        .closing-takeaway, .quote-box, .manifesto-statement,
        .feature-callout-bar, .grid-callout-bar, .closing-cta,
        .timeline-callout, .feature-cards, .stats-cards, .grid-tiles {
            page-break-inside: avoid;
        }
        h1, h2, h3, .eyebrow { page-break-after: avoid; }
    """)

    html_doc = f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<style>{global_css}</style>
</head><body>{"".join(pages_html)}</body></html>"""

    return html_doc, pagination_css


@app.post("/docs/generate/pdf")
async def generate_pdf(payload: dict):
    prompt = payload["prompt"]
    page_count = requested_count(prompt, "page", default=5, minimum=1, maximum=15)
    structure = await call_ai_pdf(prompt, page_count)
    html_doc, pagination_css = render_pdf_html(structure, page_count)
    buffer = io.BytesIO()
    HTML(string=html_doc).write_pdf(buffer, stylesheets=[pagination_css])
    buffer.seek(0)
    blob = await upload_to_vercel_blob(buffer, "output.pdf", "application/pdf")
    return blob


# ══════════════════════════════════════════════════════════════════════════════
# PPTX GENERATION
# ══════════════════════════════════════════════════════════════════════════════

async def call_ai_pptx(prompt: str, slide_count: int) -> dict:
    system = f"""You are an elite presentation strategist. Return ONLY valid JSON, no markdown fences.
Generate content for exactly {slide_count} slides. slides[0] is ALWAYS the title slide.

PALETTE OPTIONS (pick one that fits the topic):
midnight=deep navy professional, forest=green nature, coral=energetic warm, 
terracotta=warm cultural, ocean=deep blue tech, charcoal=minimal elegant,
berry=premium luxury, cherry=bold dramatic

SLIDE TYPES:
- "title"      → slides[0] only: big title, subtitle, tagline
- "bullets"    → standard: heading + 4-5 rich bullet points  
- "two_col"    → heading + left_points list + right_points list (two equal columns)
- "stat"       → heading + 3 big stat cards (number + label + context)
- "quote"      → big pull-quote + attribution + supporting points
- "image_text" → heading + body paragraph + highlight sidebar
- "closing"    → final slide: closing headline + call to action + tagline

JSON schema:
{{
  "title": "Deck title",
  "subtitle": "Subtitle line",
  "palette": "<palette name>",
  "slides": [
    {{
      "type": "title|bullets|two_col|stat|quote|image_text|closing",
      "title": "Slide headline",
      "subtitle": "Only for title slide",
      "tagline": "Only for title or closing slide",
      "bullets": ["Bullet with enough detail to be useful (15+ words each)", "..."],
      "left_points": ["Point A", "Point B", "Point C"],
      "right_points": ["Point D", "Point E", "Point F"],
      "stats": [{{"number":"big value","label":"what it is","context":"one sentence"}}],
      "quote": "Full memorable quote text",
      "attribution": "— Source or speaker",
      "body": "Rich paragraph for image_text slides",
      "highlights": ["Sidebar point 1", "Sidebar point 2", "Sidebar point 3"],
      "cta": "Call to action text for closing slide"
    }}
  ]
}}

RULES:
- slides[0].type MUST be "title"
- slides[{slide_count-1}].type MUST be "closing"
- Interior slides: mix bullets, two_col, stat, quote, image_text — never same type twice in a row
- stat slides: always exactly 3 stat objects
- two_col: always 3 items in left_points and 3 in right_points
- bullets: 4-5 items, each 15+ words
- The slides array MUST contain exactly {slide_count} items"""

    raw = await _deepinfra_call([
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ], temperature=0.4)
    data = json.loads(raw)
    slides = data.get("slides", [])
    while len(slides) < slide_count:
        i = len(slides)
        slides.append({
            "type": "bullets", "title": f"Key Insight {i+1}",
            "bullets": ["This slide presents essential detail on the topic with clear, actionable perspective.",
                        "Each point is crafted to deliver maximum value to the audience.",
                        "Visual clarity and content depth define this presentation's approach.",
                        "Strategic thinking drives every recommendation made here."]
        })
    data["slides"] = slides[:slide_count]
    return data


PPTX_PALETTES = {
    "midnight":  {"bg": (0x1E,0x27,0x61), "primary": (0x1E,0x27,0x61), "accent": (0x49,0x8F,0xD6), "light": (0xCA,0xDC,0xFC), "white": (0xFF,0xFF,0xFF), "gray": (0x8A,0x9A,0xBF), "card": (0x28,0x34,0x76)},
    "forest":    {"bg": (0x2C,0x5F,0x2D), "primary": (0x2C,0x5F,0x2D), "accent": (0x97,0xBC,0x62), "light": (0xD4,0xED,0xC1), "white": (0xFF,0xFF,0xFF), "gray": (0x7A,0x9A,0x5A), "card": (0x36,0x72,0x37)},
    "coral":     {"bg": (0xF9,0x61,0x67), "primary": (0x2F,0x3C,0x7E), "accent": (0xF9,0x61,0x67), "light": (0xFD,0xE0,0xE0), "white": (0xFF,0xFF,0xFF), "gray": (0x88,0x77,0xAA), "card": (0x3A,0x4A,0x9A)},
    "terracotta":{"bg": (0xB8,0x50,0x42), "primary": (0xB8,0x50,0x42), "accent": (0xE7,0xE8,0xD1), "light": (0xF5,0xF0,0xE8), "white": (0xFF,0xFF,0xFF), "gray": (0xA7,0xBE,0xAE), "card": (0xD4,0x62,0x52)},
    "ocean":     {"bg": (0x06,0x5A,0x82), "primary": (0x06,0x5A,0x82), "accent": (0x02,0xC3,0x9A), "light": (0xC8,0xF0,0xE8), "white": (0xFF,0xFF,0xFF), "gray": (0x7A,0xB8,0xD4), "card": (0x1C,0x72,0x93)},
    "charcoal":  {"bg": (0x36,0x45,0x4F), "primary": (0x36,0x45,0x4F), "accent": (0x7A,0xC0,0xD4), "light": (0xE8,0xF2,0xF5), "white": (0xFF,0xFF,0xFF), "gray": (0x8A,0x9A,0xA8), "card": (0x48,0x58,0x64)},
    "berry":     {"bg": (0x6D,0x2E,0x46), "primary": (0x6D,0x2E,0x46), "accent": (0xEC,0xE2,0xD0), "light": (0xF5,0xEC,0xE5), "white": (0xFF,0xFF,0xFF), "gray": (0xA2,0x67,0x69), "card": (0x84,0x3A,0x58)},
    "cherry":    {"bg": (0x99,0x00,0x11), "primary": (0x99,0x00,0x11), "accent": (0xFC,0xF6,0xF5), "light": (0xFF,0xE8,0xE8), "white": (0xFF,0xFF,0xFF), "gray": (0xCC,0x88,0x88), "card": (0xB0,0x10,0x20)},
}


def _rgb(t: tuple) -> PPTXRGBColor:
    return PPTXRGBColor(*t)

def _solid(shape, color: PPTXRGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

def _add_rect(slide, l, t, w, h, color, line=False):
    from pptx.util import Inches as I, Pt
    shape = slide.shapes.add_shape(1, I(l), I(t), I(w), I(h))
    _solid(shape, color)
    if not line:
        shape.line.fill.background()
    return shape

def _txt(slide, text, l, t, w, h, size=16, bold=False, color=None, align=PP_ALIGN.LEFT,
         italic=False, wrap=True, font="Calibri"):
    tb = slide.shapes.add_textbox(PPTXInches(l), PPTXInches(t), PPTXInches(w), PPTXInches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = PPTXPt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    if color:
        run.font.color.rgb = color
    return tb

def _add_bullets_textbox(slide, bullets, l, t, w, h, size=14, color=None, marker="•",
                          indent_color=None, font="Calibri"):
    tb = slide.shapes.add_textbox(PPTXInches(l), PPTXInches(t), PPTXInches(w), PPTXInches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = PPTXPt(6)
        run = p.add_run()
        run.text = f"{marker}  {b}"
        run.font.size = PPTXPt(size)
        run.font.name = font
        if color:
            run.font.color.rgb = color


@app.post("/docs/generate/pptx")
async def generate_pptx(payload: dict):
    prompt = payload["prompt"]
    slide_count = requested_count(prompt, "slide", default=8, minimum=3, maximum=20)
    structure = await call_ai_pptx(prompt, slide_count)

    prs = Presentation()
    prs.slide_width  = PPTXInches(13.33)
    prs.slide_height = PPTXInches(7.5)

    pname = str(structure.get("palette","midnight")).lower()
    pal = PPTX_PALETTES.get(pname, PPTX_PALETTES["midnight"])
    BG    = _rgb(pal["bg"])
    PRI   = _rgb(pal["primary"])
    ACC   = _rgb(pal["accent"])
    LIGHT = _rgb(pal["light"])
    WHITE = _rgb(pal["white"])
    GRAY  = _rgb(pal["gray"])
    CARD  = _rgb(pal["card"])

    deck_title = structure.get("title","Presentation")

    for si, slide_data in enumerate(structure["slides"][:slide_count]):
        stype  = str(slide_data.get("type","bullets")).lower()
        stitle = str(slide_data.get("title",""))
        slide  = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        if stype == "title":
            _add_rect(slide, 0, 0, 13.33, 7.5, BG)
            _add_rect(slide, 0, 0, 0.18, 7.5, ACC)
            circ = slide.shapes.add_shape(9, PPTXInches(9.8), PPTXInches(-1.5), PPTXInches(5.5), PPTXInches(5.5))
            circ.fill.solid(); circ.fill.fore_color.rgb = CARD; circ.line.fill.background()
            circ2 = slide.shapes.add_shape(9, PPTXInches(10.8), PPTXInches(4.5), PPTXInches(3.5), PPTXInches(3.5))
            circ2.fill.solid(); circ2.fill.fore_color.rgb = _rgb(pal["card"]); circ2.line.fill.background()

            subtitle_txt = str(slide_data.get("subtitle",""))
            tagline_txt  = str(slide_data.get("tagline",""))

            _txt(slide, stitle, 0.5, 1.6, 9.5, 2.8, size=48, bold=True, color=WHITE, align=PP_ALIGN.LEFT, font="Calibri")
            if subtitle_txt:
                _txt(slide, subtitle_txt, 0.5, 4.6, 9, 0.7, size=22, color=LIGHT, align=PP_ALIGN.LEFT, font="Calibri")
            _add_rect(slide, 0, 6.8, 13.33, 0.7, CARD)
            if tagline_txt:
                _txt(slide, tagline_txt, 0.5, 6.82, 10, 0.55, size=12, color=GRAY, align=PP_ALIGN.LEFT, font="Calibri")
            _txt(slide, f"01/{slide_count:02d}", 12.1, 6.82, 1.0, 0.55, size=11, color=GRAY, align=PP_ALIGN.RIGHT, font="Calibri")

        elif stype == "closing":
            _add_rect(slide, 0, 0, 13.33, 7.5, BG)
            _add_rect(slide, 0, 0, 0.18, 7.5, ACC)
            circ = slide.shapes.add_shape(9, PPTXInches(8), PPTXInches(0.5), PPTXInches(6), PPTXInches(6))
            circ.fill.solid(); circ.fill.fore_color.rgb = CARD; circ.line.fill.background()

            cta = str(slide_data.get("cta","Thank you."))
            tagline = str(slide_data.get("tagline",""))

            _txt(slide, "CLOSING", 0.5, 0.6, 5, 0.4, size=10, color=ACC, bold=True, font="Calibri")
            _txt(slide, stitle, 0.5, 1.2, 9.5, 2.2, size=40, bold=True, color=WHITE, font="Calibri")
            cta_box = slide.shapes.add_shape(1, PPTXInches(0.5), PPTXInches(3.7), PPTXInches(7.5), PPTXInches(1.0))
            _solid(cta_box, ACC)
            ctf = cta_box.text_frame
            ctf.paragraphs[0].text = cta
            ctf.paragraphs[0].alignment = PP_ALIGN.CENTER
            ctf.paragraphs[0].runs[0].font.size = PPTXPt(18)
            ctf.paragraphs[0].runs[0].font.bold = True
            ctf.paragraphs[0].runs[0].font.color.rgb = BG
            ctf.paragraphs[0].runs[0].font.name = "Calibri"
            if tagline:
                _txt(slide, tagline, 0.5, 5.0, 9, 0.6, size=14, color=LIGHT, italic=True, font="Calibri")
            _add_rect(slide, 0, 6.8, 13.33, 0.7, CARD)
            _txt(slide, deck_title, 0.5, 6.82, 10, 0.55, size=11, color=GRAY, font="Calibri")
            _txt(slide, f"{si+1:02d}/{slide_count:02d}", 12.1, 6.82, 1.0, 0.55, size=11, color=GRAY, align=PP_ALIGN.RIGHT, font="Calibri")

        elif stype == "bullets":
            _add_rect(slide, 0, 0, 13.33, 7.5, WHITE)
            _add_rect(slide, 0, 0, 13.33, 1.5, PRI)
            _txt(slide, stitle, 0.5, 0.2, 12, 1.1, size=28, bold=True, color=WHITE, font="Calibri")
            bullets = slide_data.get("bullets", [])
            _add_bullets_textbox(slide, bullets, 0.5, 1.8, 12.3, 5.3, size=15, color=PRI, marker="▸", font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, ACC)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

        elif stype == "two_col":
            _add_rect(slide, 0, 0, 13.33, 7.5, _rgb((0xF8,0xFA,0xFC)))
            _add_rect(slide, 0, 0, 13.33, 1.35, PRI)
            _txt(slide, stitle, 0.45, 0.18, 12.4, 1.0, size=26, bold=True, color=WHITE, font="Calibri")
            lcard = slide.shapes.add_shape(1, PPTXInches(0.35), PPTXInches(1.55), PPTXInches(6.1), PPTXInches(5.65))
            _solid(lcard, WHITE)
            _add_rect(slide, 0.35, 1.55, 6.1, 0.18, ACC)
            left_pts = slide_data.get("left_points", [])
            _add_bullets_textbox(slide, left_pts, 0.6, 1.9, 5.7, 5.0, size=14, color=_rgb((0x1E,0x29,0x3B)), marker="→", font="Calibri")
            rcard = slide.shapes.add_shape(1, PPTXInches(6.9), PPTXInches(1.55), PPTXInches(6.1), PPTXInches(5.65))
            _solid(rcard, _rgb(pal["card"]))
            right_pts = slide_data.get("right_points", [])
            _add_bullets_textbox(slide, right_pts, 7.1, 1.9, 5.7, 5.0, size=14, color=WHITE, marker="→", font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, ACC)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

        elif stype == "stat":
            _add_rect(slide, 0, 0, 13.33, 7.5, BG)
            _add_rect(slide, 0, 0, 0.18, 7.5, ACC)
            _txt(slide, stitle, 0.45, 0.3, 12.4, 1.0, size=30, bold=True, color=WHITE, font="Calibri")
            _add_rect(slide, 0.45, 1.35, 8, 0.06, ACC)
            stats_data = slide_data.get("stats", [])[:3]
            card_w, card_h = 3.9, 4.0
            starts = [0.45, 4.6, 8.75]
            for ci, st in enumerate(stats_data):
                cx = starts[ci]
                card = slide.shapes.add_shape(1, PPTXInches(cx), PPTXInches(1.65), PPTXInches(card_w), PPTXInches(card_h))
                _solid(card, CARD)
                _add_rect(slide, cx, 1.65, card_w, 0.2, ACC)
                num  = str(st.get("number","—"))
                lbl  = str(st.get("label",""))
                ctx  = str(st.get("context",""))
                _txt(slide, num, cx+0.2, 2.05, card_w-0.4, 1.4, size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Calibri")
                _txt(slide, lbl, cx+0.2, 3.55, card_w-0.4, 0.6, size=13, bold=True, color=ACC, align=PP_ALIGN.CENTER, font="Calibri")
                _txt(slide, ctx, cx+0.2, 4.2, card_w-0.4, 1.3, size=11, color=LIGHT, align=PP_ALIGN.CENTER, font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, CARD)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

        elif stype == "quote":
            _add_rect(slide, 0, 0, 13.33, 7.5, BG)
            _add_rect(slide, 0, 0, 0.18, 7.5, ACC)
            _txt(slide, "\u201C", 0.4, -0.3, 3, 2.5, size=120, color=CARD, bold=True, font="Georgia")
            quote_txt  = str(slide_data.get("quote",""))
            attrib_txt = str(slide_data.get("attribution",""))
            qbox = slide.shapes.add_shape(1, PPTXInches(0.8), PPTXInches(1.2), PPTXInches(9.5), PPTXInches(2.8))
            _solid(qbox, CARD)
            _txt(slide, quote_txt, 1.1, 1.4, 8.9, 2.5, size=20, color=WHITE, italic=True, font="Georgia")
            if attrib_txt:
                _txt(slide, attrib_txt, 1.1, 4.2, 9, 0.5, size=13, color=ACC, bold=True, font="Calibri")
            hi = slide_data.get("highlights",[]) or slide_data.get("bullets",[])
            for hj, h in enumerate(hi[:3]):
                hy = 1.5 + hj * 1.6
                hbox = slide.shapes.add_shape(1, PPTXInches(10.6), PPTXInches(hy), PPTXInches(2.4), PPTXInches(1.3))
                _solid(hbox, CARD)
                _txt(slide, str(h), 10.7, hy+0.1, 2.2, 1.1, size=11, color=WHITE, font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, CARD)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

        elif stype == "image_text":
            _add_rect(slide, 0, 0, 13.33, 7.5, _rgb((0xF8,0xFA,0xFC)))
            _add_rect(slide, 0, 0, 13.33, 1.35, PRI)
            _txt(slide, stitle, 0.45, 0.18, 12.4, 1.0, size=26, bold=True, color=WHITE, font="Calibri")
            body_txt = str(slide_data.get("body",""))
            _txt(slide, body_txt, 0.45, 1.6, 8.5, 5.5, size=14, color=_rgb((0x1E,0x29,0x3B)), font="Calibri")
            sbox = slide.shapes.add_shape(1, PPTXInches(9.3), PPTXInches(1.55), PPTXInches(3.7), PPTXInches(5.65))
            _solid(sbox, _rgb(pal["card"]))
            _add_rect(slide, 9.3, 1.55, 3.7, 0.2, ACC)
            _txt(slide, "KEY POINTS", 9.4, 1.6, 3.5, 0.5, size=10, bold=True, color=ACC, font="Calibri")
            hi_pts = slide_data.get("highlights",[])[:4]
            _add_bullets_textbox(slide, hi_pts, 9.4, 2.15, 3.5, 4.9, size=12, color=WHITE, marker="◆", font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, ACC)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

        else:
            _add_rect(slide, 0, 0, 13.33, 7.5, WHITE)
            _add_rect(slide, 0, 0, 13.33, 1.35, PRI)
            _txt(slide, stitle, 0.45, 0.18, 12.4, 1.0, size=26, bold=True, color=WHITE, font="Calibri")
            bullets = slide_data.get("bullets",[])
            _add_bullets_textbox(slide, bullets, 0.5, 1.7, 12.3, 5.3, size=14, color=_rgb((0x1E,0x29,0x3B)), font="Calibri")
            _add_rect(slide, 0, 7.2, 13.33, 0.3, ACC)
            _txt(slide, f"{si+1:02d}/{slide_count:02d}  •  {deck_title}", 0.5, 7.22, 12, 0.28, size=9, color=GRAY, font="Calibri")

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    blob = await upload_to_vercel_blob(
        buffer, "output.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    return blob


# ══════════════════════════════════════════════════════════════════════════════
# DOCX GENERATION
# ══════════════════════════════════════════════════════════════════════════════

async def call_ai_docx(prompt: str) -> dict:
    system = """You are a professional document writer and editor. Return ONLY valid JSON, no markdown fences.

Create a polished, comprehensive document. Each section should be detailed and professional.

JSON schema:
{
  "title": "Document Title",
  "subtitle": "Compelling subtitle",
  "author": "Prepared by AI Document Service",
  "date": "2025",
  "abstract": "A 2-3 sentence executive summary or abstract of the whole document.",
  "palette": "indigo|teal|crimson|emerald|slate|violet|amber|rose",
  "sections": [
    {
      "heading": "Section Heading",
      "body": "Rich detailed paragraph content. At least 80 words per section with specific, valuable information.",
      "callout": "Optional: one important pull-quote or highlighted note for this section.",
      "subsections": [
        { "heading": "Subsection Heading", "body": "Subsection content, at least 40 words." }
      ]
    }
  ],
  "conclusion": "A compelling 2-3 sentence conclusion that ties everything together."
}

RULES:
- Every section body: 80+ words
- 4-7 sections minimum
- subsections: optional, 0-3 per section
- palette: pick one that emotionally fits the topic"""

    raw = await _deepinfra_call([
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ], temperature=0.38)
    return json.loads(raw)


def _set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.lstrip('#'))
    tcPr.append(shd)


def _add_docx_heading_band(doc: Document, text: str, bg_hex: str, fg_hex: str, level_pt: int = 14):
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    _set_cell_bg(cell, bg_hex)
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after  = Pt(5)
    p.paragraph_format.left_indent  = Inches(0.15)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(level_pt)
    run.font.color.rgb = RGBColor(
        int(fg_hex[1:3],16), int(fg_hex[3:5],16), int(fg_hex[5:7],16)
    )
    doc.add_paragraph()


def _add_pull_quote(doc: Document, text: str, accent_hex: str):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Inches(0.4)
    p.paragraph_format.right_indent = Inches(0.4)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(8)
    run = p.add_run(f'"{text}"')
    run.italic = True
    run.font.size = Pt(12)
    r,g,b = int(accent_hex[1:3],16),int(accent_hex[3:5],16),int(accent_hex[5:7],16)
    run.font.color.rgb = RGBColor(r,g,b)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '24')
    left.set(qn('w:space'), '12')
    left.set(qn('w:color'), accent_hex.lstrip('#'))
    pBdr.append(left)
    pPr.append(pBdr)


@app.post("/docs/generate/docx")
async def generate_docx(payload: dict):
    structure = await call_ai_docx(payload["prompt"])

    pname = str(structure.get("palette","indigo")).lower()
    pal   = PALETTES.get(pname, PALETTES["indigo"])
    PRI_HEX  = pal["primary"]
    ACC_HEX  = pal["accent"]

    def hexrgb(h):
        h = h.lstrip('#')
        return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

    PRI_RGB  = hexrgb(PRI_HEX)
    ACC_RGB  = hexrgb(ACC_HEX)
    GRAY_RGB = RGBColor(0x6B,0x72,0x80)
    BODY_RGB = RGBColor(0x1F,0x2A,0x3C)

    doc = Document()

    for sec in doc.sections:
        sec.top_margin    = Inches(1.1)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin   = Inches(1.2)
        sec.right_margin  = Inches(1.2)

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_before = Pt(40)
    title_para.paragraph_format.space_after  = Pt(10)
    tr = title_para.add_run(structure.get("title","Document"))
    tr.bold = True
    tr.font.size = Pt(32)
    tr.font.color.rgb = PRI_RGB

    if structure.get("subtitle"):
        sp = doc.add_paragraph()
        sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp.paragraph_format.space_after = Pt(6)
        sr = sp.add_run(structure["subtitle"])
        sr.italic = True
        sr.font.size = Pt(15)
        sr.font.color.rgb = ACC_RGB

    sep_table = doc.add_table(rows=1, cols=1)
    sep_table.style = 'Table Grid'
    sep_cell = sep_table.cell(0,0)
    _set_cell_bg(sep_cell, ACC_HEX.lstrip('#'))
    sep_tbl = sep_table._tbl
    sep_tblPr = sep_tbl.find(qn('w:tblPr')) or OxmlElement('w:tblPr')
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '3000')
    tblW.set(qn('w:type'), 'dxa')
    sep_tblPr.append(tblW)
    sep_cell.paragraphs[0].paragraph_format.space_before = Pt(1)
    sep_cell.paragraphs[0].paragraph_format.space_after  = Pt(1)

    doc.add_paragraph()

    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_para.paragraph_format.space_after = Pt(4)
    mr = meta_para.add_run(structure.get("author",""))
    mr.font.size = Pt(11)
    mr.font.color.rgb = GRAY_RGB

    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dr = date_para.add_run(structure.get("date",""))
    dr.font.size = Pt(11)
    dr.font.color.rgb = GRAY_RGB

    doc.add_page_break()

    if structure.get("abstract"):
        _add_docx_heading_band(doc, "Executive Summary", PRI_HEX, "#FFFFFF", level_pt=14)
        abs_para = doc.add_paragraph(structure["abstract"])
        abs_para.paragraph_format.space_after = Pt(10)
        abs_para.paragraph_format.left_indent = Inches(0.1)
        for run in abs_para.runs:
            run.font.size = Pt(11.5)
            run.font.color.rgb = BODY_RGB
            run.italic = True
        doc.add_paragraph()

    for si, sec in enumerate(structure.get("sections", [])):
        is_alt = (si % 2 == 1)
        h_bg  = ACC_HEX if is_alt else PRI_HEX
        _add_docx_heading_band(doc, sec["heading"], h_bg, "#FFFFFF", level_pt=14)

        body_para = doc.add_paragraph(sec.get("body",""))
        body_para.paragraph_format.space_after  = Pt(10)
        body_para.paragraph_format.space_before = Pt(2)
        body_para.paragraph_format.left_indent  = Inches(0.1)
        for run in body_para.runs:
            run.font.size = Pt(11)
            run.font.color.rgb = BODY_RGB

        if sec.get("callout"):
            _add_pull_quote(doc, sec["callout"], ACC_HEX)

        for sub in sec.get("subsections", []):
            sub_para = doc.add_paragraph()
            sub_para.paragraph_format.space_before = Pt(10)
            sub_para.paragraph_format.space_after  = Pt(4)
            sub_para.paragraph_format.left_indent  = Inches(0.1)
            sub_run = sub_para.add_run(sub["heading"])
            sub_run.bold = True
            sub_run.font.size = Pt(12.5)
            sub_run.font.color.rgb = ACC_RGB

            sub_body = doc.add_paragraph(sub.get("body",""))
            sub_body.paragraph_format.space_after = Pt(8)
            sub_body.paragraph_format.left_indent = Inches(0.2)
            for run in sub_body.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = BODY_RGB

        doc.add_paragraph()

    if structure.get("conclusion"):
        _add_docx_heading_band(doc, "Conclusion", PRI_HEX, "#FFFFFF", level_pt=14)
        conc_para = doc.add_paragraph(structure["conclusion"])
        conc_para.paragraph_format.space_after = Pt(10)
        conc_para.paragraph_format.left_indent = Inches(0.1)
        for run in conc_para.runs:
            run.font.size = Pt(11)
            run.font.color.rgb = BODY_RGB
            run.italic = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    blob = await upload_to_vercel_blob(
        buffer, "output.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return blob