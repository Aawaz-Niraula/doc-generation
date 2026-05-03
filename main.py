from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Pt as PPTXPt, Inches as PPTXInches
from pptx.dml.color import RGBColor as PPTXRGBColor
from pptx.enum.text import PP_ALIGN
from weasyprint import HTML
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
import httpx, io, json, os, re, uuid
from html import escape
from vercel.blob import AsyncBlobClient

app = FastAPI()

DEEPINFRA_KEY = os.getenv("DEEPINFRA_KEY")
VERCEL_BLOB_TOKEN = os.getenv("VERCEL_BLOB_TOKEN") or os.getenv("BLOB_READ_WRITE_TOKEN")


# ─── Health check (for cron-job.org keep-alive ping) ────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok"}


# ─── DeepInfra helpers ───────────────────────────────────────────────────────

async def call_deepinfra_json(prompt: str, fmt: str) -> dict:
    """Ask AI for structured JSON content."""
    if fmt in ["docx", "pdf"]:
        schema = """{
  "title": "Document Title",
  "subtitle": "Optional subtitle or description",
  "sections": [
    {
      "heading": "Section Heading",
      "body": "Several paragraphs of detailed content here. Write at least 3-4 sentences per section with rich information.",
      "subsections": [
        { "heading": "Subsection", "body": "Subsection content..." }
      ]
    }
  ]
}"""
    else:
        schema = """{
  "title": "Presentation Title",
  "subtitle": "Presentation subtitle",
  "slides": [
    {
      "title": "Slide Title",
      "bullets": ["Bullet point one", "Bullet point two", "Bullet point three"]
    }
  ]
}"""

    system = f"""You are a professional document writer. Return ONLY valid JSON, no markdown fences, no explanation.
The user will specify how many pages or slides they want — you MUST generate enough content to fill that.
For documents: each section body must be at least 150 words. Add subsections if needed to fill pages.
For presentations: generate exactly as many slides as requested, each with 4-6 bullet points.
Use this exact JSON schema:
{schema}"""

    async with httpx.AsyncClient(timeout=90) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={
                "model": "deepseek-ai/DeepSeek-V4-Flash",
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt}
                ]
            }
        )
    raw = res.json()["choices"][0]["message"]["content"]
    # Strip markdown fences if model slips them in
    raw = re.sub(r"```json|```", "", raw).strip()
    return json.loads(raw)


async def call_deepinfra_html(prompt: str) -> str:
    """Ask AI to generate a beautifully styled HTML document for PDF export."""
    system = """You are a professional document designer. Return ONLY a complete HTML document with inline CSS.
Rules:
- Use beautiful colors, gradients, typography. Make it look like a premium designed report.
- Use a color palette with a strong primary color (deep purple, navy, teal — pick one per document).
- Style headings with colored backgrounds or left border accents.
- Add a styled cover/header section at the top with the title, subtitle, date.
- Use proper spacing, font sizes, line heights.
- The user will specify how many pages — generate ENOUGH text content to fill that many pages when printed.
  A standard A4 page fits roughly 400-500 words. Multiply accordingly and write that much content.
- Use <div style="page-break-before: always"> to force page breaks at appropriate points.
- No external fonts or images — use system fonts like Georgia, Arial, or monospace only.
- Return the full HTML document starting with <!DOCTYPE html>."""

    async with httpx.AsyncClient(timeout=120) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={
                "model": "deepseek-ai/DeepSeek-V4-Flash",
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt}
                ]
            }
        )
    return res.json()["choices"][0]["message"]["content"]


NUMBER_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
}


def requested_count(prompt: str, unit: str, default: int, minimum: int, maximum: int) -> int:
    normalized = prompt.lower()
    match = re.search(r"\b(\d{1,2})\s*(?:-|\s)?" + re.escape(unit) + r"s?\b", normalized)
    if match:
        return max(minimum, min(maximum, int(match.group(1))))
    for word, value in NUMBER_WORDS.items():
        if re.search(r"\b" + word + r"\s+(?:-|\s)?" + re.escape(unit) + r"s?\b", normalized):
            return max(minimum, min(maximum, value))
    return default


async def call_deepinfra_pdf_pages(prompt: str, page_count: int) -> dict:
    """Ask AI for bounded, page-by-page content so PDF page count is deterministic."""
    system = f"""You are a premium editorial report writer. Return ONLY valid JSON, no markdown fences, no explanation.
Create content for exactly {page_count} PDF pages. The renderer will place one JSON page on one physical PDF page.
Keep each page concise: 90-135 words total. Do not overfill pages. No long tables. No markdown.
Make the content polished, specific, visually segmentable, and presentation-grade.
Use this exact schema:
{{
  "title": "Report title",
  "subtitle": "Short subtitle",
  "palette": "indigo" | "teal" | "crimson" | "emerald",
  "pages": [
    {{
      "eyebrow": "Short section label",
      "heading": "Page headline",
      "body": "One compact paragraph with concrete, high-value content.",
      "highlights": ["Short punchy point", "Short punchy point", "Short punchy point"],
      "callout": "One memorable takeaway sentence"
    }}
  ]
}}
The pages array MUST contain exactly {page_count} items."""

    async with httpx.AsyncClient(timeout=120) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={
                "model": "deepseek-ai/DeepSeek-V4-Flash",
                "temperature": 0.35,
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt},
                ],
            },
        )
    raw = res.json()["choices"][0]["message"]["content"]
    raw = re.sub(r"```json|```", "", raw).strip()
    data = json.loads(raw)
    pages = data.get("pages", [])
    if len(pages) < page_count:
        pages.extend({
            "eyebrow": "Continuation",
            "heading": f"Key Perspective {i + 1}",
            "body": "This page develops the requested topic with concise, presentation-ready detail while preserving the exact page count.",
            "highlights": ["Focused insight", "Clear implication", "Practical takeaway"],
            "callout": "The strongest documents stay precise, structured, and visually deliberate.",
        } for i in range(len(pages), page_count))
    data["pages"] = pages[:page_count]
    return data


def render_premium_pdf_html(structure: dict, page_count: int) -> str:
    palettes = {
        "indigo": ("#312E81", "#6366F1", "#A5B4FC", "#EEF2FF"),
        "teal": ("#134E4A", "#14B8A6", "#99F6E4", "#ECFEFF"),
        "crimson": ("#7F1D1D", "#E11D48", "#FDA4AF", "#FFF1F2"),
        "emerald": ("#064E3B", "#10B981", "#A7F3D0", "#ECFDF5"),
    }
    primary, accent, soft, wash = palettes.get(str(structure.get("palette", "indigo")).lower(), palettes["indigo"])
    title = escape(str(structure.get("title") or "Generated Report"))
    subtitle = escape(str(structure.get("subtitle") or "A polished, AI-generated document"))

    page_html = []
    for index, page in enumerate(structure.get("pages", [])[:page_count], start=1):
        eyebrow = escape(str(page.get("eyebrow") or f"Page {index}"))
        heading = escape(str(page.get("heading") or title))
        body = escape(str(page.get("body") or ""))
        callout = escape(str(page.get("callout") or ""))
        highlights = page.get("highlights") or []
        highlight_items = "".join(f"<li>{escape(str(item))}</li>" for item in highlights[:4])
        cover_class = " cover" if index == 1 else ""
        page_html.append(f"""
        <section class="page{cover_class}">
          <div class="page-bg-mark">{index:02d}</div>
          <header class="topline"><span>{eyebrow}</span><strong>{index}/{page_count}</strong></header>
          <div class="hero-band">
            <p class="kicker">{title}</p>
            <h1>{heading}</h1>
            <p class="subtitle">{subtitle if index == 1 else body}</p>
          </div>
          <div class="content-grid">
            <article class="body-card"><p>{body if index == 1 else callout}</p></article>
            <aside class="highlight-card"><h2>Highlights</h2><ul>{highlight_items}</ul></aside>
          </div>
          <footer>{callout if index == 1 else title}</footer>
        </section>
        """)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8" />
<style>
@page {{ size: A4; margin: 0; }}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; color: #111827; }}
.page {{ width: 210mm; height: 297mm; padding: 16mm; position: relative; overflow: hidden; page-break-after: always; background: linear-gradient(135deg, #ffffff 0%, {wash} 54%, #ffffff 100%); }}
.page:last-child {{ page-break-after: auto; }}
.page::before {{ content: ""; position: absolute; inset: 0; background: radial-gradient(circle at 86% 12%, {soft} 0, transparent 31%), linear-gradient(90deg, {primary} 0 7mm, transparent 7mm); opacity: 0.92; }}
.page-bg-mark {{ position: absolute; right: 9mm; bottom: 4mm; font-size: 82pt; font-weight: 900; color: {primary}; opacity: 0.055; }}
.topline {{ position: relative; z-index: 1; display: flex; justify-content: space-between; align-items: center; color: {primary}; font-size: 9pt; letter-spacing: 0.12em; text-transform: uppercase; }}
.topline strong {{ color: white; background: {primary}; border-radius: 999px; padding: 4px 10px; }}
.hero-band {{ position: relative; z-index: 1; margin-top: 18mm; padding: 12mm; border-radius: 10mm; background: linear-gradient(135deg, {primary}, {accent}); color: white; box-shadow: 0 18px 42px rgba(17, 24, 39, 0.18); }}
.kicker {{ margin: 0 0 5mm; font-size: 8.5pt; letter-spacing: 0.14em; text-transform: uppercase; opacity: 0.78; }}
h1 {{ margin: 0; font-size: 29pt; line-height: 1.02; letter-spacing: -0.02em; }}
.subtitle {{ margin: 7mm 0 0; font-size: 12pt; line-height: 1.48; color: rgba(255,255,255,0.88); }}
.content-grid {{ position: relative; z-index: 1; display: grid; grid-template-columns: 1.18fr 0.82fr; gap: 8mm; margin-top: 10mm; }}
.body-card, .highlight-card {{ border: 1px solid rgba(17,24,39,0.08); border-radius: 7mm; background: rgba(255,255,255,0.86); padding: 8mm; box-shadow: 0 10px 24px rgba(17,24,39,0.08); }}
.body-card p {{ margin: 0; font-size: 12pt; line-height: 1.58; }}
.highlight-card h2 {{ margin: 0 0 5mm; color: {primary}; font-size: 12pt; text-transform: uppercase; letter-spacing: 0.09em; }}
ul {{ margin: 0; padding-left: 5mm; }}
li {{ margin: 0 0 3.4mm; font-size: 10.5pt; line-height: 1.35; }}
footer {{ position: absolute; z-index: 1; left: 16mm; right: 16mm; bottom: 12mm; color: {primary}; font-size: 9.5pt; border-top: 1px solid rgba(17,24,39,0.12); padding-top: 5mm; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cover .hero-band {{ margin-top: 24mm; padding: 15mm; }}
.cover h1 {{ font-size: 34pt; }}
</style></head><body>{''.join(page_html)}</body></html>"""

# ─── Vercel Blob upload ──────────────────────────────────────────────────────

async def upload_to_vercel_blob(buffer: io.BytesIO, filename: str, content_type: str) -> dict:
    """Upload generated bytes to Vercel Blob and return URL metadata for BotBot."""
    if not VERCEL_BLOB_TOKEN:
        raise RuntimeError("VERCEL_BLOB_TOKEN or BLOB_READ_WRITE_TOKEN is not configured")

    buffer.seek(0)
    pathname = f"generated/{uuid.uuid4().hex}-{filename}"

    async with AsyncBlobClient(token=VERCEL_BLOB_TOKEN) as client:
        blob = await client.put(
            pathname,
            buffer.read(),
            access="public",
            content_type=content_type,
            add_random_suffix=False,
            overwrite=True,
        )

    return {
        "url": blob.url,
        "downloadUrl": getattr(blob, "download_url", blob.url),
        "pathname": blob.pathname,
        "contentType": getattr(blob, "content_type", content_type),
    }


# ─── DOCX generation ────────────────────────────────────────────────────────

@app.post("/docs/generate/docx")
async def generate_docx(payload: dict):
    structure = await call_deepinfra_json(payload["prompt"], "docx")

    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(structure["title"])
    title_run.bold = True
    title_run.font.size = Pt(28)
    title_run.font.color.rgb = RGBColor(0x4F, 0x46, 0xE5)  # indigo

    # Subtitle
    if structure.get("subtitle"):
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub_para.add_run(structure["subtitle"])
        sub_run.font.size = Pt(13)
        sub_run.font.color.rgb = RGBColor(0x6B, 0x7A, 0x99)
        sub_run.italic = True

    doc.add_paragraph()  # spacer

    # Divider
    divider = doc.add_paragraph("─" * 65)
    divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
    divider.runs[0].font.color.rgb = RGBColor(0xC7, 0xD2, 0xFE)

    doc.add_paragraph()

    # Sections
    for sec in structure.get("sections", []):
        # Section heading
        h = doc.add_heading(sec["heading"], level=1)
        h.paragraph_format.space_before = Pt(16)
        h.paragraph_format.space_after = Pt(6)
        for run in h.runs:
            run.font.color.rgb = RGBColor(0x4F, 0x46, 0xE5)
            run.font.size = Pt(16)
            run.bold = True

        # Body
        body_para = doc.add_paragraph(sec["body"])
        body_para.paragraph_format.space_after = Pt(10)
        body_para.paragraph_format.line_spacing = Pt(14)
        for run in body_para.runs:
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x1F, 0x2A, 0x44)

        # Subsections
        for sub in sec.get("subsections", []):
            sh = doc.add_heading(sub["heading"], level=2)
            for run in sh.runs:
                run.font.color.rgb = RGBColor(0x7C, 0x3A, 0xED)
                run.font.size = Pt(13)

            sp = doc.add_paragraph(sub["body"])
            sp.paragraph_format.space_after = Pt(8)
            for run in sp.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(0x1F, 0x2A, 0x44)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    blob = await upload_to_vercel_blob(
        buffer, "output.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return blob


# ─── PPTX generation ────────────────────────────────────────────────────────

@app.post("/docs/generate/pptx")
async def generate_pptx(payload: dict):
    structure = await call_deepinfra_json(payload["prompt"], "pptx")

    prs = Presentation()
    prs.slide_width = PPTXInches(13.33)
    prs.slide_height = PPTXInches(7.5)

    # Color palette
    PRIMARY = PPTXRGBColor(0x4F, 0x46, 0xE5)    # indigo
    ACCENT = PPTXRGBColor(0x7C, 0x3A, 0xED)      # purple
    LIGHT = PPTXRGBColor(0xEE, 0xF2, 0xFF)       # very light indigo
    DARK = PPTXRGBColor(0x1E, 0x1B, 0x4B)        # dark navy
    WHITE = PPTXRGBColor(0xFF, 0xFF, 0xFF)

    def set_bg(slide, color: PPTXRGBColor):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_text_box(slide, text, left, top, width, height,
                     font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
        txBox = slide.shapes.add_textbox(
            PPTXInches(left), PPTXInches(top),
            PPTXInches(width), PPTXInches(height)
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.alignment = align
        run = p.runs[0]
        run.font.size = PPTXPt(font_size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        return txBox

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_bg(title_slide, DARK)

    # Accent bar
    bar = title_slide.shapes.add_shape(
        1,  # rectangle
        PPTXInches(0), PPTXInches(3.2),
        PPTXInches(13.33), PPTXInches(0.08)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()

    add_text_box(title_slide, structure["title"],
                 1, 2, 11, 1.5, font_size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    if structure.get("subtitle"):
        add_text_box(title_slide, structure["subtitle"],
                     1, 3.6, 11, 1, font_size=20, color=PPTXRGBColor(0xC7, 0xD2, 0xFE),
                     align=PP_ALIGN.CENTER)

    # Content slides
    for i, slide_data in enumerate(structure.get("slides", [])):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        # Alternate background: white vs very light
        bg_color = WHITE if i % 2 == 0 else LIGHT
        set_bg(slide, bg_color)

        # Left accent bar
        left_bar = slide.shapes.add_shape(
            1,
            PPTXInches(0), PPTXInches(0),
            PPTXInches(0.12), PPTXInches(7.5)
        )
        left_bar.fill.solid()
        left_bar.fill.fore_color.rgb = PRIMARY
        left_bar.line.fill.background()

        # Slide number badge
        badge = slide.shapes.add_shape(
            1,
            PPTXInches(11.8), PPTXInches(0.2),
            PPTXInches(1.2), PPTXInches(0.4)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT
        badge.line.fill.background()
        badge_tf = badge.text_frame
        badge_tf.paragraphs[0].text = f"{i + 1}"
        badge_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        badge_tf.paragraphs[0].runs[0].font.color.rgb = WHITE
        badge_tf.paragraphs[0].runs[0].font.size = PPTXPt(12)
        badge_tf.paragraphs[0].runs[0].font.bold = True

        # Title
        add_text_box(slide, slide_data["title"],
                     0.3, 0.3, 11, 1, font_size=28, bold=True, color=DARK)

        # Divider line
        line = slide.shapes.add_shape(
            1,
            PPTXInches(0.3), PPTXInches(1.4),
            PPTXInches(10), PPTXInches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = PPTXRGBColor(0xC7, 0xD2, 0xFE)
        line.line.fill.background()

        # Bullets
        txBox = slide.shapes.add_textbox(
            PPTXInches(0.5), PPTXInches(1.6),
            PPTXInches(12), PPTXInches(5.5)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for j, bullet in enumerate(slide_data.get("bullets", [])):
            p = tf.add_paragraph() if j > 0 else tf.paragraphs[0]
            p.text = f"  •  {bullet}"
            p.space_before = PPTXPt(8)
            run = p.runs[0]
            run.font.size = PPTXPt(18)
            run.font.color.rgb = PPTXRGBColor(0x1F, 0x2A, 0x44)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    blob = await upload_to_vercel_blob(
        buffer, "output.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    return blob


# ─── PDF generation (WeasyPrint from AI-generated HTML) ─────────────────────

@app.post("/docs/generate/pdf")
async def generate_pdf(payload: dict):
    prompt = payload["prompt"]
    page_count = requested_count(prompt, "page", default=5, minimum=1, maximum=12)
    structure = await call_deepinfra_pdf_pages(prompt, page_count)
    html_content = render_premium_pdf_html(structure, page_count)

    buffer = io.BytesIO()
    HTML(string=html_content).write_pdf(buffer)
    buffer.seek(0)

    blob = await upload_to_vercel_blob(buffer, "output.pdf", "application/pdf")
    return blob