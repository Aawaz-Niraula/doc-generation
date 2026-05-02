from fastapi import FastAPI
from docx import Document
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import httpx, io, json, os

app = FastAPI()

DEEPINFRA_KEY = os.getenv("DEEPINFRA_KEY")
VERCEL_BLOB_TOKEN = os.getenv("VERCEL_BLOB_TOKEN")


async def call_deepinfra(prompt: str, fmt: str) -> dict:
    if fmt in ["docx", "pdf"]:
        schema = '{"title": "...", "sections": [{"heading": "...", "body": "..."}]}'
    else:
        schema = '{"title": "...", "slides": [{"title": "...", "bullets": ["...", "..."]}]}'

    async with httpx.AsyncClient(timeout=60) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={
                "model": "deepseek-ai/DeepSeek-V4-Flash",
                "messages": [
                    {
                        "role": "system",
                        "content": f"Return only valid JSON, no markdown, no explanation. Use this exact schema: {schema}"
                    },
                    {"role": "user", "content": prompt}
                ]
            }
        )
    return json.loads(res.json()["choices"][0]["message"]["content"])


async def upload_to_vercel_blob(buffer: io.BytesIO, filename: str, content_type: str) -> str:
    async with httpx.AsyncClient() as client:
        res = await client.put(
            f"https://blob.vercel-storage.com/{filename}",
            content=buffer.read(),
            headers={
                "authorization": f"Bearer {VERCEL_BLOB_TOKEN}",
                "x-content-type": content_type
            }
        )
    return res.json()["url"]


@app.post("/docs/generate/docx")
async def generate_docx(payload: dict):
    structure = await call_deepinfra(payload["prompt"], "docx")

    doc = Document()
    doc.add_heading(structure["title"], 0)
    for section in structure["sections"]:
        doc.add_heading(section["heading"], level=1)
        doc.add_paragraph(section["body"])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    url = await upload_to_vercel_blob(buffer, "output.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return {"url": url}


@app.post("/docs/generate/pptx")
async def generate_pptx(payload: dict):
    structure = await call_deepinfra(payload["prompt"], "pptx")

    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = structure["title"]

    for slide_data in structure["slides"]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_data["title"]
        tf = slide.placeholders[1].text_frame
        tf.text = slide_data["bullets"][0]
        for bullet in slide_data["bullets"][1:]:
            tf.add_paragraph().text = bullet

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    url = await upload_to_vercel_blob(buffer, "output.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    return {"url": url}


@app.post("/docs/generate/pdf")
async def generate_pdf(payload: dict):
    structure = await call_deepinfra(payload["prompt"], "pdf")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    story.append(Paragraph(f"<b>{structure['title']}</b>"))
    story.append(Spacer(1, 12))
    for section in structure["sections"]:
        story.append(Paragraph(f"<b>{section['heading']}</b>"))
        story.append(Spacer(1, 6))
        story.append(Paragraph(section["body"]))
        story.append(Spacer(1, 12))
    doc.build(story)
    buffer.seek(0)

    url = await upload_to_vercel_blob(buffer, "output.pdf", "application/pdf")
    return {"url": url}