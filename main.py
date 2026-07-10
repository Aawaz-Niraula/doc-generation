"""main.py — orchestrator: parse input → build model → theme → render → QA → ship.

Two front doors, one pipeline:

- FastAPI service (Vercel):  POST /docs/generate  {"prompt": "...", "format": "pdf|pptx|docx|both"}
  Renders in memory, QA-gates, then uploads to Vercel Blob under a unique name.
  When several formats are requested, each format gets its OWN format-tuned AI
  content pass (run concurrently) so every file's content is unique to its medium.

- CLI (local):  python main.py --input examples/sample_content.json --formats pptx,pdf
  Writes uniquely named files into outputs/ (temp file first; moved only after
  QA passes) and prints the absolute path + a one-line summary per file.

A generation that fails QA raises loudly — nothing half-written ever lands in
outputs/ or the blob store.
"""

from __future__ import annotations

import asyncio
import os
import traceback
import uuid
from dataclasses import dataclass, field
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse

import assets
import qa as qa_mod
from ai_content import (AIContentError, generate_document_spec, normalize_topic_prompt,
                        parse_output_format, requested_count)
from content_model import ContentError, DocumentSpec
from render_docx import render_docx
from render_pdf import render_pdf
from render_pptx import render_pptx
from theme import PALETTES, Theme, resolve_theme

app = FastAPI()

VERCEL_BLOB_TOKEN = os.getenv("VERCEL_BLOB_TOKEN")

CONTENT_TYPES = {
    "pdf": "application/pdf",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

_COUNT_UNITS = {"pdf": "page", "pptx": "slide", "docx": "section"}

# Hard deadline for HTTP requests: callers (BotBot) abort at ~290s, so the
# service must answer — success or a clear error — before that.
GENERATION_BUDGET_SECONDS = 270


# ─── Error handling ───────────────────────────────────────────────────────────

def _exception_detail(exc: Exception) -> str:
    trace_lines = traceback.format_exception(type(exc), exc, exc.__traceback__)
    tail = "".join(trace_lines[-8:]).strip()
    return f"{type(exc).__name__}: {str(exc) or 'Document generation failed'}\n{tail}"


@app.exception_handler(Exception)
async def unhandled_exception_handler(request, exc):
    return JSONResponse(status_code=500, content={"detail": _exception_detail(exc)})


def _http_error(exc: Exception):
    if isinstance(exc, HTTPException):
        raise exc
    if isinstance(exc, AIContentError):
        raise HTTPException(status_code=exc.status, detail=str(exc)) from exc
    if isinstance(exc, (qa_mod.QAError, ContentError)):
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    raise HTTPException(status_code=500, detail=_exception_detail(exc)) from exc


@app.get("/health")
async def health():
    return {"status": "ok"}


# ─── Config ───────────────────────────────────────────────────────────────────

@dataclass
class GenerationConfig:
    formats: List[str] = field(default_factory=lambda: ["pdf"])
    theme_name: Optional[str] = None        # named palette override
    page_size: str = "a4"                   # a4 | letter (PDF only)
    out_dir: str = "outputs"
    seed: Optional[int] = None              # deterministic layouts/theme when set
    section_count: Optional[int] = None
    upload: bool = False                    # push to Vercel Blob after QA

    @classmethod
    def from_payload(cls, payload: dict, fmt: str) -> "GenerationConfig":
        return cls(
            formats=[fmt],
            theme_name=payload.get("theme"),
            page_size=str(payload.get("page_size") or "a4").lower(),
            seed=payload.get("seed"),
            section_count=payload.get("sections"),
            upload=True,
        )


@dataclass
class GenerationResult:
    format: str
    spec: DocumentSpec
    theme: Theme
    data: bytes
    qa: qa_mod.QAReport
    path: Optional[str] = None
    blob: Optional[dict] = None

    def summary(self) -> str:
        unit = "slides" if self.format == "pptx" else "pages" if self.format == "pdf" else "paragraphs"
        return (f"{self.format.upper()}: {self.qa.units} {unit}, "
                f"theme '{self.theme.name}' ({self.theme.aesthetic_label}), "
                f"{self.qa.size_kb:.0f} KB")

    def as_response(self) -> dict:
        out = {
            "filename": os.path.basename(self.path) if self.path else None,
            "path": self.path,
            "qa": self.qa.as_dict(),
            "theme": self.theme.name,
            "summary": self.summary(),
        }
        if self.blob:
            out.update(self.blob)
        return out


# ─── Vercel Blob upload ───────────────────────────────────────────────────────

def _blob_value(blob, key: str, default=None):
    if isinstance(blob, dict):
        return blob.get(key, default)
    return getattr(blob, key, default)


async def upload_to_vercel_blob(data: bytes, filename: str, content_type: str) -> dict:
    if not VERCEL_BLOB_TOKEN:
        raise HTTPException(status_code=500,
                            detail="VERCEL_BLOB_TOKEN is not configured on the document service")
    from vercel.blob import AsyncBlobClient

    pathname = f"generated/{uuid.uuid4().hex[:8]}-{filename}"
    client = AsyncBlobClient(token=VERCEL_BLOB_TOKEN)
    blob = await client.put(
        pathname, data, access="public",
        content_type=content_type, add_random_suffix=False, overwrite=False,
    )
    url = _blob_value(blob, "url")
    if not url:
        raise HTTPException(status_code=502, detail="Vercel Blob upload returned no URL")
    download_url = _blob_value(blob, "download_url") or _blob_value(blob, "downloadUrl") or url
    return {
        "url": url,
        "downloadUrl": download_url,
        "pathname": _blob_value(blob, "pathname", pathname),
        "contentType": _blob_value(blob, "content_type") or _blob_value(blob, "contentType") or content_type,
    }


# ─── The pipeline ─────────────────────────────────────────────────────────────

def _writable_out_dir(cfg: GenerationConfig) -> str:
    """Serverless filesystems are read-only outside /tmp."""
    if os.getenv("VERCEL") or os.getenv("AWS_LAMBDA_FUNCTION_NAME"):
        return os.path.join("/tmp", "outputs")
    return cfg.out_dir


async def _render_one(fmt: str, spec: DocumentSpec, theme: Theme,
                      cfg: GenerationConfig) -> GenerationResult:
    """Render → QA (strict; temp-file semantics) → finalize → optional upload."""
    if fmt == "pdf":
        font_css = await assets.build_embedded_font_css(theme)
        data = render_pdf(spec, theme, page_size=cfg.page_size, font_face_css=font_css)
    elif fmt == "pptx":
        data = render_pptx(spec, theme)
    elif fmt == "docx":
        data = render_docx(spec, theme)
    else:
        raise ValueError(f"unknown format: {fmt}")

    report = qa_mod.run_qa(fmt, data, spec, page_size=cfg.page_size, strict=True)
    path = qa_mod.finalize_output(data, _writable_out_dir(cfg), spec.slug(), fmt)

    smoke = qa_mod.visual_smoke_test(path)
    if smoke:
        report.warnings.append(smoke)

    result = GenerationResult(format=fmt, spec=spec, theme=theme, data=data,
                              qa=report, path=path)
    if cfg.upload:
        result.blob = await upload_to_vercel_blob(
            data, os.path.basename(path), CONTENT_TYPES[fmt])
    return result


async def _generate_format(prompt: str, fmt: str, cfg: GenerationConfig,
                           spec: Optional[DocumentSpec] = None) -> GenerationResult:
    """Full pipeline for ONE format. Unless a pre-built spec is supplied, the
    content is generated fresh and format-tuned, so each format is unique."""
    if spec is None:
        unit = _COUNT_UNITS[fmt]
        count = cfg.section_count or requested_count(prompt, unit, default=10, minimum=8, maximum=12)
        spec = await generate_document_spec(prompt, count, doc_kind=fmt)
    theme = resolve_theme(spec.design, topic=spec.topic_prompt or prompt,
                          requested=cfg.theme_name, seed=cfg.seed)
    if fmt in ("pdf", "pptx"):     # docx is prose-first; images not used there
        await assets.prefetch_images(spec)
    return await _render_one(fmt, spec, theme, cfg)


async def generate(prompt: str, cfg: GenerationConfig) -> List[GenerationResult]:
    """Generate every requested format concurrently, each with its own
    format-tuned content pass."""
    tasks = [_generate_format(prompt, fmt, cfg) for fmt in cfg.formats]
    return list(await asyncio.gather(*tasks))


async def _generate_with_deadline(coro):
    """Bound an HTTP-triggered generation so the service always responds
    before the caller's timeout instead of being cut off by the gateway."""
    try:
        return await asyncio.wait_for(coro, timeout=GENERATION_BUDGET_SECONDS)
    except asyncio.TimeoutError:
        raise HTTPException(
            status_code=504,
            detail=(f"Document generation exceeded the {GENERATION_BUDGET_SECONDS}s budget. "
                    "Try fewer sections or a simpler prompt."))


# ─── HTTP endpoints (backward-compatible) ─────────────────────────────────────

@app.post("/docs/generate")
async def generate_document(payload: dict):
    """Unified endpoint — routes on payload.format or prompt wording."""
    prompt = normalize_topic_prompt(str(payload.get("prompt", "")))
    if not prompt:
        raise HTTPException(status_code=422, detail="payload.prompt is required")
    fmt = str(payload.get("format") or parse_output_format(prompt)).lower().strip()
    try:
        if fmt == "both":
            cfg = GenerationConfig.from_payload(payload, "pdf")
            cfg.formats = ["pdf", "pptx"]
            results = await _generate_with_deadline(generate(prompt, cfg))
            return {r.format: r.as_response() for r in results}
        if fmt in ("pptx", "presentation", "slides"):
            return await generate_pptx(payload)
        if fmt in ("docx", "word", "document"):
            return await generate_docx(payload)
        return await generate_pdf(payload)
    except Exception as exc:
        _http_error(exc)


@app.post("/docs/generate/pdf")
async def generate_pdf(payload: dict):
    try:
        prompt = normalize_topic_prompt(str(payload["prompt"]))
        cfg = GenerationConfig.from_payload(payload, "pdf")
        result = await _generate_with_deadline(_generate_format(prompt, "pdf", cfg))
        return result.as_response()
    except Exception as exc:
        _http_error(exc)


@app.post("/docs/generate/pptx")
async def generate_pptx(payload: dict):
    try:
        prompt = normalize_topic_prompt(str(payload["prompt"]))
        cfg = GenerationConfig.from_payload(payload, "pptx")
        result = await _generate_with_deadline(_generate_format(prompt, "pptx", cfg))
        return result.as_response()
    except Exception as exc:
        _http_error(exc)


@app.post("/docs/generate/docx")
async def generate_docx(payload: dict):
    try:
        prompt = normalize_topic_prompt(str(payload["prompt"]))
        cfg = GenerationConfig.from_payload(payload, "docx")
        result = await _generate_with_deadline(_generate_format(prompt, "docx", cfg))
        return result.as_response()
    except Exception as exc:
        _http_error(exc)


# ─── CLI ──────────────────────────────────────────────────────────────────────

def _cli():
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate designed PDF / PPTX / DOCX documents.")
    src = parser.add_mutually_exclusive_group(required=True)
    src.add_argument("--prompt", help="topic prompt (needs DEEPINFRA_KEY)")
    src.add_argument("--input", help="path to a content-definition JSON file")
    parser.add_argument("--formats", default="pdf",
                        help="comma-separated: pdf,pptx,docx (default: pdf)")
    parser.add_argument("--theme", default=None,
                        help=f"named palette: {', '.join(sorted(PALETTES))}")
    parser.add_argument("--page-size", default="a4", choices=["a4", "letter"],
                        help="PDF page size (default: a4)")
    parser.add_argument("--out-dir", default="outputs",
                        help="output directory (default: outputs/)")
    parser.add_argument("--seed", type=int, default=None,
                        help="deterministic theme/layout seed")
    parser.add_argument("--sections", type=int, default=None,
                        help="section/slide/page count (default: parsed from prompt or 10)")
    parser.add_argument("--upload", action="store_true",
                        help="also upload to Vercel Blob (needs VERCEL_BLOB_TOKEN)")
    args = parser.parse_args()

    formats = [f.strip().lower() for f in args.formats.split(",") if f.strip()]
    for f in formats:
        if f not in CONTENT_TYPES:
            parser.error(f"unknown format '{f}' (choose from pdf, pptx, docx)")

    cfg = GenerationConfig(formats=formats, theme_name=args.theme,
                           page_size=args.page_size, out_dir=args.out_dir,
                           seed=args.seed, section_count=args.sections,
                           upload=args.upload)

    async def run() -> List[GenerationResult]:
        if args.input:
            spec = DocumentSpec.from_json_file(args.input)
            if cfg.section_count:
                spec.ensure_length(cfg.section_count)
            spec.validate()
            results = []
            for fmt in cfg.formats:      # one fixed spec → every format
                theme = resolve_theme(spec.design, topic=spec.topic_prompt,
                                      requested=cfg.theme_name, seed=cfg.seed)
                if fmt in ("pdf", "pptx"):
                    await assets.prefetch_images(spec)
                results.append(await _render_one(fmt, spec, theme, cfg))
            return results
        return await generate(normalize_topic_prompt(args.prompt), cfg)

    try:
        results = asyncio.run(run())
    except (qa_mod.QAError, ContentError, AIContentError) as exc:
        raise SystemExit(f"GENERATION FAILED: {exc}")

    for r in results:
        print(r.path)
        print(f"  {r.summary()}")
        for w in r.qa.warnings:
            print(f"  note: {w}")
        if r.blob:
            print(f"  uploaded: {r.blob['url']}")


if __name__ == "__main__":
    _cli()
