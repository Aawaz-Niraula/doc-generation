# doc-generation — designed PDF / PPTX / DOCX engine

Turns a topic prompt (or a JSON content definition) into **presentation-grade
documents**: a 16:9 PowerPoint deck with animations, native editable charts,
real photos and speaker notes; an A4/Letter editorial PDF with embedded fonts
and vector graphics; and a Word document with a live table of contents.

Every generation is validated by an automated QA gate before the file is
allowed to exist, and every run writes a **fresh, uniquely named file** — never
appending to or overwriting anything.

## Architecture

| Module | Responsibility |
|---|---|
| `content_model.py` | Typed `DocumentSpec` (sections, stats, charts, tables, steps, tiles, quotes, speaker notes). **The single model all renderers consume.** |
| `theme.py` | Design system: 10 named palettes with topic keywords, typography scale, WCAG contrast guard, metric-safe PPTX font pairs, layout registry, font-metric estimation for the overflow guard. |
| `ai_content.py` | One creative-director LLM call → `DocumentSpec`. Content is **format-tuned**: slides get punchy copy + spoken notes, PDFs get long-form analysis, DOCX gets flowing prose with subsections. |
| `assets.py` | Time-boxed asset fetch: Google fonts as base64 `@font-face` (Fontsource CDN) for the PDF, AI topic images (Pollinations) for PDF/PPTX. Fails soft. |
| `render_pdf.py` | WeasyPrint HTML→PDF: content-faithful page archetypes, SVG data panels, bookmarks, metadata. |
| `render_pptx.py` | python-pptx: explicit inch positioning, auto-fit overflow guard, speaker notes, OOXML entrance animations + transitions, clickable navigation, cover-cropped photos, native charts. |
| `render_docx.py` | python-docx: live TOC field, outline levels, banded tables, running footer with PAGE field. |
| `qa.py` | Post-generation validation + atomic temp-file → `outputs/` finalization. |
| `main.py` | Orchestrator: FastAPI service (Vercel) **and** local CLI. |

## Install

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

WeasyPrint (PDF only) needs native libraries:

- **macOS**: `brew install glib pango cairo gdk-pixbuf`, then run with
  `DYLD_FALLBACK_LIBRARY_PATH=/opt/homebrew/lib`.
- **Debian/Ubuntu/Vercel**: see `build.sh` (libpango, libcairo, gdk-pixbuf).

PPTX and DOCX generation are pure Python and work everywhere.

## Worked example

The repo ships a complete content definition:
[`examples/sample_content.json`](examples/sample_content.json). Generate the
deck **and** the report from it:

```bash
DYLD_FALLBACK_LIBRARY_PATH=/opt/homebrew/lib \
python main.py --input examples/sample_content.json --formats pptx,pdf
```

Output:

```
/abs/path/outputs/the-state-of-grid-scale-energy-storage_20260709-072607.pptx
  PPTX: 10 slides, theme 'bauhaus_data' (Bauhaus Data), 168 KB
/abs/path/outputs/the-state-of-grid-scale-energy-storage_20260709-072609.pdf
  PDF: 10 pages, theme 'bauhaus_data' (Bauhaus Data), 178 KB
```

From a prompt instead (needs `DEEPINFRA_KEY`):

```bash
python main.py --prompt "10 slides on the economics of vertical farming" --formats pptx
```

Useful flags: `--theme boardroom` (named palette), `--page-size letter`,
`--seed 42` (deterministic layouts), `--sections 8`, `--out-dir build/`,
`--upload` (push to Vercel Blob).

## Running the service

```bash
DEEPINFRA_KEY=... VERCEL_BLOB_TOKEN=... uvicorn main:app --reload
```

```
POST /docs/generate         {"prompt": "...", "format": "pdf|pptx|docx|both"}
POST /docs/generate/pdf     {"prompt": "..."}
POST /docs/generate/pptx    {"prompt": "..."}
POST /docs/generate/docx    {"prompt": "..."}
```

Responses include the blob `url`, `filename`, `qa` report, and a one-line
`summary`. With `"format": "both"` each format gets its **own** format-tuned
content pass, run concurrently — two unique files, one request, same latency.

## Adding a theme

Append a `Palette` to `PALETTES` in [`theme.py`](theme.py):

```python
Palette("orchard", "#2D4A1E", "#3E6429", "#9CC069", "#E8590C",
        ("orchard", "farming", "produce"), "verdant green with harvest orange"),
```

One dominant color (60–70% of visual weight), a secondary card tone, a
supporting mid tone, one sharp accent, and the topic keywords that make the
auto-matcher pick it. Tints, washes, and readable ink colors are derived —
and every pairing is pushed through the WCAG contrast guard automatically.

## Defining content by hand

Any JSON matching the shape of `examples/sample_content.json` works. Sections
declare a narrative `type` (`cover`, `context`, `deep_dive_1..3`,
`comparative`, `case_study`, `future`, `takeaways`, `closing`) plus whatever
blocks they have: `stats`, `chart_data`, `data_table`, `steps`, `tiles`,
`left_points`/`right_points`, `callout`/`attribution`, `image_keyword`,
`speaker_notes`. The layout planners assign each section a compatible visual
layout per run — same content never renders the same way twice, and no two
consecutive pages/slides share a layout.

## QA gate (runs on every generation)

1. **Integrity** — the artifact is reopened with the real parser
   (python-pptx / pypdf / python-docx); slide/page count must match the model.
2. **Content** — all text is extracted and scanned for placeholder patterns
   (`lorem`, `TODO`, `[insert`, `XXX`), empty slides, duplicates.
3. **Overflow** — every PPTX text frame is re-estimated against its container
   with the same font metrics the renderer's auto-fit guard used.
4. **Visual smoke test** *(optional)* — set `QA_VISUAL=1` with LibreOffice
   installed to headless-convert the deck as a render check.

Only after QA passes is the temp file atomically moved into `outputs/`
(or uploaded). A failed run raises loudly and leaves nothing behind.
