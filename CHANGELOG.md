# Changelog

## 2026-07-09 — Production-grade engine restructure

**Why:** the generator was a single 2,800-line `main.py` with design decisions
scattered through render code, no output-file management, PDF-only QA, and
several "AI-generated deck" visual tells baked in.

### Architecture
- **Split into single-responsibility modules**: `content_model.py`, `theme.py`,
  `ai_content.py`, `assets.py`, `render_pdf.py`, `render_pptx.py`,
  `render_docx.py`, `qa.py`, with `main.py` reduced to orchestration
  (FastAPI service + new local CLI).
- **One typed content model** (`DocumentSpec`) consumed by all three renderers —
  one content definition can produce any format. A `fallback_highlights()`
  chain guarantees no layout ever renders an empty rail.
- **One file per generation**: `outputs/{slug}_{YYYYMMDD-HHMMSS}.{ext}`, with
  `_v2/_v3` collision suffixes. Files are written to a temp name and atomically
  renamed **only after QA passes** — a crashed run can never leave a corrupt
  file. Blob uploads get unique names too (`overwrite=False`).
- **Format-unique content**: each requested format gets its own format-tuned
  LLM pass (slide copy + speaker notes vs. long-form report prose vs. flowing
  document), run concurrently for `format: both` — so latency stays inside the
  Vercel window while every file reads native to its medium.

### Design system (`theme.py`)
- 10 named palettes with **topic keywords** (finance ≠ wellness ≠ tech), each
  with one dominant color, supporting tones, and one sharp accent; tints/washes
  derived, never beige.
- **WCAG contrast enforcement**: every fg/bg pairing (including AI-proposed
  colors) passes through `ensure_contrast` before renderers see it.
- Typography scale encoded once (titles 36–44 bold, section headers 20–24,
  body 14–16, captions 10–12); **metric-safe Office fonts** for PPTX/DOCX
  (Cambria/Times display + Calibri/Arial body + Consolas data — never Aptos);
  the PDF keeps AI-chosen Google fonts because they are embedded as base64 TTFs.
- Layout registry + shared no-consecutive-repeat rule; planners now also prefer
  layouts unused so far in the document.

### Anti-patterns removed (now impossible in render code)
- Accent rules/underlines beneath titles → whitespace hierarchy.
- Edge stripes, card top-strips, footer hairlines, cover accent bar →
  background tints, gradients, and soft shadows.
- Bordered pull quotes → tinted quote blocks.
- Replaced by ONE repeated motif: the numbered accent circle (PDF cards,
  timeline nodes, PPTX stat/tile/quote cards, geo panels).

### PPTX
- **Auto-fit overflow guard**: every text box estimates rendered height from
  font metrics and steps the size down (then word-clips) before it can
  overflow; internal margins zeroed; word-wrap on everywhere.
- **Speaker notes on every slide**, document properties set.
- **Real photos** (prefetched topic images) cover-cropped via srcRect — never
  distorted; new **timeline layout** so `steps` content is no longer dropped.
- Kept and modularized: OOXML entrance animations + per-deck transition,
  clickable prev/home/next navigation, native editable charts, gradient +
  shadow depth, dark sandwich (title/closing) or full-dark premium mode.

### PDF
- Kept the WeasyPrint pipeline **deliberately** (spec suggested ReportLab):
  it is what production runs (`build.sh` installs pango/cairo), it embeds real
  fonts, and it delivers the flow/TOC-bookmarks/page-numbers/vector goals of
  Platypus with far richer art direction. Platypus can't match the current
  visual bar.
- New: comparison duo-cards on feature pages, charts embedded on stat pages
  (content-faithful, no two stat pages alike), dark closing page with CTA pill
  bookending the cover, PDF metadata (`<title>`, author, subject), A4/Letter.

### QA (`qa.py`) — now for all three formats
- Reopen with the real parser; slide/page count vs. model; placeholder regex;
  duplicate-slide detection; per-text-frame overflow re-check (PPTX); page
  dimension check (PDF); optional LibreOffice visual smoke test (`QA_VISUAL=1`).

### Misc
- `requirements.txt` now pins exact versions; `reportlab` dropped (unused).
- Dead code removed (~15 unused SVG/chart helpers from earlier iterations).
- Added `examples/sample_content.json` worked example, README, this changelog.
