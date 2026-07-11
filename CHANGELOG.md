# Changelog

## 2026-07-11 (fourth pass) — True per-generation design randomness

**Why:** audit of "is every generation actually different?" found two
deterministic paths: the LLM gravitated to the same aesthetic for the same
topic at temp 0.38, and `match_palette()` always returned the same named
palette for the same keywords.

- Every generation now DEALS a random aesthetic direction (of 10) into the
  creative-director prompt — the model must commit to it or pick the nearest
  fit, never its habitual favourite; content temperature raised to 0.55.
- `match_palette()` picks randomly among all top-scoring palettes instead of
  first-best; unmatched topics were already random.
- PPTX font pool widened 3 → 5 pairs (adds Georgia/Tahoma, Georgia/Verdana —
  metric-safe on Office and macOS).
- Verified over 300 simulated runs: palettes, font pairs, dark/light split and
  dealt directions all spread; layouts/motion were already seed-shuffled.

## 2026-07-11 (third pass) — No dead space, relevant imagery, anchored typography

**Why:** screenshot review found an empty "WHAT IT MEANS" chart rail, bento
tiles with text pinned to the top of tall cards, a split layout whose short
body left half a slide white, and an off-topic vintage drawing served as a
photo band (Openverse relevance lottery).

- **`_rail_content()` guarantee**: side rails / support panels fill from
  highlights → stat contexts → tiles/takeaways → body/notes sentences →
  callout; a hollow container can no longer ship. Quote side-cards skip
  content identical to the quote rather than duplicating it.
- **Vertical anchoring**: card/tile text centers in its cell
  (`MSO_ANCHOR.MIDDLE`) — no more top-heavy tiles with dead space below.
- **Split layout structure**: body paragraph capped, accent divider + up to 3
  numbered insight rows fill the lower half (deduped against the photo chip).
- **Image relevance reorder**: keyed stock (Pexels/Unsplash) → Pollinations AI
  generation (`enhance=true`, cinematic style prompt) → Openverse
  (`category=photograph`) → Wikimedia. Keyless CC search is now last — a
  wrong real photo is worse than a right generated one.
- **Generation rate-limit gate**: Pollinations downloads run behind a
  3-permit semaphore (10 concurrent anonymous requests trip rate limiting);
  per-section budget 75s. Test renders: 10/10 slides with photos.

## 2026-07-11 (later) — Imagery everywhere, vector graphics, bento takeaways, stronger covers

**Why:** covers/quotes/splits got photos in the first pass but interior slides
(stats, timelines, comparisons, takeaways) still read as bland text-and-card
walls; user asked for imagery throughout, SVG-style graphics, and harder-
hitting front pages (refs: three Canva template families).

- **Photo band headers**: stat / two_col / timeline / cards / chart slides lay
  a full-width photo strip with scrim behind the eyebrow+title zone whenever
  the section has an image — every layout can now carry photography, and the
  LLM is asked for an `image_keyword` on EVERY section (varied subjects).
- **Stronger cover/closing bleed**: near-black scrim over the text side
  punching to 28% over the photo, plus a bottom scrim for tagline/footer
  legibility; display title +8pt on photo covers; accent kicker chip; floating
  outline-ring and dot-grid vector accents.
- **Vector ("SVG-style") graphics** — native shapes, crisp at any zoom, survive
  Keynote import (literal SVG embedding does not): dot grids, outline rings,
  timeline connector spine, percentage gauge bars on `%` stats, smoothed native
  line charts. (python-pptx cannot embed real .svg; Keynote drops svgBlip.)
- **Bento takeaways**: monotonous 2x2 card grid replaced by an asymmetric bento
  — featured accent tile + wide + two squares, mirrored per deck seed.
- Openverse candidate list now caps origin URLs at 2 so its always-up thumbnail
  proxy is reachable within the 3-candidate budget (origin links are often
  stale Flickr 502s). Result: 10/10 slides with photos in test renders.
- 21st.dev Magic MCP was consulted for component patterns; its inspiration
  endpoint currently returns schema-invalid MCP responses (server-side bug).

## 2026-07-11 — Real web photography, cinematic photo layouts, varied Keynote-safe motion

**Why:** generated decks read as bland text-and-card walls — images only came
from AI generation (Pollinations), the LLM was told to avoid requesting them,
only one layout ever displayed one (in a small side box), and half the injected
animations/transitions used effects Keynote silently drops.

### Imagery (`assets.py`, `ai_content.py`)
- **Real-photo provider chain** per section: Pexels → Unsplash (when
  `PEXELS_API_KEY` / `UNSPLASH_ACCESS_KEY` are configured) → Openverse →
  Wikimedia Commons (both keyless) → Pollinations AI generation only as last
  resort. Candidates are validated (magic bytes, min/max size), stale origin
  links fall through to the next candidate, and Openverse's thumbnail proxy
  backs up dead links. Per-section chains run concurrently and are time-boxed.
- **LLM now asked for imagery**: `image_keyword` is required on cover,
  deep_dive_2, case_study and closing, recommended elsewhere, and must be a
  concrete stock-photo-style search query (2-4 words).
- Default fetch size raised to 1600px for full-bleed use.

### Photo layouts (`render_pptx.py`)
- **Full-bleed photo covers, closings and quote slides** behind a
  dominant-color gradient scrim (per-stop alpha via raw OOXML) — dense over the
  text, lifting to let the photo show through.
- **Editorial split layout** for `image_text`: full-height photo bleeding off
  the right 40% of the slide, with the highlight chip overlaid; footer and
  title respect the split.
- Geometric fallbacks unchanged when no image is available.

### Motion
- Entrance/transition pools trimmed to effects that survive the
  PowerPoint → Keynote round-trip (fade, wipes, dissolve, circle, diamond;
  fade/push/wipe transitions) — previously blinds/strips/cut/zoom imported as
  "no animation".
- **Varied per generation AND per slide**: both pools are shuffled with the
  deck seed then cycled, so consecutive slides never share an effect and two
  runs never share a sequence; stagger rhythm and entrance speed also vary
  per deck. Full-bleed photos and scrims stay static — only content animates.



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
