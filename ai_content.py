"""ai_content.py — turns a user prompt into a DocumentSpec via the LLM.

ONE creative-director call produces a unified content structure that all three
renderers consume, so requesting "both" formats costs a single LLM round-trip
(important inside serverless time budgets) and the deck and the report tell
the same story with the same design system.
"""

from __future__ import annotations

import json
import os
import re
from typing import Optional

import httpx

from content_model import DocumentSpec

DEEPINFRA_KEY = os.getenv("DEEPINFRA_KEY")
_MODEL = "deepseek-ai/DeepSeek-V4-Flash"


class AIContentError(RuntimeError):
    def __init__(self, message: str, status: int = 502):
        super().__init__(message)
        self.status = status


# ─── Prompt-input parsing ─────────────────────────────────────────────────────

NUMBER_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15,
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


def normalize_topic_prompt(prompt: str) -> str:
    """Strip template placeholders from the user prompt."""
    cleaned = prompt
    for placeholder in (
        r"\[INSERT TOPIC\]", r"\[INSERT TOPIC HERE\]",
        r"\[PDF\s*/\s*PPTX\s*/\s*DOCS/BOTH\]", r"\[PDF\s*/\s*PPTX\]",
        r"Topic:\s*\[.*?\]", r"Output format:\s*\[.*?\]",
    ):
        cleaned = re.sub(placeholder, "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip() or prompt.strip()


def parse_output_format(prompt: str) -> str:
    """Detect desired output from prompt text. Returns pdf|pptx|docx|both."""
    normalized = prompt.lower()
    if re.search(r"\b(?:both|pdf\s*\+\s*pptx|pdf\s*and\s*pptx|all\s+formats)\b", normalized):
        return "both"
    if re.search(r"\b(?:pptx|powerpoint|presentation|slides?)\b", normalized):
        return "pptx"
    if re.search(r"\b(?:docx|word|document)\b", normalized) and "pdf" not in normalized:
        return "docx"
    if re.search(r"\b(?:pdf|report)\b", normalized):
        return "pdf"
    return "pdf"


# ─── LLM transport ────────────────────────────────────────────────────────────

async def _deepinfra_call(messages: list, temperature: float = 0.4) -> str:
    if not DEEPINFRA_KEY:
        raise AIContentError("DEEPINFRA_KEY is not configured on the document service", status=500)

    # Must finish well inside the 270s service deadline (main.py) so images,
    # rendering, QA and blob upload still fit in the caller's 5-minute budget.
    async with httpx.AsyncClient(timeout=httpx.Timeout(210, connect=10)) as client:
        res = await client.post(
            "https://api.deepinfra.com/v1/openai/chat/completions",
            headers={"Authorization": f"Bearer {DEEPINFRA_KEY}"},
            json={"model": _MODEL, "temperature": temperature, "messages": messages},
        )
    if res.status_code >= 400:
        raise AIContentError(f"DeepInfra request failed ({res.status_code}): {res.text[:500]}")
    try:
        raw = res.json()["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
        raise AIContentError("DeepInfra returned an invalid response") from exc
    return re.sub(r"```json|```", "", raw).strip()


def _as_json(raw: str, label: str) -> dict:
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AIContentError(f"{label} returned invalid JSON") from exc


# ─── Creative-director prompt (unified schema) ────────────────────────────────

AESTHETIC_DIRECTIONS = [
    "editorial_magazine", "scientific_brutalism", "cinematic_dark", "bauhaus_data",
    "japanese_minimalism", "retro_futurism", "nature_organic", "neon_cyberpunk",
    "luxury_editorial", "abstract_expressionism",
]


# Format-specific writing direction: the SAME schema, but the content is written
# for its medium — so a deck, a report, and a doc from one prompt are each unique.
_FORMAT_DIRECTIVES = {
    "pdf": """WRITING FOR: a printed A4 editorial report.
- Long-form authority: every body 80+ words of analysis, not summary.
- Lean on data_table and chart_data; the reader studies these pages.
- Callouts are editorial pull-quotes; eyebrows read like magazine section labels.""",
    "pptx": """WRITING FOR: a projected 16:9 presentation.
- Punchy and scannable: headings ≤ 8 words, highlights 10-18 words each.
- The DETAIL lives in speaker_notes (3-4 spoken sentences per slide) — slides carry the hook.
- Prefer stats, comparisons (left/right points), and one-line callouts over paragraphs.
- bodies max 60 words; closing cta must be a call the presenter can say out loud.""",
    "docx": """WRITING FOR: a flowing Word document.
- Narrative prose with embedded data; 6+ sections must include 1-2 subsections (40+ words each).
- abstract and conclusion are mandatory and substantive.
- Prefer data woven into sentences plus one data_table; avoid list-heavy sections.""",
}


def _system_prompt(section_count: int, doc_kind: str = "pdf") -> str:
    aesthetics = ", ".join(AESTHETIC_DIRECTIONS)
    directive = _FORMAT_DIRECTIVES.get(doc_kind, _FORMAT_DIRECTIVES["pdf"])
    return f"""You are a world-class creative director, data visualization expert, and document strategist.
Return ONLY valid JSON — no markdown fences, no commentary.

CORE DIRECTIVE:
- Do NOT reuse generic templates. Pick ONE aesthetic direction from: {aesthetics}
  (or invent a better one). State which direction you chose and why it fits the topic.
- Commit fully to that aesthetic across all {section_count} sections.
- Use REAL topic-specific data: authentic statistics, dates, names, figures. No lorem ipsum,
  no TODO, no [insert ...] placeholders anywhere.

{directive}

DESIGN SYSTEM (define before content — lock it in):
- colors: exactly 5 named hex values — dominant, secondary, supporting_1, supporting_2, accent
  (dominant carries 60-70% of visual weight; accent used sparingly; must fit the topic emotionally)
- typography: 2-3 Google Fonts with exact usage (h1_font, h2_font, body_font, data_font)
- layout_philosophy: one sentence describing grid rotation strategy
- signature_element: ONE visual motif repeated for cohesion

FORBIDDEN:
- Generic blue gradient defaults, identical layouts repeated, bullet lists as primary content,
  centered body paragraphs, placeholder or generic content.

EVERY SECTION MUST include at least ONE of:
- chart_data (real numbers), hero_stat (single powerful stat), data_table (styled table),
  stats (3 metric cards), steps (timeline), or image_keyword (real-world visual).
- Prefer data-native visuals over fake-looking AI photos. Use image_keyword only when a
  real-world visual materially helps; never request fake people, fake screenshots, fake logos,
  or readable text inside an image.

NARRATIVE STRUCTURE for {section_count} sections (adapt if count differs — preserve order):
1.  "cover"       — cinematic opening; declare the aesthetic; subtitle + tagline
2.  "context"     — stakes via data (3 stats or a chart), not bullets
3.  "deep_dive_1" — richest analysis; comparison columns (left_points/right_points) + data
4.  "deep_dive_2" — visual storytelling; body narrative + highlights + image_keyword
5.  "deep_dive_3" — quantitative; exactly 3 stats with REAL numbers + chart_data
6.  "comparative" — timeline or before/after; steps with dates/figures
7.  "case_study"  — pull quote (callout) + attribution + supporting highlights
8.  "future"      — forward-looking; 3 bold statements + abstract tone
9.  "takeaways"   — 4 visual tiles — NOT a bullet list
10. "closing"     — as impactful as the cover; cta + tagline

JSON schema (one structure drives the PDF report, the PPTX deck, and the DOCX doc):
{{
  "title": "Document title",
  "subtitle": "Compelling subtitle",
  "author": "Prepared for [audience] — 2026",
  "date": "2026",
  "abstract": "2-3 sentence executive summary with at least one real statistic.",
  "conclusion": "Forward-looking 2-3 sentence close.",
  "aesthetic_direction": {{"name": "cinematic_dark", "label": "Readable name", "rationale": "2 sentences"}},
  "design_system": {{
    "colors": {{"dominant": "#1A1A2E", "secondary": "#16213E", "supporting_1": "#0F3460",
                "supporting_2": "#E8E8E8", "accent": "#E94560"}},
    "typography": {{"h1_font": "Playfair Display", "h2_font": "DM Sans",
                    "body_font": "Source Sans 3", "data_font": "JetBrains Mono"}},
    "layout_philosophy": "one sentence",
    "signature_element": "one repeated motif"
  }},
  "sections": [
    {{
      "type": "cover|context|deep_dive_1|deep_dive_2|deep_dive_3|comparative|case_study|future|takeaways|closing",
      "eyebrow": "SECTION LABEL",
      "heading": "Headline",
      "subtitle": "cover only",
      "tagline": "cover/closing only — include the aesthetic label",
      "body": "Rich paragraph, 70+ words, expert content with real figures.",
      "highlights": ["Insight with real data", "…", "…"],
      "callout": "Memorable quotable sentence.",
      "attribution": "— Source, Year (case_study only)",
      "statements": ["Bold sentence."],
      "stats": [{{"value": "847M", "label": "METRIC", "context": "one sentence"}}],
      "steps": [{{"step": "Phase", "desc": "what happened", "date": "2024"}}],
      "tiles": [{{"label": "01", "value": "Takeaway text"}}],
      "takeaways": ["closing card text"],
      "chart_data": [["Label", 42]],
      "hero_stat": {{"value": "847M", "label": "METRIC", "context": "one sentence"}},
      "data_table": {{"headers": ["Col"], "rows": [["cell"]]}},
      "left_title": "Perspective A", "left_points": ["15+ word point"],
      "right_title": "Perspective B", "right_points": ["15+ word point"],
      "cta": "closing call to action",
      "image_keyword": "photorealistic topic image (only when it materially helps)",
      "speaker_notes": "2-3 sentences a presenter would say on this section.",
      "subsections": [{{"heading": "Sub", "body": "40+ words (deep dives only, max 2)"}}]
    }}
  ]
}}

RULES:
- sections[0].type MUST be "cover"; sections[{section_count - 1}].type MUST be "closing"
- sections array MUST have exactly {section_count} items, following the narrative order
- every section: speaker_notes filled in; body 70+ words where present; topic-specific
- deep_dive_3/context: exactly 3 stats with real numbers
- comparative/deep_dive_1: 3 items in BOTH left_points and right_points
- takeaways: exactly 4 tiles
- Omit fields that don't apply to a section (empty string / empty list)."""


async def generate_document_spec(prompt: str, section_count: int,
                                 doc_kind: str = "pdf") -> DocumentSpec:
    """One LLM call → validated DocumentSpec (padded/trimmed to section_count).

    `doc_kind` (pdf|pptx|docx) tunes HOW the content is written for its medium;
    when the caller requests several formats, each gets its own call (run
    concurrently by main.py) so every file's content is unique to its format.
    """
    raw = await _deepinfra_call(
        [{"role": "system", "content": _system_prompt(section_count, doc_kind)},
         {"role": "user", "content": prompt}],
        temperature=0.38,
    )
    data = _as_json(raw, "Document content generator")
    spec = DocumentSpec.from_ai_dict(data, topic_prompt=prompt)
    spec.ensure_length(section_count)
    return spec.validate()
