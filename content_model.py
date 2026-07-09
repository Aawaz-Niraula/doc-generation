"""content_model.py — the single typed document model consumed by every renderer.

One `DocumentSpec` describes a document independent of output format; render_pdf,
render_pptx, and render_docx all consume the same instance, so one content
definition can produce any format. Instances come from the AI generator
(`DocumentSpec.from_ai_dict`), from a JSON file (`DocumentSpec.from_json_file`),
or are built directly in Python for tests and worked examples.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


# ─── Leaf blocks ──────────────────────────────────────────────────────────────

@dataclass
class Stat:
    value: str = ""
    label: str = ""
    context: str = ""

    @classmethod
    def from_dict(cls, d) -> Optional["Stat"]:
        if not isinstance(d, dict):
            return None
        value = str(d.get("value") or d.get("number") or "").strip()
        if not value:
            return None
        return cls(
            value=value,
            label=str(d.get("label") or "").strip(),
            context=str(d.get("context") or d.get("sub") or "").strip(),
        )


@dataclass
class Step:
    title: str = ""
    description: str = ""
    date: str = ""

    @classmethod
    def from_dict(cls, d) -> Optional["Step"]:
        if not isinstance(d, dict):
            return None
        title = str(d.get("step") or d.get("title") or "").strip()
        desc = str(d.get("desc") or d.get("description") or "").strip()
        if not (title or desc):
            return None
        return cls(title=title, description=desc, date=str(d.get("date") or "").strip())


@dataclass
class Tile:
    label: str = ""
    value: str = ""

    @classmethod
    def from_dict(cls, d) -> Optional["Tile"]:
        if not isinstance(d, dict):
            return None
        label = str(d.get("label") or "").strip()
        value = str(d.get("value") or "").strip()
        if not (label or value):
            return None
        return cls(label=label, value=value)


@dataclass
class TableBlock:
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)

    @classmethod
    def from_dict(cls, d) -> Optional["TableBlock"]:
        if not isinstance(d, dict):
            return None
        headers = [str(h) for h in (d.get("headers") or []) if str(h).strip()]
        rows = [[str(c) for c in row] for row in (d.get("rows") or []) if isinstance(row, (list, tuple))]
        if not headers or not rows:
            return None
        return cls(headers=headers, rows=rows)

    def is_empty(self) -> bool:
        return not (self.headers and self.rows)


@dataclass
class ChartBlock:
    """Label/value series. `kind` is a rendering hint (column|line|bar)."""
    points: List[Tuple[str, float]] = field(default_factory=list)
    title: str = ""
    kind: str = "column"

    @classmethod
    def from_raw(cls, raw, title: str = "", kind: str = "column") -> Optional["ChartBlock"]:
        points: List[Tuple[str, float]] = []
        if isinstance(raw, list):
            for item in raw[:8]:
                try:
                    label, val = item[0], float(item[1])
                    points.append((str(label), val))
                except (TypeError, ValueError, IndexError):
                    continue
        if len(points) < 2:
            return None
        return cls(points=points, title=title, kind=kind)


@dataclass
class ColumnBlock:
    title: str = ""
    points: List[str] = field(default_factory=list)


# ─── Section (one slide / one page / one chapter) ─────────────────────────────

# Narrative roles the AI produces, in canonical order.
SECTION_KINDS = (
    "cover", "context", "deep_dive_1", "deep_dive_2", "deep_dive_3",
    "comparative", "case_study", "future", "takeaways", "closing",
)


@dataclass
class Subsection:
    heading: str = ""
    body: str = ""


@dataclass
class Section:
    kind: str = "context"              # narrative role, one of SECTION_KINDS
    eyebrow: str = ""                  # small section label
    heading: str = ""
    body: str = ""
    highlights: List[str] = field(default_factory=list)
    callout: str = ""                  # quotable sentence
    attribution: str = ""              # quote source (case_study)
    statements: List[str] = field(default_factory=list)   # manifesto lines
    stats: List[Stat] = field(default_factory=list)
    steps: List[Step] = field(default_factory=list)
    tiles: List[Tile] = field(default_factory=list)
    takeaways: List[str] = field(default_factory=list)
    chart: Optional[ChartBlock] = None
    table: Optional[TableBlock] = None
    hero_stat: Optional[Stat] = None
    left: Optional[ColumnBlock] = None
    right: Optional[ColumnBlock] = None
    subtitle: str = ""                 # cover only
    tagline: str = ""                  # cover / closing
    cta: str = ""                      # closing call to action
    image_keyword: str = ""
    image_data: str = ""               # data: URI, filled by assets.prefetch_images
    speaker_notes: str = ""
    subsections: List[Subsection] = field(default_factory=list)
    # Layout assignment (set per run by the renderers' layout planners)
    layout: str = ""
    mirror: bool = False
    geo_seed: int = 0

    def has_data_visual(self) -> bool:
        return bool(self.chart or (self.table and not self.table.is_empty()) or self.hero_stat)

    def fallback_highlights(self) -> List[str]:
        """Best available short-point content, so no layout ever renders an
        empty rail: highlights → statements → tiles → takeaways → columns."""
        if self.highlights:
            return self.highlights
        if self.statements:
            return self.statements
        if self.tiles:
            return [t.value or t.label for t in self.tiles if (t.value or t.label)]
        if self.takeaways:
            return self.takeaways
        points = []
        for col in (self.left, self.right):
            if col and col.points:
                points.extend(col.points[:2])
        return points

    @classmethod
    def from_dict(cls, d: dict) -> "Section":
        kind = str(d.get("type") or d.get("kind") or "context").strip().lower()
        stats = [s for s in (Stat.from_dict(x) for x in d.get("stats") or []) if s]
        steps = [s for s in (Step.from_dict(x) for x in d.get("steps") or []) if s]
        tiles = [t for t in (Tile.from_dict(x) for x in d.get("tiles") or []) if t]
        left = right = None
        if d.get("left_points"):
            left = ColumnBlock(
                title=str(d.get("left_title") or "").strip(),
                points=[str(p) for p in d["left_points"] if str(p).strip()],
            )
        if d.get("right_points"):
            right = ColumnBlock(
                title=str(d.get("right_title") or "").strip(),
                points=[str(p) for p in d["right_points"] if str(p).strip()],
            )
        subsections = [
            Subsection(heading=str(s.get("heading") or ""), body=str(s.get("body") or ""))
            for s in (d.get("subsections") or []) if isinstance(s, dict) and s.get("body")
        ]
        return cls(
            kind=kind if kind in SECTION_KINDS else "context",
            eyebrow=str(d.get("eyebrow") or "").strip(),
            heading=str(d.get("heading") or d.get("title") or "").strip(),
            body=str(d.get("body") or "").strip(),
            highlights=[str(h) for h in (d.get("highlights") or d.get("bullets") or []) if str(h).strip()],
            callout=str(d.get("callout") or d.get("quote") or "").strip(),
            attribution=str(d.get("attribution") or "").strip(),
            statements=[str(s) for s in (d.get("statements") or []) if str(s).strip()],
            stats=stats,
            steps=steps,
            tiles=tiles,
            takeaways=[str(t) for t in (d.get("takeaways") or []) if str(t).strip()],
            chart=ChartBlock.from_raw(d.get("chart_data"), title=str(d.get("heading") or d.get("title") or "")),
            table=TableBlock.from_dict(d.get("data_table")),
            hero_stat=Stat.from_dict(d.get("hero_stat")),
            left=left,
            right=right,
            subtitle=str(d.get("subtitle") or "").strip(),
            tagline=str(d.get("tagline") or "").strip(),
            cta=str(d.get("cta") or "").strip(),
            image_keyword=str(d.get("image_keyword") or "").strip(),
            speaker_notes=str(d.get("speaker_notes") or d.get("notes") or "").strip(),
            subsections=subsections,
        )


# ─── Design intent from the AI (validated & resolved by theme.py) ─────────────

@dataclass
class DesignSpec:
    aesthetic_name: str = ""
    aesthetic_label: str = ""
    rationale: str = ""
    colors: dict = field(default_factory=dict)       # dominant/secondary/supporting_1/supporting_2/accent
    typography: dict = field(default_factory=dict)   # h1_font/h2_font/body_font/data_font
    signature_element: str = ""
    layout_philosophy: str = ""
    palette_name: str = ""                            # legacy named-palette fallback

    @classmethod
    def from_dict(cls, data: dict) -> "DesignSpec":
        aesthetic = data.get("aesthetic_direction") or {}
        ds = data.get("design_system") or {}
        return cls(
            aesthetic_name=str(aesthetic.get("name") or "").strip(),
            aesthetic_label=str(aesthetic.get("label") or aesthetic.get("name") or "").strip(),
            rationale=str(aesthetic.get("rationale") or "").strip(),
            colors={k: str(v) for k, v in (ds.get("colors") or {}).items()},
            typography={k: str(v) for k, v in (ds.get("typography") or {}).items()},
            signature_element=str(ds.get("signature_element") or "").strip(),
            layout_philosophy=str(ds.get("layout_philosophy") or "").strip(),
            palette_name=str(data.get("palette") or "").strip().lower(),
        )


# ─── Document ─────────────────────────────────────────────────────────────────

class ContentError(ValueError):
    """Raised when a document spec is structurally unusable."""


@dataclass
class DocumentSpec:
    title: str = "Document"
    subtitle: str = ""
    author: str = ""
    date: str = ""
    abstract: str = ""
    conclusion: str = ""
    design: DesignSpec = field(default_factory=DesignSpec)
    sections: List[Section] = field(default_factory=list)
    topic_prompt: str = ""             # the original user ask, for theme matching

    # ── construction ──────────────────────────────────────────────────────────

    @classmethod
    def from_ai_dict(cls, data: dict, topic_prompt: str = "") -> "DocumentSpec":
        raw_sections = data.get("sections") or data.get("pages") or data.get("slides") or []
        sections = [Section.from_dict(s) for s in raw_sections if isinstance(s, dict)]
        spec = cls(
            title=str(data.get("title") or "Document").strip() or "Document",
            subtitle=str(data.get("subtitle") or "").strip(),
            author=str(data.get("author") or "AI Document Studio").strip(),
            date=str(data.get("date") or "").strip(),
            abstract=str(data.get("abstract") or "").strip(),
            conclusion=str(data.get("conclusion") or "").strip(),
            design=DesignSpec.from_dict(data),
            sections=sections,
            topic_prompt=topic_prompt,
        )
        return spec

    @classmethod
    def from_json_file(cls, path: str) -> "DocumentSpec":
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return cls.from_ai_dict(data, topic_prompt=str(data.get("topic_prompt") or ""))

    # ── normalisation & validation ────────────────────────────────────────────

    def ensure_length(self, count: int) -> "DocumentSpec":
        """Pad or trim to exactly `count` sections and pin cover/closing roles."""
        while len(self.sections) < count:
            i = len(self.sections)
            kind = SECTION_KINDS[min(i, len(SECTION_KINDS) - 1)]
            self.sections.append(Section(
                kind=kind,
                eyebrow=f"Section {i + 1}",
                heading=f"Key Insight {i + 1}",
                body=("This section presents essential analysis and actionable perspective "
                      "on the topic, offering structured insight to guide informed "
                      "decision-making and strategic planning."),
                highlights=["Structured insight", "Clear implication", "Practical takeaway"],
                speaker_notes="Walk the audience through the key point of this section.",
            ))
        self.sections = self.sections[:count]
        if count >= 2:
            self.sections[0].kind = "cover"
            self.sections[-1].kind = "closing"
        return self

    def validate(self) -> "DocumentSpec":
        if not self.title.strip():
            raise ContentError("Document has no title")
        if len(self.sections) < 2:
            raise ContentError(f"Document has only {len(self.sections)} section(s); need at least 2")
        for i, sec in enumerate(self.sections):
            if not (sec.heading or sec.body or sec.callout or sec.stats or sec.statements):
                raise ContentError(f"Section {i + 1} is empty (no heading, body, or content blocks)")
        return self

    # ── output naming ─────────────────────────────────────────────────────────

    def slug(self, max_len: int = 48) -> str:
        s = re.sub(r"[^a-z0-9]+", "-", self.title.lower()).strip("-")
        s = re.sub(r"-{2,}", "-", s)
        return (s[:max_len].rstrip("-")) or "document"
