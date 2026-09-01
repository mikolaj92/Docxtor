"""Typed, domain-blind projection for DOCX review consumers."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .docx import DocxDocument
from .docx_facts import docx_facts
from .docx_inline import _advances_offset, paragraph_to_inline_segments
from .docx_models import AddressableComment, AddressableSpan
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import ReviewCoverage, ReviewDiagnostic


@dataclass(frozen=True)
class ReviewParagraphProjection:
    locator: str
    text: str
    paragraph_index: int | None
    story_kind: str
    is_heading: bool
    opaque_ranges: tuple[tuple[int, int], ...] = ()


@dataclass(frozen=True)
class ReviewNoteProjection:
    note_id: str
    kind: str
    text: str


@dataclass(frozen=True)
class DocxReviewProjection:
    paragraphs: tuple[ReviewParagraphProjection, ...]
    spans: tuple[AddressableSpan, ...]
    comments: tuple[AddressableComment, ...]
    notes: tuple[ReviewNoteProjection, ...]
    table_count: int
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...]


def project_docx_for_review(source: str | Path | bytes) -> DocxReviewProjection:
    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    document = DocxDocument.open_bytes(data)
    inventory = inventory_review_markup(data)
    facts = docx_facts(data)
    by_locator = {fact.container_id: fact for fact in facts.paragraphs}
    paragraphs = []
    for segment in document.segments:
        locator = segment.container_id or ""
        if not locator or locator.startswith(("comment:", "footnote:", "endnote:")):
            continue
        paragraph = document.resolve_paragraph(locator)
        fact = by_locator.get(locator)
        style_id = fact.style_id if fact is not None else None
        paragraphs.append(
            ReviewParagraphProjection(
                locator,
                segment.text,
                segment.paragraph_index,
                locator.split(":", 1)[0],
                bool(style_id and (style_id.startswith("Heading") or style_id == "Title")),
                _opaque_ranges(paragraph),
            )
        )
    notes = tuple(
        ReviewNoteProjection(
            item.value or item.fact_id, item.kind.split("_", 1)[0], item.target or ""
        )
        for item in facts.notes
        if item.kind in {"footnote_user", "endnote_user"}
    )
    table_ids = {
        fact.coordinate.table_index
        for fact in facts.paragraphs
        if fact.coordinate.table_index is not None
    }
    return DocxReviewProjection(
        tuple(paragraphs),
        document.spans,
        inventory.comments,
        notes,
        len(table_ids),
        inventory.coverage,
        inventory.diagnostics,
    )


def _opaque_ranges(paragraph: object) -> tuple[tuple[int, int], ...]:
    if paragraph is None:
        return ()
    segments = paragraph_to_inline_segments(paragraph)
    raw = "".join(segment.text for segment in segments if _advances_offset(segment))
    lead = len(raw) - len(raw.lstrip())
    limit = len(raw.strip())
    offset = 0
    ranges = []
    for segment in segments:
        length = len(segment.text) if _advances_offset(segment) else 0
        if segment.kind == "opaque" and length:
            start = max(offset - lead, 0)
            end = min(offset + length - lead, limit)
            if start < end:
                ranges.append((start, end))
        offset += length
    return tuple(ranges)
