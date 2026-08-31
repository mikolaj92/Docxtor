from __future__ import annotations

import copy
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .docx_models import InlineSegment
from .docx_ns import _TEXT_NODE_TAGS
from .docx_xml import _is_text_box_container


def _advances_offset(segment: InlineSegment) -> bool:
    """Whether the segment contributes to visible/offset space (base mechanical view)."""
    return segment.kind in ("text", "opaque")


def _visible_text(segments: list[InlineSegment]) -> str:
    return "".join(segment.text for segment in segments if _advances_offset(segment))


def _visible_len(segments: list[InlineSegment]) -> int:
    return len(_visible_text(segments))


def _copy_segment(segment: InlineSegment, text: str) -> InlineSegment:
    return InlineSegment(
        kind=segment.kind,
        text=text,
        rpr=copy.deepcopy(segment.rpr),
        element=copy.deepcopy(segment.element) if segment.element is not None else None,
    )


def _rpr_at(segments: list[InlineSegment], offset: int) -> Any | None:
    """Return a deepcopy of rpr active at the given visible offset."""
    cursor = 0
    previous: Any | None = None
    for segment in segments:
        if segment.kind != "text":
            if _advances_offset(segment):
                cursor += len(segment.text)
            continue
        next_cursor = cursor + len(segment.text)
        if cursor <= offset <= next_cursor:
            return copy.deepcopy(segment.rpr)
        previous = segment.rpr
        cursor = next_cursor
    return copy.deepcopy(previous)


def _index_at_visible_offset(segments: list[InlineSegment], offset: int) -> int:
    cursor = 0
    for index, segment in enumerate(segments):
        if not _advances_offset(segment):
            continue
        if cursor >= offset:
            return index
        cursor += len(segment.text)
        if cursor >= offset:
            return index + 1
    return len(segments)


def _split_visible_offset(segments: list[InlineSegment], offset: int) -> list[InlineSegment]:
    """Split a text segment at the visible character offset. Pure mechanical."""
    if offset <= 0:
        return segments

    result: list[InlineSegment] = []
    cursor = 0
    split_done = False
    for segment in segments:
        if segment.kind != "text":
            result.append(segment)
            if _advances_offset(segment):
                cursor += len(segment.text)
            continue
        next_cursor = cursor + len(segment.text)
        if not split_done and cursor < offset < next_cursor:
            split_at = offset - cursor
            result.append(_copy_segment(segment, segment.text[:split_at]))
            result.append(_copy_segment(segment, segment.text[split_at:]))
            split_done = True
        else:
            result.append(segment)
        cursor = next_cursor
    return result


def _insert_visible(
    segments: list[InlineSegment], offset: int, insert: InlineSegment
) -> list[InlineSegment]:
    """Insert at visible offset. Pure mechanical."""
    segments = _split_visible_offset(segments, offset)
    index = _index_at_visible_offset(segments, offset)
    return [*segments[:index], insert, *segments[index:]]


def _replace_visible_range(
    segments: list[InlineSegment],
    start: int,
    end: int,
    replacement: list[InlineSegment],
) -> list[InlineSegment]:
    """Replace [start, end) visible range. Pure mechanical."""
    segments = _split_visible_offset(_split_visible_offset(segments, end), start)
    result: list[InlineSegment] = []
    inserted = False
    offset = 0
    for segment in segments:
        next_offset = offset + (len(segment.text) if _advances_offset(segment) else 0)
        if segment.kind == "text" and start <= offset and next_offset <= end:
            if not inserted:
                result.extend(s for s in replacement if s.text)
                inserted = True
            offset = next_offset
            continue
        result.append(segment)
        offset = next_offset
    if not inserted:
        index = _index_at_visible_offset(result, start)
        result[index:index] = [s for s in replacement if s.text]
    return result


def _inline_width(child: Any) -> str:
    """Visible contribution of a non-text inline child (tab, break, etc.)."""
    if child.tag == qn("w:tab"):
        return "\t"
    if child.tag in (qn("w:br"), qn("w:cr")):
        return "\n"
    return ""


def _descendant_visible_text(element: Any) -> str:
    """Visible characters contributed by an opaque subtree (for offset accounting).

    Text-box subtrees are excluded: those paragraphs belong to ``txbx:N``
    containers, not the wrapping body/header run.
    """
    parts: list[str] = []

    def walk(node: Any) -> None:
        if node is not element and _is_text_box_container(node.tag):
            return
        if node.tag in _TEXT_NODE_TAGS and node.text:
            parts.append(node.text)
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            parts.append("\n")
        for child in node:
            walk(child)

    walk(element)
    return "".join(parts)


def _wrap_run_child(rpr: Any | None, child: Any) -> Any:
    """Wrap a non-text run child back into a run element, preserving rpr if present.
    Used for opaque preservation.
    """
    run = OxmlElement("w:r")  # type: ignore[name-defined]
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    run.append(copy.deepcopy(child))
    return run


def _run_segments(run: Any) -> list[InlineSegment]:
    """Decompose a single <w:r> into text + opaque segments. Pure mechanical."""
    rpr = run.find(qn("w:rPr"))
    result: list[InlineSegment] = []
    for child in run:
        tag = child.tag
        if tag == qn("w:rPr"):
            continue
        if tag in _TEXT_NODE_TAGS:
            if child.text:
                result.append(InlineSegment("text", child.text, copy.deepcopy(rpr)))
            continue
        # Non-text run content (tab, break, drawing, field char, ...)
        # re-wrapped so rpr survives on re-emit. Text-box drawings contribute
        # no body-visible width; their text lives on txbx:N segments.
        width = (
            ""
            if _is_text_box_container(child.tag)
            or any(_is_text_box_container(n.tag) for n in child.iter())
            else _inline_width(child)
        )
        result.append(
            InlineSegment(
                "opaque",
                width,
                element=_wrap_run_child(rpr, child),
            )
        )
    return result



def paragraph_to_inline_segments(paragraph: Paragraph) -> list[InlineSegment]:
    """Canonical decomposition of a python-docx Paragraph into ordered InlineSegments.

    This is the single mechanical source for:
    - separating editable text runs from opaque inline content
    - preserving rPr on text runs
    - keeping non-text elements (images, tabs, breaks, fields, hyperlinks, ...)
      as opaque with their visible width contribution for offset math.

    Consumers (especially reviewkit) MUST use this instead of re-walking XML.
    """
    segments: list[InlineSegment] = []
    for child in paragraph._p:
        tag = child.tag
        if tag == qn("w:pPr"):
            continue
        if tag == qn("w:r"):
            segments.extend(_run_segments(child))
            continue
        # Opaque top-level element inside paragraph (e.g. a drawing outside a run, or other).
        segments.append(
            InlineSegment(
                "opaque",
                _descendant_visible_text(child),
                element=copy.deepcopy(child),
            )
        )
    return segments


def rebuild_paragraph_from_inline(paragraph: Paragraph, segments: list[InlineSegment]) -> None:
    """Neutral rebuild: replace paragraph children with the given segments.

    - Preserves existing <w:pPr>.
    - Text segments become <w:r><w:rPr>...</w:rPr><w:t>...</w:t></w:r> (best effort rPr).
    - Opaque segments re-emit their original element.
    - No tracked markup, no comments. Pure mechanical roundtrip for non-review use.

    Review-specific rebuild (with ins/del, revision stamping, comment ranges) stays in
    the review layer (reviewkit) which can use this for base then overlay, or keep its
    own emission for tracked semantics.
    """
    parent = paragraph._p
    # Remove existing non-pPr children
    for child in list(parent):
        if child.tag != qn("w:pPr"):
            parent.remove(child)

    for seg in segments:
        if not seg.text and seg.kind != "opaque":
            continue
        if seg.kind == "text":
            run = OxmlElement("w:r")
            if seg.rpr is not None:
                run.append(copy.deepcopy(seg.rpr))
            t = OxmlElement("w:t")
            if seg.text[:1].isspace() or seg.text[-1:].isspace():
                t.set(qn("xml:space"), "preserve")
            t.text = seg.text
            run.append(t)
            parent.append(run)
        elif seg.kind == "opaque" and seg.element is not None:
            parent.append(copy.deepcopy(seg.element))

