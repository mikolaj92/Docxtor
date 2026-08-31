from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .docx_models import AddressableSpan, SpanRole
from .docx_ns import _TEXT_NODE_TAGS, R_ID
from .docx_xml import _is_text_box_container, _local_tag, _w_get


@dataclass(frozen=True)
class _TextUnit:
    text: str
    node: Any | None
    role: SpanRole
    revision_id: str | None = None
    revision_author: str | None = None
    revision_date: str | None = None
    hyperlink_anchor: str | None = None
    hyperlink_rel_id: str | None = None

    @property
    def key(self) -> tuple[object, ...]:
        return (
            self.role,
            self.revision_id,
            self.revision_author,
            self.revision_date,
            self.hyperlink_anchor,
            self.hyperlink_rel_id,
        )


def _wrapper_context(node: Any) -> dict[str, Any]:
    role: SpanRole = "run"
    revision_id = revision_author = revision_date = None
    hyperlink_anchor = hyperlink_rel_id = None
    for ancestor in node.iterancestors():
        local = _local_tag(ancestor.tag)
        if local == "p":
            break
        if local == "hyperlink":
            if role == "run":
                role = "hyperlink"
            if hyperlink_anchor is None and hyperlink_rel_id is None:
                hyperlink_anchor = _w_get(ancestor, "anchor")
                hyperlink_rel_id = ancestor.get(R_ID)
        elif local == "ins":
            if role == "run":
                role = "insertion"
            if revision_id is None:
                revision_id = _w_get(ancestor, "id")
                revision_author = _w_get(ancestor, "author")
                revision_date = _w_get(ancestor, "date")
        elif local == "del":
            if role == "run":
                role = "deletion"
            if revision_id is None:
                revision_id = _w_get(ancestor, "id")
                revision_author = _w_get(ancestor, "author")
                revision_date = _w_get(ancestor, "date")
    return {
        "role": role,
        "revision_id": revision_id,
        "revision_author": revision_author,
        "revision_date": revision_date,
        "hyperlink_anchor": hyperlink_anchor,
        "hyperlink_rel_id": hyperlink_rel_id,
    }


def _paragraph_units(p_element: Any) -> list[_TextUnit]:
    """Document-order text contributions inside one paragraph element."""
    units: list[_TextUnit] = []

    def walk(node: Any) -> None:
        if node is not p_element and _is_text_box_container(node.tag):
            return
        tag = node.tag
        if tag in _TEXT_NODE_TAGS:
            units.append(_TextUnit(text=node.text or "", node=node, **_wrapper_context(node)))
            return
        if tag == qn("w:tab"):
            units.append(_TextUnit(text="\t", node=None, **_wrapper_context(node)))
            return
        if tag in (qn("w:br"), qn("w:cr")):
            units.append(_TextUnit(text="\n", node=None, **_wrapper_context(node)))
            return
        for child in node:
            walk(child)

    walk(p_element)
    return units


def _set_text_node(node: Any, text: str) -> None:
    node.text = text
    space = qn("xml:space")
    if text[:1].isspace() or text[-1:].isspace():
        node.set(space, "preserve")
    elif space in node.attrib:
        del node.attrib[space]


def _writable_units(p_element: Any) -> list[_TextUnit]:
    """Text nodes that can receive an in-place character replacement."""
    return [unit for unit in _paragraph_units(p_element) if unit.node is not None]


def _replace_plain_range(p_element: Any, start: int, end: int, replacement: str) -> None:
    """Replace [start, end) in paragraph plain text without unwrapping OOXML."""
    units = _paragraph_units(p_element)
    writable_ids = {id(unit.node) for unit in _writable_units(p_element)}
    cursor = 0
    overlapping: list[tuple[_TextUnit, int]] = []
    for unit in units:
        unit_start = cursor
        cursor += len(unit.text)
        if unit.node is None or id(unit.node) not in writable_ids:
            continue
        if unit_start < end and cursor > start and unit.text:
            overlapping.append((unit, unit_start))
    if not overlapping:
        raise ValueError(f"could not map replacement offsets to runs for segment: {start}:{end}")

    first_unit, first_start = overlapping[0]
    last_unit, last_start = overlapping[-1]
    first_node = first_unit.node
    last_node = last_unit.node
    first_text = first_node.text or ""
    last_text = last_node.text or ""
    prefix = first_text[: max(0, start - first_start)]
    suffix = last_text[max(0, end - last_start) :]
    if first_node is last_node:
        _set_text_node(first_node, prefix + replacement + suffix)
        return
    _set_text_node(first_node, prefix + replacement)
    for unit, _unit_start in overlapping[1:-1]:
        _set_text_node(unit.node, "")
    _set_text_node(last_node, suffix)


def _paragraph_spans(
    paragraph: Paragraph, container_id: str, paragraph_index: int
) -> list[AddressableSpan]:
    spans: list[AddressableSpan] = []
    buf: list[str] = []
    current: _TextUnit | None = None
    start = 0
    cursor = 0

    def flush() -> None:
        nonlocal current, start
        if current is None:
            return
        text = "".join(buf)
        buf.clear()
        if not text:
            current = None
            return
        spans.append(
            AddressableSpan(
                span_id=f"{container_id}:span:{len(spans)}",
                container_id=container_id,
                role=current.role,
                text=text,
                start_offset=start,
                end_offset=start + len(text),
                paragraph_index=paragraph_index,
                revision_id=current.revision_id,
                revision_author=current.revision_author,
                revision_date=current.revision_date,
                hyperlink_anchor=current.hyperlink_anchor,
                hyperlink_rel_id=current.hyperlink_rel_id,
            )
        )
        current = None

    for unit in _paragraph_units(paragraph._p):
        if not unit.text:
            continue
        if current is None:
            current = unit
            start = cursor
        elif unit.key != current.key:
            flush()
            current = unit
            start = cursor
        buf.append(unit.text)
        cursor += len(unit.text)
    flush()
    return spans


def _paragraph_visible_text(paragraph: Paragraph) -> str:
    """Return paragraph text including nested ins/del/hyperlink/sdt text nodes."""
    return "".join(unit.text for unit in _paragraph_units(paragraph._p))
