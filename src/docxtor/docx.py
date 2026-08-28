from __future__ import annotations

import copy
from collections.abc import Iterable
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document as PyDocxDocument
from docx.document import Document as DocxDocumentType
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .common import DocumentError

W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")
W_DEL_TEXT = qn("w:delText")
W_SDT = qn("w:sdt")
W_SDT_CONTENT = qn("w:sdtContent")
W_TBL = qn("w:tbl")
W_TXBX_CONTENT = qn("w:txbxContent")
V_TEXTBOX = "{urn:schemas-microsoft-com:vml}textbox"
WPS_TXBX = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}txbx"
R_ID = qn("r:id")
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_COMMENTS_EXTENDED_PART = "word/commentsExtended.xml"
_COMMENTS_IDS_PART = "word/commentsIds.xml"
_COMMENTS_EXTENSIBLE_PART = "word/commentsExtensible.xml"
_PEOPLE_PART = "word/people.xml"
_DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"
_CONTENT_TYPES_PART = "[Content_Types].xml"
_THREAD_PARTS = {
    _COMMENTS_EXTENDED_PART,
    _COMMENTS_IDS_PART,
    _COMMENTS_EXTENSIBLE_PART,
    _PEOPLE_PART,
}
_THREAD_REL_BY_TARGET = {
    "commentsExtended.xml": (
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "application/vnd.ms-word.commentsExtended+xml",
    ),
    "commentsIds.xml": (
        "http://schemas.microsoft.com/office/2016/relationships/commentsIds",
        "application/vnd.ms-word.commentsIds+xml",
    ),
    "commentsExtensible.xml": (
        "http://schemas.microsoft.com/office/2018/relationships/commentsExtensible",
        "application/vnd.ms-word.commentsExtensible+xml",
    ),
    "people.xml": (
        "http://schemas.microsoft.com/office/2011/relationships/people",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
    ),
}

_TEXT_NODE_TAGS = {W_T, W_DEL_TEXT}
_UNSUPPORTED_MOVE_TAGS = {
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
}


@dataclass(frozen=True)
class TextSegment:
    """Stable textual segment inside a DOCX.

    container_id examples:
      body:p:0
      header:0:p:0
      table:0:r:0:c:0:p:0   (for table cells)
      txbx:0:p:0            (floating text box)

    paragraph_index is the global order index (body + tables + headers/footers
    + text boxes), counting ALL paragraphs including empty ones for stable
    addressing.

    run_indices: indices (within the python-docx paragraph.runs) of runs that
    contributed non-empty text at parse time. Useful for domain anchors.
    """

    id: str
    text: str
    part: str
    index: int
    container_id: str | None = None
    paragraph_index: int | None = None
    run_indices: list[int] | None = None


@dataclass(frozen=True)
class SegmentReplacement:
    """Replacement targeting a segment or a sub-range inside it.

    Offsets are in characters of the segment's text.
    If start_offset and end_offset are None -> whole segment.
    """

    container_id: str | None = None
    id: str | None = None
    span_id: str | None = None
    text: str = ""
    start_offset: int | None = None
    end_offset: int | None = None


class UnsupportedRevisionError(DocumentError):
    """Raised when a DOCX uses a revision form that cannot be written in place."""


SpanRole = Literal["run", "insertion", "deletion", "hyperlink"]
InlineSegmentKind = Literal["text", "opaque"]


@dataclass(frozen=True)
class AddressableSpan:
    """Stable mechanical span inside a paragraph's extracted text.

    ``role`` is the OOXML wrapper that owns the characters:
    ``run``, ``insertion`` (``w:ins``), ``deletion`` (``w:del`` / ``w:delText``),
    or ``hyperlink`` (``w:hyperlink``). Nested wrappers keep both identities:
    a hyperlink inside an insertion is ``role="hyperlink"`` and still carries
    revision author/date/id.

    Offsets are in the same character space as ``TextSegment.text`` and
    ``SegmentReplacement`` (including deleted text). Consumers can project an
    after-changes view or a revision ledger; Docxtor does not.
    """

    span_id: str
    container_id: str
    role: SpanRole
    text: str
    start_offset: int
    end_offset: int
    paragraph_index: int | None = None
    revision_id: str | None = None
    revision_author: str | None = None
    revision_date: str | None = None
    hyperlink_anchor: str | None = None
    hyperlink_rel_id: str | None = None


@dataclass(frozen=True)
class AddressableComment:
    """Mechanical Word comment body with stable identity and thread metadata.

    ``container_id`` addresses the comment body paragraph that carries the text
    (``comment:{id}:p:{n}``). ``locator`` is the story paragraph that hosts the
    ``w:commentRangeStart`` marker; replies typically have none. Parent/reply
    identity comes from ``commentsExtended.xml`` when present. Docxtor does not
    interpret review meaning.
    """

    comment_id: str
    container_id: str
    text: str
    author: str = ""
    initials: str | None = None
    locator: str | None = None
    anchor_text: str = ""
    parent_id: str | None = None
    date: str | None = None


@dataclass
class InlineSegment:
    """Canonical mechanical segment for paragraph-level DOCX manipulation.

    This is the single source of truth for run/offset addressing, visible-text
    coordinate math, rPr formatting preservation, and opaque inline content
    (images, tabs, breaks, fields, hyperlinks, pre-existing revisions, ...).

    reviewkit (and other consumers) MUST delegate decomposition, splitting,
    insertion, and range replacement to this representation instead of
    reimplementing the logic.

    - kind="text": editable run text; rpr carries formatting to preserve on split/replace.
    - kind="opaque": non-text inline; element is the original XML to re-emit verbatim;
      text holds the visible contribution (e.g. "\t", "\n", or extracted t text) so
      char offsets stay aligned with parser coordinate systems.
    """

    kind: InlineSegmentKind
    text: str
    rpr: Any | None = None
    element: Any | None = None


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


def _w_get(element: Any, name: str) -> str | None:
    return element.get(qn(f"w:{name}"))


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


def _existing_comments_part(doc: DocxDocumentType) -> Any | None:
    try:
        return doc.part.part_related_by(RT.COMMENTS)
    except KeyError:
        return None


def _unsupported_revision_reason(doc: DocxDocumentType) -> str | None:
    roots: list[Any] = [doc.element]
    for section in doc.sections:
        for story in (section.header, section.footer):
            if not story._has_definition:
                continue
            roots.append(story._definition.element)
    comments_part = _existing_comments_part(doc)
    if comments_part is not None:
        roots.append(comments_part.element)
    for root in roots:
        for element in root.iter():
            local = _local_tag(element.tag)
            if local in _UNSUPPORTED_MOVE_TAGS:
                return local
            if local in {"ins", "del"}:
                for child in element:
                    if _local_tag(child.tag) in {"p", "tbl", "tr", "tc"}:
                        return f"block-{local}"
    return None


def _clark(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"


def _package_part_name(partname: Any) -> str:
    text = str(partname)
    return text[1:] if text.startswith("/") else text


def _is_thread_part_name(name: str) -> bool:
    return (
        name in _THREAD_PARTS
        or name.startswith("word/_rels/comments")
        or name.startswith("word/_rels/people")
    )


def _capture_thread_parts(package: Any) -> dict[str, tuple[bytes, str]]:
    captured: dict[str, tuple[bytes, str]] = {}
    for part in package.iter_parts():
        name = _package_part_name(part.partname)
        if _is_thread_part_name(name):
            captured[name] = (part.blob, part.content_type)
    return captured


def _ensure_thread_parts(package: Any, captured: dict[str, tuple[bytes, str]]) -> None:
    if not captured:
        return
    document_part = package.main_document_part
    existing = {_package_part_name(part.partname): part for part in package.iter_parts()}
    for name, (blob, content_type) in captured.items():
        if name.startswith("word/_rels/"):
            continue
        part = existing.get(name)
        if part is None:
            part = Part(PackURI(f"/{name}"), content_type, blob, package)
            target = name.rsplit("/", 1)[-1]
            reltype, _ctype = _THREAD_REL_BY_TARGET.get(target, (None, None))
            if reltype is None:
                continue
            document_part.relate_to(part, reltype)
            existing[name] = part
        else:
            part._blob = blob


def _restore_thread_sidecars(data: bytes, captured: dict[str, tuple[bytes, str]]) -> bytes:
    if not captured:
        return data
    with ZipFile(BytesIO(data)) as bundle:
        names = set(bundle.namelist())
        entries = [(info, bundle.read(info.filename)) for info in bundle.infolist()]
    missing = {name: blob for name, (blob, _ctype) in captured.items() if name not in names}
    if not missing:
        return data

    restored: list[tuple[Any, bytes]] = []
    for info, blob in entries:
        if info.filename == _CONTENT_TYPES_PART:
            blob = _merge_content_type_overrides(blob, missing, captured)
        elif info.filename == _DOCUMENT_RELS_PART:
            blob = _merge_document_relationships(blob, missing)
        restored.append((info, blob))
    for name, blob in sorted(missing.items()):
        restored.append((name, blob))

    output = BytesIO()
    with ZipFile(output, mode="w", compression=ZIP_DEFLATED) as bundle:
        for info, blob in restored:
            if isinstance(info, str):
                bundle.writestr(info, blob)
            else:
                bundle.writestr(info, blob)
    return output.getvalue()


def _merge_content_type_overrides(
    rendered: bytes,
    missing: dict[str, bytes],
    captured: dict[str, tuple[bytes, str]],
) -> bytes:
    root = ElementTree.fromstring(rendered)
    have = {override.get("PartName") for override in root.findall(_clark(CT_NS, "Override"))}
    changed = False
    for name in missing:
        part_name = f"/{name}"
        if part_name in have:
            continue
        _blob, content_type = captured[name]
        override = ElementTree.SubElement(root, _clark(CT_NS, "Override"))
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        changed = True
    if not changed:
        return rendered
    return ElementTree.tostring(root, encoding="UTF-8", xml_declaration=True)


def _merge_document_relationships(rendered: bytes, missing: dict[str, bytes]) -> bytes:
    root = ElementTree.fromstring(rendered)
    have_targets = {rel.get("Target") for rel in root.findall(_clark(REL_NS, "Relationship"))}
    used_ids = {rel.get("Id") for rel in root.findall(_clark(REL_NS, "Relationship"))}
    next_id = 1
    changed = False
    for name in missing:
        target = name.rsplit("/", 1)[-1]
        reltype, _ctype = _THREAD_REL_BY_TARGET.get(target, (None, None))
        if reltype is None or target in have_targets:
            continue
        while f"rId{next_id}" in used_ids:
            next_id += 1
        rel = ElementTree.SubElement(root, _clark(REL_NS, "Relationship"))
        rel.set("Id", f"rId{next_id}")
        rel.set("Type", reltype)
        rel.set("Target", target)
        used_ids.add(f"rId{next_id}")
        have_targets.add(target)
        next_id += 1
        changed = True
    if not changed:
        return rendered
    return ElementTree.tostring(root, encoding="UTF-8", xml_declaration=True)


def _comment_date_attr(element: Any) -> str | None:
    return _w_get(element, "date")


def _comment_parent_ids(comments_part: Any, package: Any) -> dict[str, str]:
    parts = {_package_part_name(part.partname): part for part in package.iter_parts()}
    extended = parts.get(_COMMENTS_EXTENDED_PART)
    if comments_part is None or extended is None:
        return {}
    para_to_comment: dict[str, str] = {}
    for comment in comments_part.element.findall(qn("w:comment")):
        comment_id = _w_get(comment, "id")
        if comment_id is None:
            continue
        for paragraph in comment.iter(qn("w:p")):
            para_id = paragraph.get(f"{{{W14_NS}}}paraId")
            if para_id:
                para_to_comment[para_id] = comment_id
                break
    if not para_to_comment:
        return {}
    try:
        root = ElementTree.fromstring(extended.blob)
    except ElementTree.ParseError:
        return {}
    parents: dict[str, str] = {}
    for element in root.iter():
        if element.tag != _clark(W15_NS, "commentEx"):
            continue
        para_id = element.get(_clark(W15_NS, "paraId"))
        parent_para = element.get(_clark(W15_NS, "paraIdParent"))
        if not para_id or not parent_para:
            continue
        comment_id = para_to_comment.get(para_id)
        parent_id = para_to_comment.get(parent_para)
        if comment_id and parent_id and comment_id != parent_id:
            parents[comment_id] = parent_id
    return parents


def _comment_anchors(
    paragraphs_by_container: dict[str, Paragraph],
) -> dict[str, tuple[str, str]]:
    open_ids: dict[str, list[str]] = {}
    started_at: dict[str, str] = {}
    finished: dict[str, tuple[str, str]] = {}
    for locator, paragraph in paragraphs_by_container.items():
        if locator.startswith("comment:"):
            continue
        p_element = paragraph._p
        for node in p_element.iter():
            local = _local_tag(node.tag)
            if local == "commentRangeStart":
                comment_id = _w_get(node, "id")
                if comment_id is None:
                    continue
                open_ids.setdefault(comment_id, [])
                started_at.setdefault(comment_id, locator)
            elif local == "commentRangeEnd":
                comment_id = _w_get(node, "id")
                if comment_id is None or comment_id not in open_ids:
                    continue
                finished[comment_id] = (
                    started_at.get(comment_id, locator),
                    "".join(open_ids.pop(comment_id)),
                )
                started_at.pop(comment_id, None)
            elif open_ids and node.tag in _TEXT_NODE_TAGS and node.text:
                for buffer in open_ids.values():
                    buffer.append(node.text)
            elif open_ids and node.tag == qn("w:tab"):
                for buffer in open_ids.values():
                    buffer.append("\t")
            elif open_ids and node.tag in (qn("w:br"), qn("w:cr")):
                for buffer in open_ids.values():
                    buffer.append("\n")
    for comment_id, buffer in open_ids.items():
        finished[comment_id] = (started_at.get(comment_id, ""), "".join(buffer))
    return finished


def _collect_comments(
    comments_part: Any | None,
    paragraphs_by_container: dict[str, Paragraph],
    package: Any,
) -> list[AddressableComment]:
    if comments_part is None:
        return []
    anchors = _comment_anchors(paragraphs_by_container)
    parents = _comment_parent_ids(comments_part, package)
    comments: list[AddressableComment] = []
    for comment_elm in comments_part.element.findall(qn("w:comment")):
        comment_id = _w_get(comment_elm, "id")
        if comment_id is None:
            continue
        prefix = f"comment:{comment_id}"
        texts: list[str] = []
        first_container: str | None = None
        for local_idx, paragraph in enumerate(comment_elm.iter(qn("w:p"))):
            text = "".join(unit.text for unit in _paragraph_units(paragraph))
            if not text:
                continue
            if first_container is None:
                first_container = f"{prefix}:p:{local_idx}"
            texts.append(text)
        if not texts:
            continue
        locator, anchor_text = anchors.get(comment_id, (None, ""))
        comments.append(
            AddressableComment(
                comment_id=comment_id,
                container_id=first_container or f"{prefix}:p:0",
                text="\n".join(texts),
                author=_w_get(comment_elm, "author") or "",
                initials=_w_get(comment_elm, "initials"),
                locator=locator or None,
                anchor_text=anchor_text,
                parent_id=parents.get(comment_id),
                date=_comment_date_attr(comment_elm),
            )
        )
    return comments


def _paragraph_visible_text(paragraph: Paragraph) -> str:
    """Return paragraph text including nested ins/del/hyperlink/sdt text nodes."""
    return "".join(unit.text for unit in _paragraph_units(paragraph._p))


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


def _local_tag(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _is_text_box_container(tag: str) -> bool:
    """True for Word/VML/DrawingML text-box wrappers.

    ``w:txbxContent`` is the actual paragraph host. VML ``v:textbox`` and
    DrawingML ``wps:txbx`` wrap that host; walk into them so a box nested
    inside a drawing is not silence.
    """
    if tag in {W_TXBX_CONTENT, V_TEXTBOX, WPS_TXBX}:
        return True
    return _local_tag(tag) in {"txbxContent", "textbox", "txbx"}


def _iter_text_box_hosts(root: Any) -> list[Any]:
    """Return each text-bearing ``w:txbxContent`` (or VML/DrawingML wrapper).

    Nested boxes keep document order. Empty decorative shapes (no descendant
    ``w:t``) are skipped so they never occupy a ``txbx:N`` slot.
    """
    hosts: list[Any] = []
    seen: set[int] = set()

    def has_text(element: Any) -> bool:
        for node in element.iter():
            if _local_tag(node.tag) in {"t", "delText"} and node.text and node.text.strip():
                return True
        return False

    for element in root.iter():
        if not _is_text_box_container(element.tag):
            continue
        identity = id(element)
        if identity in seen:
            continue
        if _local_tag(element.tag) != "txbxContent":
            inner = None
            for child in element.iter():
                if child is not element and _local_tag(child.tag) == "txbxContent":
                    inner = child
                    break
            if inner is not None:
                if id(inner) in seen:
                    continue
                if not has_text(inner):
                    seen.add(id(inner))
                    continue
                seen.add(id(inner))
                hosts.append(inner)
                continue
        if not has_text(element):
            seen.add(identity)
            continue
        seen.add(identity)
        hosts.append(element)
    return hosts


def _iter_paragraph_elements(container: Any, *, skip_text_boxes: bool = False) -> list[Any]:
    """Collect w:p elements in document order, including those nested in w:sdt.

    python-docx's ``.paragraphs`` only returns direct ``w:p`` children and therefore
    omits content-control (``w:sdt`` / ``w:sdtContent``) paragraphs. Tables are left
    to the caller's existing table walk so global ordering stays unchanged.

    Floating text boxes (``w:txbxContent``) are a separate container family. The
    body/header/table walk skips them so boxed paragraphs are not double-counted
    as body runs; ``_iter_text_box_hosts`` indexes them as ``txbx:N``.
    """
    result: list[Any] = []

    def walk(element: Any) -> None:
        for child in element:
            tag = child.tag
            if skip_text_boxes and _is_text_box_container(tag):
                continue
            if tag == W_P:
                result.append(child)
            elif tag == W_SDT:
                for content in child.iterchildren(W_SDT_CONTENT):
                    walk(content)
            # Nested tables are handled by the dedicated table enumeration path.

    walk(container)
    return result


def _paragraphs_from_container(
    container_element: Any,
    parent: Any,
    *,
    skip_text_boxes: bool = False,
) -> list[Paragraph]:
    """Wrap collected paragraph elements as python-docx Paragraph proxies."""
    return [
        Paragraph(p, parent)
        for p in _iter_paragraph_elements(container_element, skip_text_boxes=skip_text_boxes)
    ]


@dataclass
class _ParaRef:
    """Internal mapping from our segment to python-docx paragraph + metadata."""

    id: str
    container_id: str
    paragraph_index: int | None
    paragraph: Paragraph
    part_name: str  # "body", "header:0", "table:0:r:0:c:0", etc.


class DocxDocument:
    """DOCX editing surface backed by python-docx (the proper library for the format).

    - Stable container_id + paragraph_index addressing.
    - Whole segment or offset-based partial replacements (run splitting).
    - Preserves formatting because we operate on runs.
    - Roundtrips via python-docx save.
    """

    def __init__(
        self,
        doc: DocxDocumentType,
        segments: list[TextSegment],
        refs: list[_ParaRef],
    ) -> None:
        self._doc = doc
        self._segments = segments
        self._refs = refs  # index-aligned with segments
        self._spans: list[AddressableSpan] = []
        self._comments: list[AddressableComment] = []
        self._thread_parts: dict[str, tuple[bytes, str]] = {}

    @classmethod
    def open(cls, path: str | Path) -> DocxDocument:
        path = Path(path)
        doc = PyDocxDocument(str(path))
        return cls._from_pydocx(doc)

    @classmethod
    def open_bytes(cls, data: bytes) -> DocxDocument:
        doc = PyDocxDocument(BytesIO(data))
        return cls._from_pydocx(doc)

    @classmethod
    def _from_pydocx(cls, doc: DocxDocumentType) -> DocxDocument:
        segments: list[TextSegment] = []
        refs: list[_ParaRef] = []

        # Global paragraph index counts EVERY paragraph in document order
        # (body, table cells, headers, footers, text boxes), including empty
        # ones. This matches the contract expected by dike_docs locator and
        # anchors. Text boxes are appended after headers/footers so documents
        # without boxes keep existing body/table/header indices.
        global_paragraph_index = 0
        paragraphs_by_index: dict[int, Paragraph] = {}
        paragraphs_by_container: dict[str, Paragraph] = {}

        def add_paragraphs(paragraphs: list[Paragraph], prefix: str) -> None:
            nonlocal global_paragraph_index
            for local_idx, para in enumerate(paragraphs):
                paragraphs_by_index[global_paragraph_index] = para

                # container_id: body uses global index for stability (matches Dike anchors);
                # other sections use local index within their container.
                if prefix == "body":
                    cid = f"body:p:{global_paragraph_index}"
                else:
                    cid = f"{prefix}:p:{local_idx}"

                paragraphs_by_container[cid] = para

                text = _paragraph_visible_text(para)
                run_indices = (
                    [ri for ri, run in enumerate(para.runs) if run.text] if para.runs else []
                )

                if text:
                    seg_id = f"s{len(segments)}"
                    segments.append(
                        TextSegment(
                            id=seg_id,
                            text=text,
                            part=(
                                "word/comments.xml"
                                if prefix.startswith("comment:")
                                else "word/document.xml"
                                if prefix.startswith(("body", "table", "txbx"))
                                else f"word/{prefix.split(':')[0]}.xml"
                            ),
                            index=local_idx,
                            container_id=cid,
                            paragraph_index=global_paragraph_index,
                            run_indices=run_indices,
                        )
                    )
                    refs.append(
                        _ParaRef(
                            id=seg_id,
                            container_id=cid,
                            paragraph_index=global_paragraph_index,
                            paragraph=para,
                            part_name=prefix,
                        )
                    )

                global_paragraph_index += 1

        def add_comment_paragraphs(paragraphs: list[Paragraph], comment_id: str) -> None:
            for local_idx, para in enumerate(paragraphs):
                cid = f"comment:{comment_id}:p:{local_idx}"
                paragraphs_by_container[cid] = para
                text = _paragraph_visible_text(para)
                if not text:
                    continue
                seg_id = f"s{len(segments)}"
                segments.append(
                    TextSegment(
                        id=seg_id,
                        text=text,
                        part="word/comments.xml",
                        index=local_idx,
                        container_id=cid,
                        paragraph_index=None,
                        run_indices=[
                            run_index for run_index, run in enumerate(para.runs) if run.text
                        ],
                    )
                )
                refs.append(
                    _ParaRef(
                        id=seg_id,
                        container_id=cid,
                        paragraph_index=None,
                        paragraph=para,
                        part_name=f"comment:{comment_id}",
                    )
                )

        # Body (include w:sdt/w:sdtContent paragraphs omitted by python-docx).
        # Skip floating text boxes here; they are indexed as txbx:N below.
        add_paragraphs(
            _paragraphs_from_container(doc.element.body, doc._body, skip_text_boxes=True),
            "body",
        )

        # Tables
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    add_paragraphs(
                        _paragraphs_from_container(cell._tc, cell, skip_text_boxes=True),
                        f"table:{ti}:r:{ri}:c:{ci}",
                    )

        # Headers / Footers. Accessing ``._element`` on an absent or linked
        # story calls python-docx's get-or-add path and mutates the package.
        # Only inspect definitions explicitly referenced by this section.
        for si, section in enumerate(doc.sections):
            for story_name, story in (
                ("header", section.header),
                ("footer", section.footer),
            ):
                if not story._has_definition:
                    continue
                definition = story._definition
                add_paragraphs(
                    _paragraphs_from_container(definition.element, story, skip_text_boxes=True),
                    f"{story_name}:{si}",
                )

        # Floating text boxes (VML v:textbox / w:txbxContent / DrawingML wps:txbx).
        # python-docx does not surface these as paragraphs. Index after the
        # ordinary stories so documents without boxes keep stable body ids.
        for box_idx, host in enumerate(_iter_text_box_hosts(doc.element)):
            add_paragraphs(
                _paragraphs_from_container(host, doc._body, skip_text_boxes=True),
                f"txbx:{box_idx}",
            )

        comments_part = _existing_comments_part(doc)
        if comments_part is not None:
            for comment_elm in comments_part.element.findall(qn("w:comment")):
                comment_id = _w_get(comment_elm, "id")
                if comment_id is None:
                    continue
                add_comment_paragraphs(
                    _paragraphs_from_container(comment_elm, comments_part),
                    comment_id,
                )

        instance = cls(doc=doc, segments=segments, refs=refs)
        instance._paragraphs_by_index = paragraphs_by_index
        instance._paragraphs_by_container = paragraphs_by_container
        instance._spans = instance._collect_spans()
        instance._comments = _collect_comments(
            comments_part, paragraphs_by_container, doc.part.package
        )
        instance._thread_parts = _capture_thread_parts(doc.part.package)
        _ensure_thread_parts(doc.part.package, instance._thread_parts)
        return instance

    @property
    def segments(self) -> tuple[TextSegment, ...]:
        return tuple(self._segments)

    @property
    def texts(self) -> list[str]:
        return [s.text for s in self._segments]

    @property
    def spans(self) -> tuple[AddressableSpan, ...]:
        return tuple(self._spans)

    @property
    def comments(self) -> tuple[AddressableComment, ...]:
        return tuple(self._comments)

    # ------------------------------------------------------------------
    # Structure access (generic DOCX addressing - for Temida adapters)
    # ------------------------------------------------------------------

    def resolve_paragraph(self, container_id: str) -> Paragraph | None:
        """Resolve a python-docx Paragraph by stable container_id.

        container_id examples: "body:p:0", "body:p:17", "header:0:p:0",
        "table:0:r:1:c:2:p:0", "txbx:0:p:0".
        """
        if not hasattr(self, "_paragraphs_by_container"):
            return None
        return self._paragraphs_by_container.get(container_id)

    def resolve_paragraph_by_index(self, index: int) -> Paragraph | None:
        """Resolve by global paragraph index (counts every paragraph in order,
        including empty ones). Matches dike/posejdon locator contracts.
        """
        if not hasattr(self, "_paragraphs_by_index"):
            return None
        return self._paragraphs_by_index.get(index)

    def get_all_paragraphs(self) -> list[Paragraph]:
        """All paragraphs in document order (body, tables, headers, footers, boxes).
        Includes empty paragraphs to keep index stable.
        """
        if not hasattr(self, "_paragraphs_by_index") or not self._paragraphs_by_index:
            return []
        max_i = max(self._paragraphs_by_index.keys())
        return [
            self._paragraphs_by_index[i] for i in range(max_i + 1) if i in self._paragraphs_by_index
        ]

    def get_inline_segments(self, container_id: str) -> list[InlineSegment]:
        """Return the canonical rich InlineSegment decomposition for one paragraph.

        container_id examples: "body:p:0", "header:0:p:0", table cell variants.
        This is the bridge for review-specific layers to obtain the mechanical view
        (text + opaque with rpr/element) and then use the pure offset functions
        (_split_visible_offset, _insert_visible, _replace_visible_range, etc.)
        without reimplementing paragraph traversal or run decomposition.
        """
        para = self.resolve_paragraph(container_id)
        if para is None:
            return []
        return paragraph_to_inline_segments(para)

    # ------------------------------------------------------------------
    # High-level target application (WriteTarget style)
    # ------------------------------------------------------------------

    def apply_targets(
        self,
        targets: list[dict[str, Any] | SegmentReplacement],
        *,
        strict: bool = False,
    ) -> None:
        """Apply a list of replacement targets.

        Each target can be:
          - SegmentReplacement
          - dict with keys: container_id or id, text, optional start_offset/end_offset
          - object with .container_id, .start_offset, .end_offset, .text (e.g. WriteTarget)

        This is the bridge for ReplacementPlan.write_targets.
        """
        normalized: list[SegmentReplacement] = []
        for target in targets:
            if isinstance(target, SegmentReplacement):
                normalized.append(target)
                continue

            if isinstance(target, dict):
                normalized.append(
                    SegmentReplacement(
                        container_id=target.get("container_id"),
                        id=target.get("id"),
                        span_id=target.get("span_id"),
                        text=str(target.get("text", "")),
                        start_offset=target.get("start_offset"),
                        end_offset=target.get("end_offset"),
                    )
                )
                continue

            # duck-type WriteTarget-like
            normalized.append(
                SegmentReplacement(
                    container_id=getattr(target, "container_id", None),
                    id=getattr(target, "segment_id", None),
                    span_id=getattr(target, "span_id", None),
                    text=str(getattr(target, "text", getattr(target, "replacement_text", ""))),
                    start_offset=getattr(target, "start_offset", None),
                    end_offset=getattr(target, "end_offset", None),
                )
            )
        self.apply_replacements(normalized, strict=strict)

    def to_markdown(self) -> str:
        blocks = [f"<!-- docxtor:{s.id} -->\n{s.text}" for s in self._segments]
        return "\n\n".join(blocks)

    def get_indexed_paragraphs(self) -> list[tuple[int, str, Paragraph]]:
        """Return every paragraph in document order with its stable identifiers.

        Returns list of (global_paragraph_index, container_id, python-docx.Paragraph).
        Includes empty paragraphs so that paragraph_index stays in sync with
        dike/posejdon anchor contracts (body + tables + headers/footers + boxes).
        This is the canonical source of addressing.
        """
        if not hasattr(self, "_paragraphs_by_index") or not self._paragraphs_by_index:
            return []
        max_i = max(self._paragraphs_by_index.keys())
        out: list[tuple[int, str, Paragraph]] = []
        for i in range(max_i + 1):
            if i not in self._paragraphs_by_index:
                continue
            para = self._paragraphs_by_index[i]
            # find a container_id for it (prefer body: global, else scan)
            cid = f"body:p:{i}"
            if cid not in self._paragraphs_by_container:
                for c, p in self._paragraphs_by_container.items():
                    if p is para:
                        cid = c
                        break
            out.append((i, cid, para))
        return out

    # ------------------------------------------------------------------
    # Placeholder replacement (mechanical, for reinjection)
    # ------------------------------------------------------------------

    def replace_placeholder(
        self,
        container_id: str,
        placeholder: str,
        replacement: str,
    ) -> None:
        """Mechanical: in the *current* text of the paragraph identified by container_id,
        find the first occurrence of placeholder and replace it with replacement.

        This is the non-domain part of reinjection flows.
        Offsets are computed on the live paragraph text; run splitting is handled internally.
        """
        para = self.resolve_paragraph(container_id)
        if para is None:
            raise ValueError(f"no paragraph for container_id {container_id!r}")

        current_text = _paragraph_visible_text(para)
        start = current_text.find(placeholder)
        if start < 0:
            raise ValueError(f"placeholder {placeholder!r} not found in segment {container_id}")
        end = start + len(placeholder)
        self.apply_replacements(
            [
                SegmentReplacement(
                    container_id=container_id,
                    text=replacement,
                    start_offset=start,
                    end_offset=end,
                )
            ],
            strict=True,
        )

    # ------------------------------------------------------------------
    # Replacement API (supports full + offset ranges via python-docx)
    # ------------------------------------------------------------------

    def apply_texts(self, texts: Iterable[str], *, strict: bool = False) -> None:
        texts = list(texts)
        if len(texts) != len(self._segments):
            raise ValueError(f"expected {len(self._segments)} segments, got {len(texts)}")
        self._require_supported_revisions()
        for i, txt in enumerate(texts):
            self._replace_full_segment(i, txt)

    def apply_replacements(
        self,
        replacements: list[SegmentReplacement],
        *,
        strict: bool = False,
    ) -> None:
        for replacement in replacements:
            if not isinstance(replacement, SegmentReplacement):
                raise TypeError("replacements must contain only SegmentReplacement instances")
        if replacements:
            self._require_supported_revisions()
        by_container = {r.container_id: i for i, r in enumerate(self._segments)}
        by_id = {r.id: i for i, r in enumerate(self._segments)}
        resolved: list[tuple[int, int | None, int | None, str]] = []
        for replacement in replacements:
            idx, start, end = self._resolve_replacement(replacement, by_container, by_id, strict)
            if idx is None:
                continue
            full = self._segments[idx].text
            if start is not None or end is not None:
                s = 0 if start is None else start
                e = len(full) if end is None else end
                if not 0 <= s < e <= len(full):
                    raise ValueError(
                        f"invalid replacement offsets for segment "
                        f"{self._refs[idx].container_id}: expected "
                        f"0 <= start < end <= {len(full)}, got {s}:{e}"
                    )
            resolved.append((idx, start, end, replacement.text))
        for idx, start, end, text in resolved:
            self._apply_to_paragraph(idx, text, start, end)

    def apply_markdown(self, markdown: str, *, strict: bool = True) -> None:
        import re as _re

        by_id = {
            m.group("id"): m.group("text").rstrip("\n")
            for m in _re.finditer(
                r"<!-- docxtor:(?P<id>s\d+) -->\n(?P<text>.*?)"
                r"(?=\n<!-- docxtor:s\d+ -->\n|\Z)",
                markdown,
                _re.DOTALL,
            )
        }
        if strict:
            expected = {s.id for s in self._segments}
            actual = set(by_id.keys())
            missing = sorted(expected - actual)
            unknown = sorted(actual - expected)
            if missing or unknown:
                raise ValueError(f"markdown marker mismatch; missing={missing} unknown={unknown}")

        if any(seg.id in by_id for seg in self._segments):
            self._require_supported_revisions()
        for i, seg in enumerate(self._segments):
            if seg.id in by_id:
                self._replace_full_segment(i, by_id[seg.id])

    # ------------------------------------------------------------------
    # Save / bytes
    # ------------------------------------------------------------------

    def save_docx(self, path: str | Path) -> None:
        Path(path).write_bytes(self.to_bytes())

    def to_bytes(self) -> bytes:
        _ensure_thread_parts(self._doc.part.package, self._thread_parts)
        buf = BytesIO()
        self._doc.save(buf)
        return _restore_thread_sidecars(buf.getvalue(), self._thread_parts)

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _require_supported_revisions(self) -> None:
        reason = _unsupported_revision_reason(self._doc)
        if reason is not None:
            raise UnsupportedRevisionError(
                f"unsupported revision form {reason!r}; refusing to write a partial artifact"
            )

    def _collect_spans(self) -> list[AddressableSpan]:
        spans: list[AddressableSpan] = []
        for ref in self._refs:
            spans.extend(_paragraph_spans(ref.paragraph, ref.container_id, ref.paragraph_index))
        return spans

    def _refresh_after_edit(self, index: int) -> None:
        ref = self._refs[index]
        new_text = _paragraph_visible_text(ref.paragraph)
        old = self._segments[index]
        self._segments[index] = replace(old, text=new_text)
        self._spans = self._collect_spans()
        self._comments = _collect_comments(
            _existing_comments_part(self._doc),
            self._paragraphs_by_container,
            self._doc.part.package,
        )

    def _find_index(
        self,
        rep: SegmentReplacement,
        by_container: dict[str, int],
        by_id: dict[str, int],
        strict: bool,
    ) -> int | None:
        idx, _start, _end = self._resolve_replacement(rep, by_container, by_id, strict)
        return idx

    def _resolve_replacement(
        self,
        rep: SegmentReplacement,
        by_container: dict[str, int],
        by_id: dict[str, int],
        strict: bool,
    ) -> tuple[int | None, int | None, int | None]:
        if rep.span_id:
            span = next((s for s in self._spans if s.span_id == rep.span_id), None)
            if span is None:
                if strict:
                    raise ValueError(f"unknown replacement target: {rep.span_id}")
                return None, None, None
            idx = by_container.get(span.container_id)
            if idx is None:
                if strict:
                    raise ValueError(f"unknown replacement target: {rep.span_id}")
                return None, None, None
            if rep.start_offset is None and rep.end_offset is None:
                return idx, span.start_offset, span.end_offset
            local_start = 0 if rep.start_offset is None else rep.start_offset
            local_end = len(span.text) if rep.end_offset is None else rep.end_offset
            if not 0 <= local_start < local_end <= len(span.text):
                raise ValueError(
                    f"invalid replacement offsets for span {span.span_id}: "
                    f"expected 0 <= start < end <= {len(span.text)}, "
                    f"got {local_start}:{local_end}"
                )
            return idx, span.start_offset + local_start, span.start_offset + local_end
        if rep.container_id:
            idx = by_container.get(str(rep.container_id))
        elif rep.id:
            idx = by_id.get(str(rep.id))
        else:
            idx = None
        if idx is None and strict:
            raise ValueError(f"unknown replacement target: {rep.container_id or rep.id}")
        return idx, rep.start_offset, rep.end_offset

    def _replace_full_segment(self, index: int, text: str) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        full = _paragraph_visible_text(para)
        if full:
            _replace_plain_range(para._p, 0, len(full), text)
        elif para.runs:
            para.runs[0].text = text
        else:
            para.add_run(text)
        self._refresh_after_edit(index)

    def _apply_to_paragraph(
        self,
        index: int,
        replacement: str,
        start: int | None,
        end: int | None,
    ) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        full = _paragraph_visible_text(para)
        if start is None and end is None:
            self._replace_full_segment(index, replacement)
            return

        s = 0 if start is None else start
        e = len(full) if end is None else end
        if not 0 <= s < e <= len(full):
            raise ValueError(
                f"invalid replacement offsets for segment {ref.container_id}: "
                f"expected 0 <= start < end <= {len(full)}, got {s}:{e}"
            )

        if s == 0 and e == len(full):
            self._replace_full_segment(index, replacement)
            return

        _replace_plain_range(para._p, s, e, replacement)
        self._refresh_after_edit(index)

    def _build_run_ranges(self, paragraph: Paragraph) -> list[tuple[int, int, int]]:
        out: list[tuple[int, int, int]] = []
        cur = 0
        for i, run in enumerate(paragraph.runs):
            ln = len(run.text)
            out.append((i, cur, cur + ln))
            cur += ln
        return out

    def _split_run(self, paragraph: Paragraph, run_index: int, offset: int) -> int:
        if offset <= 0:
            return run_index
        run = paragraph.runs[run_index]
        if offset >= len(run.text):
            return run_index + 1
        left = run.text[:offset]
        right = run.text[offset:]
        run.text = left

        # clone the underlying XML element
        cloned = copy.deepcopy(run._element)
        run._element.addnext(cloned)

        new_run = paragraph.runs[run_index + 1]
        new_run.text = right
        return run_index + 1


# ----------------------------------------------------------------------
# Helper to expose for advanced users if needed
# ----------------------------------------------------------------------


def _paragraph_text(p: Paragraph) -> str:
    return "".join(r.text for r in p.runs)
