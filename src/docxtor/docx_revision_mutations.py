from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
from typing import Any

from docx import Document as PyDocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .common import DocumentError
from .docx_inline import (
    _descendant_visible_text,
    _inline_width,
    _visible_text,
    paragraph_to_inline_segments,
)
from .docx_ns import _TEXT_NODE_TAGS
from .docx_review_models import OperationReceipt, OperationStatus
from .docx_revisions import RevisionInventory, inventory_revisions_bytes
from .docx_stories import index_stories
from .docx_xml import _is_text_box_container

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_TRANSPARENT_WRAPPERS = {qn("w:hyperlink"), qn("w:ins")}
_MARKUP_TAGS = {
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
}
_COMMENT_REFERENCE = qn("w:commentReference")


class RevisionMutationError(DocumentError):
    """A neutral revision could not be represented without guessing."""


@dataclass(frozen=True)
class RevisionAuthor:
    author: str
    date: str | None = None


@dataclass(frozen=True)
class RevisionRange:
    locator: str
    start_offset: int
    end_offset: int
    expected_text: str | None = None


@dataclass(frozen=True)
class RevisionPosition:
    locator: str
    offset: int


@dataclass(frozen=True)
class RevisionMutationResult:
    data: bytes
    receipt: OperationReceipt
    before: RevisionInventory
    after: RevisionInventory


@dataclass(frozen=True)
class _LivePiece:
    kind: str
    text: str
    run: Any | None
    markup: bool


def insert_revision(
    data: bytes,
    position: RevisionPosition,
    text: str,
    reviewer: RevisionAuthor,
) -> RevisionMutationResult:
    if not text:
        raise RevisionMutationError("inserted revision text must not be empty")
    document, paragraph = _paragraph(data, position.locator)
    _explode_runs(paragraph._p)
    visible = _aligned_visible_text(paragraph, position.locator)
    if not 0 <= position.offset <= len(visible):
        raise RevisionMutationError(f"invalid insertion offset for {position.locator}")
    _split_at_visible(paragraph, position.offset)
    revision_id = _next_revision_id(data)
    wrapper = _revision_wrapper("ins", revision_id, reviewer)
    run = OxmlElement("w:r")
    template = _rpr_near_offset(paragraph, position.offset)
    if template is not None:
        run.append(deepcopy(template))
    node = OxmlElement("w:t")
    node.text = text
    _apply_xml_space(node, text)
    run.append(node)
    wrapper.append(run)
    reference = _insertion_reference(paragraph, position.offset)
    if reference is None:
        paragraph._p.append(wrapper)
    else:
        reference.addprevious(wrapper)
    wrapped = "".join(_run_plain_text(child) for child in wrapper if child.tag == qn("w:r"))
    if wrapped != text:
        raise RevisionMutationError(f"revision range cannot be mapped at {position.locator}")
    return _result(data, document, "insert_revision", position.locator, revision_id)


def delete_revision(
    data: bytes,
    target: RevisionRange,
    reviewer: RevisionAuthor,
) -> RevisionMutationResult:
    document, paragraph = _paragraph(data, target.locator)
    _explode_runs(paragraph._p)
    visible = _aligned_visible_text(paragraph, target.locator)
    if not 0 <= target.start_offset < target.end_offset <= len(visible):
        raise RevisionMutationError(f"invalid deletion range for {target.locator}")
    selected = visible[target.start_offset : target.end_offset]
    if target.expected_text is not None and selected != target.expected_text:
        raise RevisionMutationError(f"revision range text changed at {target.locator}")
    _split_at_visible(paragraph, target.end_offset)
    _split_at_visible(paragraph, target.start_offset)
    selected_runs = _runs_in_visible_range(
        paragraph, target.start_offset, target.end_offset, target.locator
    )
    wrapped_text = "".join(_run_plain_text(run) for run in selected_runs)
    if wrapped_text != selected:
        raise RevisionMutationError(f"revision range cannot be mapped at {target.locator}")
    revision_id = _next_revision_id(data)
    wrapper = _revision_wrapper("del", revision_id, reviewer)
    selected_runs[0].addprevious(wrapper)
    for run_element in selected_runs:
        for text_node in run_element.iter(qn("w:t")):
            text_node.tag = qn("w:delText")
        parent = run_element.getparent()
        if parent is None:
            raise RevisionMutationError(f"revision range cannot be mapped at {target.locator}")
        parent.remove(run_element)
        wrapper.append(run_element)
    confirmed = "".join(_run_plain_text(run) for run in wrapper.iterchildren(qn("w:r")))
    if confirmed != selected:
        raise RevisionMutationError(f"revision range cannot be mapped at {target.locator}")
    return _result(data, document, "delete_revision", target.locator, revision_id)


def replace_revision(
    data: bytes,
    target: RevisionRange,
    replacement: str,
    reviewer: RevisionAuthor,
) -> tuple[RevisionMutationResult, RevisionMutationResult]:
    deleted = delete_revision(data, target, reviewer)
    inserted = insert_revision(
        deleted.data,
        RevisionPosition(target.locator, target.start_offset),
        replacement,
        reviewer,
    )
    return deleted, inserted


def _paragraph(data: bytes, locator: str) -> tuple[Any, Any]:
    document = PyDocxDocument(BytesIO(data))
    paragraph = index_stories(document).paragraphs_by_container.get(locator)
    if paragraph is None:
        raise RevisionMutationError(f"unknown revision locator: {locator}")
    return document, paragraph


def _revision_wrapper(kind: str, revision_id: int, reviewer: RevisionAuthor) -> Any:
    element = OxmlElement(f"w:{kind}")
    element.set(qn("w:id"), str(revision_id))
    element.set(qn("w:author"), reviewer.author)
    if reviewer.date is not None:
        element.set(qn("w:date"), reviewer.date)
    return element


def _next_revision_id(data: bytes) -> int:
    ids = [
        int(revision.revision_id)
        for revision in inventory_revisions_bytes(data).revisions
        if revision.revision_id is not None and revision.revision_id.isdigit()
    ]
    return max(ids, default=-1) + 1


def _aligned_visible_text(paragraph: Any, locator: str) -> str:
    visible = _visible_text(paragraph_to_inline_segments(paragraph))
    from_pieces = "".join(piece.text for piece in _live_pieces(paragraph))
    if visible != from_pieces:
        raise RevisionMutationError(f"revision range cannot be mapped at {locator}")
    return visible


def _live_pieces(paragraph: Any) -> list[_LivePiece]:
    pieces: list[_LivePiece] = []

    def from_run(run: Any) -> None:
        markup_run = _run_is_markup(run)
        for child in run:
            if child.tag == qn("w:rPr"):
                continue
            if child.tag in _TEXT_NODE_TAGS:
                if child.text:
                    pieces.append(_LivePiece("text", child.text, run, False))
                continue
            width = (
                ""
                if _is_text_box_container(child.tag)
                or any(_is_text_box_container(node.tag) for node in child.iter())
                else _inline_width(child)
            )
            pieces.append(
                _LivePiece(
                    "opaque",
                    width,
                    run,
                    markup_run or child.tag == _COMMENT_REFERENCE,
                )
            )

    def walk(parent: Any) -> None:
        for child in parent:
            if child.tag == qn("w:pPr"):
                continue
            if child.tag == qn("w:r"):
                from_run(child)
                continue
            if child.tag in _TRANSPARENT_WRAPPERS:
                walk(child)
                continue
            pieces.append(
                _LivePiece(
                    "opaque",
                    _descendant_visible_text(child),
                    None,
                    child.tag in _MARKUP_TAGS,
                )
            )

    walk(paragraph._p)
    return pieces


def _run_is_markup(run: Any) -> bool:
    children = [child for child in run if child.tag != qn("w:rPr")]
    return bool(children) and all(
        child.tag in _MARKUP_TAGS or child.tag == _COMMENT_REFERENCE for child in children
    )


def _explode_runs(parent: Any) -> None:
    for child in list(parent):
        if child.tag in _TRANSPARENT_WRAPPERS:
            _explode_runs(child)
            continue
        if child.tag != qn("w:r"):
            continue
        replacements = _exploded_runs(child)
        if replacements is None:
            continue
        for run in replacements:
            child.addprevious(run)
        parent.remove(child)


def _exploded_runs(run: Any) -> list[Any] | None:
    content = [child for child in run if child.tag != qn("w:rPr")]
    if len(content) <= 1:
        return None
    rpr = run.find(qn("w:rPr"))
    exploded: list[Any] = []
    for child in content:
        new_run = OxmlElement("w:r")
        if rpr is not None:
            new_run.append(deepcopy(rpr))
        new_run.append(deepcopy(child))
        exploded.append(new_run)
    return exploded


def _split_at_visible(paragraph: Any, offset: int) -> None:
    cursor = 0
    for piece in _live_pieces(paragraph):
        width = len(piece.text)
        if (
            piece.kind == "text"
            and piece.run is not None
            and cursor < offset < cursor + width
        ):
            _split_text_run(piece.run, offset - cursor)
            return
        cursor += width


def _split_text_run(run: Any, local_offset: int) -> None:
    texts = [child for child in run if child.tag in _TEXT_NODE_TAGS]
    if len(texts) != 1:
        return
    node = texts[0]
    text = node.text or ""
    if not 0 < local_offset < len(text):
        return
    left, right = text[:local_offset], text[local_offset:]
    node.text = left
    _apply_xml_space(node, left)
    following = deepcopy(run)
    for child in following:
        if child.tag in _TEXT_NODE_TAGS:
            child.text = right
            _apply_xml_space(child, right)
    run.addnext(following)


def _runs_in_visible_range(paragraph: Any, start: int, end: int, locator: str) -> list[Any]:
    runs: list[Any] = []
    cursor = 0
    for piece in _live_pieces(paragraph):
        piece_start = cursor
        piece_end = cursor + len(piece.text)
        cursor = piece_end
        if piece.markup:
            continue
        zero_width_inside = piece_start == piece_end and start <= piece_start < end
        overlaps = piece_start < end and piece_end > start
        if piece.kind != "text":
            if overlaps or zero_width_inside:
                raise RevisionMutationError(f"revision range crosses opaque content at {locator}")
            continue
        if (
            piece.run is not None
            and piece_start >= start
            and piece_end <= end
            and piece_start < piece_end
        ):
            runs.append(piece.run)
    if not runs:
        raise RevisionMutationError(f"revision range cannot be mapped at {locator}")
    parents = {run.getparent() for run in runs}
    if len(parents) != 1 or None in parents:
        raise RevisionMutationError(f"revision range cannot be mapped at {locator}")
    return runs


def _insertion_reference(paragraph: Any, offset: int) -> Any | None:
    cursor = 0
    for piece in _live_pieces(paragraph):
        if cursor >= offset and piece.kind == "text" and piece.run is not None:
            return piece.run
        cursor += len(piece.text)
    return None


def _rpr_near_offset(paragraph: Any, offset: int) -> Any | None:
    reference = _insertion_reference(paragraph, offset)
    if reference is None:
        last = None
        for piece in _live_pieces(paragraph):
            if piece.kind == "text" and piece.run is not None:
                last = piece.run
        reference = last
    if reference is None:
        return None
    return deepcopy(reference.find(qn("w:rPr")))


def _run_plain_text(run: Any) -> str:
    return "".join((child.text or "") for child in run if child.tag in _TEXT_NODE_TAGS)


def _apply_xml_space(node: Any, text: str) -> None:
    if text[:1].isspace() or text[-1:].isspace():
        node.set(_XML_SPACE, "preserve")
    elif _XML_SPACE in node.attrib:
        del node.attrib[_XML_SPACE]


def _serialize(document: Any) -> bytes:
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _result(
    source: bytes,
    document: Any,
    operation: str,
    locator: str,
    revision_id: int,
) -> RevisionMutationResult:
    before = inventory_revisions_bytes(source)
    payload = _serialize(document)
    after = inventory_revisions_bytes(payload)
    created = [item for item in after.revisions if item.revision_id == str(revision_id)]
    if not created:
        raise RevisionMutationError("revision creation was not confirmed after round-trip")
    return RevisionMutationResult(
        data=payload,
        receipt=OperationReceipt(
            operation=operation,
            status=OperationStatus.APPLIED,
            affected_parts=tuple(sorted({item.part_name for item in created})),
            created_ids=(str(revision_id),),
            locator=locator,
            before_sha256=sha256(source).hexdigest(),
            after_sha256=sha256(payload).hexdigest(),
        ),
        before=before,
        after=after,
    )
