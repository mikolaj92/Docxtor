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
    _index_at_visible_offset,
    _split_visible_offset,
    _visible_text,
    paragraph_to_inline_segments,
    rebuild_paragraph_from_inline,
)
from .docx_review_models import OperationReceipt, OperationStatus
from .docx_revisions import RevisionInventory, inventory_revisions_bytes
from .docx_stories import index_stories


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


def insert_revision(
    data: bytes,
    position: RevisionPosition,
    text: str,
    reviewer: RevisionAuthor,
) -> RevisionMutationResult:
    if not text:
        raise RevisionMutationError("inserted revision text must not be empty")
    document, paragraph = _paragraph(data, position.locator)
    segments = paragraph_to_inline_segments(paragraph)
    visible = _visible_text(segments)
    if not 0 <= position.offset <= len(visible):
        raise RevisionMutationError(f"invalid insertion offset for {position.locator}")
    segments = _split_visible_offset(segments, position.offset)
    insertion_index = _index_at_visible_offset(segments, position.offset)
    rebuild_paragraph_from_inline(paragraph, segments)
    revision_id = _next_revision_id(data)
    wrapper = _revision_wrapper("ins", revision_id, reviewer)
    run = OxmlElement("w:r")
    template = _nearby_run(paragraph, insertion_index)
    if template is not None and template.rPr is not None:
        run.append(deepcopy(template.rPr))
    node = OxmlElement("w:t")
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run.append(node)
    wrapper.append(run)
    reference = _run_at(paragraph, insertion_index)
    if reference is None:
        paragraph._p.append(wrapper)
    else:
        reference.addprevious(wrapper)
    return _result(data, document, "insert_revision", position.locator, revision_id)


def delete_revision(
    data: bytes,
    target: RevisionRange,
    reviewer: RevisionAuthor,
) -> RevisionMutationResult:
    document, paragraph = _paragraph(data, target.locator)
    segments = paragraph_to_inline_segments(paragraph)
    visible = _visible_text(segments)
    if not 0 <= target.start_offset < target.end_offset <= len(visible):
        raise RevisionMutationError(f"invalid deletion range for {target.locator}")
    selected = visible[target.start_offset : target.end_offset]
    if target.expected_text is not None and selected != target.expected_text:
        raise RevisionMutationError(f"revision range text changed at {target.locator}")
    segments = _split_visible_offset(
        _split_visible_offset(segments, target.end_offset), target.start_offset
    )
    start = _index_at_visible_offset(segments, target.start_offset)
    end = _index_at_visible_offset(segments, target.end_offset)
    chosen = segments[start:end]
    if not chosen or any(segment.kind != "text" for segment in chosen):
        raise RevisionMutationError(f"revision range crosses opaque content at {target.locator}")
    rebuild_paragraph_from_inline(paragraph, segments)
    revision_id = _next_revision_id(data)
    wrapper = _revision_wrapper("del", revision_id, reviewer)
    runs = list(paragraph.runs)
    selected_runs = runs[start:end]
    if len(selected_runs) != len(chosen):
        raise RevisionMutationError(f"revision range cannot be mapped at {target.locator}")
    selected_runs[0]._r.addprevious(wrapper)
    for run in selected_runs:
        run_element = run._r
        for text_node in run_element.iter(qn("w:t")):
            text_node.tag = qn("w:delText")
        parent = run_element.getparent()
        parent.remove(run_element)
        wrapper.append(run_element)
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


def mark_paragraph_revision(
    data: bytes,
    locator: str,
    kind: str,
    reviewer: RevisionAuthor,
) -> RevisionMutationResult:
    if kind not in {"ins", "del"}:
        raise RevisionMutationError("paragraph mark kind must be 'ins' or 'del'")
    document, paragraph = _paragraph(data, locator)
    revision_id = _next_revision_id(data)
    ppr = paragraph._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    if any(child.tag in {qn("w:ins"), qn("w:del")} for child in rpr):
        raise RevisionMutationError(f"paragraph mark already has a revision at {locator}")
    rpr.append(_revision_wrapper(kind, revision_id, reviewer))
    return _result(data, document, f"paragraph_mark_{kind}", locator, revision_id)


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


def _run_at(paragraph: Any, index: int) -> Any | None:
    runs = list(paragraph._p.iterchildren(qn("w:r")))
    return runs[index] if index < len(runs) else None


def _nearby_run(paragraph: Any, index: int) -> Any | None:
    runs = list(paragraph.runs)
    if not runs:
        return None
    return runs[min(index, len(runs) - 1)]._r


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
