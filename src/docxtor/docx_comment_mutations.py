from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
from typing import Any

from docx import Document as PyDocxDocument
from docx.oxml.ns import qn

from .common import DocumentError
from .docx_comments import (
    _ensure_thread_parts,
    _existing_comments_part,
    _restore_thread_sidecars,
)
from .docx_models import AddressableComment
from .docx_review_models import OperationReceipt, OperationStatus
from .docx_stories import index_stories
from .docx_units import _paragraph_visible_text, _replace_plain_range


class CommentMutationError(DocumentError):
    """A comment mutation could not be proven complete."""


@dataclass(frozen=True)
class CommentRange:
    locator: str
    start_offset: int
    end_offset: int
    expected_text: str | None = None


@dataclass(frozen=True)
class CommentAuthor:
    author: str
    initials: str | None = None
    date: str | None = None


@dataclass(frozen=True)
class CommentMutationResult:
    data: bytes
    receipt: OperationReceipt
    comments: tuple[AddressableComment, ...]


def add_comment(
    data: bytes,
    target: CommentRange,
    text: str,
    author: CommentAuthor,
) -> CommentMutationResult:
    from .docx_review_geometry import (
        PhysicalCommentRange,
        PhysicalCommentSpan,
        project_docx_review_geometry,
    )
    from .docx_review_geometry_comments import add_physical_comment_bytes

    geometry = project_docx_review_geometry(data)
    paragraph = next((item for item in geometry.paragraphs if item.locator == target.locator), None)
    if paragraph is None:
        raise CommentMutationError(f"unknown comment locator: {target.locator}")
    if not 0 <= target.start_offset < target.end_offset <= len(paragraph.text):
        raise CommentMutationError(f"invalid comment range for {target.locator}")
    selected = paragraph.text[target.start_offset : target.end_offset]
    if target.expected_text is not None and selected != target.expected_text:
        raise CommentMutationError(f"comment range text changed at {target.locator}")
    return add_physical_comment_bytes(
        data,
        PhysicalCommentRange(
            geometry.document_sha256,
            (
                PhysicalCommentSpan(
                    target.locator, target.start_offset, target.end_offset, selected
                ),
            ),
        ),
        text,
        author,
    )


def update_comment(
    data: bytes,
    comment_id: str,
    text: str,
    *,
    expected_text: str | None = None,
    author: CommentAuthor | None = None,
) -> CommentMutationResult:
    from .docx_review_inventory import inventory_review_markup
    from .docx_review_models import ReviewCoverage

    before = inventory_review_markup(data)
    if before.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment update source coverage is incomplete")
    selected = next((item for item in before.comments if item.comment_id == comment_id), None)
    if selected is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    ids = tuple(
        item.comment_id
        for item in before.comments
        if item.comment_id == comment_id
        or (selected.region_set_id is not None and item.region_set_id == selected.region_set_id)
    )
    candidate = data
    affected: set[str] = set()
    for identity in ids:
        result = _update_comment(
            candidate, identity, text, expected_text=expected_text, author=author
        )
        candidate = result.data
        affected.update(result.receipt.affected_parts)
    after = inventory_review_markup(candidate)
    if after.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment update result coverage is incomplete")
    return CommentMutationResult(
        candidate,
        OperationReceipt(
            operation="update_comment",
            status=OperationStatus.APPLIED,
            affected_parts=tuple(sorted(affected)),
            locator=selected.locator,
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(candidate).hexdigest(),
        ),
        after.comments,
    )


def _update_comment(
    data: bytes,
    comment_id: str,
    text: str,
    *,
    expected_text: str | None = None,
    author: CommentAuthor | None = None,
) -> CommentMutationResult:
    if not text:
        raise CommentMutationError("comment text must not be empty")
    document = PyDocxDocument(BytesIO(data))
    stories = index_stories(document)
    existing = next(
        (comment for comment in stories.comments if comment.comment_id == comment_id),
        None,
    )
    if existing is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    canonical = existing.user_note_text if existing.user_note_text is not None else existing.text
    if expected_text is not None and canonical != expected_text:
        raise CommentMutationError(f"comment text changed at {comment_id}")
    from .docx_comment_scopes import scope_link

    comments_part = _existing_comments_part(document)
    if comments_part is None:
        raise CommentMutationError("comment body part is unavailable for update")
    comment_element = next(
        node
        for node in comments_part.element.iter(qn("w:comment"))
        if node.get(qn("w:id")) == comment_id
    )
    bridge = scope_link(comment_element)
    extra = [
        segment
        for segment in stories.segments
        if (segment.container_id or "").startswith(f"comment:{comment_id}:")
        and segment.container_id != existing.container_id
        and not (
            bridge is not None
            and segment.container_id in stories.paragraphs_by_container
            and stories.paragraphs_by_container[segment.container_id]._p is bridge.paragraph
        )
    ]
    if extra:
        raise CommentMutationError(f"comment {comment_id} spans multiple paragraphs")
    paragraph = stories.paragraphs_by_container.get(existing.container_id)
    if paragraph is None:
        raise CommentMutationError(f"unknown comment locator: {existing.container_id}")
    visible = _paragraph_visible_text(paragraph)
    if visible:
        _replace_plain_range(paragraph._p, 0, len(visible), text)
    elif paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)
    if author is not None:
        _apply_comment_author(document, comment_id, author)
    payload = _serialize(document, thread_parts=stories.thread_parts)
    after = index_stories(PyDocxDocument(BytesIO(payload)))
    updated = next((item for item in after.comments if item.comment_id == comment_id), None)
    if (
        updated is None
        or updated.text != text + ("\n" + bridge.label if bridge is not None else "")
        or (bridge is not None and updated.user_note_text != text)
        or updated.locator != existing.locator
        or updated.anchor_text != existing.anchor_text
        or updated.parent_id != existing.parent_id
    ):
        raise CommentMutationError("comment update was not confirmed after round-trip")
    if author is None and (
        updated.author != existing.author
        or updated.initials != existing.initials
        or updated.date != existing.date
    ):
        raise CommentMutationError("comment update was not confirmed after round-trip")
    return CommentMutationResult(
        data=payload,
        receipt=OperationReceipt(
            operation="update_comment",
            status=OperationStatus.APPLIED,
            affected_parts=("word/comments.xml",),
            created_ids=(),
            locator=existing.locator,
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(payload).hexdigest(),
        ),
        comments=tuple(after.comments),
    )


def remove_comments(data: bytes, comment_ids: set[str] | None = None) -> CommentMutationResult:
    """Remove complete native region sets and descendants in one transaction."""
    from .docx_comment_removal import remove_comment_sets

    return remove_comment_sets(data, comment_ids)


def _apply_comment_author(document: Any, comment_id: str, author: CommentAuthor) -> None:
    comments_part = _existing_comments_part(document)
    if comments_part is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    for comment in comments_part.element.findall(qn("w:comment")):
        if comment.get(qn("w:id")) != comment_id:
            continue
        comment.set(qn("w:author"), author.author)
        if author.initials is not None:
            comment.set(qn("w:initials"), author.initials)
        if author.date is not None:
            comment.set(qn("w:date"), author.date)
        return
    raise CommentMutationError(f"unknown comment ID: {comment_id}")


def _serialize(doc: Any, thread_parts: dict[str, tuple[bytes, str]] | None = None) -> bytes:
    if thread_parts:
        _ensure_thread_parts(doc.part.package, thread_parts)
    buffer = BytesIO()
    doc.save(buffer)
    payload = buffer.getvalue()
    if thread_parts:
        return _restore_thread_sidecars(payload, thread_parts)
    return payload
