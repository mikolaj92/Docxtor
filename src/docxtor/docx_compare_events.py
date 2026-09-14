"""Comment, revision-lifecycle, and formatting change composition."""

from __future__ import annotations

from dataclasses import replace
from difflib import SequenceMatcher
from hashlib import sha256
from math import inf

from .docx_compare_models import (
    BlockPair,
    CommentChange,
    CommentChangeStatus,
    ComparisonDiagnostic,
    DocumentBlock,
    DocumentComment,
    DocumentView,
    FormattingChange,
    RevisionEvent,
    RevisionEventStatus,
    RevisionReference,
    RevisionResolutionEvidence,
)

_MAX_FORMAT_ALIGNMENT_PRODUCT = 100_000


def compare_comments(
    left: DocumentView, right: DocumentView, comparison_id: str
) -> tuple[CommentChange, ...]:
    before = {comment.comment_id: comment for comment in left.comments}
    after = {comment.comment_id: comment for comment in right.comments}
    changes = []
    for comment_id in sorted(set(before) | set(after)):
        left_comment = before.get(comment_id)
        right_comment = after.get(comment_id)
        if left_comment is None:
            status = CommentChangeStatus.ADDED
        elif right_comment is None:
            status = CommentChangeStatus.REMOVED
        elif _comment_signature(left_comment) == _comment_signature(right_comment):
            continue
        else:
            status = CommentChangeStatus.CHANGED
        changes.append(
            CommentChange(
                _stable_id(
                    "comment-change",
                    comparison_id,
                    comment_id,
                    status.value,
                    left_comment.text if left_comment else "",
                    right_comment.text if right_comment else "",
                ),
                status,
                left_comment,
                right_comment,
            )
        )
    return tuple(changes)


def compare_revisions(
    left: tuple[RevisionReference, ...],
    right: tuple[RevisionReference, ...],
    pairs: tuple[BlockPair, ...],
    left_view: DocumentView,
    right_view: DocumentView,
    comparison_id: str,
) -> tuple[RevisionEvent, ...]:
    left = _assign_revision_pairs(left, left_view)
    right = _assign_revision_pairs(right, right_view)
    before = _index_revisions(left)
    after = _index_revisions(right)
    left_blocks = {block.id: block for block in left_view.blocks}
    right_blocks = {block.id: block for block in right_view.blocks}
    pair_by_id = {pair.pair_id: pair for pair in pairs}
    events = []
    for key in sorted(set(before) | set(after)):
        left_revision = before.get(key)
        right_revision = after.get(key)
        if left_revision is None:
            status = RevisionEventStatus.INTRODUCED
            continuity = None
        elif right_revision is None:
            status = RevisionEventStatus.RESOLVED
            continuity = _resolution_evidence(left_revision, left_blocks, right_blocks, pair_by_id)
        elif _revision_signature(left_revision) == _revision_signature(right_revision):
            continue
        else:
            status = RevisionEventStatus.CHANGED
            continuity = None
        events.append(
            RevisionEvent(
                _stable_id(
                    "revision-event",
                    comparison_id,
                    repr(key),
                    status.value,
                    left_revision.locator if left_revision else "",
                    right_revision.locator if right_revision else "",
                ),
                status,
                left_revision,
                right_revision,
                continuity,
            )
        )
    return tuple(events)


def compare_formatting(
    pairs: tuple[BlockPair, ...],
    left: DocumentView,
    right: DocumentView,
    left_style_parts: tuple[tuple[str, str], ...],
    right_style_parts: tuple[tuple[str, str], ...],
    comparison_id: str,
) -> tuple[tuple[FormattingChange, ...], tuple[ComparisonDiagnostic, ...]]:
    left_by_id = {block.id: block for block in left.blocks}
    right_by_id = {block.id: block for block in right.blocks}
    changes: list[FormattingChange] = []
    diagnostics: list[ComparisonDiagnostic] = []
    for pair in pairs:
        before = left_by_id.get(pair.left_block_id or "")
        after = right_by_id.get(pair.right_block_id or "")
        if before is None or after is None:
            continue
        if before.paragraph_formatting_sha256 != after.paragraph_formatting_sha256:
            changes.append(
                FormattingChange(
                    _stable_id("format-change", comparison_id, pair.pair_id, "paragraph"),
                    "paragraph_properties",
                    pair.pair_id,
                    before.id,
                    after.id,
                    before.paragraph_formatting_sha256,
                    after.paragraph_formatting_sha256,
                )
            )
        before_run_styles = _run_style_sequence(before)
        after_run_styles = _run_style_sequence(after)
        if before.text == after.text and before.formatting_spans != after.formatting_spans:
            changes.append(
                FormattingChange(
                    _stable_id("format-change", comparison_id, pair.pair_id, "run"),
                    "run_properties",
                    pair.pair_id,
                    before.id,
                    after.id,
                    before.formatting_sha256,
                    after.formatting_sha256,
                )
            )
        elif before.text != after.text:
            shared_text_format_changed = _shared_text_formatting_mismatch(before, after)
            range_layout_uncertain = (
                shared_text_format_changed is None
                and before.formatting_spans != after.formatting_spans
                and max(len(before_run_styles), len(after_run_styles)) > 1
            )
            if (
                before_run_styles != after_run_styles
                or shared_text_format_changed is True
                or range_layout_uncertain
            ):
                diagnostics.append(
                    ComparisonDiagnostic(
                        "run_formatting_not_isolated_with_text_change",
                        "run formatting changed or could not be isolated from accepted text edits",
                        "both",
                        locator=pair.pair_id,
                    )
                )

    before_styles = dict(left_style_parts)
    after_styles = dict(right_style_parts)
    for name in sorted(set(before_styles) | set(after_styles)):
        before_hash = before_styles.get(name)
        after_hash = after_styles.get(name)
        if before_hash == after_hash:
            continue
        changes.append(
            FormattingChange(
                _stable_id("format-change", comparison_id, "style-part", name),
                f"style_part:{name}",
                None,
                None,
                None,
                before_hash,
                after_hash,
            )
        )
    return tuple(changes), tuple(diagnostics)


def _comment_signature(comment: DocumentComment) -> tuple[object, ...]:
    anchor = comment.anchor
    return (
        comment.text,
        comment.author,
        comment.initials,
        comment.date,
        comment.parent_id,
        (anchor.pair_id or anchor.locator) if anchor else None,
        anchor.text if anchor else None,
        anchor.start_offset if anchor else None,
        anchor.end_offset if anchor else None,
    )


def _index_revisions(
    revisions: tuple[RevisionReference, ...],
) -> dict[tuple[str, ...], RevisionReference]:
    grouped: dict[tuple[str, ...], list[RevisionReference]] = {}
    for reference in revisions:
        anchor = reference.pair_id or reference.container_id or reference.locator
        group = (
            reference.part_name,
            reference.revision_id or "",
            reference.kind,
            anchor,
        )
        grouped.setdefault(group, []).append(reference)

    indexed = {}
    for group, members in grouped.items():
        members.sort(
            key=lambda item: (
                item.start_offset if item.start_offset is not None else inf,
                item.end_offset if item.end_offset is not None else inf,
                item.text,
                item.author or "",
                item.date or "",
                item.locator,
            )
        )
        for occurrence, reference in enumerate(members):
            indexed[(*group, str(occurrence))] = reference
    return indexed


def _revision_signature(reference: RevisionReference) -> tuple[object, ...]:
    return (
        reference.kind,
        reference.author,
        reference.date,
        reference.text,
        reference.start_offset,
        reference.end_offset,
    )


def _assign_revision_pairs(
    revisions: tuple[RevisionReference, ...], view: DocumentView
) -> tuple[RevisionReference, ...]:
    pair_by_block = {block.id: block.pair_id for block in view.blocks}
    return tuple(
        replace(item, pair_id=pair_by_block.get(item.block_id or "")) for item in revisions
    )


def _resolution_evidence(
    revision: RevisionReference,
    left_blocks: dict[str, DocumentBlock],
    right_blocks: dict[str, DocumentBlock],
    pairs: dict[str, BlockPair],
) -> RevisionResolutionEvidence:
    pair = pairs.get(revision.pair_id or "")
    if pair is None or pair.left_block_id is None or pair.right_block_id is None:
        return RevisionResolutionEvidence.UNVERIFIED_REMOVAL
    left_block = left_blocks.get(pair.left_block_id)
    right_block = right_blocks.get(pair.right_block_id)
    if left_block is not None and right_block is not None and left_block.text == right_block.text:
        return RevisionResolutionEvidence.ACCEPTED_TEXT_EQUAL
    return RevisionResolutionEvidence.UNVERIFIED_REMOVAL


def _run_style_sequence(block: DocumentBlock) -> tuple[str, ...]:
    return tuple(span.properties_sha256 for span in block.formatting_spans)


def _shared_text_formatting_mismatch(before: DocumentBlock, after: DocumentBlock) -> bool | None:
    if len(before.text) * len(after.text) > _MAX_FORMAT_ALIGNMENT_PRODUCT:
        return None
    before_styles = _style_by_character(before)
    after_styles = _style_by_character(after)
    matcher = SequenceMatcher(None, before.text, after.text, autojunk=False)
    has_shared_characters = False
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            continue
        has_shared_characters = True
        if before_styles[i1:i2] != after_styles[j1:j2]:
            return True
    return False if has_shared_characters else None


def _style_by_character(block: DocumentBlock) -> list[str | None]:
    result: list[str | None] = [None] * len(block.text)
    for span in block.formatting_spans:
        result[span.start_offset : span.end_offset] = [span.properties_sha256] * (
            span.end_offset - span.start_offset
        )
    return result


def _stable_id(prefix: str, *parts: str) -> str:
    return f"{prefix}-{sha256(chr(31).join(parts).encode('utf-8')).hexdigest()[:20]}"
