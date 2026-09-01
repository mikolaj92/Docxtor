"""Physical, policy-blind rendering of review revisions and comments."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .docx_comment_mutations import CommentAuthor, CommentRange, add_comment
from .docx_package import normalize_docx_timestamps, restore_semantically_unchanged_xml_parts
from .docx_publish import publish_docx
from .docx_revision_mutations import (
    RevisionAuthor,
    RevisionPosition,
    RevisionRange,
    delete_revision,
    insert_revision,
    replace_revision,
)


@dataclass(frozen=True)
class PhysicalRevisionEdit:
    locator: str
    kind: str
    start_offset: int
    end_offset: int
    text: str = ""
    expected_text: str | None = None


@dataclass(frozen=True)
class PhysicalCommentEdit:
    locator: str
    start_offset: int
    end_offset: int
    text: str


@dataclass(frozen=True)
class PhysicalReviewPlan:
    revisions: tuple[PhysicalRevisionEdit, ...] = ()
    comments: tuple[PhysicalCommentEdit, ...] = ()


@dataclass(frozen=True)
class PhysicalReviewer:
    author: str = "Reviewer"
    initials: str = "RV"
    date: str | None = None


class PhysicalReviewRenderError(ValueError):
    pass


def render_physical_review(
    source: str | Path | bytes,
    output: str | Path,
    plan: PhysicalReviewPlan,
    *,
    reviewer: PhysicalReviewer = PhysicalReviewer(),
) -> Path:
    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    rev_author = RevisionAuthor(reviewer.author, reviewer.date)
    try:
        for edit in plan.revisions:
            if edit.kind == "insert":
                data = insert_revision(
                    data, RevisionPosition(edit.locator, edit.start_offset), edit.text, rev_author
                ).data
            elif edit.kind == "delete":
                data = delete_revision(
                    data,
                    RevisionRange(
                        edit.locator, edit.start_offset, edit.end_offset, edit.expected_text
                    ),
                    rev_author,
                ).data
            elif edit.kind == "replace":
                _deleted, inserted = replace_revision(
                    data,
                    RevisionRange(
                        edit.locator, edit.start_offset, edit.end_offset, edit.expected_text
                    ),
                    edit.text,
                    rev_author,
                )
                data = inserted.data
            else:
                raise PhysicalReviewRenderError(f"unsupported physical revision kind: {edit.kind}")
        for comment in plan.comments:
            data = add_comment(
                data,
                CommentRange(comment.locator, comment.start_offset, comment.end_offset),
                comment.text,
                CommentAuthor(reviewer.author, reviewer.initials, reviewer.date),
            ).data
    except Exception as exc:
        raise PhysicalReviewRenderError(str(exc)) from exc
    target = publish_docx(data, output).path
    if not isinstance(source, bytes):
        restore_semantically_unchanged_xml_parts(source, target)
    normalize_docx_timestamps(target)
    return target
