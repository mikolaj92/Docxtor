"""Neutral immutable results for read-only DOCX document comparison."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_facts import ContainerCoordinate


class ComparisonCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class DiagnosticSeverity(StrEnum):
    WARNING = "warning"


class TextChangeKind(StrEnum):
    INSERTED = "inserted"
    DELETED = "deleted"
    REPLACED = "replaced"


class CommentChangeStatus(StrEnum):
    ADDED = "added"
    REMOVED = "removed"
    CHANGED = "changed"


class RevisionEventStatus(StrEnum):
    INTRODUCED = "introduced"
    RESOLVED = "resolved"
    CHANGED = "changed"


class RevisionResolutionEvidence(StrEnum):
    ACCEPTED_TEXT_EQUAL = "accepted_text_equal"
    UNVERIFIED_REMOVAL = "unverified_removal"


@dataclass(frozen=True)
class ComparisonDiagnostic:
    code: str
    message: str
    side: str | None = None
    part_name: str | None = None
    locator: str | None = None
    severity: DiagnosticSeverity = DiagnosticSeverity.WARNING


@dataclass(frozen=True)
class DocumentSpan:
    """One physical text span and its accepted-content offset, when present."""

    role: str
    text: str
    start_offset: int
    end_offset: int
    accepted_start_offset: int | None
    accepted_end_offset: int | None
    revision_kind: str | None = None
    revision_id: str | None = None
    revision_author: str | None = None
    revision_date: str | None = None
    hyperlink_anchor: str | None = None
    hyperlink_rel_id: str | None = None


@dataclass(frozen=True)
class FormattingSpan:
    start_offset: int
    end_offset: int
    properties_sha256: str


@dataclass(frozen=True)
class DocumentBlock:
    id: str
    pair_id: str
    order: int
    story_id: str
    coordinate: ContainerCoordinate
    style_id: str | None
    outline_level: int | None
    raw_text: str
    text: str
    spans: tuple[DocumentSpan, ...]
    formatting_sha256: str
    paragraph_formatting_sha256: str
    formatting_spans: tuple[FormattingSpan, ...]


@dataclass(frozen=True)
class CommentAnchor:
    """Comment selection anchored to accepted-content offsets within a block."""

    locator: str | None
    block_id: str | None
    pair_id: str | None
    text: str
    start_offset: int | None = None
    end_offset: int | None = None


@dataclass(frozen=True)
class DocumentComment:
    comment_id: str
    text: str
    author: str
    initials: str | None
    date: str | None
    parent_id: str | None
    anchor: CommentAnchor | None


@dataclass(frozen=True)
class DocumentView:
    sha256: str
    blocks: tuple[DocumentBlock, ...]
    comments: tuple[DocumentComment, ...]
    coverage: ComparisonCoverage
    diagnostics: tuple[ComparisonDiagnostic, ...]


@dataclass(frozen=True)
class BlockPair:
    pair_id: str
    story_id: str
    left_block_id: str | None
    right_block_id: str | None


@dataclass(frozen=True)
class TextAnchor:
    block_id: str
    pair_id: str
    start_offset: int
    end_offset: int
    text: str


@dataclass(frozen=True)
class TextChange:
    change_id: str
    pair_id: str
    kind: TextChangeKind
    left: TextAnchor | None
    right: TextAnchor | None


@dataclass(frozen=True)
class CommentChange:
    change_id: str
    status: CommentChangeStatus
    left: DocumentComment | None
    right: DocumentComment | None


@dataclass(frozen=True)
class RevisionReference:
    part_name: str
    locator: str
    kind: str
    revision_id: str | None
    author: str | None
    date: str | None
    container_id: str | None
    block_id: str | None
    pair_id: str | None
    text: str
    start_offset: int | None = None
    end_offset: int | None = None


@dataclass(frozen=True)
class RevisionEvent:
    change_id: str
    status: RevisionEventStatus
    left: RevisionReference | None
    right: RevisionReference | None
    resolution_evidence: RevisionResolutionEvidence | None = None


@dataclass(frozen=True)
class FormattingChange:
    change_id: str
    scope: str
    pair_id: str | None
    left_block_id: str | None
    right_block_id: str | None
    left_sha256: str | None
    right_sha256: str | None


@dataclass(frozen=True)
class DocxDocumentComparison:
    comparison_id: str
    left: DocumentView
    right: DocumentView
    block_pairs: tuple[BlockPair, ...]
    text_changes: tuple[TextChange, ...]
    comment_changes: tuple[CommentChange, ...]
    revision_events: tuple[RevisionEvent, ...]
    formatting_changes: tuple[FormattingChange, ...]
    coverage: ComparisonCoverage
    diagnostics: tuple[ComparisonDiagnostic, ...]


__all__ = [
    "BlockPair",
    "CommentAnchor",
    "CommentChange",
    "CommentChangeStatus",
    "ComparisonCoverage",
    "ComparisonDiagnostic",
    "DiagnosticSeverity",
    "DocumentBlock",
    "DocumentComment",
    "DocumentSpan",
    "DocumentView",
    "DocxDocumentComparison",
    "FormattingChange",
    "FormattingSpan",
    "RevisionEvent",
    "RevisionEventStatus",
    "RevisionReference",
    "RevisionResolutionEvidence",
    "TextAnchor",
    "TextChange",
    "TextChangeKind",
]
