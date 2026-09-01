from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_models import AddressableComment, AddressableSpan
from .docx_revisions import Revision


class ReviewCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class ReviewDiagnosticSeverity(StrEnum):
    ERROR = "error"
    WARNING = "warning"


class OperationStatus(StrEnum):
    APPLIED = "applied"
    NOOP = "noop"
    REJECTED = "rejected"


@dataclass(frozen=True)
class ReviewDiagnostic:
    code: str
    message: str
    part_name: str | None = None
    severity: ReviewDiagnosticSeverity = ReviewDiagnosticSeverity.ERROR


@dataclass(frozen=True)
class CommentRevisionAssociation:
    comment_id: str
    revision_kinds: tuple[str, ...]
    part_names: tuple[str, ...]
    locator: str | None = None


@dataclass(frozen=True)
class ReviewMarkupInventory:
    revisions: tuple[Revision, ...]
    comments: tuple[AddressableComment, ...]
    comment_revision_associations: tuple[CommentRevisionAssociation, ...]
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...]

    @property
    def has_revisions(self) -> bool:
        return bool(self.revisions)

    @property
    def has_comments(self) -> bool:
        return bool(self.comments)


@dataclass(frozen=True)
class ReviewParagraph:
    """Domain-blind paragraph projection in canonical DOCX address space."""

    locator: str
    paragraph_index: int | None
    text: str
    part_name: str
    style_name: str | None
    opaque_ranges: tuple[tuple[int, int], ...] = ()


@dataclass(frozen=True)
class ReviewDocumentProjection:
    """Physical review-document facts safe for policy layers to consume."""

    paragraphs: tuple[ReviewParagraph, ...]
    spans: tuple[AddressableSpan, ...]
    comments: tuple[AddressableComment, ...]
    markup: ReviewMarkupInventory

    def paragraph(self, locator: str) -> ReviewParagraph | None:
        return next((item for item in self.paragraphs if item.locator == locator), None)


@dataclass(frozen=True)
class ReviewPurityInspection:
    """Package-level review residue; marker values are supplied by the caller."""

    revision_parts: tuple[str, ...]
    comment_parts: tuple[str, ...]
    comment_count: int
    marker_parts: tuple[str, ...]
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...] = ()

    @property
    def is_pure(self) -> bool:
        return not (
            self.revision_parts or self.comment_parts or self.comment_count or self.marker_parts
        ) and self.coverage is ReviewCoverage.COMPLETE


@dataclass(frozen=True)
class OperationReceipt:
    operation: str
    status: OperationStatus
    affected_parts: tuple[str, ...]
    created_ids: tuple[str, ...] = ()
    locator: str | None = None
    before_sha256: str | None = None
    after_sha256: str | None = None
    diagnostics: tuple[ReviewDiagnostic, ...] = ()


@dataclass(frozen=True)
class ReviewBatchReceipt:
    data: bytes
    operations: tuple[OperationReceipt, ...]
    inventory_before: ReviewMarkupInventory
    inventory_after: ReviewMarkupInventory
