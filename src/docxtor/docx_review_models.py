from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_models import AddressableComment
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
class ReviewMarkupInventory:
    revisions: tuple[Revision, ...]
    comments: tuple[AddressableComment, ...]
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...]

    @property
    def has_revisions(self) -> bool:
        return bool(self.revisions)

    @property
    def has_comments(self) -> bool:
        return bool(self.comments)


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
