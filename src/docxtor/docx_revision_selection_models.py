"""Immutable, policy-neutral records for existing revision disposition."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_revisions import RevisionOperationError


class RevisionDisposition(StrEnum):
    ACCEPT = "accept"
    REJECT = "reject"


class RevisionDispositionCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class RevisionSelectionError(RevisionOperationError):
    """An existing revision cannot be selected or resolved without guessing."""


@dataclass(frozen=True)
class RevisionDispositionDiagnostic:
    code: str
    detail: str
    part_name: str | None = None
    locator: str | None = None


@dataclass(frozen=True)
class RevisionDispositionTarget:
    target_id: str
    part_name: str
    namespace: str
    kind: str
    revision_id: str | None
    locator: str
    node_sha256: str
    author: str | None
    date: str | None
    text: str
    paragraph_locator: str | None
    paragraph_text: str | None
    supported: bool
    unsupported_reason: str | None = None
    group_id: str | None = None
    required_target_ids: tuple[str, ...] = ()


@dataclass(frozen=True)
class RevisionDispositionInventory:
    profile_id: str
    input_sha256: str
    inventory_sha256: str
    revisions: tuple[RevisionDispositionTarget, ...]
    coverage: RevisionDispositionCoverage
    diagnostics: tuple[RevisionDispositionDiagnostic, ...]
    scanned_parts: tuple[str, ...]

    @property
    def count(self) -> int:
        return len(self.revisions)


@dataclass(frozen=True)
class RevisionDecision:
    target_id: str
    disposition: RevisionDisposition


@dataclass(frozen=True)
class RevisionDispositionReceipt:
    target: RevisionDispositionTarget
    disposition: RevisionDisposition
    coverage: RevisionDispositionCoverage
    resolved: bool
    retained_text: str


@dataclass(frozen=True)
class RevisionDispositionResult:
    data: bytes
    output_sha256: str
    before: RevisionDispositionInventory
    after: RevisionDispositionInventory
    receipts: tuple[RevisionDispositionReceipt, ...]
    changed_parts: tuple[str, ...]
