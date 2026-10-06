"""Neutral data model for DOCX revision inventories and operations."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_package import PackageError

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


def _tag(name: str) -> str:
    return f"{{{_W}}}{name}"


class RevisionKind(StrEnum):
    INSERTION = "ins"
    DELETION = "del"
    MOVE_FROM = "moveFrom"
    MOVE_TO = "moveTo"
    RUN_PROPERTIES = "rPrChange"
    PARAGRAPH_PROPERTIES = "pPrChange"
    SECTION_PROPERTIES = "sectPrChange"
    TABLE_PROPERTIES = "tblPrChange"
    ROW_PROPERTIES = "trPrChange"
    CELL_PROPERTIES = "tcPrChange"
    CELL_INSERTION = "cellIns"
    CELL_DELETION = "cellDel"
    CELL_MERGE = "cellMerge"
    TABLE_GRID = "tblGridChange"
    TABLE_EXCEPTION_PROPERTIES = "tblPrExChange"
    NUMBERING = "numberingChange"
    RANGE_MARKER = "range_marker"
    CONFLICT = "conflict"


class RevisionInventoryCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class RevisionOperation(StrEnum):
    ACCEPT_ALL = "accept_all"
    REJECT_ALL = "reject_all"


@dataclass(frozen=True)
class RevisionCoverageDiagnostic:
    part_name: str
    code: str
    detail: str
    locator: str | None = None


@dataclass(frozen=True)
class Revision:
    kind: RevisionKind
    raw_kind: str
    part_name: str
    revision_id: str | None
    author: str | None
    date: str | None
    locator: str
    paragraph_mark: bool


@dataclass(frozen=True)
class RevisionInventory:
    revisions: tuple[Revision, ...]
    coverage: RevisionInventoryCoverage
    diagnostics: tuple[RevisionCoverageDiagnostic, ...]

    @property
    def count(self) -> int:
        return len(self.revisions)


@dataclass(frozen=True)
class RevisionOperationReceipt:
    operation: RevisionOperation
    output_bytes: bytes
    before: RevisionInventory
    after: RevisionInventory
    comments_dropped: bool


class RevisionOperationError(PackageError):
    """The package cannot be flattened without guessing."""


# Compatibility-specific names make callers' exception handling intention clear.
class AcceptRevisionsError(RevisionOperationError):
    pass


class RejectRevisionsError(RevisionOperationError):
    pass
