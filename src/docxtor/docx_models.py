from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal

from .common import DocumentError


@dataclass(frozen=True)
class TextSegment:
    """Stable textual segment inside a DOCX.

    container_id examples:
      body:p:0
      header:0:p:0
      table:0:r:0:c:0:p:0   (for table cells)
      txbx:0:p:0            (floating text box)

    paragraph_index is the global order index (body + tables + headers/footers
    + text boxes), counting ALL paragraphs including empty ones for stable
    addressing.

    run_indices: indices (within the python-docx paragraph.runs) of runs that
    contributed non-empty text at parse time. Useful for domain anchors.
    """

    id: str
    text: str
    part: str
    index: int
    container_id: str | None = None
    paragraph_index: int | None = None
    run_indices: list[int] | None = None


@dataclass(frozen=True)
class SegmentReplacement:
    """Replacement targeting a segment or a sub-range inside it.

    Offsets are in characters of the segment's text.
    If start_offset and end_offset are None -> whole segment.
    """

    container_id: str | None = None
    id: str | None = None
    span_id: str | None = None
    text: str = ""
    start_offset: int | None = None
    end_offset: int | None = None


class UnsupportedRevisionError(DocumentError):
    """Raised when a DOCX uses a revision form that cannot be written in place."""


SpanRole = Literal["run", "insertion", "deletion", "hyperlink"]
InlineSegmentKind = Literal["text", "opaque"]


@dataclass(frozen=True)
class AddressableSpan:
    """Stable mechanical span inside a paragraph's extracted text.

    ``role`` is the OOXML wrapper that owns the characters:
    ``run``, ``insertion`` (``w:ins``), ``deletion`` (``w:del`` / ``w:delText``),
    or ``hyperlink`` (``w:hyperlink``). Nested wrappers keep both identities:
    a hyperlink inside an insertion is ``role="hyperlink"`` and still carries
    revision author/date/id.

    Offsets are in the same character space as ``TextSegment.text`` and
    ``SegmentReplacement`` (including deleted text). Consumers can project an
    after-changes view or a revision ledger; Docxtor does not.
    """

    span_id: str
    container_id: str
    role: SpanRole
    text: str
    start_offset: int
    end_offset: int
    paragraph_index: int | None = None
    revision_id: str | None = None
    revision_author: str | None = None
    revision_date: str | None = None
    hyperlink_anchor: str | None = None
    hyperlink_rel_id: str | None = None


@dataclass(frozen=True)
class AddressableComment:
    """Mechanical Word comment body with stable identity and thread metadata.

    ``container_id`` addresses the comment body paragraph that carries the text
    (``comment:{id}:p:{n}``). ``locator`` is the story paragraph that hosts the
    ``w:commentRangeStart`` marker; replies typically have none. Parent/reply
    identity comes from ``commentsExtended.xml`` when present. Docxtor does not
    interpret review meaning.
    """

    comment_id: str
    container_id: str
    text: str
    author: str = ""
    initials: str | None = None
    locator: str | None = None
    anchor_text: str = ""
    parent_id: str | None = None
    date: str | None = None


@dataclass
class InlineSegment:
    """Canonical mechanical segment for paragraph-level DOCX manipulation.

    This is the single source of truth for run/offset addressing, visible-text
    coordinate math, rPr formatting preservation, and opaque inline content
    (images, tabs, breaks, fields, hyperlinks, pre-existing revisions, ...).

    reviewkit (and other consumers) MUST delegate decomposition, splitting,
    insertion, and range replacement to this representation instead of
    reimplementing the logic.

    - kind="text": editable run text; rpr carries formatting to preserve on split/replace.
    - kind="opaque": non-text inline; element is the original XML to re-emit verbatim;
      text holds the visible contribution (e.g. "\t", "\n", or extracted t text) so
      char offsets stay aligned with parser coordinate systems.
    """

    kind: InlineSegmentKind
    text: str
    rpr: Any | None = None
    element: Any | None = None
