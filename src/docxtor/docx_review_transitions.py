from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256

from .docx_review_geometry import (
    PhysicalCommentSpan,
    PhysicalReviewGeometry,
    project_docx_review_geometry,
)
from .docx_review_models import OperationReceipt, PhysicalTextInsertion, ReviewCoverage


@dataclass(frozen=True)
class PhysicalSpanTransition:
    source: PhysicalReviewGeometry
    result: PhysicalReviewGeometry
    insertions: tuple[PhysicalTextInsertion, ...]
    proven: bool


def record_span_transition(
    source: bytes, result: bytes, receipt: OperationReceipt
) -> PhysicalSpanTransition:
    before, after = project_docx_review_geometry(source), project_docx_review_geometry(result)
    old = {item.locator: item.text for item in before.paragraphs if item.story_kind != "comment"}
    new = {item.locator: item.text for item in after.paragraphs if item.story_kind != "comment"}
    proven = (
        before.coverage is ReviewCoverage.COMPLETE
        and after.coverage is ReviewCoverage.COMPLETE
        and receipt.before_sha256 == sha256(source).hexdigest()
        and receipt.after_sha256 == sha256(result).hexdigest()
        and old.keys() == new.keys()
    )
    recovered = dict(new)
    for insertion in reversed(receipt.text_insertions):
        value = recovered.get(insertion.locator)
        if (
            value is None
            or insertion.length < 1
            or insertion.start_offset < 0
            or insertion.start_offset + insertion.length > len(value)
        ):
            proven = False
            break
        recovered[insertion.locator] = (
            value[: insertion.start_offset] + value[insertion.start_offset + insertion.length :]
        )
    proven = proven and recovered == old
    return PhysicalSpanTransition(before, after, receipt.text_insertions, proven)


def map_spans(
    transitions: tuple[PhysicalSpanTransition, ...],
    source_sha256: str,
    current: PhysicalReviewGeometry,
    spans: tuple[PhysicalCommentSpan, ...],
) -> tuple[PhysicalCommentSpan, ...]:
    if not spans:
        raise ValueError("review span mapping requires a nonempty captured scope")
    if source_sha256 == current.document_sha256:
        source = current
        chain = ()
    else:
        start = next(
            (
                index
                for index, item in enumerate(transitions)
                if item.source.document_sha256 == source_sha256
            ),
            None,
        )
        if start is None:
            raise ValueError("review source snapshot is absent from native transition history")
        chain = transitions[start:]
        source = chain[0].source
    paragraphs = {item.locator: item for item in source.paragraphs}
    offsets: list[tuple[str, int, int]] = []
    for span in spans:
        paragraph = paragraphs.get(span.locator)
        if (
            paragraph is None
            or not 0 <= span.start_offset <= span.end_offset <= len(paragraph.text)
            or paragraph.text[span.start_offset : span.end_offset] != span.expected_text
        ):
            raise ValueError("review source span does not match its captured snapshot")
        offsets.append((span.locator, span.start_offset, span.end_offset))
    digest = source_sha256
    for transition in chain:
        if not transition.proven or transition.source.document_sha256 != digest:
            raise ValueError("review scope crosses an unproved native text transition")
        for insertion in transition.insertions:
            offsets = [
                (
                    locator,
                    start + insertion.length if insertion.start_offset < start else start,
                    end + insertion.length if insertion.start_offset <= end else end,
                )
                if locator == insertion.locator
                else (locator, start, end)
                for locator, start, end in offsets
            ]
        digest = transition.result.document_sha256
    if digest != current.document_sha256 or current.coverage is not ReviewCoverage.COMPLETE:
        raise ValueError("review transition history does not reach the current complete snapshot")
    current_paragraphs = {item.locator: item for item in current.paragraphs}
    return tuple(
        PhysicalCommentSpan(locator, start, end, current_paragraphs[locator].text[start:end])
        for locator, start, end in offsets
    )
