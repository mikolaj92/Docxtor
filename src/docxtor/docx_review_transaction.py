from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from .docx_review_inventory import inventory_review_markup
from .docx_review_models import OperationReceipt, ReviewBatchReceipt, ReviewCoverage


class ReviewTransactionError(ValueError):
    """A batch failed; no partial document is returned or published."""


ReviewMutation = Callable[[bytes], tuple[bytes, OperationReceipt]]


@dataclass(frozen=True)
class ReviewCommand:
    operation_id: str
    mutate: ReviewMutation


def apply_review_batch(data: bytes, commands: Sequence[ReviewCommand]) -> ReviewBatchReceipt:
    """Run mutations in memory and return only a completely validated batch.

    ``data`` is immutable and callers receive no intermediate bytes. If any operation
    fails or leaves incomplete review coverage, this function raises and the caller's
    source and destination remain untouched.
    """
    before = inventory_review_markup(data)
    if before.coverage is ReviewCoverage.INCOMPLETE:
        raise ReviewTransactionError("source review markup coverage is incomplete")
    working = data
    receipts: list[OperationReceipt] = []
    try:
        for command in commands:
            working, receipt = command.mutate(working)
            receipts.append(receipt)
        after = inventory_review_markup(working)
        if after.coverage is ReviewCoverage.INCOMPLETE:
            raise ReviewTransactionError("result review markup coverage is incomplete")
    except ReviewTransactionError:
        raise
    except Exception as exc:
        raise ReviewTransactionError(f"review batch failed before publication: {exc}") from exc
    return ReviewBatchReceipt(
        data=working,
        operations=tuple(receipts),
        inventory_before=before,
        inventory_after=after,
    )
