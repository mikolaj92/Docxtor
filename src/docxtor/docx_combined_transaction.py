from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from .docx import DocxDocument
from .docx_models import SegmentReplacement
from .docx_package_transaction import (
    PackageMutation,
    PackageTransactionReceipt,
    PackageValidator,
    apply_package_transaction,
)


@dataclass(frozen=True)
class CombinedTransactionReceipt:
    data: bytes
    text_targets: tuple[SegmentReplacement, ...]
    package: PackageTransactionReceipt


def apply_docx_transaction(
    data: bytes,
    *,
    text_targets: Sequence[SegmentReplacement] = (),
    package_mutations: Sequence[PackageMutation] = (),
    validators: Sequence[PackageValidator] = (),
    require_complete_coverage: bool = True,
) -> CombinedTransactionReceipt:
    """Apply addressable text and package operations as one publishable result.

    Text targets are strict. Package validation and global dispositions cover the
    final bytes; an exception exposes no intermediate result.
    """
    document = DocxDocument.open_bytes(data)
    normalized = tuple(text_targets)
    document.apply_replacements(list(normalized), strict=True)
    after_text = document.to_bytes()
    package = apply_package_transaction(
        after_text,
        package_mutations,
        validators=validators,
        require_complete_coverage=require_complete_coverage,
    )
    return CombinedTransactionReceipt(package.data, normalized, package)
