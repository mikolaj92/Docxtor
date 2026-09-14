"""Read-only structural projection and text comparison for DOCX artifacts."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path

from .docx_compare_alignment import align_blocks, diff_text
from .docx_compare_events import compare_comments, compare_formatting, compare_revisions
from .docx_compare_models import ComparisonCoverage, DocxDocumentComparison
from .docx_compare_projection import project_document


def compare_docx_documents(
    left: str | Path | bytes, right: str | Path | bytes
) -> DocxDocumentComparison:
    """Compare two DOCX inputs without modifying them or inferring authorship.

    Text differences use the accepted-content projection: tracked insertions are
    included and tracked deletions are excluded. The original revision wrappers
    remain in each block's spans and are reported separately as lifecycle events.
    """
    before = project_document(left, side="left")
    after = project_document(right, side="right")
    comparison_id = _comparison_id(before.view.sha256, after.view.sha256)
    alignment = align_blocks(before.view, after.view, comparison_id)
    text_changes, text_diagnostics = diff_text(alignment.pairs, alignment.left, alignment.right)
    formatting_changes, formatting_diagnostics = compare_formatting(
        alignment.pairs,
        alignment.left,
        alignment.right,
        before.style_parts,
        after.style_parts,
        comparison_id,
    )
    diagnostics = (
        *alignment.left.diagnostics,
        *alignment.right.diagnostics,
        *alignment.diagnostics,
        *text_diagnostics,
        *formatting_diagnostics,
    )
    return DocxDocumentComparison(
        comparison_id=comparison_id,
        left=alignment.left,
        right=alignment.right,
        block_pairs=alignment.pairs,
        text_changes=text_changes,
        comment_changes=compare_comments(alignment.left, alignment.right, comparison_id),
        revision_events=compare_revisions(
            before.revisions,
            after.revisions,
            alignment.pairs,
            alignment.left,
            alignment.right,
            comparison_id,
        ),
        formatting_changes=formatting_changes,
        coverage=(ComparisonCoverage.INCOMPLETE if diagnostics else ComparisonCoverage.COMPLETE),
        diagnostics=diagnostics,
    )


def _comparison_id(left_sha: str, right_sha: str) -> str:
    material = f"docxtor-compare-v1\x1f{left_sha}\x1f{right_sha}"
    return f"comparison-{sha256(material.encode('utf-8')).hexdigest()[:24]}"


__all__ = ["compare_docx_documents"]
