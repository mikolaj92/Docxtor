from __future__ import annotations

from io import BytesIO
from typing import Any

from docx import Document as PyDocxDocument
from lxml import etree

from .docx_models import AddressableComment
from .docx_package import PackageError, parse_package_xml, read_package_entries
from .docx_review_models import ReviewCoverage, ReviewDiagnostic, ReviewMarkupInventory
from .docx_revisions import RevisionOperationError, inventory_revisions_bytes
from .docx_stories import index_stories


def inventory_review_markup(data: bytes) -> ReviewMarkupInventory:
    """Inventory revisions and comments without equating unreadable with empty."""
    diagnostics: list[ReviewDiagnostic] = []
    comments: tuple[AddressableComment, ...] = ()
    try:
        entries = read_package_entries(data)
        revision_inventory = inventory_revisions_bytes(data)
    except (PackageError, RevisionOperationError) as exc:
        return ReviewMarkupInventory(
            revisions=(),
            comments=(),
            coverage=ReviewCoverage.INCOMPLETE,
            diagnostics=(ReviewDiagnostic("package_unreadable", str(exc)),),
        )
    for diagnostic in revision_inventory.diagnostics:
        diagnostics.append(
            ReviewDiagnostic(
                diagnostic.code,
                diagnostic.detail,
                diagnostic.part_name,
            )
        )
    try:
        comments = tuple(index_stories(PyDocxDocument(BytesIO(data))).comments)
    except (KeyError, ValueError, TypeError, etree.XMLSyntaxError) as exc:
        diagnostics.append(ReviewDiagnostic("comments_unreadable", str(exc), "word/comments.xml"))
    _validate_comment_markers(entries, comments, diagnostics)
    return ReviewMarkupInventory(
        revisions=revision_inventory.revisions,
        comments=comments,
        coverage=ReviewCoverage.INCOMPLETE if diagnostics else ReviewCoverage.COMPLETE,
        diagnostics=tuple(diagnostics),
    )


def _validate_comment_markers(
    entries: tuple[Any, ...],
    comments: tuple[Any, ...],
    diagnostics: list[ReviewDiagnostic],
) -> None:
    bodies = {comment.comment_id for comment in comments}
    replies = {comment.comment_id for comment in comments if comment.parent_id is not None}
    counts: dict[str, dict[str, int]] = {}
    for entry in entries:
        if not entry.name.startswith("word/") or not entry.name.endswith(".xml"):
            continue
        root = parse_package_xml(entry.data, part_name=entry.name)
        for element in root.iter():
            local = _local(element.tag)
            if local not in {"commentRangeStart", "commentRangeEnd", "commentReference"}:
                continue
            comment_id = _attr(element, "id")
            if comment_id is None:
                diagnostics.append(ReviewDiagnostic("comment_marker_without_id", local, entry.name))
                continue
            row = counts.setdefault(comment_id, {})
            row[local] = row.get(local, 0) + 1
    for comment_id in sorted(bodies | set(counts)):
        row = counts.get(comment_id, {})
        if comment_id not in bodies:
            diagnostics.append(ReviewDiagnostic("orphan_comment_marker", comment_id))
        names = ("commentRangeStart", "commentRangeEnd", "commentReference")
        expected = comment_id not in replies or bool(row)
        if expected and any(row.get(name, 0) != 1 for name in names):
            diagnostics.append(ReviewDiagnostic("incomplete_comment_range", comment_id))


def _attr(element: etree._Element, local: str) -> str | None:
    return next(
        (value for name, value in element.attrib.items() if _local(name) == local),
        None,
    )


def _local(tag: object) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""
