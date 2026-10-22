from __future__ import annotations

from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from docx import Document

from docxtor import (
    CommentAuthor,
    CommentMutationError,
    CommentRange,
    DocxDocument,
    PublishError,
    RevisionAuthor,
    RevisionPosition,
    RevisionRange,
)
from docxtor.docx_review_inventory import inventory_review_markup
from docxtor.docx_review_models import OperationStatus, ReviewCoverage


def _docx(text: str = "Hello world") -> bytes:
    buffer = BytesIO()
    document = Document()
    document.add_paragraph(text)
    document.save(buffer)
    return buffer.getvalue()


def test_add_read_and_remove_exact_comment_range() -> None:
    document = DocxDocument.open_bytes(_docx())
    added = document.add_comment(
        CommentRange("body:p:0", 6, 11, "world"),
        "Neutral note",
        CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z"),
    )
    assert added.receipt.status is OperationStatus.APPLIED
    assert added.receipt.created_ids == ("0",)
    assert len(added.comments) == 1
    comment = added.comments[0]
    assert comment.locator == "body:p:0"
    assert comment.anchor_text == "world"
    assert comment.text == "Neutral note"
    assert comment.author == "Reviewer"
    assert comment.date == "2024-01-01T00:00:00Z"

    inventory = inventory_review_markup(document.to_bytes())
    assert inventory.coverage is ReviewCoverage.COMPLETE
    assert inventory.comments == added.comments

    removed = document.remove_comments({"0"})
    assert removed.receipt.status is OperationStatus.APPLIED
    assert removed.comments == ()
    assert inventory_review_markup(removed.data).comments == ()


def test_comment_preflight_failure_returns_no_partial_bytes() -> None:
    document = DocxDocument.open_bytes(_docx())
    before = document.to_bytes()
    with pytest.raises(CommentMutationError, match="text changed"):
        document.add_comment(
            CommentRange("body:p:0", 6, 11, "stale"),
            "Note",
            CommentAuthor("Reviewer"),
        )
    assert sha256(document.to_bytes()).hexdigest() == sha256(before).hexdigest()


def test_publish_preserves_existing_target_on_validator_failure(tmp_path: Path) -> None:
    path = tmp_path / "keep.docx"
    path.write_bytes(_docx())
    original = path.read_bytes()
    document = DocxDocument.open(path)

    def reject(_path: Path) -> None:
        raise ValueError("rejected")

    with pytest.raises(PublishError, match="rejected"):
        document.publish(validators=(reject,))
    assert path.read_bytes() == original


def test_publish_returns_receipt_and_normalizes_zip_timestamps(tmp_path: Path) -> None:
    destination = tmp_path / "published.docx"
    destination.write_bytes(_docx())
    document = DocxDocument.open(destination)
    receipt = document.publish()
    assert receipt.destination == destination
    assert receipt.size == destination.stat().st_size
    assert receipt.sha256 == sha256(destination.read_bytes()).hexdigest()
    with ZipFile(destination) as archive:
        assert {info.date_time for info in archive.infolist()} == {(1980, 1, 1, 0, 0, 0)}


def test_inventory_is_incomplete_for_malformed_story_xml() -> None:
    source = _docx()
    output = BytesIO()
    with ZipFile(BytesIO(source)) as before, ZipFile(output, "w", ZIP_DEFLATED) as after:
        for info in before.infolist():
            data = before.read(info.filename)
            if info.filename == "word/document.xml":
                data = b"<broken"
            after.writestr(info, data)
    inventory = inventory_review_markup(output.getvalue())
    assert inventory.coverage is ReviewCoverage.INCOMPLETE
    assert inventory.diagnostics


def test_batch_failure_does_not_return_partial_document() -> None:
    from docxtor.docx_review_models import OperationReceipt
    from docxtor.docx_review_transaction import (
        ReviewCommand,
        ReviewTransactionError,
        apply_review_batch,
    )

    source = _docx()

    def first(data: bytes) -> tuple[bytes, OperationReceipt]:
        return data + b"partial", OperationReceipt("first", OperationStatus.APPLIED, ())

    def fail(_data: bytes) -> tuple[bytes, OperationReceipt]:
        raise ValueError("boom")

    with pytest.raises(ReviewTransactionError, match="boom"):
        apply_review_batch(source, (ReviewCommand("first", first), ReviewCommand("fail", fail)))
    assert source.startswith(b"PK")


def test_create_inline_revisions() -> None:
    reviewer = RevisionAuthor("Reviewer", "2024-01-01T00:00:00Z")
    inserted = DocxDocument.open_bytes(_docx()).insert_revision(
        RevisionPosition("body:p:0", 5),
        " NEW",
        reviewer,
    )
    assert inserted.after.revisions[0].raw_kind == "ins"
    deleted = DocxDocument.open_bytes(_docx()).delete_revision(
        RevisionRange("body:p:0", 6, 11, "world"),
        reviewer,
    )
    assert deleted.after.revisions[0].raw_kind == "del"


def test_update_comment_rewrites_body_and_keeps_identity() -> None:
    document = DocxDocument.open_bytes(_docx())
    added = document.add_comment(
        CommentRange("body:p:0", 6, 11, "world"),
        "Neutral note",
        CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z"),
    )
    updated = document.update_comment(
        added.receipt.created_ids[0],
        "Revised note",
        expected_text="Neutral note",
    )
    assert updated.receipt.status is OperationStatus.APPLIED
    assert updated.receipt.operation == "update_comment"
    comment = updated.comments[0]
    assert comment.comment_id == added.receipt.created_ids[0]
    assert comment.text == "Revised note"
    assert comment.locator == "body:p:0"
    assert comment.anchor_text == "world"
    assert comment.author == "Reviewer"
    assert comment.initials == "RV"
    assert comment.date == "2024-01-01T00:00:00Z"


def test_update_comment_preflight_failure_returns_no_partial_bytes() -> None:
    document = DocxDocument.open_bytes(_docx())
    added = document.add_comment(
        CommentRange("body:p:0", 6, 11, "world"),
        "Neutral note",
        CommentAuthor("Reviewer"),
    )
    digest = sha256(document.to_bytes()).hexdigest()
    with pytest.raises(CommentMutationError, match="changed"):
        document.update_comment(
            added.receipt.created_ids[0], "Revised", expected_text="stale"
        )
    with pytest.raises(CommentMutationError, match="unknown comment ID"):
        document.update_comment("999", "Revised")
    with pytest.raises(CommentMutationError, match="must not be empty"):
        document.update_comment(added.receipt.created_ids[0], "")
    assert sha256(document.to_bytes()).hexdigest() == digest
