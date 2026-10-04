from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

import pytest
from lxml import etree

from docxtor import (
    CommentAuthor,
    CommentMutationError,
    CommentRange,
    DocxDocument,
    RevisionAuthor,
    RevisionPosition,
    RevisionRange,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def _docx(text: str = "Hello world") -> bytes:
    from docx import Document

    buffer = BytesIO()
    document = Document()
    document.add_paragraph(text)
    document.save(buffer)
    return buffer.getvalue()


def _part_xml(data: bytes, part: str = "word/document.xml") -> etree._Element:
    with ZipFile(BytesIO(data)) as archive:
        return etree.fromstring(archive.read(part))


def test_open_handle_comments_and_native_tracked_edits() -> None:
    document = DocxDocument.open_bytes(_docx())
    author = CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z")
    reviewer = RevisionAuthor("Reviewer", "2024-01-01T00:00:00Z")

    added = document.add_comment(
        CommentRange("body:p:0", 6, 11, "world"),
        "Neutral note",
        author,
    )
    comment_id = added.receipt.created_ids[0]
    assert document.comments[0].text == "Neutral note"
    assert document.comments[0].anchor_text == "world"

    updated = document.update_comment(comment_id, "Revised note", expected_text="Neutral note")
    assert updated.receipt.operation == "update_comment"
    assert document.comments[0].text == "Revised note"
    assert document.comments[0].comment_id == comment_id
    assert document.comments[0].author == "Reviewer"

    removed = document.delete_comment(comment_id)
    assert removed.receipt.operation == "remove_comments"
    assert document.comments == ()

    inserted = document.insert_revision(RevisionPosition("body:p:0", 5), " NEW", reviewer)
    assert inserted.after.revisions[0].raw_kind == "ins"

    body = next(segment.text for segment in document.segments if segment.container_id == "body:p:0")
    hello_start = body.index("Hello")
    deleted = document.delete_revision(
        RevisionRange("body:p:0", hello_start, hello_start + 5, "Hello"),
        reviewer,
    )
    assert any(item.raw_kind == "del" for item in deleted.after.revisions)

    body = next(segment.text for segment in document.segments if segment.container_id == "body:p:0")
    world_start = body.index("world")
    _deleted, replaced = document.replace_revision(
        RevisionRange("body:p:0", world_start, world_start + 5, "world"),
        "there",
        reviewer,
    )
    assert any(item.raw_kind == "ins" for item in replaced.after.revisions)
    assert any(item.raw_kind == "del" for item in replaced.after.revisions)

    payload = document.to_bytes()
    root = _part_xml(payload)
    assert root.xpath("count(//w:ins)", namespaces=NS) >= 2
    assert root.xpath("count(//w:del)", namespaces=NS) >= 2
    with ZipFile(BytesIO(payload)) as archive:
        comments_xml = (
            archive.read("word/comments.xml") if "word/comments.xml" in archive.namelist() else b""
        )
    assert b"Revised note" not in comments_xml
    assert not root.xpath("//w:commentReference", namespaces=NS)


def test_open_handle_keeps_comment_across_tracked_insert() -> None:
    document = DocxDocument.open_bytes(_docx())
    added = document.add_comment(
        CommentRange("body:p:0", 0, 5, "Hello"),
        "Neutral note",
        CommentAuthor("Reviewer"),
    )
    document.update_comment(added.receipt.created_ids[0], "Revised note")
    document.insert_revision(
        RevisionPosition("body:p:0", 5),
        " NEW",
        RevisionAuthor("Reviewer", "2024-01-01T00:00:00Z"),
    )
    assert document.comments[0].text == "Revised note"
    assert document.comments[0].anchor_text == "Hello"
    assert any(span.role == "insertion" for span in document.spans)
    root = _part_xml(document.to_bytes())
    assert root.xpath("count(//w:ins)", namespaces=NS) == 1
    assert root.xpath("count(//w:commentReference)", namespaces=NS) == 1


def test_open_handle_keeps_state_when_comment_update_fails() -> None:
    document = DocxDocument.open_bytes(_docx())
    added = document.add_comment(
        CommentRange("body:p:0", 6, 11, "world"),
        "Neutral note",
        CommentAuthor("Reviewer"),
    )
    comments_before = document.comments
    texts_before = document.texts
    with pytest.raises(CommentMutationError, match="unknown comment ID"):
        document.update_comment("999", "Revised")
    assert document.comments == comments_before
    assert document.texts == texts_before
    assert document.comments[0].comment_id == added.receipt.created_ids[0]
