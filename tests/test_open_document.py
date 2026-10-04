from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree

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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def _docx(*paragraphs: str) -> bytes:
    from docx import Document

    texts = paragraphs or ("Hello world",)
    document = Document()
    if document.paragraphs:
        document.paragraphs[0].text = texts[0]
        rest = texts[1:]
    else:
        rest = texts
    for text in rest:
        document.add_paragraph(text)
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _part_xml(data: bytes, part: str = "word/document.xml") -> etree._Element:
    with ZipFile(BytesIO(data)) as archive:
        return etree.fromstring(archive.read(part))


def test_open_handle_comments_and_native_tracked_edits() -> None:
    document = DocxDocument.open_bytes(_docx("Hello", "world", "extra"))
    author = CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z")
    reviewer = RevisionAuthor("Reviewer", "2024-01-01T00:00:00Z")

    added = document.add_comment(
        CommentRange("body:p:0", 0, 5, "Hello"),
        "Neutral note",
        author,
    )
    comment_id = added.receipt.created_ids[0]
    assert document.comments[0].text == "Neutral note"
    assert document.comments[0].anchor_text == "Hello"

    updated = document.update_comment(comment_id, "Revised note", expected_text="Neutral note")
    assert updated.receipt.operation == "update_comment"
    assert document.comments[0].text == "Revised note"
    assert document.comments[0].comment_id == comment_id
    assert document.comments[0].author == "Reviewer"

    removed = document.remove_comments({comment_id})
    assert removed.receipt.operation == "remove_comments"
    assert document.comments == ()

    inserted = document.insert_revision(RevisionPosition("body:p:0", 5), "!", reviewer)
    assert inserted.after.revisions[0].raw_kind == "ins"
    deleted = document.delete_revision(RevisionRange("body:p:1", 0, 5, "world"), reviewer)
    assert any(item.raw_kind == "del" for item in deleted.after.revisions)
    _deleted, replaced = document.replace_revision(
        RevisionRange("body:p:2", 0, 5, "extra"),
        "gone",
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


def test_open_path_handle_publish_writes_back_to_source_file(tmp_path: Path) -> None:
    path = tmp_path / "review.docx"
    path.write_bytes(_docx("Hello", "world", "extra"))
    original = path.read_bytes()
    document = DocxDocument.open(path)
    author = CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z")
    reviewer = RevisionAuthor("Reviewer", "2024-01-01T00:00:00Z")

    added = document.add_comment(
        CommentRange("body:p:0", 0, 5, "Hello"),
        "Neutral note",
        author,
    )
    document.update_comment(added.receipt.created_ids[0], "Revised note")
    document.remove_comments({added.receipt.created_ids[0]})
    document.insert_revision(RevisionPosition("body:p:0", 5), "!", reviewer)
    document.delete_revision(RevisionRange("body:p:1", 0, 5, "world"), reviewer)
    document.replace_revision(RevisionRange("body:p:2", 0, 5, "extra"), "gone", reviewer)

    receipt = document.publish()
    assert receipt.destination == path
    assert path.read_bytes() != original
    reopened = DocxDocument.open(path)
    root = _part_xml(path.read_bytes())
    assert root.xpath("count(//w:ins)", namespaces=NS) >= 2
    assert root.xpath("count(//w:del)", namespaces=NS) >= 2
    assert not root.xpath("//w:commentReference", namespaces=NS)
    assert reopened.comments == ()


def test_publish_without_source_path_fails_closed() -> None:
    document = DocxDocument.open_bytes(_docx())
    with pytest.raises(PublishError, match="destination path"):
        document.publish()


def test_publish_validator_failure_leaves_opened_file_unchanged(tmp_path: Path) -> None:
    path = tmp_path / "keep.docx"
    path.write_bytes(_docx())
    original = path.read_bytes()
    document = DocxDocument.open(path)
    document.add_comment(
        CommentRange("body:p:0", 0, 5, "Hello"),
        "Neutral note",
        CommentAuthor("Reviewer"),
    )

    def reject(_path: Path) -> None:
        raise ValueError("rejected")

    with pytest.raises(PublishError, match="rejected"):
        document.publish(validators=(reject,))
    assert path.read_bytes() == original


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


def _comment_ids(root: etree._Element, local: str) -> list[str]:
    return [str(value) for value in root.xpath(f"//w:{local}/@w:id", namespaces=NS)]


def test_open_path_adds_repeated_native_comments_on_one_sentence(tmp_path: Path) -> None:
    path = tmp_path / "sentence.docx"
    sentence = "The same sentence."
    path.write_bytes(_docx(sentence))
    document = DocxDocument.open(path)
    author = CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z")
    target = CommentRange("body:p:0", 0, len(sentence), sentence)
    notes = ("first pass", "second pass", "third pass")

    for note in notes:
        document.add_comment(target, note, author)
    document.publish()

    reopened = DocxDocument.open(path)
    assert [comment.text for comment in reopened.comments] == list(notes)
    assert all(
        comment.locator == "body:p:0" and comment.anchor_text == sentence
        for comment in reopened.comments
    )
    root = _part_xml(path.read_bytes())
    ids = ["0", "1", "2"]
    assert sorted(_comment_ids(root, "commentRangeStart")) == ids
    assert sorted(_comment_ids(root, "commentRangeEnd")) == ids
    assert sorted(_comment_ids(root, "commentReference")) == ids
    comments_xml = ZipFile(path).read("word/comments.xml")
    for note in notes:
        assert note.encode() in comments_xml


def test_open_path_adds_comment_over_a_span_that_already_has_markers(tmp_path: Path) -> None:
    path = tmp_path / "overlap.docx"
    sentence = "The same sentence."
    path.write_bytes(_docx(sentence))
    document = DocxDocument.open(path)
    author = CommentAuthor("Reviewer", "RV", "2024-01-01T00:00:00Z")

    inner = document.add_comment(
        CommentRange("body:p:0", 4, 8, "same"),
        "inner pass",
        author,
    )
    outer = document.add_comment(
        CommentRange("body:p:0", 0, len(sentence), sentence),
        "outer pass",
        author,
    )
    document.publish()

    reopened = DocxDocument.open(path)
    by_id = {comment.comment_id: comment for comment in reopened.comments}
    inner_id = inner.receipt.created_ids[0]
    outer_id = outer.receipt.created_ids[0]
    assert by_id[inner_id].text == "inner pass"
    assert by_id[inner_id].anchor_text == "same"
    assert by_id[outer_id].text == "outer pass"
    assert by_id[outer_id].anchor_text == sentence
    root = _part_xml(path.read_bytes())
    assert sorted(_comment_ids(root, "commentRangeStart")) == sorted([inner_id, outer_id])
    assert sorted(_comment_ids(root, "commentRangeEnd")) == sorted([inner_id, outer_id])
    assert sorted(_comment_ids(root, "commentReference")) == sorted([inner_id, outer_id])
