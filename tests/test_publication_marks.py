from io import BytesIO

from docx import Document

from docxtor import (
    BodyAppendix,
    DocumentMark,
    append_body_appendix,
    has_body_appendix,
    has_document_mark,
    remove_body_appendix,
    stamp_document_mark,
)


def _data() -> bytes:
    document = Document()
    document.add_paragraph("original")
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def test_mark_and_appendix_are_typed_physical_operations() -> None:
    mark = DocumentMark("final", "Final", ("one", "two"), "Final artifact")
    marked = stamp_document_mark(_data(), mark)
    assert has_document_mark(marked, mark)
    appendix = BodyAppendix("Review trail", ("Intro", "1. [A] note"))
    appended = append_body_appendix(marked, appendix)
    assert has_body_appendix(appended, heading=appendix.heading)
    cleaned, changed = remove_body_appendix(appended, heading=appendix.heading)
    assert changed and not has_body_appendix(cleaned, heading=appendix.heading)
    assert Document(BytesIO(cleaned)).paragraphs[0].text == "original"


def test_unreadable_mark_package_has_public_error() -> None:
    import pytest

    from docxtor import PublicationMarkError

    with pytest.raises(PublicationMarkError, match="unreadable"):
        has_body_appendix(b"not zip", heading="x")


def test_append_body_appendix_can_set_paragraph_pagination_controls() -> None:
    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry", "Second entry"),
        keep_paragraphs_together=True,
        keep_heading_with_next=True,
    )

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )

    assert paragraphs[heading_index].paragraph_format.keep_with_next is True
    appendix_paragraphs = paragraphs[heading_index + 1 :]
    assert all(
        paragraph.paragraph_format.keep_together is True for paragraph in appendix_paragraphs
    )
    assert all(
        paragraph.paragraph_format.keep_with_next is not True for paragraph in appendix_paragraphs
    )


def test_append_body_appendix_preserves_original_paragraph_content_and_formatting() -> None:
    source_document = Document()
    original = source_document.add_paragraph()
    original.add_run("Original ")
    original.add_run("bold").bold = True
    original.add_run(" and ")
    original.add_run("italic").italic = True
    source = BytesIO()
    source_document.save(source)

    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry"),
        keep_paragraphs_together=True,
        keep_heading_with_next=True,
        keep_first_paragraph_with_next=True,
    )
    appended = append_body_appendix(source.getvalue(), appendix)
    retained = Document(BytesIO(appended)).paragraphs[0]

    assert retained.text == "Original bold and italic"
    assert retained.runs[1].bold is True
    assert retained.runs[3].italic is True


def test_append_body_appendix_can_keep_only_first_paragraph_with_next() -> None:
    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry", "Second entry"),
        keep_first_paragraph_with_next=True,
    )

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )
    content = paragraphs[heading_index + 1 :]

    assert paragraphs[heading_index].paragraph_format.keep_with_next is None
    assert content[0].paragraph_format.keep_with_next is True
    assert all(
        paragraph.paragraph_format.keep_with_next is not True for paragraph in content[1:]
    )


def test_append_body_appendix_keeps_default_pagination_behavior() -> None:
    appendix = BodyAppendix("Review trail", ("Intro", "First entry"))

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )

    assert [paragraph.text for paragraph in paragraphs] == [
        "original",
        "",
        "Review trail",
        "Intro",
        "First entry",
    ]
    assert paragraphs[heading_index].paragraph_format.keep_with_next is None
    assert all(
        paragraph.paragraph_format.keep_together is None
        for paragraph in paragraphs[heading_index + 1 :]
    )
    assert all(
        paragraph.paragraph_format.keep_with_next is not True
        for paragraph in paragraphs[heading_index + 1 :]
    )
