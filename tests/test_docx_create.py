from io import BytesIO

from docx import Document

from docxtor import create_docx_from_paragraphs


def test_create_docx_from_paragraphs() -> None:
    data = create_docx_from_paragraphs(("one", "two"))
    assert [item.text for item in Document(BytesIO(data)).paragraphs] == ["one", "two"]
