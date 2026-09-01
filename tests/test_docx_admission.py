from io import BytesIO

from docx import Document

from docxtor import DocumentPackageKind, inspect_docx_admission


def test_valid_admission_has_page_count() -> None:
    stream = BytesIO()
    Document().save(stream)
    result = inspect_docx_admission(stream.getvalue())
    assert result.package_kind is DocumentPackageKind.VALID
    assert result.has_main_document is True
    assert result.page_count == 1


def test_encrypted_ole_admission() -> None:
    result = inspect_docx_admission(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1rest")
    assert result.package_kind is DocumentPackageKind.ENCRYPTED
