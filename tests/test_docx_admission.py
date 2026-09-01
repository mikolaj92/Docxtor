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


def test_pages_survive_unrelated_incomplete_relationships() -> None:
    from zipfile import ZIP_DEFLATED, ZipFile

    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        archive.writestr(
            "word/document.xml",
            '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        )
        archive.writestr(
            "docProps/app.xml",
            '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Pages>12</Pages></Properties>',
        )
    assert inspect_docx_admission(stream.getvalue()).page_count == 12
