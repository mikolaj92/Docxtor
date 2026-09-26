from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from docxtor.docx_package_diff import PackagePartChangeKind as Kind
from docxtor.docx_package_diff import compare_docx_packages


def _package(parts: list[tuple[str, bytes]]) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        for name, data in parts:
            archive.writestr(name, data)
    return stream.getvalue()


def test_distinguishes_xml_semantics_binary_and_add_remove() -> None:
    before = _package(
        [
            ("word/document.xml", b'<root a="1"><v>one</v></root>'),
            ("word/styles.xml", b"<styles />"),
            ("word/media/img.bin", b"old"),
            ("word/removed.xml", b"<gone/>"),
        ]
    )
    after = _package(
        [
            ("word/document.xml", b'<root a="1"><v>two</v></root>'),
            ("word/styles.xml", b"<styles></styles>"),
            ("word/media/img.bin", b"new"),
            ("word/added.bin", b"new-part"),
        ]
    )
    report = compare_docx_packages(before, after)
    changes = {item.part_name: item for item in report.parts}
    assert changes["word/document.xml"].kind is Kind.SEMANTIC_XML_CHANGED
    assert changes["word/styles.xml"].kind is Kind.BYTE_CHANGED
    assert changes["word/media/img.bin"].kind is Kind.BYTE_CHANGED
    assert changes["word/added.bin"].kind is Kind.ADDED
    assert changes["word/removed.xml"].kind is Kind.REMOVED
    assert changes["word/document.xml"].before_sha256 != changes["word/document.xml"].after_sha256


def test_byte_identical_parts_and_malformed_xml_are_not_semantically_equal() -> None:
    before = _package([("same.bin", b"x"), ("bad.xml", b"<root>")])
    after = _package([("same.bin", b"x"), ("bad.xml", b"<root />")])
    report = compare_docx_packages(before, after)
    changes = {item.part_name: item for item in report.parts}
    assert changes["same.bin"].kind is Kind.BYTE_IDENTICAL
    assert changes["bad.xml"].kind is Kind.MALFORMED_XML
    assert changes["bad.xml"].before_sha256 and changes["bad.xml"].after_sha256


def test_report_order_is_deterministic_and_comparison_is_before_preservation() -> None:
    left = _package([("z.xml", b"<x a='1'/ >".replace(b"/ >", b"/>")), ("a", b"1")])
    right = _package([("a", b"1"), ("z.xml", b'<x a="1"></x>')])
    report = compare_docx_packages(left, right)
    assert [item.part_name for item in report.parts] == ["a", "z.xml"]
    assert report.parts[0].kind is Kind.BYTE_IDENTICAL
    assert report.parts[1].kind is Kind.BYTE_CHANGED
