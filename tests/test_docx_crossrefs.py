from io import BytesIO
from zipfile import ZipFile

from docxtor.docx_crossrefs import (
    CrossReferenceSeverity,
    diagnose_docx_cross_references,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _package(body: str) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        archive.writestr("[Content_Types].xml", f'<Types xmlns="{CT}"/>')
        archive.writestr(
            "word/document.xml",
            f'<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}</w:body></w:document>',
        )
    return stream.getvalue()


def test_resolves_bookmark_anchor_and_ref_field_and_paired_id() -> None:
    data = _package(
        '<w:bookmarkStart w:id="7" w:name="target"/><w:hyperlink w:anchor="target"/>'
        '<w:fldSimple w:instr=" REF target "/><w:bookmarkEnd w:id="7"/>'
    )
    report = diagnose_docx_cross_references(data)
    assert report.valid
    assert report.complete
    assert report.diagnostics == ()


def test_reports_dangling_anchor_ref_and_unpaired_bookmark() -> None:
    report = diagnose_docx_cross_references(
        _package(
            '<w:bookmarkStart w:id="8" w:name="lonely"/><w:hyperlink w:anchor="absent"/>'
            '<w:fldSimple w:instr=" PAGEREF nowhere "/>'
        )
    )
    findings = {(item.code, item.referenced_value) for item in report.diagnostics}
    assert ("bookmark_pair_mismatch", "8") in findings
    assert ("hyperlink_bookmark_missing", "absent") in findings
    assert ("field_bookmark_missing", "nowhere") in findings
    assert all(
        item.locator and item.part_name == "word/document.xml" for item in report.diagnostics
    )


def test_unsupported_relationship_hyperlink_is_not_clean_coverage() -> None:
    report = diagnose_docx_cross_references(_package('<w:hyperlink r:id="rId7"/>'))
    assert not report.complete
    assert "relationship-backed hyperlink (w:hyperlink/@r:id)" in report.unsupported_forms
    assert any(item.severity is CrossReferenceSeverity.UNKNOWN for item in report.diagnostics)
