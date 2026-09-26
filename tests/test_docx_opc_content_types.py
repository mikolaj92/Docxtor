from io import BytesIO
from zipfile import ZipFile

from docxtor.docx_opc_content_types import (
    PROFILE_ID,
    ContentTypeSeverity,
    validate_docx_content_types,
)

CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_CT = "application/vnd.openxmlformats-package.relationships+xml"
MAIN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"


def _package(declarations: str, parts: dict[str, bytes] | None = None) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        archive.writestr("[Content_Types].xml", declarations.encode())
        for name, data in (parts or {"word/document.xml": b"<doc/>"}).items():
            archive.writestr(name, data)
    return stream.getvalue()


def _types(*children: str) -> str:
    return f'<Types xmlns="{CT}">{"".join(children)}</Types>'


def test_valid_known_part_types_are_accepted() -> None:
    declarations = _types(
        f'<Default Extension="rels" ContentType="{REL_CT}"/>',
        f'<Override PartName="/word/document.xml" ContentType="{MAIN_CT}"/>',
    )
    result = validate_docx_content_types(_package(declarations))
    assert result.profile_id == PROFILE_ID
    assert result.valid
    assert result.complete
    assert result.diagnostics == ()


def test_reports_missing_duplicate_and_conflicting_declarations() -> None:
    declarations = _types(
        '<Override PartName="/word/document.xml" ContentType="application/xml"/>',
        f'<Override PartName="/word/document.xml" ContentType="{MAIN_CT}"/>',
    )
    result = validate_docx_content_types(_package(declarations))
    codes = {d.code for d in result.diagnostics}
    assert "content_type_declaration_conflict" in codes
    assert "known_part_content_type_mismatch" in codes
    assert any(d.part_name == "word/document.xml" and d.declaration for d in result.diagnostics)


def test_duplicate_identical_declaration_is_reported_and_unknown_part_is_coverage_gap() -> None:
    declarations = _types(
        f'<Default Extension="rels" ContentType="{REL_CT}"/>',
        f'<Default Extension="rels" ContentType="{REL_CT}"/>',
        f'<Override PartName="/word/document.xml" ContentType="{MAIN_CT}"/>',
    )
    result = validate_docx_content_types(
        _package(declarations, {"word/document.xml": b"<doc/>", "custom/foo.bin": b"opaque"})
    )
    assert any(d.code == "content_type_declaration_duplicate" for d in result.diagnostics)
    assert any(d.code == "unknown_part_content_type" for d in result.diagnostics)
    assert any(d.severity is ContentTypeSeverity.UNKNOWN for d in result.diagnostics)
    assert not result.complete


def test_reports_unresolved_declaration_and_missing_content_types_part() -> None:
    unresolved = _types(
        f'<Override PartName="/word/document.xml" ContentType="{MAIN_CT}"/>',
        '<Override PartName="/custom/ghost.xml" ContentType="application/xml"/>',
    )
    result = validate_docx_content_types(_package(unresolved))
    assert any(d.code == "content_type_declaration_unresolved" for d in result.diagnostics)

    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        archive.writestr("word/document.xml", b"<doc/>")
    missing = validate_docx_content_types(stream.getvalue())
    assert any(d.code == "content_types_part_missing" for d in missing.diagnostics)
