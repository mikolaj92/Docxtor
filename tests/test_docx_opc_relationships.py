from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from docxtor.docx_opc_relationships import (
    PROFILE_ID,
    RelationshipSeverity,
    validate_docx_relationships,
)

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def _package(root_rels: str | None = None, document_rels: str = "") -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        archive.writestr(
            "word/document.xml",
            b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
        )
        if root_rels is not None:
            archive.writestr("_rels/.rels", root_rels.encode())
        if document_rels:
            archive.writestr("word/_rels/document.xml.rels", document_rels.encode())
    return stream.getvalue()


def _root(*children: str) -> str:
    return f'<Relationships xmlns="{REL}">{"".join(children)}</Relationships>'


def _relationship(rid: str, kind: str, target: str) -> str:
    return f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"/>'


def test_relationship_profile_accepts_minimal_valid_docx() -> None:
    result = validate_docx_relationships(
        _package(_root(_relationship("r1", OFFICE + "officeDocument", "word/document.xml")))
    )
    assert result.profile_id == PROFILE_ID
    assert result.valid
    assert result.complete
    assert result.diagnostics == ()


def test_reports_missing_required_main_part_and_office_relationship() -> None:
    payload = BytesIO()
    with ZipFile(payload, "w") as archive:
        archive.writestr("[Content_Types].xml", b"<Types/>")
        archive.writestr("_rels/.rels", _root())
    result = validate_docx_relationships(payload.getvalue())
    assert {d.code for d in result.diagnostics} >= {
        "required_part_missing",
        "required_relationship_cardinality",
    }


def test_reports_duplicate_singleton_and_source_policy_violation() -> None:
    rels = _root(
        _relationship("r1", OFFICE + "officeDocument", "word/document.xml")
        + _relationship("s1", OFFICE + "styles", "word/styles.xml")
        + _relationship("s2", OFFICE + "styles", "word/styles.xml")
    )
    result = validate_docx_relationships(_package(rels))
    codes = {d.code for d in result.diagnostics}
    assert "relationship_cardinality_exceeded" in codes
    assert "disallowed_relationship_source" in codes
    assert "missing_relationship_target" in codes
    assert all(d.constraint_id.startswith(PROFILE_ID + ":") for d in result.diagnostics)


def test_unknown_relationship_type_is_explicit_incomplete_coverage() -> None:
    root = _root(
        _relationship("r1", OFFICE + "officeDocument", "word/document.xml"),
        _relationship("x1", "urn:vendor:future", "opaque.bin"),
    )
    result = validate_docx_relationships(_package(root))
    assert any(d.code == "unprofiled_relationship_type" for d in result.diagnostics)
    assert not result.complete
    assert any(d.severity is RelationshipSeverity.UNKNOWN for d in result.diagnostics)
