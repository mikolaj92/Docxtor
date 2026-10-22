from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml

from docxtor import flatten_link_fields

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _replace_parts(path: Path, replacements: dict[str, bytes]) -> None:
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts.update(replacements)
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


@pytest.mark.parametrize("as_string", [False, True])
def test_unchanged_package_keeps_exact_bytes(tmp_path: Path, as_string: bool) -> None:
    path = tmp_path / "unchanged.docx"
    document = Document()
    document.add_paragraph("Plain text — bez linków.")
    document.save(path)
    before = path.read_bytes()
    modified_ns = path.stat().st_mtime_ns

    assert flatten_link_fields(str(path) if as_string else path) is False

    assert path.read_bytes() == before
    assert path.stat().st_mtime_ns == modified_ns


def test_rewritten_package_keeps_display_and_is_idempotent(tmp_path: Path) -> None:
    path = tmp_path / "links.docx"
    document = Document()
    paragraphs = [
        document.add_paragraph(),
        document.sections[0].header.paragraphs[0],
        document.sections[0].footer.paragraphs[0],
    ]
    for paragraph in paragraphs:
        rid = paragraph.part.relate_to("https://example.invalid", RT.HYPERLINK, is_external=True)
        paragraph._p.append(
            parse_xml(
                f'<w:fldSimple xmlns:w="{_W}" w:instr=\'HYPERLINK "https://example.invalid"\'>'
                "<w:r><w:t>Field label</w:t></w:r></w:fldSimple>"
            )
        )
        paragraph._p.append(
            parse_xml(
                f'<w:hyperlink xmlns:w="{_W}" xmlns:r="{_R}" r:id="{rid}">'
                "<w:r><w:t> — Link label</w:t></w:r></w:hyperlink>"
            )
        )
    document.save(path)
    before = path.read_bytes()

    assert flatten_link_fields(path) is True
    rewritten = path.read_bytes()
    assert rewritten != before
    reopened = Document(path)
    assert [
        reopened.paragraphs[0].text,
        reopened.sections[0].header.paragraphs[0].text,
        reopened.sections[0].footer.paragraphs[0].text,
    ] == ["Field label — Link label"] * 3
    with ZipFile(path) as archive:
        for story in ("document", "header1", "footer1"):
            root = ET.fromstring(archive.read(f"word/{story}.xml"))
            assert root.find(f".//{{{_W}}}fldSimple") is None
            assert root.find(f".//{{{_W}}}hyperlink") is None
            rels = ET.fromstring(archive.read(f"word/_rels/{story}.xml.rels"))
            assert not any(rel.get("Type") == RT.HYPERLINK for rel in rels)

    assert flatten_link_fields(path) is False
    assert path.read_bytes() == rewritten


def test_malformed_document_preserves_source(tmp_path: Path) -> None:
    path = tmp_path / "malformed.docx"
    document = Document()
    document.save(path)
    _replace_parts(
        path,
        {
            "word/document.xml": b"<broken>",
            "customXml/item1.xml": b"<item>https://example.invalid</item>",
        },
    )
    before = path.read_bytes()

    with pytest.raises(ET.ParseError):
        flatten_link_fields(path)

    assert path.read_bytes() == before


def test_unreadable_custom_xml_survives_package_rewrite(tmp_path: Path) -> None:
    path = tmp_path / "custom.docx"
    document = Document()
    document.add_paragraph("Display text")
    document.save(path)
    _replace_parts(
        path,
        {
            "customXml/item1.xml": b"<broken>",
            "customXml/item2.xml": b"<item>https://example.invalid</item>",
        },
    )

    assert flatten_link_fields(path) is True

    with ZipFile(path) as archive:
        assert archive.read("customXml/item1.xml") == b"<broken>"
        assert ET.fromstring(archive.read("customXml/item2.xml")).text is None
    assert Document(path).paragraphs[0].text == "Display text"
    rewritten = path.read_bytes()
    assert flatten_link_fields(path) is False
    assert path.read_bytes() == rewritten
