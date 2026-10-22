from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docxtor import flatten_link_fields


def _linked_document(path: Path) -> None:
    document = Document()
    paragraphs = [
        document.add_paragraph(),
        document.sections[0].header.paragraphs[0],
        document.sections[0].footer.paragraphs[0],
    ]
    for paragraph in paragraphs:
        paragraph.add_run("Before ")
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), ' HYPERLINK "https://example.invalid/field" ')
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "Field label"
        run.append(text)
        field.append(run)
        paragraph._p.append(field)
        paragraph.add_run(" / ")
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(
            qn("r:id"),
            paragraph.part.relate_to(
                "https://example.invalid/link", RT.HYPERLINK, is_external=True
            ),
        )
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "Link label"
        run.append(text)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
        paragraph.add_run(" after")
    document.save(path)


@pytest.mark.parametrize("unreadable_custom_xml", [False, True])
def test_unchanged_package_is_not_written(tmp_path: Path, unreadable_custom_xml: bool) -> None:
    path = tmp_path / "unchanged.docx"
    document = Document()
    document.add_paragraph("Ordinary text")
    document.save(path)
    if unreadable_custom_xml:
        with ZipFile(path, "a") as archive:
            archive.writestr("customXml/unreadable.xml", b"<broken")
    before = path.read_bytes()
    stat = path.stat()

    assert flatten_link_fields(str(path)) is False

    assert path.read_bytes() == before
    assert path.stat().st_mtime_ns == stat.st_mtime_ns
    assert path.stat().st_ino == stat.st_ino


def test_rewritten_package_preserves_display_and_is_idempotent(tmp_path: Path) -> None:
    path = tmp_path / "linked.docx"
    _linked_document(path)
    with ZipFile(path, "a") as archive:
        archive.writestr("customXml/unreadable.xml", b"<broken")
    before = path.read_bytes()

    assert flatten_link_fields(path) is True
    assert path.read_bytes() != before
    document = Document(path)
    paragraphs = [
        document.paragraphs[0],
        document.sections[0].header.paragraphs[0],
        document.sections[0].footer.paragraphs[0],
    ]
    assert [paragraph.text for paragraph in paragraphs] == [
        "Before Field label / Link label after"
    ] * 3
    for paragraph in paragraphs:
        assert all(rel.reltype != RT.HYPERLINK for rel in paragraph.part.rels.values())
    with ZipFile(path) as archive:
        assert archive.read("customXml/unreadable.xml") == b"<broken"
    rewritten = path.read_bytes()
    stat = path.stat()

    assert flatten_link_fields(path) is False
    assert path.read_bytes() == rewritten
    assert path.stat().st_mtime_ns == stat.st_mtime_ns
    assert path.stat().st_ino == stat.st_ino


@pytest.mark.parametrize("part", ["word/document.xml", "word/header1.xml", "word/footer1.xml"])
def test_malformed_story_preserves_source(tmp_path: Path, part: str) -> None:
    path = tmp_path / "malformed.docx"
    _linked_document(path)
    with ZipFile(path) as archive:
        members = [(info, archive.read(info)) for info in archive.infolist()]
    with ZipFile(path, "w") as archive:
        for info, data in members:
            archive.writestr(info, b"<broken" if info.filename == part else data)
    before = path.read_bytes()
    stat = path.stat()

    with pytest.raises(ET.ParseError):
        flatten_link_fields(path)

    assert path.read_bytes() == before
    assert path.stat().st_mtime_ns == stat.st_mtime_ns
    assert path.stat().st_ino == stat.st_ino
