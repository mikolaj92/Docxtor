from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docxtor import (
    DocxDocument,
    flatten_link_fields,
    instruction_is_strippable_hyperlink_field,
)


@pytest.mark.parametrize("token", ["HYPERLINK", "LINK", "hyperlink", "lInK"])
@pytest.mark.parametrize(
    "uri",
    [
        "mailto:a@example.com",
        "http://example.com",
        "HTTPS://example.com",
        "FiLe:///tmp/document.docx",
        "FtP://example.com/document.docx",
    ],
)
def test_instruction_is_strippable_hyperlink_field(token: str, uri: str) -> None:
    assert instruction_is_strippable_hyperlink_field(f' {token} "{uri}" ') is True


@pytest.mark.parametrize(
    "instruction",
    [
        r'HYPERLINK \l "_Toc123"',
        r" LINK \l _Toc456 ",
        r'hyperlink \L "_tOc789"',
        "",
        "   ",
        '"https://example.com"',
        'INCLUDETEXT "https://example.com"',
        'NOTHYPERLINK "https://example.com"',
        'LINKED "https://example.com"',
        "HYPERLINK",
        'LINK "relative/document.docx"',
        'HYPERLINK "urn:example:document"',
    ],
)
def test_instruction_is_not_strippable_hyperlink_field(instruction: str) -> None:
    assert instruction_is_strippable_hyperlink_field(instruction) is False


def test_flatten_hyperlink_field_keeps_display(tmp_path: Path) -> None:
    path = tmp_path / "x.docx"
    doc = Document()
    paragraph = doc.add_paragraph("Contact: ")

    def run_with(node):
        run = OxmlElement("w:r")
        run.append(node)
        return run

    for kind in ("begin",):
        node = OxmlElement("w:fldChar")
        node.set(qn("w:fldCharType"), kind)
        paragraph._p.append(run_with(node))
    instr = OxmlElement("w:instrText")
    instr.text = ' HYPERLINK "mailto:a@example.com" '
    paragraph._p.append(run_with(instr))
    node = OxmlElement("w:fldChar")
    node.set(qn("w:fldCharType"), "separate")
    paragraph._p.append(run_with(node))
    text = OxmlElement("w:t")
    text.text = "a@example.com"
    paragraph._p.append(run_with(text))
    node = OxmlElement("w:fldChar")
    node.set(qn("w:fldCharType"), "end")
    paragraph._p.append(run_with(node))
    doc.save(path)
    assert flatten_link_fields(path)
    document = DocxDocument.open(path)
    assert any("a@example.com" in segment.text for segment in document.segments)
