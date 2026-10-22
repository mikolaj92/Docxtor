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


@pytest.mark.parametrize("token", ["HYPERLINK", "LINK", "hyperlink", "LiNk"])
@pytest.mark.parametrize(
    "uri",
    [
        "mailto:a@example.com",
        "http://example.com",
        "https://example.com",
        "file:/example.docx",
        "ftp://example.com/file",
        "HtTpS://example.com",
    ],
)
def test_instruction_recognizes_hyperlink_uri(token: str, uri: str) -> None:
    assert instruction_is_strippable_hyperlink_field(f' {token} "{uri}" ') is True


@pytest.mark.parametrize(
    "instruction",
    [
        r"HYPERLINK \l _Toc123",
        r' LINK \l "_Toc123" ',
        r'hyperlink \L "_toc123"',
        "",
        '"https://example.com"',
        'NOTHYPERLINK "https://example.com"',
        'HYPERLINK "relative/path"',
        'LINK "urn:example:file"',
    ],
)
def test_instruction_rejects_toc_or_missing_token_or_uri(instruction: str) -> None:
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
