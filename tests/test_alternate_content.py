from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

from docx import Document as PyDocxDocument
from docx.oxml import OxmlElement
from lxml import etree

from docxtor import DocxDocument, SegmentReplacement

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
UNKNOWN = "urn:example:unknown-feature"


def _branch_paragraph(text: str):
    paragraph = OxmlElement("w:p")
    run = OxmlElement("w:r")
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    paragraph.append(run)
    return paragraph


def _docx_with_alternate_content(*, unknown_requirement: bool = False) -> bytes:
    document = PyDocxDocument()
    document.settings.element.set(
        "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable",
        "w14",
    )
    document.add_paragraph("Before")
    document.add_paragraph("placeholder")
    document.add_paragraph("After")
    body = document.element.body
    original = body[1]
    alternate = etree.Element(
        f"{{{MC}}}AlternateContent",
        nsmap={"mc": MC, "w": W, "w14": W14, "x": UNKNOWN},
    )
    choice = etree.SubElement(alternate, f"{{{MC}}}Choice")
    choice.set("Requires", "x" if unknown_requirement else "w14")
    choice.append(_branch_paragraph("Choice text"))
    fallback = etree.SubElement(alternate, f"{{{MC}}}Fallback")
    fallback.append(_branch_paragraph("Fallback text"))
    body.replace(original, alternate)
    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _part(data: bytes, name: str) -> etree._Element:
    with ZipFile(BytesIO(data)) as package:
        return etree.fromstring(package.read(name))


def test_read_indexes_only_one_supported_alternate_content_branch_and_keeps_ids() -> None:
    document = DocxDocument.open_bytes(_docx_with_alternate_content())

    actual = [
        (segment.container_id, segment.paragraph_index, segment.text)
        for segment in document.segments
    ]
    assert actual == [
        ("body:p:0", 0, "Before"),
        ("body:p:1", 1, "Choice text"),
        ("body:p:2", 2, "After"),
    ]
    assert document.alternate_content_coverage.regions == 1
    assert document.alternate_content_coverage.choice_selected == 1
    assert document.alternate_content_coverage.fallback_selected == 0
    assert document.alternate_content_coverage.unknown_requirements == ()
    assert document.alternate_content_coverage.unsupported_requirements == ()


def test_unknown_requires_uses_fallback_and_reports_coverage_gap() -> None:
    document = DocxDocument.open_bytes(_docx_with_alternate_content(unknown_requirement=True))

    assert [segment.text for segment in document.segments] == ["Before", "Fallback text", "After"]
    assert document.alternate_content_coverage.regions == 1
    assert document.alternate_content_coverage.choice_selected == 0
    assert document.alternate_content_coverage.fallback_selected == 1
    assert document.alternate_content_coverage.unknown_requirements == ()
    assert document.alternate_content_coverage.unsupported_requirements == (
        "urn:example:unknown-feature",
    )
    assert document.alternate_content_coverage.complete is False


def test_mutation_updates_selected_branch_and_preserves_unselected_branch_verbatim() -> None:
    source = _docx_with_alternate_content()
    document = DocxDocument.open_bytes(source)
    document.apply_replacements(
        [SegmentReplacement(container_id="body:p:1", text="Updated choice")], strict=True
    )

    root = _part(document.to_bytes(), "word/document.xml")
    alternate = root.find(f".//{{{MC}}}AlternateContent")
    assert alternate is not None
    choice_text = alternate.find(f"{{{MC}}}Choice//{{{W}}}t").text
    fallback_text = alternate.find(f"{{{MC}}}Fallback//{{{W}}}t").text
    assert choice_text == "Updated choice"
    assert fallback_text == "Updated choice"
    reopened = DocxDocument.open_bytes(document.to_bytes())
    assert [segment.container_id for segment in reopened.segments] == [
        "body:p:0",
        "body:p:1",
        "body:p:2",
    ]
