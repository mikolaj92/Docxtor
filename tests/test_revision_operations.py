from __future__ import annotations

# ruff: noqa: E501
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from lxml import etree

from docxtor.docx_revisions import (
    AcceptRevisionsError,
    RejectRevisionsError,
    RevisionInventoryCoverage,
    RevisionKind,
    accept_all_revisions_bytes,
    inventory_revisions_bytes,
    reject_all_revisions_bytes,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W}


def _package(document: str, **parts: str) -> bytes:
    comment_override = (
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        if "word/comments.xml" in parts
        else ""
    )
    content_types = f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{comment_override}</Types>'
    root_rels = f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document)
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return output.getvalue()


def _document(body: str) -> str:
    return f'<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>'


def _xml(data: bytes, part: str = "word/document.xml") -> etree._Element:
    with ZipFile(BytesIO(data)) as archive:
        return etree.fromstring(archive.read(part))


def _text(data: bytes, part: str = "word/document.xml") -> str:
    return "".join(_xml(data, part).xpath("//w:t/text()", namespaces=NS))


def test_inventory_and_inline_accept_reject() -> None:
    data = _package(
        _document(
            '<w:p><w:r><w:t>A</w:t></w:r><w:del w:id="7" w:author="Ada" w:date="2025-01-02T03:04:05Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="8" w:author="Bob"><w:r><w:t>new</w:t></w:r></w:ins></w:p>'
        )
    )
    inventory = inventory_revisions_bytes(data)
    assert inventory.coverage is RevisionInventoryCoverage.COMPLETE
    assert [(item.kind, item.revision_id, item.author) for item in inventory.revisions] == [
        (RevisionKind.DELETION, "7", "Ada"),
        (RevisionKind.INSERTION, "8", "Bob"),
    ]
    assert all(item.part_name == "word/document.xml" for item in inventory.revisions)
    assert all(item.locator.startswith("/w:document/") for item in inventory.revisions)
    accepted = accept_all_revisions_bytes(data, drop_comments=False)
    rejected = reject_all_revisions_bytes(data, drop_comments=False)
    assert _text(accepted.output_bytes) == "Anew"
    assert _text(rejected.output_bytes) == "Aold"
    assert accepted.before.count == rejected.before.count == 2
    assert accepted.after.count == rejected.after.count == 0


@pytest.mark.parametrize("operation", [accept_all_revisions_bytes, reject_all_revisions_bytes])
def test_revision_disposition_preserves_untouched_word_xml_bytes(operation) -> None:
    unrelated = (
        b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        + f'<w:fonts xmlns:w="{W}"><w:font w:name="Exact"/></w:fonts>'.encode()
    )
    data = _package(
        _document('<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins></w:p>'),
        **{"word/fontTable.xml": unrelated.decode()},
    )

    result = operation(data).output_bytes

    with ZipFile(BytesIO(result)) as archive:
        assert archive.read("word/fontTable.xml") == unrelated


def test_paragraph_mark_deletion_accepts_by_joining_and_reject_preserves() -> None:
    body = '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="A"/></w:rPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>'
    data = _package(_document(body))
    assert inventory_revisions_bytes(data).revisions[0].paragraph_mark is True
    accepted = accept_all_revisions_bytes(data)
    rejected = reject_all_revisions_bytes(data)
    assert _xml(accepted.output_bytes).xpath("count(//w:p)", namespaces=NS) == 1
    assert _text(accepted.output_bytes) == "firstsecond"
    assert _xml(rejected.output_bytes).xpath("count(//w:p)", namespaces=NS) == 2


def test_header_and_footnote_parts_are_transformed() -> None:
    header = f'<w:hdr xmlns:w="{W}"><w:p><w:ins w:id="2" w:author="H"><w:r><w:t>head</w:t></w:r></w:ins></w:p></w:hdr>'
    notes = f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p><w:del w:id="3"><w:r><w:delText>note</w:delText></w:r></w:del></w:p></w:footnote></w:footnotes>'
    data = _package(
        _document("<w:p/>"), **{"word/header1.xml": header, "word/footnotes.xml": notes}
    )
    assert {item.part_name for item in inventory_revisions_bytes(data).revisions} == {
        "word/header1.xml",
        "word/footnotes.xml",
    }
    accepted = accept_all_revisions_bytes(data).output_bytes
    assert _text(accepted, "word/header1.xml") == "head"
    assert _text(accepted, "word/footnotes.xml") == ""
    rejected = reject_all_revisions_bytes(data).output_bytes
    assert _text(rejected, "word/header1.xml") == ""
    assert _text(rejected, "word/footnotes.xml") == "note"


def test_comments_can_be_preserved_or_dropped() -> None:
    document = _document(
        '<w:p><w:commentRangeStart w:id="0"/><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>'
    )
    comments = f'<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>remark</w:t></w:r></w:p></w:comment></w:comments>'
    rels = f'<Relationships xmlns="{REL}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'
    data = _package(
        document, **{"word/comments.xml": comments, "word/_rels/document.xml.rels": rels}
    )
    preserved = accept_all_revisions_bytes(data, drop_comments=False).output_bytes
    with ZipFile(BytesIO(preserved)) as archive:
        assert "word/comments.xml" in archive.namelist()
    assert _xml(preserved).xpath("count(//w:commentReference)", namespaces=NS) == 1
    dropped = accept_all_revisions_bytes(data).output_bytes
    with ZipFile(BytesIO(dropped)) as archive:
        assert "word/comments.xml" not in archive.namelist()
        assert b"comments" not in archive.read("word/_rels/document.xml.rels")
        assert b"comments" not in archive.read("[Content_Types].xml")
    assert _xml(dropped).xpath("count(//w:commentReference)", namespaces=NS) == 0


@pytest.mark.parametrize("operation", [accept_all_revisions_bytes, reject_all_revisions_bytes])
def test_unsupported_move_range_fails_before_return(operation) -> None:
    data = _package(_document('<w:p><w:moveFromRangeStart w:id="1"/></w:p>'))
    assert inventory_revisions_bytes(data).coverage is RevisionInventoryCoverage.INCOMPLETE
    with pytest.raises((AcceptRevisionsError, RejectRevisionsError), match="unsupported"):
        operation(data)


def test_operation_rejects_malformed_package() -> None:
    with pytest.raises(AcceptRevisionsError):
        accept_all_revisions_bytes(b"not a zip")


def test_reject_fails_closed_for_structural_cell_revision() -> None:
    data = _package(
        _document(
            '<w:tbl><w:tr><w:trPr><w:cellIns w:id="1"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
        )
    )
    with pytest.raises(RejectRevisionsError, match="cellIns"):
        reject_all_revisions_bytes(data)
