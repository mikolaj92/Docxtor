from base64 import b64decode
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.section import WD_SECTION
from lxml import etree

from docxtor import (
    BodyAppendix,
    DocumentMark,
    append_body_appendix,
    has_body_appendix,
    has_document_mark,
    remove_body_appendix,
    stamp_document_mark,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_V = "urn:schemas-microsoft-com:vml"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_FOOTER_PARTS = tuple(f"word/footer{i}.xml" for i in range(1, 5))
_PNG = b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def _qname(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _synthetic_multi_footer_docx() -> bytes:
    document = Document()
    first = document.sections[0]
    first.different_first_page_header_footer = True
    document.settings.odd_and_even_pages_header_footer = True
    first.footer.paragraphs[0].text = "default footer"
    first.first_page_footer.paragraphs[0].text = "first-page footer"
    first.even_page_footer.paragraphs[0].text = "even-page footer"

    second = document.add_section(WD_SECTION.NEW_PAGE)
    second.footer.is_linked_to_previous = False
    second.footer.paragraphs[0].text = "linked-source footer"
    document.add_section(WD_SECTION.NEW_PAGE)

    stream = BytesIO()
    document.save(stream)
    source = stream.getvalue()
    with ZipFile(BytesIO(source)) as archive:
        members = [(info.filename, archive.read(info.filename)) for info in archive.infolist()]

    rewritten: list[tuple[str, bytes]] = []
    for name, payload in members:
        if name == "[Content_Types].xml":
            root = etree.fromstring(payload)
            if not any(
                child.get("Extension") == "png"
                for child in root.findall(_qname(_CT, "Default"))
            ):
                default = etree.SubElement(root, _qname(_CT, "Default"))
                default.set("Extension", "png")
                default.set("ContentType", "image/png")
                payload = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            rewritten.append((name, payload))
            continue
        if name not in _FOOTER_PARTS:
            rewritten.append((name, payload))
            continue
        index = int(name.removeprefix("word/footer").removesuffix(".xml"))
        root = etree.fromstring(payload)
        for child in list(root):
            root.remove(child)
        if index < 4:
            paragraph = etree.SubElement(root, _qname(_W, "p"))
            field_run = etree.SubElement(paragraph, _qname(_W, "r"))
            begin = etree.SubElement(field_run, _qname(_W, "fldChar"))
            begin.set(_qname(_W, "fldCharType"), "begin")
            instruction_run = etree.SubElement(paragraph, _qname(_W, "r"))
            instruction = etree.SubElement(instruction_run, _qname(_W, "instrText"))
            instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            instruction.text = " PAGE \\* MERGEFORMAT "
            end_run = etree.SubElement(paragraph, _qname(_W, "r"))
            end = etree.SubElement(end_run, _qname(_W, "fldChar"))
            end.set(_qname(_W, "fldCharType"), "end")

            drawing_run = etree.SubElement(paragraph, _qname(_W, "r"))
            alternate = etree.SubElement(drawing_run, _qname(_MC, "AlternateContent"))
            choice = etree.SubElement(alternate, _qname(_MC, "Choice"))
            choice.set("Requires", "wps")
            drawing = etree.SubElement(choice, _qname(_W, "drawing"))
            inline = etree.SubElement(drawing, _qname(_WP, "inline"))
            graphic = etree.SubElement(inline, _qname(_A, "graphic"))
            graphic_data = etree.SubElement(graphic, _qname(_A, "graphicData"))
            graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
            picture = etree.SubElement(graphic_data, _qname(_PIC, "pic"))
            blip_fill = etree.SubElement(picture, _qname(_PIC, "blipFill"))
            blip = etree.SubElement(blip_fill, _qname(_A, "blip"))
            blip.set(_qname(_R, "embed"), f"rIdGraphic{index}")
            fallback = etree.SubElement(alternate, _qname(_MC, "Fallback"))
            pict = etree.SubElement(fallback, _qname(_W, "pict"))
            shape = etree.SubElement(pict, _qname(_V, "shape"))
            imagedata = etree.SubElement(shape, _qname(_V, "imagedata"))
            imagedata.set(_qname(_R, "id"), f"rIdGraphic{index}")
            marker = etree.SubElement(root, _qname(_W, "p"))
            marker_run = etree.SubElement(marker, _qname(_W, "r"))
            marker_text = etree.SubElement(marker_run, _qname(_W, "t"))
            marker_text.text = f"legacy-footer-{index}"
        rewritten.append(
            (name, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))
        )

    for index in range(1, 5):
        rels_name = f"word/_rels/footer{index}.xml.rels"
        rels = etree.Element(_qname(_REL, "Relationships"), nsmap={None: _REL})
        relationship = etree.SubElement(rels, _qname(_REL, "Relationship"))
        relationship.set("Id", f"rIdGraphic{index}")
        relationship.set("Type", f"{_R}/image")
        relationship.set("Target", f"media/footer{index}.png")
        rewritten.append(
            (
                rels_name,
                etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True),
            )
        )
        rewritten.append((f"word/media/footer{index}.png", _PNG))

    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, payload in rewritten:
            archive.writestr(name, payload)
    return output.getvalue()


def _package_member(data: bytes, name: str) -> bytes:
    with ZipFile(BytesIO(data)) as archive:
        return archive.read(name)


def _xml_count(data: bytes, name: str, local_name: str) -> int:
    root = etree.fromstring(_package_member(data, name))
    return len(root.xpath(f".//*[local-name()='{local_name}']"))


def _package_names(data: bytes) -> tuple[str, ...]:
    with ZipFile(BytesIO(data)) as archive:
        return tuple(archive.namelist())


def test_python_docx_roundtrip_alone_keeps_complex_footer_payload() -> None:
    source = _synthetic_multi_footer_docx()
    stream = BytesIO()
    Document(BytesIO(source)).save(stream)
    roundtripped = stream.getvalue()

    for name in _FOOTER_PARTS[:3]:
        after = _package_member(roundtripped, name)
        assert b"legacy-footer-" in after
        assert (
            _xml_count(roundtripped, name, "instrText")
            == _xml_count(source, name, "instrText")
            == 1
        )
        assert (
            _xml_count(roundtripped, name, "AlternateContent")
            == _xml_count(source, name, "AlternateContent")
            == 1
        )
        assert (
            _xml_count(roundtripped, name, "drawing")
            == _xml_count(source, name, "drawing")
            == 1
        )
        rels_name = name.replace("word/", "word/_rels/") + ".rels"
        assert _package_member(roundtripped, rels_name) == (
            _package_member(source, rels_name)
        )


def test_stamp_document_mark_preserves_complex_footer_payload_and_relationships() -> None:
    source = _synthetic_multi_footer_docx()
    mark = DocumentMark("final", "Final", ("one", "two"), "Publication mark")

    marked = stamp_document_mark(source, mark)

    assert set(_package_names(marked)) == set(_package_names(source))
    assert _package_member(marked, "word/document.xml") == _package_member(
        source, "word/document.xml"
    )
    assert has_document_mark(marked, mark)
    for name in _FOOTER_PARTS[:3]:
        after = _package_member(marked, name)
        assert b"legacy-footer-" in after
        assert b"Publication mark" in after
        assert (
            _xml_count(marked, name, "instrText")
            == _xml_count(source, name, "instrText")
            == 1
        )
        assert (
            _xml_count(marked, name, "AlternateContent")
            == _xml_count(source, name, "AlternateContent")
            == 1
        )
        assert (
            _xml_count(marked, name, "drawing")
            == _xml_count(source, name, "drawing")
            == 1
        )
        rels_name = name.replace("word/", "word/_rels/") + ".rels"
        assert _package_member(marked, rels_name) == _package_member(source, rels_name)
        image_name = (
            name.replace("word/footer", "word/media/footer").removesuffix(".xml") + ".png"
        )
        assert _package_member(marked, image_name) == _package_member(
            source, image_name
        )
    empty_footer = _package_member(marked, "word/footer4.xml")
    assert empty_footer.count(b"Publication mark") == 1
    assert _package_member(marked, "word/_rels/footer4.xml.rels") == _package_member(
        source, "word/_rels/footer4.xml.rels"
    )


def test_stamp_document_mark_is_idempotent_for_linked_multi_footer_package() -> None:
    source = _synthetic_multi_footer_docx()
    mark = DocumentMark("final", "Final", ("one", "two"), "Publication mark")

    marked = stamp_document_mark(source, mark)
    stamped_again = stamp_document_mark(marked, mark)

    assert set(_package_names(stamped_again)) == set(_package_names(marked))
    assert _package_member(stamped_again, "word/document.xml") == _package_member(
        marked, "word/document.xml"
    )
    for name in _FOOTER_PARTS:
        first_stamp = _package_member(marked, name)
        second_stamp = _package_member(stamped_again, name)
        assert second_stamp == first_stamp
        assert second_stamp.count(b"Publication mark") == 1


def _data() -> bytes:
    document = Document()
    document.add_paragraph("original")
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def test_mark_and_appendix_are_typed_physical_operations() -> None:
    mark = DocumentMark("final", "Final", ("one", "two"), "Final artifact")
    marked = stamp_document_mark(_data(), mark)
    assert has_document_mark(marked, mark)
    appendix = BodyAppendix("Review trail", ("Intro", "1. [A] note"))
    appended = append_body_appendix(marked, appendix)
    assert has_body_appendix(appended, heading=appendix.heading)
    cleaned, changed = remove_body_appendix(appended, heading=appendix.heading)
    assert changed and not has_body_appendix(cleaned, heading=appendix.heading)
    assert Document(BytesIO(cleaned)).paragraphs[0].text == "original"


def test_unreadable_mark_package_has_public_error() -> None:
    import pytest

    from docxtor import PublicationMarkError

    with pytest.raises(PublicationMarkError, match="unreadable"):
        has_body_appendix(b"not zip", heading="x")


def test_append_body_appendix_can_set_paragraph_pagination_controls() -> None:
    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry", "Second entry"),
        keep_paragraphs_together=True,
        keep_heading_with_next=True,
    )

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )

    assert paragraphs[heading_index].paragraph_format.keep_with_next is True
    appendix_paragraphs = paragraphs[heading_index + 1 :]
    assert all(
        paragraph.paragraph_format.keep_together is True for paragraph in appendix_paragraphs
    )
    assert all(
        paragraph.paragraph_format.keep_with_next is not True for paragraph in appendix_paragraphs
    )


def test_append_body_appendix_preserves_original_paragraph_content_and_formatting() -> None:
    source_document = Document()
    original = source_document.add_paragraph()
    original.add_run("Original ")
    original.add_run("bold").bold = True
    original.add_run(" and ")
    original.add_run("italic").italic = True
    source = BytesIO()
    source_document.save(source)

    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry"),
        keep_paragraphs_together=True,
        keep_heading_with_next=True,
        keep_first_paragraph_with_next=True,
    )
    appended = append_body_appendix(source.getvalue(), appendix)
    retained = Document(BytesIO(appended)).paragraphs[0]

    assert retained.text == "Original bold and italic"
    assert retained.runs[1].bold is True
    assert retained.runs[3].italic is True


def test_append_body_appendix_can_keep_only_first_paragraph_with_next() -> None:
    appendix = BodyAppendix(
        "Review trail",
        ("Intro", "First entry", "Second entry"),
        keep_first_paragraph_with_next=True,
    )

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )
    content = paragraphs[heading_index + 1 :]

    assert paragraphs[heading_index].paragraph_format.keep_with_next is None
    assert content[0].paragraph_format.keep_with_next is True
    assert all(
        paragraph.paragraph_format.keep_with_next is not True for paragraph in content[1:]
    )


def test_append_body_appendix_keeps_default_pagination_behavior() -> None:
    appendix = BodyAppendix("Review trail", ("Intro", "First entry"))

    appended = append_body_appendix(_data(), appendix)
    paragraphs = Document(BytesIO(appended)).paragraphs
    heading_index = next(
        i for i, paragraph in enumerate(paragraphs) if paragraph.text == appendix.heading
    )

    assert [paragraph.text for paragraph in paragraphs] == [
        "original",
        "",
        "Review trail",
        "Intro",
        "First entry",
    ]
    assert paragraphs[heading_index].paragraph_format.keep_with_next is None
    assert all(
        paragraph.paragraph_format.keep_together is None
        for paragraph in paragraphs[heading_index + 1 :]
    )
    assert all(
        paragraph.paragraph_format.keep_with_next is not True
        for paragraph in paragraphs[heading_index + 1 :]
    )
