from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document as PyDocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from docxtor import (
    DocxDocument,
    InventoryCoverage,
    SurfaceCapability,
    SurfaceKind,
    SurfaceVisibility,
)


def _write_docx(path: Path) -> None:
    source = PyDocxDocument()
    source.add_paragraph("Visible person")
    source.core_properties.author = "Hidden Author"
    source.part.relate_to("mailto:hidden@example.test", RT.HYPERLINK, is_external=True)
    source.save(path)


def _add_unknown_part(path: Path) -> None:
    buffer = BytesIO()
    with ZipFile(path) as source, ZipFile(buffer, "w", ZIP_DEFLATED) as output:
        for item in source.infolist():
            if item.filename not in {"[Content_Types].xml", "word/_rels/document.xml.rels"}:
                output.writestr(item, source.read(item.filename))
        output.writestr("word/embeddings/opaque1.bin", b"opaque payload")
        content_types = source.read("[Content_Types].xml").decode("utf-8")
        content_types = content_types.replace(
            "</Types>",
            '<Override PartName="/word/embeddings/opaque1.bin" '
            'ContentType="application/vnd.example.opaque"/></Types>',
        )
        output.writestr("[Content_Types].xml", content_types)
        rels = source.read("word/_rels/document.xml.rels").decode("utf-8")
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdOpaque" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" '
            'Target="embeddings/opaque1.bin"/></Relationships>',
        )
        output.writestr("word/_rels/document.xml.rels", rels)
    path.write_bytes(buffer.getvalue())


def test_inventory_enumerates_visible_hidden_and_relationship_values(tmp_path: Path) -> None:
    path = tmp_path / "inventory.docx"
    _write_docx(path)

    inventory = DocxDocument.open(path).inventory()

    assert inventory.coverage is InventoryCoverage.COMPLETE
    visible = [surface for surface in inventory.surfaces if surface.value == "Visible person"]
    assert visible
    assert visible[0].kind is SurfaceKind.TEXT
    assert visible[0].visibility is SurfaceVisibility.VISIBLE
    assert visible[0].capability is SurfaceCapability.VALUE_REPLACE

    author = [surface for surface in inventory.surfaces if surface.value == "Hidden Author"]
    assert author
    assert author[0].visibility is SurfaceVisibility.HIDDEN
    assert author[0].part_name == "docProps/core.xml"

    external = [
        surface
        for surface in inventory.surfaces
        if surface.kind is SurfaceKind.RELATIONSHIP and surface.external
    ]
    assert len(external) == 1
    assert external[0].value == "mailto:hidden@example.test"
    assert external[0].capability is SurfaceCapability.VALUE_REPLACE
    assert external[0].relationship_id
    assert external[0].relationship_type == RT.HYPERLINK


def test_inventory_has_stable_ids_and_value_hashes(tmp_path: Path) -> None:
    path = tmp_path / "stable.docx"
    _write_docx(path)

    first = DocxDocument.open(path).inventory()
    second = DocxDocument.open(path).inventory()

    first_identity = [(surface.surface_id, surface.value_sha256) for surface in first.surfaces]
    second_identity = [(surface.surface_id, surface.value_sha256) for surface in second.surfaces]
    assert first_identity == second_identity
    assert len({surface.surface_id for surface in first.surfaces}) == len(first.surfaces)


def test_inventory_reports_unknown_binary_part_as_incomplete(tmp_path: Path) -> None:
    path = tmp_path / "opaque.docx"
    _write_docx(path)
    _add_unknown_part(path)

    inventory = DocxDocument.open(path).inventory()

    assert inventory.coverage is InventoryCoverage.INCOMPLETE
    assert inventory.unknown_parts == ("word/embeddings/opaque1.bin",)
    part = next(part for part in inventory.parts if part.name == inventory.unknown_parts[0])
    assert part.understood is False
    assert part.is_xml is False
    assert part.sha256


def test_inventory_reports_orphan_unknown_part_as_incomplete(tmp_path: Path) -> None:
    path = tmp_path / "orphan.docx"
    _write_docx(path)
    buffer = BytesIO()
    with ZipFile(path) as source, ZipFile(buffer, "w", ZIP_DEFLATED) as output:
        for item in source.infolist():
            output.writestr(item, source.read(item.filename))
        output.writestr("word/orphan/opaque.bin", b"unreachable but still carried")
    path.write_bytes(buffer.getvalue())

    inventory = DocxDocument.open(path).inventory()

    assert inventory.coverage is InventoryCoverage.INCOMPLETE
    assert "word/orphan/opaque.bin" in inventory.unknown_parts


def test_internal_relationships_are_preserve_only(tmp_path: Path) -> None:
    path = tmp_path / "internal.docx"
    _write_docx(path)

    inventory = DocxDocument.open(path).inventory()
    internal = [
        surface
        for surface in inventory.surfaces
        if surface.kind is SurfaceKind.RELATIONSHIP and not surface.external
    ]

    assert internal
    assert {surface.capability for surface in internal} == {SurfaceCapability.PRESERVE_ONLY}


def test_inventory_exposes_html_text_surface(tmp_path: Path) -> None:
    path = tmp_path / "text-part.docx"
    _write_docx(path)
    buffer = BytesIO()
    with ZipFile(path) as source, ZipFile(buffer, "w", ZIP_DEFLATED) as output:
        for item in source.infolist():
            output.writestr(item, source.read(item.filename))
        output.writestr("word/afchunk.html", b'<a href="mailto:a@example.com">x</a>')
    path.write_bytes(buffer.getvalue())
    inventory = DocxDocument.open(path).inventory()
    surface = next(item for item in inventory.surfaces if item.part_name == "word/afchunk.html")
    assert surface.role == "text_part_content"
    assert "mailto:a@example.com" in surface.value
