from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from docx import Document as PyDocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from docxtor import (
    DocxDocument,
    PackageDispositionStatus,
    PackageMutation,
    PackageMutationError,
    PackageMutationKind,
    SegmentReplacement,
    SurfaceKind,
    apply_docx_transaction,
    apply_package_transaction,
)


def _document(path: Path) -> bytes:
    document = PyDocxDocument()
    document.add_paragraph("Visible person")
    document.core_properties.author = "Hidden Author"
    document.part.relate_to("mailto:hidden@example.test", RT.HYPERLINK, is_external=True)
    document.save(path)
    return path.read_bytes()


def _add_related_part(data: bytes) -> bytes:
    source = BytesIO(data)
    output = BytesIO()
    with ZipFile(source) as before, ZipFile(output, "w", ZIP_DEFLATED) as after:
        for item in before.infolist():
            payload = before.read(item.filename)
            if item.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"</Types>",
                    b'<Override PartName="/customXml/item42.xml" '
                    b'ContentType="application/xml"/></Types>',
                )
            if item.filename == "_rels/.rels":
                payload = payload.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdExtra" Type="urn:example" '
                    b'Target="customXml/item42.xml"/></Relationships>',
                )
            after.writestr(item, payload)
        after.writestr(
            "customXml/item42.xml", b"<?xml version='1.0'?><root><value>secret</value></root>"
        )
    return output.getvalue()


def test_inventory_exposes_qualified_context_and_opc_graph(tmp_path: Path) -> None:
    data = _add_related_part(_document(tmp_path / "source.docx"))

    inventory = DocxDocument.open_bytes(data).inventory()
    author = next(item for item in inventory.surfaces if item.value == "Hidden Author")

    assert author.element_qname
    assert author.ancestor_qnames
    assert author.role
    assert author.locator_version == "docxtor-surface-v1"
    assert "customXml/item42.xml" in inventory.graph.reachable_parts
    assert any(item.target_part == "customXml/item42.xml" for item in inventory.graph.relationships)


def test_remove_part_cascades_relationship_and_content_type(tmp_path: Path) -> None:
    data = _add_related_part(_document(tmp_path / "source.docx"))
    before = DocxDocument.open_bytes(data).inventory()
    part = next(item for item in before.parts if item.name == "customXml/item42.xml")

    receipt = apply_package_transaction(
        data,
        [
            PackageMutation(
                "remove-extra",
                PackageMutationKind.REMOVE_PART,
                part.name,
                part.sha256,
                cascade=True,
            )
        ],
    )

    assert "customXml/item42.xml" not in {item.name for item in receipt.inventory_after.parts}
    assert not any(
        item.target_part == "customXml/item42.xml"
        for item in receipt.inventory_after.graph.relationships
    )
    assert (
        next(item for item in receipt.parts if item.identity == part.name).status
        is PackageDispositionStatus.REMOVED
    )


def test_remove_relationship_has_global_surface_dispositions(tmp_path: Path) -> None:
    data = _document(tmp_path / "source.docx")
    before = DocxDocument.open_bytes(data).inventory()
    relationship = next(
        item for item in before.surfaces if item.kind is SurfaceKind.RELATIONSHIP and item.external
    )

    receipt = apply_package_transaction(
        data,
        [
            PackageMutation(
                "remove-link",
                PackageMutationKind.REMOVE_RELATIONSHIP,
                relationship.surface_id,
                relationship.value_sha256,
            )
        ],
    )

    removed = next(item for item in receipt.surfaces if item.identity == relationship.surface_id)
    assert removed.status is PackageDispositionStatus.REMOVED
    assert all(item.status is not PackageDispositionStatus.UNEXPECTED for item in receipt.surfaces)


def test_validator_blocks_result_before_return(tmp_path: Path) -> None:
    data = _document(tmp_path / "source.docx")

    def reject(_receipt) -> None:
        raise ValueError("consumer rejected")

    with pytest.raises(ValueError, match="consumer rejected"):
        apply_package_transaction(data, [], validators=[reject])


def test_combined_transaction_applies_strict_text_and_package_surface(tmp_path: Path) -> None:
    data = _document(tmp_path / "source.docx")
    document = DocxDocument.open_bytes(data)
    segment = next(item for item in document.segments if item.text == "Visible person")
    author = next(item for item in document.inventory().surfaces if item.value == "Hidden Author")

    receipt = apply_docx_transaction(
        data,
        text_targets=[SegmentReplacement(container_id=segment.container_id, text="Masked person")],
        package_mutations=[
            PackageMutation(
                "replace-author",
                PackageMutationKind.REPLACE_SURFACE,
                author.surface_id,
                author.value_sha256,
                "Anonymous",
            )
        ],
    )

    reopened = DocxDocument.open_bytes(receipt.data)
    assert "Masked person" in reopened.texts
    assert "Anonymous" in {item.value for item in reopened.inventory().surfaces}


def test_unknown_target_and_incomplete_source_fail_closed(tmp_path: Path) -> None:
    data = _document(tmp_path / "source.docx")
    with pytest.raises(PackageMutationError, match="unknown package part"):
        apply_package_transaction(
            data,
            [PackageMutation("bad", PackageMutationKind.REMOVE_PART, "missing.bin", "0" * 64)],
        )
