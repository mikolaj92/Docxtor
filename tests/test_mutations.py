from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document as PyDocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from docxtor import (
    DocxDocument,
    SurfaceDispositionStatus,
    SurfaceMutationError,
    SurfaceReplacement,
)


def _write_docx(path: Path) -> None:
    source = PyDocxDocument()
    source.add_paragraph("Visible person")
    source.core_properties.author = "Hidden Author"
    source.part.relate_to("mailto:hidden@example.test", RT.HYPERLINK, is_external=True)
    source.save(path)


def _replacement(document: DocxDocument, value: str, replacement: str) -> SurfaceReplacement:
    surface = next(item for item in document.inventory().surfaces if item.value == value)
    return SurfaceReplacement(
        surface_id=surface.surface_id,
        value=replacement,
        expected_value_sha256=surface.value_sha256,
    )


def test_mutates_xml_text_and_confirms_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "text.docx"
    _write_docx(path)
    document = DocxDocument.open(path)

    result = document.apply_surface_replacements(
        [_replacement(document, "Visible person", "Masked person")]
    )

    assert result.unresolved == ()
    assert result.dispositions[0].status is SurfaceDispositionStatus.REWRITTEN
    reopened = DocxDocument.open_bytes(result.data)
    assert "Masked person" in reopened.texts
    assert "Visible person" not in reopened.texts


def test_mutates_hidden_metadata_and_external_relationship(tmp_path: Path) -> None:
    path = tmp_path / "hidden.docx"
    _write_docx(path)
    document = DocxDocument.open(path)

    result = document.apply_surface_replacements(
        [
            _replacement(document, "Hidden Author", "Anonymous"),
            _replacement(
                document,
                "mailto:hidden@example.test",
                "mailto:anonymous@example.invalid",
            ),
        ]
    )

    values = {surface.value for surface in result.inventory.surfaces}
    assert "Anonymous" in values
    assert "mailto:anonymous@example.invalid" in values
    assert "Hidden Author" not in values
    assert "mailto:hidden@example.test" not in values


def test_rejects_stale_hash_before_writing(tmp_path: Path) -> None:
    path = tmp_path / "stale.docx"
    _write_docx(path)
    document = DocxDocument.open(path)
    replacement = _replacement(document, "Hidden Author", "Anonymous")

    with pytest.raises(SurfaceMutationError, match="changed before mutation"):
        document.apply_surface_replacements(
            [
                SurfaceReplacement(
                    surface_id=replacement.surface_id,
                    value=replacement.value,
                    expected_value_sha256="0" * 64,
                )
            ]
        )


def test_rejects_preserve_only_internal_relationship(tmp_path: Path) -> None:
    path = tmp_path / "internal.docx"
    _write_docx(path)
    document = DocxDocument.open(path)
    surface = next(
        item
        for item in document.inventory().surfaces
        if item.relationship_id and not item.external
    )

    with pytest.raises(SurfaceMutationError, match="not value-replaceable"):
        document.apply_surface_replacements(
            [
                SurfaceReplacement(
                    surface_id=surface.surface_id,
                    value="other.xml",
                    expected_value_sha256=surface.value_sha256,
                )
            ]
        )
