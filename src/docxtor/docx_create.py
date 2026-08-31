from __future__ import annotations

from collections.abc import Iterable
from io import BytesIO
from pathlib import Path

from docx import Document

from .docx_inventory import InventoryCoverage, inventory_docx
from .docx_publication_marks import write_publication_bytes


class DocxCreationError(ValueError):
    pass


def create_docx_from_paragraphs(paragraphs: Iterable[str]) -> bytes:
    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    stream = BytesIO()
    document.save(stream)
    data = stream.getvalue()
    if inventory_docx(data).coverage is not InventoryCoverage.COMPLETE:
        raise DocxCreationError("created DOCX inventory coverage is incomplete")
    return data


def write_docx_from_paragraphs(path: str | Path, paragraphs: Iterable[str]) -> Path:
    return write_publication_bytes(path, create_docx_from_paragraphs(paragraphs))
