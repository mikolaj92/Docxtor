from __future__ import annotations

from .detection import detect_document_type
from .docx import DocxDocument
from .engines import engine_for
from .pdf import PdfDocument
from .text import PlainTextDocument

Document = DocxDocument | PdfDocument | PlainTextDocument


def load_document(filename: str, content_type: str, data: bytes) -> Document:
    detection = detect_document_type(filename, content_type, data)
    return engine_for(detection.kind).open(detection, data)


def document_to_bytes(document: Document):
    """Serialize with the engine-owned filename and content-type contract."""
    return document.to_document_bytes()
