from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn

from .common import DocumentError


class DocxParagraphMutationOperations:
    """Physical paragraph mutations shared by the public DocxDocument handle.

    Reuse the handle's canonical physical projections and story indexing.
    This private mixin adds no separate public import surface.
    """

    filename: str
    _doc: Any
    _source_bytes: bytes | None

    def remove_inserted_paragraph(self, container_id: str, *, expected_text: str) -> None:
        """Remove an entirely inserted paragraph after exact physical text validation.

        Mixed source content and opaque payload outside insertion wrappers fail closed.
        Reindex surviving paragraphs, spans and stories after removal.
        """
        paragraph = self.resolve_paragraph(container_id)
        if paragraph is None:
            raise DocumentError("inserted paragraph locator does not resolve")
        spans = [span for span in self.spans if span.container_id == container_id]
        inserted = "".join(span.text for span in spans if span.role == "insertion")
        segments = self.get_inline_segments(container_id)
        if (
            not inserted
            or inserted != expected_text
            or any(span.role != "insertion" and span.text for span in spans)
            or any(
                segment.kind != "opaque"
                or segment.element is None
                or segment.element.tag != qn("w:ins")
                for segment in segments
            )
        ):
            raise DocumentError("paragraph is not the expected entirely inserted content")
        parent = paragraph._p.getparent()
        if parent is None or parent.tag != qn("w:body"):
            raise DocumentError("inserted paragraph is not attached to a supported body story")
        self._require_supported_revisions()
        parent.remove(paragraph._p)
        replacement = self._from_pydocx(self._doc, filename=self.filename)
        replacement._source_bytes = self._source_bytes
        self.__dict__.update(replacement.__dict__)
