from __future__ import annotations

from docx.text.paragraph import Paragraph

from .docx_models import (
    ParagraphLocator,
    ParagraphResolution,
    RunLocator,
    RunResolution,
)
from .docx_units import _paragraph_visible_text


class DocxLocatorOperations:
    """Typed paragraph/run projections shared by ``DocxDocument``.

    The mixin stays inside Docxtor's physical layer. Public results contain no
    python-docx objects, so consumers do not need to traverse that object graph.
    """

    _paragraphs_by_index: dict[int, Paragraph]
    _paragraphs_by_container: dict[str, Paragraph]

    @property
    def paragraph_resolutions(self) -> tuple[ParagraphResolution, ...]:
        """Return typed projections for every globally indexed paragraph.

        Empty paragraphs are included so global paragraph coordinates remain
        stable even though empty paragraphs are not text ``segments``.
        """
        container_by_identity = {
            id(paragraph): container_id
            for container_id, paragraph in self._paragraphs_by_container.items()
        }
        return tuple(
            self._paragraph_resolution(
                index,
                container_by_identity[id(paragraph)],
                paragraph,
            )
            for index, paragraph in sorted(self._paragraphs_by_index.items())
        )

    def resolve_paragraph_locator(self, locator: ParagraphLocator) -> ParagraphResolution | None:
        """Resolve a typed paragraph locator without exposing python-docx."""
        paragraph = self._paragraphs_by_container.get(locator.container_id)
        if paragraph is None:
            return None
        paragraph_index = next(
            (
                index
                for index, candidate in self._paragraphs_by_index.items()
                if candidate is paragraph
            ),
            None,
        )
        if paragraph_index is None:
            return None
        return self._paragraph_resolution(paragraph_index, locator.container_id, paragraph)

    def resolve_run_locator(self, locator: RunLocator) -> RunResolution | None:
        """Resolve a typed run locator to its identity and current text value."""
        paragraph = self.resolve_paragraph_locator(locator.paragraph)
        if paragraph is None or not 0 <= locator.run_index < len(paragraph.runs):
            return None
        return paragraph.runs[locator.run_index]

    @staticmethod
    def _paragraph_resolution(
        paragraph_index: int,
        container_id: str,
        paragraph: Paragraph,
    ) -> ParagraphResolution:
        locator = ParagraphLocator(container_id)
        return ParagraphResolution(
            identity=locator,
            value=_paragraph_visible_text(paragraph),
            paragraph_index=paragraph_index,
            runs=tuple(
                RunResolution(identity=RunLocator(locator, index), value=run.text)
                for index, run in enumerate(paragraph.runs)
            ),
        )
