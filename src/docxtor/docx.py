from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import replace
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from docx import Document as PyDocxDocument
from docx.document import Document as DocxDocumentType
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.oxml import serialize_part_xml
from docx.text.paragraph import Paragraph

from .common import DOCX_MIME, DocumentBytes, DocumentError, output_filename
from .docx_comments import (
    _collect_comments,
    _ensure_thread_parts,
    _existing_comments_part,
    _restore_thread_sidecars,
    _unsupported_revision_reason,
)
from .docx_inline import (
    _advances_offset,
    _copy_segment,
    _index_at_visible_offset,
    _insert_visible,
    _replace_visible_range,
    _rpr_at,
    _split_visible_offset,
    _visible_len,
    _visible_text,
    paragraph_to_inline_segments,
    rebuild_paragraph_from_inline,
)
from .docx_inventory import DocxInventory, inventory_docx
from .docx_models import (
    AddressableComment,
    AddressableSpan,
    InlineSegment,
    InlineSegmentKind,
    SegmentReplacement,
    SpanRole,
    TextSegment,
    UnsupportedRevisionError,
)
from .docx_mutations import (
    SurfaceMutationResult,
    SurfaceReplacement,
    apply_surface_replacements,
)
from .docx_publish import PublishReceipt, publish_docx
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import ReviewMarkupInventory
from .docx_review_transaction import ReviewCommand, apply_review_batch
from .docx_stories import _ParaRef, index_stories
from .docx_units import _paragraph_spans, _paragraph_visible_text, _replace_plain_range

__all__ = [
    "AddressableComment",
    "AddressableSpan",
    "DocxDocument",
    "InlineSegment",
    "InlineSegmentKind",
    "SegmentReplacement",
    "SpanRole",
    "TextSegment",
    "UnsupportedRevisionError",
    "_advances_offset",
    "_copy_segment",
    "_index_at_visible_offset",
    "_insert_visible",
    "_replace_visible_range",
    "_rpr_at",
    "_split_visible_offset",
    "_visible_len",
    "_visible_text",
    "paragraph_to_inline_segments",
    "rebuild_paragraph_from_inline",
]


class DocxDocument:
    """DOCX editing surface backed by python-docx."""

    def __init__(
        self,
        doc: DocxDocumentType,
        segments: list[TextSegment],
        refs: list[_ParaRef],
        *,
        filename: str = "document.docx",
    ) -> None:
        self.filename = filename
        self._doc = doc
        self._segments = segments
        self._refs = refs  # index-aligned with segments
        self._spans: list[AddressableSpan] = []
        self._comments: list[AddressableComment] = []
        self._thread_parts: dict[str, tuple[bytes, str]] = {}
        self._note_parts: dict[str, tuple[Any, Any]] = {}
        self._source_bytes: bytes | None = None

    @classmethod
    def open(cls, path: str | Path) -> DocxDocument:
        path = Path(path)
        return cls.open_bytes(path.read_bytes(), filename=path.name)

    @classmethod
    def open_bytes(cls, data: bytes, *, filename: str = "document.docx") -> DocxDocument:
        try:
            doc = PyDocxDocument(BytesIO(data))
        except (BadZipFile, PackageNotFoundError, KeyError, ValueError) as exc:
            raise DocumentError(f"unreadable DOCX package: {exc}") from exc
        instance = cls._from_pydocx(doc, filename=filename)
        instance._source_bytes = data
        return instance

    @classmethod
    def _from_pydocx(
        cls, doc: DocxDocumentType, *, filename: str = "document.docx"
    ) -> DocxDocument:
        stories = index_stories(doc)
        instance = cls(doc=doc, segments=stories.segments, refs=stories.refs, filename=filename)
        instance._paragraphs_by_index = stories.paragraphs_by_index
        instance._paragraphs_by_container = stories.paragraphs_by_container
        instance._spans = instance._collect_spans()
        instance._comments = stories.comments
        instance._thread_parts = stories.thread_parts
        instance._note_parts = stories.note_parts
        return instance

    @property
    def segments(self) -> tuple[TextSegment, ...]:
        return tuple(self._segments)

    @property
    def texts(self) -> list[str]:
        return [s.text for s in self._segments]

    @property
    def spans(self) -> tuple[AddressableSpan, ...]:
        return tuple(self._spans)

    @property
    def comments(self) -> tuple[AddressableComment, ...]:
        return tuple(self._comments)

    def inventory(self) -> DocxInventory:
        """Return a domain-blind inventory of every value carried by the DOCX package."""
        payload = self._source_bytes if self._source_bytes is not None else self.to_bytes()
        return inventory_docx(payload)

    def apply_surface_replacements(
        self,
        replacements: list[SurfaceReplacement],
    ) -> SurfaceMutationResult:
        """Apply exact package-surface mutations and return verified output bytes."""
        payload = self._source_bytes if self._source_bytes is not None else self.to_bytes()
        return apply_surface_replacements(payload, replacements)

    def resolve_paragraph(self, container_id: str) -> Paragraph | None:
        """Resolve a python-docx Paragraph by stable container_id.

        container_id examples: "body:p:0", "body:p:17", "header:0:p:0",
        "table:0:r:1:c:2:p:0", "txbx:0:p:0".
        """
        if not hasattr(self, "_paragraphs_by_container"):
            return None
        return self._paragraphs_by_container.get(container_id)

    def resolve_paragraph_by_index(self, index: int) -> Paragraph | None:
        """Resolve by global paragraph index (counts every paragraph in order,
        including empty ones). Matches dike/posejdon locator contracts.
        """
        if not hasattr(self, "_paragraphs_by_index"):
            return None
        return self._paragraphs_by_index.get(index)

    def get_all_paragraphs(self) -> list[Paragraph]:
        """All paragraphs in document order (body, tables, headers, footers, boxes).
        Includes empty paragraphs to keep index stable.
        """
        if not hasattr(self, "_paragraphs_by_index") or not self._paragraphs_by_index:
            return []
        max_i = max(self._paragraphs_by_index.keys())
        return [
            self._paragraphs_by_index[i] for i in range(max_i + 1) if i in self._paragraphs_by_index
        ]

    def get_inline_segments(self, container_id: str) -> list[InlineSegment]:
        """Return the canonical rich InlineSegment decomposition for one paragraph.

        container_id examples: "body:p:0", "header:0:p:0", table cell variants.
        This is the bridge for review-specific layers to obtain the mechanical view
        (text + opaque with rpr/element) and then use the pure offset functions
        (_split_visible_offset, _insert_visible, _replace_visible_range, etc.)
        without reimplementing paragraph traversal or run decomposition.
        """
        para = self.resolve_paragraph(container_id)
        if para is None:
            return []
        return paragraph_to_inline_segments(para)

    def apply_targets(
        self,
        targets: list[dict[str, Any] | SegmentReplacement],
        *,
        strict: bool = False,
    ) -> None:
        """Apply a list of replacement targets.

        Each target can be:
          - SegmentReplacement
          - dict with keys: container_id or id, text, optional start_offset/end_offset
          - object with .container_id, .start_offset, .end_offset, .text (e.g. WriteTarget)

        This is the bridge for ReplacementPlan.write_targets.
        """
        normalized: list[SegmentReplacement] = []
        for target in targets:
            if isinstance(target, SegmentReplacement):
                normalized.append(target)
                continue

            if isinstance(target, dict):
                normalized.append(
                    SegmentReplacement(
                        container_id=target.get("container_id"),
                        id=target.get("id"),
                        span_id=target.get("span_id"),
                        text=str(target.get("text", "")),
                        start_offset=target.get("start_offset"),
                        end_offset=target.get("end_offset"),
                    )
                )
                continue

            # duck-type WriteTarget-like
            normalized.append(
                SegmentReplacement(
                    container_id=getattr(target, "container_id", None),
                    id=getattr(target, "segment_id", None),
                    span_id=getattr(target, "span_id", None),
                    text=str(getattr(target, "text", getattr(target, "replacement_text", ""))),
                    start_offset=getattr(target, "start_offset", None),
                    end_offset=getattr(target, "end_offset", None),
                )
            )
        self.apply_replacements(normalized, strict=strict)

    def to_markdown(self) -> str:
        blocks = [f"<!-- docxtor:{s.id} -->\n{s.text}" for s in self._segments]
        return "\n\n".join(blocks)

    def get_indexed_paragraphs(self) -> list[tuple[int, str, Paragraph]]:
        """Return every paragraph in document order with its stable identifiers.

        Returns list of (global_paragraph_index, container_id, python-docx.Paragraph).
        Includes empty paragraphs so that paragraph_index stays in sync with
        dike/posejdon anchor contracts (body + tables + headers/footers + boxes).
        This is the canonical source of addressing.
        """
        if not hasattr(self, "_paragraphs_by_index") or not self._paragraphs_by_index:
            return []
        max_i = max(self._paragraphs_by_index.keys())
        out: list[tuple[int, str, Paragraph]] = []
        for i in range(max_i + 1):
            if i not in self._paragraphs_by_index:
                continue
            para = self._paragraphs_by_index[i]
            # Resolve by paragraph identity. A body ID that happens to equal the
            # global index can belong to a different paragraph when a table comes first.
            cid = next(
                (
                    container_id
                    for container_id, candidate in self._paragraphs_by_container.items()
                    if candidate is para
                ),
                f"body:p:{i}",
            )
            out.append((i, cid, para))
        return out

    # ------------------------------------------------------------------
    # Placeholder replacement (mechanical, for reinjection)
    # ------------------------------------------------------------------

    def replace_placeholder(
        self,
        container_id: str,
        placeholder: str,
        replacement: str,
    ) -> None:
        """Mechanical: in the *current* text of the paragraph identified by container_id,
        find the first occurrence of placeholder and replace it with replacement.

        This is the non-domain part of reinjection flows.
        Offsets are computed on the live paragraph text; run splitting is handled internally.
        """
        para = self.resolve_paragraph(container_id)
        if para is None:
            raise ValueError(f"no paragraph for container_id {container_id!r}")

        current_text = _paragraph_visible_text(para)
        start = current_text.find(placeholder)
        if start < 0:
            raise ValueError(f"placeholder {placeholder!r} not found in segment {container_id}")
        end = start + len(placeholder)
        self.apply_replacements(
            [
                SegmentReplacement(
                    container_id=container_id,
                    text=replacement,
                    start_offset=start,
                    end_offset=end,
                )
            ],
            strict=True,
        )

    # ------------------------------------------------------------------
    # Replacement API (supports full + offset ranges via python-docx)
    # ------------------------------------------------------------------

    def apply_texts(self, texts: Iterable[str], *, strict: bool = False) -> None:
        texts = list(texts)
        if len(texts) != len(self._segments):
            raise ValueError(f"expected {len(self._segments)} segments, got {len(texts)}")
        self._require_supported_revisions()
        for i, txt in enumerate(texts):
            self._replace_full_segment(i, txt)

    def apply_replacements(
        self,
        replacements: list[SegmentReplacement],
        *,
        strict: bool = False,
    ) -> None:
        for replacement in replacements:
            if not isinstance(replacement, SegmentReplacement):
                raise TypeError("replacements must contain only SegmentReplacement instances")
        if replacements:
            self._require_supported_revisions()
        by_container = {r.container_id: i for i, r in enumerate(self._segments)}
        by_id = {r.id: i for i, r in enumerate(self._segments)}
        resolved: list[tuple[int, int | None, int | None, str]] = []
        for replacement in replacements:
            idx, start, end = self._resolve_replacement(replacement, by_container, by_id, strict)
            if idx is None:
                continue
            full = self._segments[idx].text
            if start is not None or end is not None:
                s = 0 if start is None else start
                e = len(full) if end is None else end
                if not 0 <= s < e <= len(full):
                    raise ValueError(
                        f"invalid replacement offsets for segment "
                        f"{self._refs[idx].container_id}: expected "
                        f"0 <= start < end <= {len(full)}, got {s}:{e}"
                    )
            resolved.append((idx, start, end, replacement.text))
        for idx, start, end, text in resolved:
            self._apply_to_paragraph(idx, text, start, end)

    def apply_markdown(self, markdown: str, *, strict: bool = True) -> None:
        import re as _re

        by_id = {
            m.group("id"): m.group("text").rstrip("\n")
            for m in _re.finditer(
                r"<!-- docxtor:(?P<id>s\d+) -->\n(?P<text>.*?)"
                r"(?=\n<!-- docxtor:s\d+ -->\n|\Z)",
                markdown,
                _re.DOTALL,
            )
        }
        if strict:
            expected = {s.id for s in self._segments}
            actual = set(by_id.keys())
            missing = sorted(expected - actual)
            unknown = sorted(actual - expected)
            if missing or unknown:
                raise ValueError(f"markdown marker mismatch; missing={missing} unknown={unknown}")

        if any(seg.id in by_id for seg in self._segments):
            self._require_supported_revisions()
        for i, seg in enumerate(self._segments):
            if seg.id in by_id:
                self._replace_full_segment(i, by_id[seg.id])

    # ------------------------------------------------------------------
    # Save / bytes
    # ------------------------------------------------------------------
    def publish(
        self,
        path: str | Path,
        *,
        validators: Iterable[Any] = (),
    ) -> PublishReceipt:
        """Publish through preservation, validation, and one atomic replace."""
        return publish_docx(self.to_bytes(), path, source=self._source_bytes, validators=validators)

    def save_docx(self, path: str | Path) -> None:
        self.publish(path)

    def review_inventory(self) -> ReviewMarkupInventory:
        return inventory_review_markup(self.to_bytes())

    def facts(self) -> Any:
        from .docx_facts import docx_facts

        payload = self._source_bytes if self._source_bytes is not None else self.to_bytes()
        return docx_facts(payload)

    def apply_review_batch(self, commands: Sequence[ReviewCommand]) -> None:
        receipt = apply_review_batch(self.to_bytes(), commands)
        replacement = self.open_bytes(receipt.data, filename=self.filename)
        self.__dict__.update(replacement.__dict__)

    def to_bytes(self) -> bytes:
        _ensure_thread_parts(self._doc.part.package, self._thread_parts)
        buf = BytesIO()
        self._doc.save(buf)
        return _restore_thread_sidecars(buf.getvalue(), self._thread_parts)

    def to_document_bytes(self) -> DocumentBytes:
        return DocumentBytes(
            filename=output_filename(self.filename, "docx"),
            content_type=DOCX_MIME,
            data=self.to_bytes(),
        )

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _require_supported_revisions(self) -> None:
        reason = _unsupported_revision_reason(self._doc)
        if reason is not None:
            raise UnsupportedRevisionError(
                f"unsupported revision form {reason!r}; refusing to write a partial artifact"
            )

    def _collect_spans(self) -> list[AddressableSpan]:
        spans: list[AddressableSpan] = []
        for ref in self._refs:
            spans.extend(_paragraph_spans(ref.paragraph, ref.container_id, ref.paragraph_index))
        return spans

    def _refresh_after_edit(self, index: int) -> None:
        ref = self._refs[index]
        new_text = _paragraph_visible_text(ref.paragraph)
        old = self._segments[index]
        self._segments[index] = replace(old, text=new_text)
        story = ref.part_name.split(":", 1)[0]
        note_part = self._note_parts.get(story)
        if note_part is not None:
            part, root = note_part
            part._blob = serialize_part_xml(root)
        self._spans = self._collect_spans()
        self._comments = _collect_comments(
            _existing_comments_part(self._doc),
            self._paragraphs_by_container,
            self._doc.part.package,
        )

    def _resolve_replacement(
        self,
        rep: SegmentReplacement,
        by_container: dict[str, int],
        by_id: dict[str, int],
        strict: bool,
    ) -> tuple[int | None, int | None, int | None]:
        if rep.span_id:
            span = next((s for s in self._spans if s.span_id == rep.span_id), None)
            if span is None:
                if strict:
                    raise ValueError(f"unknown replacement target: {rep.span_id}")
                return None, None, None
            idx = by_container.get(span.container_id)
            if idx is None:
                if strict:
                    raise ValueError(f"unknown replacement target: {rep.span_id}")
                return None, None, None
            if rep.start_offset is None and rep.end_offset is None:
                return idx, span.start_offset, span.end_offset
            local_start = 0 if rep.start_offset is None else rep.start_offset
            local_end = len(span.text) if rep.end_offset is None else rep.end_offset
            if not 0 <= local_start < local_end <= len(span.text):
                raise ValueError(
                    f"invalid replacement offsets for span {span.span_id}: "
                    f"expected 0 <= start < end <= {len(span.text)}, "
                    f"got {local_start}:{local_end}"
                )
            return idx, span.start_offset + local_start, span.start_offset + local_end
        if rep.container_id:
            idx = by_container.get(str(rep.container_id))
        elif rep.id:
            idx = by_id.get(str(rep.id))
        else:
            idx = None
        if idx is None and strict:
            raise ValueError(f"unknown replacement target: {rep.container_id or rep.id}")
        return idx, rep.start_offset, rep.end_offset

    def _replace_full_segment(self, index: int, text: str) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        full = _paragraph_visible_text(para)
        if full:
            _replace_plain_range(para._p, 0, len(full), text)
        elif para.runs:
            para.runs[0].text = text
        else:
            para.add_run(text)
        self._refresh_after_edit(index)

    def _apply_to_paragraph(
        self,
        index: int,
        replacement: str,
        start: int | None,
        end: int | None,
    ) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        full = _paragraph_visible_text(para)
        if start is None and end is None:
            self._replace_full_segment(index, replacement)
            return

        s = 0 if start is None else start
        e = len(full) if end is None else end
        if not 0 <= s < e <= len(full):
            raise ValueError(
                f"invalid replacement offsets for segment {ref.container_id}: "
                f"expected 0 <= start < end <= {len(full)}, got {s}:{e}"
            )

        if s == 0 and e == len(full):
            self._replace_full_segment(index, replacement)
            return

        _replace_plain_range(para._p, s, e, replacement)
        self._refresh_after_edit(index)
