from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from enum import StrEnum
from functools import lru_cache
from pathlib import Path

import pymupdf as fitz

from .common import PDF_MIME, DocumentBytes, DocumentError, output_filename
from .docx import TextSegment

A4_WIDTH = 595
A4_HEIGHT = 842
PAGE_MARGIN = 48
TEXT_FONT_SIZE = 10
TEXT_LINE_HEIGHT = 12.5
# labels-style placeholders, e.g. [OSOBA_1], [PESEL_2], [FIRMA_Ą]
_BRACKETED_LABEL_RE = re.compile(r"\[[A-ZĄĆĘŁŃÓŚŹŻ][A-ZĄĆĘŁŃÓŚŹŻ0-9_]*\]")
UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/local/share/fonts/DejaVuSans.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
)


@dataclass(frozen=True, slots=True)
class _TextChange:
    source_start: int
    source_end: int
    source_text: str
    replacement_text: str


@dataclass(frozen=True, slots=True)
class _PdfLine:
    text: str
    rects: list[fitz.Rect]
    bbox: fitz.Rect


class PdfExtractionMode(StrEnum):
    TEXT_LAYER = "text_layer"


@dataclass
class PdfDocument:
    filename: str
    pages: list[str]
    source_bytes: bytes = field(default=b"", repr=False)
    source_pages: list[str] | None = None
    extraction_mode: PdfExtractionMode = PdfExtractionMode.TEXT_LAYER

    def __post_init__(self) -> None:
        if self.source_pages is None:
            self.source_pages = list(self.pages)

    @classmethod
    def open_bytes(cls, data: bytes, *, filename: str = "document.pdf") -> PdfDocument:
        try:
            pdf = fitz.open(stream=data, filetype="pdf")
            pages = [_page_text_with_rects(page)[0] for page in pdf]
        except Exception as error:
            raise DocumentError("Nie udało się odczytać PDF.") from error

        if not any(page.strip() for page in pages):
            raise DocumentError("PDF nie ma warstwy tekstowej. Ten plik wymaga OCR.")
        return cls(
            filename=filename,
            pages=list(pages),
            source_bytes=data,
            source_pages=list(pages),
        )

    @property
    def segments(self) -> tuple[TextSegment, ...]:
        return tuple(
            TextSegment(id=f"p{index}", text=text, part=f"page:{index}", index=index)
            for index, text in enumerate(self.pages)
            if text.strip()
        )

    @property
    def texts(self) -> list[str]:
        return [segment.text for segment in self.segments]

    def apply_texts(self, texts) -> None:
        texts = list(texts)
        segment_indexes = [segment.index for segment in self.segments]
        if len(texts) != len(segment_indexes):
            raise ValueError(f"expected {len(segment_indexes)} text segments, got {len(texts)}")
        for index, text in zip(segment_indexes, texts, strict=True):
            self.pages[index] = text

    def to_bytes(self) -> bytes:
        source_pages = self.source_pages or []
        if self.source_bytes and self.pages == source_pages:
            return self.source_bytes
        if not self.source_bytes:
            return _render_text_pdf(self.pages)
        if len(self.pages) != len(source_pages):
            raise DocumentError("Nie udało się zapisać PDF: zmieniła się liczba stron.")

        try:
            # Prefer layout-preserving edits on the original pages:
            # 1) in-place redaction/overlay for localized replacements
            # 2) per-page text rebuild when redaction is unsafe or structure changed
            # 3) full flowing rebuild only when edited text cannot fit the original pages
            pdf = fitz.open(stream=self.source_bytes, filetype="pdf")
            for page_index, (source_text, anonymized_text) in enumerate(
                zip(source_pages, self.pages, strict=True)
            ):
                if source_text == anonymized_text:
                    continue

                page = pdf[page_index]
                if not _requires_page_text_rebuild(
                    source_text, anonymized_text
                ) and _redact_page_changes(page, source_text, anonymized_text):
                    continue

                if not _replace_page_with_text(page, anonymized_text):
                    return _render_reflowed_text_pdf(self.pages)

            return pdf.tobytes(garbage=4, deflate=True)
        except DocumentError:
            raise
        except Exception as error:
            raise DocumentError("Nie udało się zapisać PDF.") from error

    def to_document_bytes(self) -> DocumentBytes:
        return DocumentBytes(
            filename=output_filename(self.filename, "pdf"),
            content_type=PDF_MIME,
            data=self.to_bytes(),
        )


def _redact_page_changes(page, source_text: str, anonymized_text: str) -> bool:
    changes = [
        change
        for change in _changed_text_spans(source_text, anonymized_text)
        if change.source_text.strip()
    ]
    if not changes:
        return True
    if _redact_page_changes_by_offsets(page, source_text, changes):
        return True

    added_redaction = False
    for change in changes:
        source = change.source_text.strip()
        if not source:
            continue
        if _unsafe_short_redaction_source(source):
            return False

        rects = _search_text_rects(page, source)
        if not rects:
            return False

        label = _redaction_label(change.replacement_text)
        for rect in rects:
            expanded = _expand_rect(page, rect)
            page.add_redact_annot(
                expanded,
                text=label,
                fill=(1, 1, 1),
                text_color=(0, 0, 0),
                fontsize=_redaction_font_size(expanded),
            )
            added_redaction = True

    if added_redaction:
        page.apply_redactions()
    return True


def _redact_page_changes_by_offsets(
    page,
    source_text: str,
    changes: list[_TextChange],
) -> bool:
    char_rects = _page_char_rects(page, source_text)
    if char_rects is None:
        return False

    added_redaction = False
    for change in changes:
        source = change.source_text.strip()
        if not source:
            continue
        if _unsafe_short_redaction_source(source):
            return False

        rects = _rects_for_source_range(
            source_text,
            char_rects,
            start=change.source_start,
            end=change.source_end,
            page=page,
        )
        if not rects:
            return False

        label = _redaction_label(change.replacement_text)
        for rect in rects:
            page.add_redact_annot(
                _expand_rect(page, rect),
                text=label,
                fill=(1, 1, 1),
                text_color=(0, 0, 0),
                fontsize=_redaction_font_size(rect),
            )
            added_redaction = True

    if added_redaction:
        page.apply_redactions()
    return True


def _changed_text_spans(source_text: str, anonymized_text: str) -> list[_TextChange]:
    matcher = SequenceMatcher(None, source_text, anonymized_text, autojunk=False)
    changes: list[_TextChange] = []
    for tag, source_start, source_end, replacement_start, replacement_end in matcher.get_opcodes():
        if tag in {"equal", "insert"}:
            continue
        changes.append(
            _TextChange(
                source_start=source_start,
                source_end=source_end,
                source_text=source_text[source_start:source_end],
                replacement_text=anonymized_text[replacement_start:replacement_end],
            )
        )
    return changes


def _requires_page_text_rebuild(source_text: str, target_text: str) -> bool:
    redaction_target = _looks_like_redaction_target(target_text)
    if len(target_text) > len(source_text) and not redaction_target:
        return True

    matcher = SequenceMatcher(None, source_text, target_text, autojunk=False)
    for tag, source_start, source_end, target_start, target_end in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "insert" and not redaction_target:
            return True

        source = source_text[source_start:source_end]
        replacement = target_text[target_start:target_end]
        if tag == "delete" and not redaction_target and source.strip():
            return True
        if tag == "replace" and _replacement_requires_reflow(
            source,
            replacement,
            redaction_target=redaction_target,
        ):
            return True
    return False


def _looks_like_redaction_target(text: str) -> bool:
    # Treat mask/angle placeholders and labels-style [TOKEN] markers as in-place
    # redaction targets so longer replacements do not force full-page reflow.
    return (
        "****" in text
        or ("<" in text and ">" in text)
        or _BRACKETED_LABEL_RE.search(text) is not None
    )


def _replacement_requires_reflow(
    source: str,
    replacement: str,
    *,
    redaction_target: bool,
) -> bool:
    if redaction_target:
        return False
    source_label = " ".join(source.split())
    replacement_label = " ".join(replacement.split())
    if not replacement_label:
        return False
    if len(replacement_label) > len(source_label):
        return True
    if any(ord(character) > 127 for character in replacement_label):
        return True
    return "\n" in replacement and "\n" not in source


def _page_char_rects(page, source_text: str) -> list[fitz.Rect | None] | None:
    raw_text, raw_rects = _page_raw_text_with_rects(page)
    if raw_text == source_text:
        return raw_rects

    char_rects: list[fitz.Rect | None] = [None] * len(source_text)
    matcher = SequenceMatcher(None, raw_text, source_text, autojunk=False)
    matched_chars = 0
    source_non_space = sum(not character.isspace() for character in source_text)
    for tag, raw_start, raw_end, source_start, source_end in matcher.get_opcodes():
        if tag != "equal":
            continue
        for raw_index, source_index in zip(
            range(raw_start, raw_end),
            range(source_start, source_end),
            strict=True,
        ):
            char_rects[source_index] = raw_rects[raw_index]
            if not source_text[source_index].isspace() and raw_rects[raw_index] is not None:
                matched_chars += 1

    if source_non_space and matched_chars / source_non_space < 0.9:
        return None
    return char_rects


def _page_raw_text_with_rects(page) -> tuple[str, list[fitz.Rect | None]]:
    return _page_text_with_rects(page)


def _page_text_with_rects(page) -> tuple[str, list[fitz.Rect | None]]:
    text_parts: list[str] = []
    rects: list[fitz.Rect | None] = []
    raw = page.get_text("rawdict")
    for block in raw.get("blocks", []):
        lines: list[_PdfLine] = []
        for line in block.get("lines", []):
            line_text_parts: list[str] = []
            line_rects: list[fitz.Rect] = []
            for span in line.get("spans", []):
                for char in span.get("chars", []):
                    line_text_parts.append(char.get("c", ""))
                    line_rects.append(fitz.Rect(char["bbox"]))
            if not line_text_parts:
                continue
            lines.append(
                _PdfLine(
                    text="".join(line_text_parts),
                    rects=line_rects,
                    bbox=fitz.Rect(line["bbox"]),
                )
            )
        for index, line in enumerate(lines):
            if index:
                previous = lines[index - 1]
                separator = " " if _is_soft_wrapped_line(previous, line, page) else "\n"
                text_parts.append(separator)
                rects.append(None)
            text_parts.append(line.text)
            rects.extend(line.rects)
        if lines:
            text_parts.append("\n")
            rects.append(None)
    return "".join(text_parts), rects


def _is_soft_wrapped_line(previous: _PdfLine, current: _PdfLine, page) -> bool:
    previous_text = previous.text.rstrip()
    if not previous_text:
        return False
    if previous_text.endswith((".", "!", "?", ":", ";")):
        return False

    page_right = page.rect.x1
    near_page_edge = previous.bbox.x1 >= page_right - 12
    same_left_edge = abs(previous.bbox.x0 - current.bbox.x0) <= 24
    long_previous_line = previous.bbox.width >= page.rect.width * 0.72

    return near_page_edge or (same_left_edge and long_previous_line)


def _rects_for_source_range(
    source_text: str,
    char_rects: list[fitz.Rect | None],
    *,
    start: int,
    end: int,
    page,
) -> list[fitz.Rect]:
    rects: list[fitz.Rect] = []
    for index in range(start, end):
        if index >= len(char_rects):
            return []
        rect = char_rects[index]
        if rect is None:
            if source_text[index].isspace():
                continue
            return []
        if source_text[index].isspace():
            continue
        rects.append(rect)

    if not rects:
        return []
    return _merge_line_rects(rects, page)


def _merge_line_rects(rects: list[fitz.Rect], page) -> list[fitz.Rect]:
    merged: list[fitz.Rect] = []
    for rect in sorted(rects, key=lambda item: (round(item.y0, 1), item.x0)):
        if not merged or not _same_text_line(merged[-1], rect):
            merged.append(fitz.Rect(rect))
            continue
        merged[-1].include_rect(rect)
    return [_expand_rect(page, rect) for rect in merged]


def _same_text_line(left: fitz.Rect, right: fitz.Rect) -> bool:
    tolerance = max(2.0, min(left.height, right.height) * 0.45)
    return abs(left.y0 - right.y0) <= tolerance or abs(left.y1 - right.y1) <= tolerance


def _unsafe_short_redaction_source(text: str) -> bool:
    alnum_count = sum(character.isalnum() for character in text)
    return 0 < alnum_count < 2


def _search_text_rects(page, text: str):
    for candidate in _search_candidates(text):
        rects = page.search_for(candidate)
        if rects:
            return rects

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) < 2:
        return []

    rects = []
    for line in lines:
        if _unsafe_short_redaction_source(line):
            return []
        line_rects = page.search_for(line)
        if not line_rects:
            return []
        rects.extend(line_rects)
    return rects


def _search_candidates(text: str) -> list[str]:
    candidates = [text.strip(), " ".join(text.split())]
    unique = []
    for candidate in candidates:
        if candidate and candidate not in unique:
            unique.append(candidate)
    return unique


def _redaction_label(text: str) -> str:
    label = " ".join(text.split())
    if not label:
        return ""
    if len(label) > 64 or any(ord(character) > 127 for character in label):
        return "****"
    return label


def _expand_rect(page, rect) -> fitz.Rect:
    expanded = fitz.Rect(rect)
    expanded.x0 = max(page.rect.x0, expanded.x0 - 0.75)
    expanded.y0 = max(page.rect.y0, expanded.y0 - 0.75)
    expanded.x1 = min(page.rect.x1, expanded.x1 + 0.75)
    expanded.y1 = min(page.rect.y1, expanded.y1 + 0.75)
    return expanded


def _redaction_font_size(rect: fitz.Rect) -> float:
    return max(5, min(10, rect.height * 0.7))


def _replace_page_with_text(page, text: str) -> bool:
    """Replace a page's text layer in place. Returns False when text does not fit."""
    page.add_redact_annot(page.rect, fill=(1, 1, 1))
    page.apply_redactions()
    return _insert_page_text(page, text)


def _render_text_pdf(pages: list[str]) -> bytes:
    pdf = fitz.open()
    for page_text in pages:
        page = pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
        _insert_page_text(page, page_text)
    return pdf.tobytes(garbage=4, deflate=True)


def _render_reflowed_text_pdf(pages: list[str]) -> bytes:
    text = "\n".join(page.rstrip("\n") for page in pages)
    return _render_flowing_text_pdf(text)


def _render_flowing_text_pdf(text: str) -> bytes:
    pdf = fitz.open()
    font = _text_font()
    kwargs = _text_insert_kwargs()
    page = pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
    x = page.rect.x0 + PAGE_MARGIN
    y = page.rect.y0 + PAGE_MARGIN + TEXT_FONT_SIZE
    bottom = page.rect.y1 - PAGE_MARGIN
    max_width = page.rect.width - (PAGE_MARGIN * 2)

    wrote_anything = False
    sections = text.replace("\r\n", "\n").replace("\r", "\n").split("\f")
    for section_index, section in enumerate(sections):
        if section_index:
            page = pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
            y = page.rect.y0 + PAGE_MARGIN + TEXT_FONT_SIZE
        for line in _wrap_text_lines(section, font=font, max_width=max_width):
            if y + TEXT_LINE_HEIGHT > bottom:
                page = pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                y = page.rect.y0 + PAGE_MARGIN + TEXT_FONT_SIZE
            if line:
                page.insert_text((x, y), line, **kwargs)
                wrote_anything = True
            y += TEXT_LINE_HEIGHT

    if not wrote_anything and pdf.page_count == 0:
        pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
    return pdf.tobytes(garbage=4, deflate=True)


def _wrap_text_lines(text: str, *, font: fitz.Font, max_width: float) -> list[str]:
    lines: list[str] = []
    for raw_line in text.split("\n"):
        if not raw_line:
            lines.append("")
            continue
        lines.extend(_wrap_text_line(raw_line, font=font, max_width=max_width))
    return lines


def _wrap_text_line(line: str, *, font: fitz.Font, max_width: float) -> list[str]:
    words = line.expandtabs(4).split()
    if not words:
        return [""]

    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if _text_width(candidate, font) <= max_width:
            current = candidate
            continue

        if current:
            lines.append(current)
            current = ""

        if _text_width(word, font) <= max_width:
            current = word
            continue

        split_word_lines = _split_long_word(word, font=font, max_width=max_width)
        lines.extend(split_word_lines[:-1])
        current = split_word_lines[-1]

    if current:
        lines.append(current)
    return lines


def _split_long_word(word: str, *, font: fitz.Font, max_width: float) -> list[str]:
    lines: list[str] = []
    current = ""
    for character in word:
        candidate = f"{current}{character}"
        if current and _text_width(candidate, font) > max_width:
            lines.append(current)
            current = character
            continue
        current = candidate
    if current:
        lines.append(current)
    return lines or [word]


def _text_width(text: str, font: fitz.Font) -> float:
    return font.text_length(text, fontsize=TEXT_FONT_SIZE)


def _insert_page_text(page, text: str) -> bool:
    """Insert page text. Returns True when all text fits the page box."""
    target = fitz.Rect(
        page.rect.x0 + PAGE_MARGIN,
        page.rect.y0 + PAGE_MARGIN,
        page.rect.x1 - PAGE_MARGIN,
        page.rect.y1 - PAGE_MARGIN,
    )
    # PyMuPDF: non-negative residual height means the full string fit.
    residual = page.insert_textbox(
        target,
        text or "",
        align=fitz.TEXT_ALIGN_LEFT,
        **_text_insert_kwargs(),
    )
    return residual >= 0


def _text_insert_kwargs() -> dict:
    kwargs = {
        "fontsize": TEXT_FONT_SIZE,
        "color": (0, 0, 0),
    }
    font_path = _unicode_font_path()
    if font_path:
        kwargs["fontfile"] = font_path
        kwargs["fontname"] = "docxtorunicode"
    return kwargs


def _text_font() -> fitz.Font:
    font_path = _unicode_font_path()
    if font_path:
        return fitz.Font(fontfile=font_path)
    return fitz.Font("helv")


@lru_cache(maxsize=1)
def _unicode_font_path() -> str | None:
    for path in UNICODE_FONT_CANDIDATES:
        if Path(path).exists():
            return path
    return None
