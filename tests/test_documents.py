from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import fitz
import pytest
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from test_docx import write_simple_docx

from docxtor import (
    DOCX_MIME,
    MD_MIME,
    PDF_MIME,
    TXT_MIME,
    DocumentError,
    DocumentKind,
    PdfExtractionMode,
    SegmentReplacement,
    detect_document_type,
    document_to_bytes,
    load_document,
)

UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/local/share/fonts/DejaVuSans.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
)


def _pdf_bytes(*pages: str, font_name: str = "Helvetica") -> bytes:
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setFont(font_name, 12)
    for index, text in enumerate(pages):
        if index:
            pdf.showPage()
            pdf.setFont(font_name, 12)
        pdf.drawString(48, 760, text)
    pdf.save()
    return output.getvalue()


def _blank_pdf_bytes() -> bytes:
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.showPage()
    pdf.save()
    return output.getvalue()


def _multipage_anonymization_pdf_bytes(page_count: int = 30) -> bytes:
    """Dense multi-page fixture similar to anonymization round-trips."""
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    for index in range(page_count):
        if index:
            pdf.showPage()
        pdf.setFont("Helvetica", 11)
        page_no = index + 1
        lines = [
            f"Page {page_no} of {page_count}. Contract between Jan Kowalski and ACME Sp. z o.o.",
            f"PESEL 44051401359 appears on page {page_no}. Email: jan.kowalski@example.com",
            "Address: ul. Testowa 12, 00-001 Warszawa. Phone +48 514 222 333.",
            "Account PL61 1140 2004 0000 3102 1234 5678 remains confidential.",
            (
                "Additional body text for density. Lorem ipsum dolor sit amet, "
                "consectetur adipiscing elit."
            ),
            f"Closing clause on page {page_no} with reference ID REF-{page_no:04d}.",
        ]
        y = 800
        for line in lines:
            pdf.drawString(48, y, line)
            y -= 16
    pdf.save()
    return output.getvalue()


def _positioned_pdf_bytes(*lines: str) -> bytes:
    pdf = fitz.open()
    page = pdf.new_page(width=595, height=842)
    y = 100
    for line in lines:
        page.insert_text((48, y), line, fontsize=10)
        y += 13
    return pdf.tobytes()


def _pdf_text(data: bytes) -> str:
    pdf = fitz.open(stream=data, filetype="pdf")
    return _normalize_text("\n".join(page.get_text("text") or "" for page in pdf))


def _pdf_page_count(data: bytes) -> int:
    pdf = fitz.open(stream=data, filetype="pdf")
    return pdf.page_count


def _normalize_text(text: str) -> str:
    return text.replace("\xa0", " ")


def _unicode_font_name() -> str:
    font_path = next((path for path in UNICODE_FONT_CANDIDATES if Path(path).exists()), None)
    if font_path is None:
        pytest.skip("Unicode font unavailable for PDF fixture")

    font_name = "DocxtorTestUnicode"
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    return font_name


def test_detects_docx_from_bytes_before_metadata(tmp_path: Path) -> None:
    input_path = tmp_path / "input.docx"
    write_simple_docx(input_path)
    data = input_path.read_bytes()

    detection = detect_document_type("upload.bin", "application/octet-stream", data)
    document = load_document("upload.bin", "application/octet-stream", data)

    assert detection.kind == DocumentKind.DOCX
    assert detection.source == "signature"
    assert document.texts == ["Hello world", "Second paragraph", "Header text"]


def test_detects_pdf_from_bytes_before_metadata() -> None:
    data = _pdf_bytes("Jan Kowalski")

    detection = detect_document_type("upload.bin", "application/octet-stream", data)
    document = load_document("upload.bin", "application/octet-stream", data)

    assert detection.kind == DocumentKind.PDF
    assert detection.source == "signature"
    assert document.texts == ["Jan Kowalski\n"]


def test_rejects_unknown_binary_document() -> None:
    with pytest.raises(DocumentError, match="Nieobsługiwany typ dokumentu"):
        load_document("upload.bin", "application/octet-stream", b"\x00\x01\x02\x03")


def test_load_docx_document_and_write_docx_bytes(tmp_path: Path) -> None:
    input_path = tmp_path / "input.docx"
    write_simple_docx(input_path)

    document = load_document("input.docx", DOCX_MIME, input_path.read_bytes())
    document.apply_replacements(
        [
            SegmentReplacement(container_id=segment.container_id, text=text)
            for segment, text in zip(document.segments, ["One", "Two", "Three"], strict=True)
        ],
        strict=True,
    )
    output = document_to_bytes(document, "input.docx")

    assert output.filename == "input.anonimizowany.docx"
    assert output.content_type == DOCX_MIME
    with ZipFile(BytesIO(output.data)) as docx:
        assert "word/document.xml" in docx.namelist()


def test_load_pdf_document_and_write_pdf_bytes() -> None:
    document = load_document("input.pdf", PDF_MIME, _pdf_bytes("Jan Kowalski"))

    assert document.extraction_mode == PdfExtractionMode.TEXT_LAYER
    assert document.texts == ["Jan Kowalski\n"]
    document.apply_texts(["<PERSON>"])
    output = document_to_bytes(document, "input.pdf")

    assert output.filename == "input.anonimizowany.pdf"
    assert output.content_type == PDF_MIME
    assert output.data.startswith(b"%PDF")
    output_text = _pdf_text(output.data)
    assert "<PERSON>" in output_text
    assert "Jan Kowalski" not in output_text


def test_load_pdf_joins_soft_wrapped_lines_with_space() -> None:
    data = _positioned_pdf_bytes(
        "X" * 20 + " Dane testowe obejmuja adres e-mail x, numer telefonu +48 514 222 333, rachun",
        "1140 2004 0000 3102 1234 5678, pojazd KR 7MZ18.",
    )

    raw_text = fitz.open(stream=data, filetype="pdf")[0].get_text("text")
    document = load_document("input.pdf", PDF_MIME, data)

    assert "rachun\n1140" in raw_text
    assert "rachun 1140" in document.texts[0]
    assert "rachun\n1140" not in document.texts[0]


def test_pdf_write_preserves_polish_text_and_page_count() -> None:
    data = _pdf_bytes(
        "Dane nie są fikcyjne. Zażółć gęślą jaźń. Jan Kowalski PESEL 44051401359",
        font_name=_unicode_font_name(),
    )
    document = load_document("input.pdf", PDF_MIME, data)

    assert "Dane nie są fikcyjne" in _normalize_text(document.texts[0])
    assert "Zażółć gęślą jaźń" in _normalize_text(document.texts[0])

    anonymized = document.texts[0].replace("Jan Kowalski", "****").replace("44051401359", "****")
    document.apply_texts([anonymized])
    output = document_to_bytes(document, "input.pdf")

    output_text = _pdf_text(output.data)
    assert _pdf_page_count(output.data) == _pdf_page_count(data)
    assert "Dane nie są fikcyjne" in output_text
    assert "Zażółć gęślą jaźń" in output_text
    assert "Jan Kowalski" not in output_text
    assert "44051401359" not in output_text


def test_pdf_write_keeps_original_page_count() -> None:
    data = _pdf_bytes("Jan Kowalski", "Anna Nowak")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts(["<PERSON>\n", "<PERSON>\n"])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 2
    assert "<PERSON>" in output_text
    assert "Jan Kowalski" not in output_text
    assert "Anna Nowak" not in output_text


def test_pdf_anonymization_round_trip_preserves_page_count() -> None:
    """PDF→PDF anonymization keeps page count (no ReportLab-style reflow)."""
    page_count = 30
    data = _multipage_anonymization_pdf_bytes(page_count)
    document = load_document("fixture.pdf", PDF_MIME, data)

    assert len(document.segments) == page_count
    assert _pdf_page_count(data) == page_count

    anonymized_texts = [
        (
            text.replace("Jan Kowalski", "<PERSON>")
            .replace("44051401359", "****")
            .replace("jan.kowalski@example.com", "<EMAIL>")
            .replace("+48 514 222 333", "<PHONE>")
            .replace("PL61 1140 2004 0000 3102 1234 5678", "<ACCOUNT>")
        )
        for text in document.texts
    ]
    document.apply_texts(anonymized_texts)
    output = document_to_bytes(document, "fixture.pdf")
    output_text = _pdf_text(output.data)

    assert output.filename == "fixture.anonimizowany.pdf"
    assert output.content_type == PDF_MIME
    assert output.data.startswith(b"%PDF")
    assert _pdf_page_count(output.data) == page_count
    assert "Jan Kowalski" not in output_text
    assert "44051401359" not in output_text
    assert "jan.kowalski@example.com" not in output_text
    assert "+48 514 222 333" not in output_text
    assert "PL61 1140 2004 0000 3102 1234 5678" not in output_text
    assert "<PERSON>" in output_text or "****" in output_text
    # Page-local content should still map to the original page index.
    assert "Page 1 of 30" in fitz.open(stream=output.data, filetype="pdf")[0].get_text()
    assert "Page 30 of 30" in fitz.open(stream=output.data, filetype="pdf")[29].get_text()


def test_pdf_write_redacts_changed_occurrence_by_offset() -> None:
    data = _pdf_bytes("Jan Kowalski oraz Jan Kowalski")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("Jan Kowalski", "<PERSON>", 1)])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert "<PERSON>" in output_text
    assert output_text.count("Jan Kowalski") == 1


def test_pdf_write_redacts_bracketed_labels_in_place() -> None:
    """Longer [OSOBA_1]-style labels use in-place redaction (no reflow)."""
    # Source is shorter than the label — previously forced whole-document reflow.
    data = _pdf_bytes("Jan signed the contract with ACME.")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("Jan", "[OSOBA_1]")])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)
    # Redaction text may wrap inside a narrow source rect; join for matching.
    compact_text = re.sub(r"\s+", "", output_text)

    assert _pdf_page_count(output.data) == 1
    assert "[OSOBA_1]" in compact_text or "****" in output_text
    assert re.search(r"\bJan\b", output_text) is None
    assert "signed the contract with ACME" in output_text


def test_pdf_label_style_anonymization_preserves_page_count() -> None:
    """labels-style replacements keep page count like mask/angle redaction."""
    page_count = 30
    data = _multipage_anonymization_pdf_bytes(page_count)
    document = load_document("fixture.pdf", PDF_MIME, data)

    anonymized_texts = [
        (
            text.replace("Jan Kowalski", "[OSOBA_1]")
            .replace("44051401359", "[PESEL_1]")
            .replace("jan.kowalski@example.com", "[EMAIL_1]")
            .replace("+48 514 222 333", "[PHONE_1]")
            .replace("PL61 1140 2004 0000 3102 1234 5678", "[KONTO_1]")
        )
        for text in document.texts
    ]
    document.apply_texts(anonymized_texts)
    output = document_to_bytes(document, "fixture.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == page_count
    assert "Jan Kowalski" not in output_text
    assert "44051401359" not in output_text
    assert "jan.kowalski@example.com" not in output_text
    assert "+48 514 222 333" not in output_text
    assert "PL61 1140 2004 0000 3102 1234 5678" not in output_text
    assert "[OSOBA_1]" in output_text or "[PESEL_1]" in output_text or "****" in output_text
    assert "Page 1 of 30" in fitz.open(stream=output.data, filetype="pdf")[0].get_text()
    assert "Page 30 of 30" in fitz.open(stream=output.data, filetype="pdf")[29].get_text()


def test_pdf_write_rebuilds_page_when_text_is_inserted() -> None:
    data = _pdf_bytes("Pierwsze zdanie.")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts(
        [document.texts[0].replace("Pierwsze zdanie.", "Pierwsze zdanie. Drugie zdanie.")]
    )
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Pierwsze zdanie. Drugie zdanie." in output_text


def test_pdf_write_reflows_inserted_text_across_original_page_boundaries() -> None:
    data = _pdf_bytes("Poczatek.", "Tresc z oryginalnej drugiej strony.")
    document = load_document("input.pdf", PDF_MIME, data)
    inserted_text = " ".join(f"dodatkowy{index}" for index in range(1, 1_501))
    document.apply_texts(
        [
            document.texts[0].replace("Poczatek.", f"Poczatek. {inserted_text}"),
            document.texts[1],
        ]
    )
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) > _pdf_page_count(data)
    assert output_text.index("Poczatek") < output_text.index("Tresc z oryginalnej drugiej strony")
    assert "dodatkowy1500" in output_text


def test_pdf_write_rebuilds_page_when_replacement_is_longer() -> None:
    data = _pdf_bytes("Status: OK")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("OK", "bardzo dobrze")])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Status: bardzo dobrze" in output_text
    assert "Status: OK" not in output_text


def test_pdf_write_removes_deleted_text() -> None:
    data = _pdf_bytes("Alpha Beta Gamma")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("Beta ", "")])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Alpha" in output_text
    assert "Gamma" in output_text
    assert "Beta" not in output_text


def test_pdf_without_text_layer_requires_ocr() -> None:
    with pytest.raises(DocumentError, match="wymaga OCR"):
        load_document("scan.pdf", PDF_MIME, _blank_pdf_bytes())


def test_load_text_document_and_write_txt_bytes() -> None:
    document = load_document("input.txt", "text/plain", "Zażółć".encode("cp1250"))

    assert document.texts == ["Zażółć"]
    document.apply_texts(["<TEXT>"])
    output = document_to_bytes(document, "input.txt")

    assert output.filename == "input.anonimizowany.txt"
    assert output.content_type == TXT_MIME
    assert output.data == b"<TEXT>"


def test_load_markdown_document_and_write_markdown_bytes() -> None:
    document = load_document("notes.md", "", b"# Title\n\nOld")

    assert document.texts == ["# Title\n\nOld"]
    document.apply_texts(["# Title\n\nNew"])
    output = document_to_bytes(document, "notes.md")

    assert output.filename == "notes.anonimizowany.md"
    assert output.content_type == MD_MIME
    assert output.data == b"# Title\n\nNew"
