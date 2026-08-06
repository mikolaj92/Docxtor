# Docxtor

Extract editable text segments from documents, modify them in memory, and write
the changed text back as document bytes.

Docxtor is useful when you need one small interface for document text
round-trips across plain text files, DOCX files, and text-layer PDFs.

## Status

Early library. The public API is small, but not stable yet.

## Supported Formats
| Format | Read | Write | Notes |
| --- | --- | --- | --- |
| TXT/Markdown/text files | yes | same text format when known | Decodes UTF-8, UTF-16, CP1250, and Latin-1 fallback. |
| DOCX | yes | DOCX | High-fidelity editing powered by `python-docx` (the standard library for Microsoft's .docx format). Stable `container_id` + `paragraph_index`, whole-segment and offset-based partial replacements with run splitting. |
| PDF | text layer only | PDF | Layout-preserving redaction/overlays when possible; per-page rebuild next; full text reflow only if content no longer fits. OCR is not bundled. |

`load_document` detects the document kind from bytes first, then falls back to
MIME type and file extension.

## Installation

From GitHub:

```bash
python -m pip install git+https://github.com/mikolaj92/Docxtor.git
```

With `uv`:

```bash
uv add git+https://github.com/mikolaj92/Docxtor.git
```

## Basic Usage

```python
from docxtor import DOCX_MIME, document_to_bytes, load_document

document = load_document("input.docx", DOCX_MIME, input_bytes)

updated_texts = [
    text.replace("old", "new")
    for text in document.texts
]

document.apply_texts(updated_texts)
output = document_to_bytes(document, "input.docx")

output.filename      # input.anonimizowany.docx
output.content_type  # application/vnd.openxmlformats-officedocument.wordprocessingml.document
output.data          # bytes
```

All loaded documents expose the same editing surface:

```python
document.texts
document.segments
document.apply_texts([...])
```

## Type Detection

```python
from docxtor import DocumentKind, detect_document_type

detection = detect_document_type(
    "upload.bin",
    "application/octet-stream",
    input_bytes,
)

if detection.kind == DocumentKind.DOCX:
    ...
```

## DOCX Markdown Bridge

DOCX documents can be exported to marker-based Markdown, edited, then applied
back to the original document structure.

```python
from docxtor import DocxDocument

document = DocxDocument.open("input.docx")
markdown = document.to_markdown()

# Keep docxtor markers intact.
edited_markdown = markdown.replace("old", "new")

document.apply_markdown(edited_markdown)
document.save_docx("output.docx")
```

## Development

```bash
uv sync --extra dev
uv run ruff check .
uv run pytest
```

## Limits

- DOCX output preserves archive entries and edits only selected Word XML story
  parts.
- Edited DOCX XML parts are rewritten, so output is not byte-identical.
- Replacement text inherits formatting through original text-node spans. Large
  length changes can move style boundaries.
- Document layout can change when replacement text length changes.
- PDF input must have a text layer. Scanned or image-only PDFs return an
  OCR-required error.
- PDF output is best-effort because PDF is fixed-layout, not an editable text
  document.
- Layout-preserving path (preferred): localized replacements are applied with
  redaction/overlay boxes on the original pages so page count and positions
  stay stable (typical anonymization labels such as `<PERSON>` / `****`).
- If a change cannot be located safely for redaction, that page is rebuilt in
  place (white-out + new text layer) instead of leaking the original text.
- Full flowing rebuild is a last resort when edited text no longer fits the
  original page boxes (large inserts/expansions). Page count may then shrink
  or grow, and original geometry is not preserved.
- Exact glyph metrics, fonts, columns, tables, headers/footers, and complex
  multi-column layouts are not guaranteed after any non-trivial edit.
- Non-ASCII redaction labels may be rendered as `****` when a matching Unicode
  font is unavailable for the overlay.

## License

MIT

## Sole mechanical DOCX layer (v0.3.0+)

**Docxtor is the single source of truth for mechanical DOCX manipulation.**

It owns:
- stable addressing (`container_id`, global `paragraph_index` counting empties)
- rich decomposition (`InlineSegment` with text + opaque, `rpr`, original element)
- pure offset primitives (`_split_visible_offset`, `_insert_visible`, `_replace_visible_range`, `_rpr_at`, `_visible_text`, ...)
- mutation (`apply_targets`, `apply_replacements`, `replace_placeholder`)
- access and rebuild (`get_inline_segments`, `paragraph_to_inline_segments`, `rebuild_paragraph_from_inline`)

Reviewkit (and Dike via it) delegates base paragraph/run/offset work to Docxtor and only layers review semantics (tracked changes as decision trace, comments, `apply_to_corrected`, `RenderIntegrityError`, policy, purity).

Temida consumers (posejdon_docs, dike_docs, anonimizator3000, ...) are thin adapters or high-level users. They contain **no** custom run splitting, offset math, or paragraph-mutation logic.

Docxtor 0.3.0+ uses `python-docx` (the standard, mature library for Microsoft's .docx / WordprocessingML format) as its internal DOCX engine.

Key features:
- Stable `container_id` (e.g. `"body:p:0"`, `"header:0"`, `"table:0:r:0:c:0:p:0"`) and `paragraph_index`.
- `SegmentReplacement` for structured edits (full segment or sub-range by character offsets).
- `apply_replacements(..., strict=True)` — fail-closed on unknown targets or bad offsets.
- Run splitting for partial replacements inside paragraphs (keeps surrounding run formatting where possible).
- Backward compatible: `apply_texts([...])`, `apply_markdown(...)`, and legacy dict form still work.

Example with offsets (similar to `WriteTarget` style used in Temida/Posejdon):

```python
from docxtor import DocxDocument, SegmentReplacement

doc = DocxDocument.open("input.docx")

print([s.container_id for s in doc.segments])
# ['body:p:0', 'body:p:1', 'header:0', ...]

# Partial replacement inside a segment (character offsets)
doc.apply_replacements([
    SegmentReplacement(
        container_id="body:p:0",
        text="REDACTED",
        start_offset=6,
        end_offset=12,
    )
], strict=True)

# Or full segment by id
doc.apply_replacements([
    {"id": "s1", "text": "New second paragraph"},
], strict=True)

doc.save_docx("output.docx")
```

`apply_replacements(..., strict=True)` raises on unknown targets or structural drift.

Legacy paths (`apply_texts`, `apply_markdown`, whole-segment dicts) remain supported for simple cases (e.g. anonimizator3000-style flows).
