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
| TXT/Markdown/text files | yes | same text format when known | Decodes UTF-8, UTF-16, CP1250, or Latin-1 input. |
| DOCX | yes | DOCX | High-fidelity editing powered by `python-docx` (the standard library for Microsoft's .docx format). Stable `container_id` + `paragraph_index`, whole-segment and offset-based partial replacements with run splitting. |
| PDF | text layer only | PDF | Layout-preserving redaction/overlays when possible; per-page rebuild next; full text reflow only if content no longer fits. OCR is not bundled. |

`load_document` identifies the document kind from bytes first, then consults
the MIME type and file extension as secondary signals.

## Installation

From GitHub:

```bash
python -m pip install git+https://github.com/mikolaj92/Docxtor.git
```

With `uv`:

```bash
uv add git+https://github.com/mikolaj92/Docxtor.git
```

Pin `v0.5.0` or later for complete package inventory, neutral surface
capabilities, and verified surface mutations. Earlier `v0.4.x` tags include
stable text/revision/comment addressing but not the complete inventory contract.
Tag `v0.4.1` still ships distribution version `0.4.0`; `v0.4.4` was the first
correctly versioned stable-addressing pin.

## Basic Usage

```python
from docxtor import DocxDocument, SegmentReplacement, document_to_bytes

document = DocxDocument.open("input.docx")
document.apply_replacements(
    [
        SegmentReplacement(
            container_id=segment.container_id,
            text=segment.text.replace("old", "new"),
        )
        for segment in document.segments
    ],
    strict=True,
)
output = document_to_bytes(document, "input.docx")

output.filename      # input.anonimizowany.docx
output.content_type  # application/vnd.openxmlformats-officedocument.wordprocessingml.document
output.data          # bytes
```

Each replacement targets a stable segment identifier. Offset-based partial
replacements are available when a whole segment should not be rewritten.

### Complete package inventory

`DocxDocument.inventory()` enumerates every ZIP/OPC member, including orphan
parts that `python-docx` cannot reach through relationships. It exposes neutral
`DocumentSurface` records for XML text, XML attributes, and relationship targets.
Each surface has a stable ID, value hash, visibility, and mechanical capability
such as `value_replace` or `preserve_only`.

```python
from docxtor import DocxDocument, InventoryCoverage

document = DocxDocument.open("input.docx")
inventory = document.inventory()
if inventory.coverage is not InventoryCoverage.COMPLETE:
    raise RuntimeError(
        f"uncovered DOCX parts: {inventory.unknown_parts + inventory.unreadable_parts}"
    )

for surface in inventory.surfaces:
    print(surface.surface_id, surface.kind, surface.capability)
```

Docxtor reports physical values and capabilities only. It does not decide
whether a value is PII, legally relevant, or review content. Consumers must
fail closed when their policy requires complete coverage.

Exact `SurfaceReplacement` mutations require the value hash observed during
inventory. Docxtor reopens the output and returns a disposition for every
requested surface. Unknown targets, stale hashes, preserve-only surfaces, and
unconfirmed writes fail before partial output is returned.

```python
from docxtor import SurfaceReplacement

surface = next(item for item in inventory.surfaces if item.external)
result = document.apply_surface_replacements([
    SurfaceReplacement(
        surface_id=surface.surface_id,
        expected_value_sha256=surface.value_sha256,
        value="https://example.invalid/redacted",
    )
])
assert not result.unresolved
result.data  # verified DOCX bytes
```

Text nested in `w:ins`, `w:del` / `w:delText`, and inline `w:hyperlink` is
addressable through `document.spans`. Each `AddressableSpan` has a stable
`span_id`, a mechanical `role` (`run`, `insertion`, `deletion`, `hyperlink`),
and the same character offsets as `TextSegment.text`. Replacing through
`SegmentReplacement(span_id=...)` or paragraph offsets writes into the
existing text nodes and keeps the wrappers, `w:rPr`, revision author/date/id,
and hyperlink anchor/relationship. Unsupported structural revisions such as
`w:moveFrom` or block-level `w:ins` raise `UnsupportedRevisionError` before
any partial output is written. Docxtor does not interpret review meaning.

User-authored footnotes and endnotes are mechanical stories with stable
`footnote:{id}:p:{n}` and `endnote:{id}:p:{n}` ids. Separator and continuation-
separator notes are not user segments, and missing note parts are never created.

Word comment bodies are addressable as `comment:{id}:p:{n}` segments and through
`document.comments`. Each `AddressableComment` carries the comment id, body text,
author/initials/date, the story locator and anchored range text when present, and
`parent_id` for replies named by `commentsExtended.xml`. Replacing comment text
writes into the existing comment paragraphs and keeps comment IDs, range anchors,
authorship attributes, and thread sidecar parts (`commentsExtended.xml`,
`commentsIds.xml`, `people.xml`). Empty comment parts and separator-only notes do
not become user-authored segments. Unknown comment targets fail before any edit.
ReviewKit remains responsible for review semantics.

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
- Stable `container_id` (e.g. `"body:p:0"`, `"header:0"`, `"table:0:r:0:c:0:p:0"`, `"txbx:0:p:0"`) and `paragraph_index`.
- `SegmentReplacement` for structured edits (full segment or sub-range by character offsets).
- `apply_replacements(..., strict=True)` — fail-closed on unknown targets or bad offsets.
- Run splitting for partial replacements inside paragraphs (keeps surrounding run formatting where possible).
- Structured `SegmentReplacement` records and `apply_replacements(..., strict=True)` are the canonical DOCX editing path.

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

# Or replace a whole segment by its stable identifier
doc.apply_replacements([
    SegmentReplacement(
        container_id="body:p:1",
        text="New second paragraph",
    ),
], strict=True)

doc.save_docx("output.docx")
```

`apply_replacements(..., strict=True)` raises on unknown targets or structural drift.

Example with nested revision / hyperlink spans:

```python
from docxtor import DocxDocument, SegmentReplacement, UnsupportedRevisionError

doc = DocxDocument.open("input.docx")
for span in doc.spans:
    print(span.span_id, span.role, span.text)
    # body:p:0:span:1 insertion Alice
    # body:p:0:span:3 hyperlink contract

doc.apply_replacements(
    [SegmentReplacement(span_id="body:p:0:span:1", text="Alina")],
    strict=True,
)
```
