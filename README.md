# Docxtor

Extract editable text segments from documents, modify them in memory, and write
the changed text back as document bytes.

Docxtor is useful when you need one small interface for document text
round-trips across plain text files, DOCX files, and text-layer PDFs.

## Status

Early library. The public API exports 216 names. The public API is not stable yet.

## Public API

Import these names from `docxtor`. Groups describe their mechanical purpose.

| Group | Exported names |
| --- | --- |
| Version | `__version__` |
| Document types and loading | `DOCX_MIME`, `MD_MIME`, `PDF_MIME`, `TXT_MIME`, `DocumentBytes`, `DocumentError`, `DocumentKind`, `DetectedDocumentType`, `detect_document_type`, `document_to_bytes`, `load_document`, `PdfDocument`, `PdfExtractionMode`, `PlainTextDocument` |
| DOCX text and inline editing | `AddressableComment`, `AddressableSpan`, `DocxDocument`, `InlineSegment`, `InlineSegmentKind`, `SegmentReplacement`, `SpanRole`, `TextSegment`, `UnsupportedRevisionError`, `paragraph_to_inline_segments`, `rebuild_paragraph_from_inline` |
| Paragraph and run locators | `ParagraphLocator`, `ParagraphResolution`, `RunLocator`, `RunResolution` |
| Comment mutations | `CommentScopeBookmark`, `CommentAuthor`, `CommentMutationError`, `CommentMutationResult`, `CommentRange` |
| Document comparison | `compare_docx_documents`, `BlockPair`, `CommentAnchor`, `CommentChange`, `CommentChangeStatus`, `ComparisonCoverage`, `ComparisonDiagnostic`, `DiagnosticSeverity`, `DocumentBlock`, `DocumentComment`, `DocumentSpan`, `DocumentView`, `DocxDocumentComparison`, `FormattingChange`, `FormattingSpan`, `RevisionEvent`, `RevisionEventStatus`, `RevisionReference`, `RevisionResolutionEvidence`, `TextAnchor`, `TextChange`, `TextChangeKind` |
| Cross-reference diagnostics | `DOCX_CROSS_REFERENCE_PROFILE_ID`, `CrossReferenceDiagnostic`, `CrossReferenceReport`, `CrossReferenceSeverity`, `diagnose_docx_cross_references` |
| Mechanical facts and structural comparison | `ChangeKind`, `ContainerCoordinate`, `DocxComparison`, `DocxFactsSnapshot`, `DocxStructureSnapshot`, `FactChange`, `FactDiagnostic`, `FactsCoverage`, `NamedFact`, `PageBreakFact`, `PageLayoutFacts`, `ParagraphFact`, `RelationshipFact`, `StoryFact`, `TransformPolicy`, `UnreadablePartFact`, `compare_docx`, `docx_facts` |
| Package inventory and graph | `DocumentSurface`, `DocxInventory`, `InventoryCoverage`, `PackagePart`, `SurfaceCapability`, `SurfaceKind`, `SurfaceVisibility`, `inventory_docx`, `SURFACE_LOCATOR_VERSION`, `PackageGraph`, `PackageRelationship` |
| Surface mutations | `SurfaceDisposition`, `SurfaceDispositionStatus`, `SurfaceMutationError`, `SurfaceMutationResult`, `SurfaceReplacement`, `apply_surface_replacements` |
| OPC content-type validation | `DOCX_CONTENT_TYPES_PROFILE_ID`, `ContentTypeDiagnostic`, `ContentTypeSeverity`, `ContentTypeValidation`, `validate_docx_content_types` |
| OPC relationship validation | `DOCX_RELATIONSHIP_PROFILE_ID`, `RelationshipDiagnostic`, `RelationshipSeverity`, `RelationshipValidation`, `validate_docx_relationships` |
| Package reading and publication primitives | `DEFAULT_PACKAGE_LIMITS`, `PackageEntry`, `PackageError`, `PackageLimits`, `normalize_docx_timestamps`, `parse_package_xml`, `read_package_entries`, `restore_semantically_unchanged_xml_parts`, `write_package_atomically` |
| Package comparison | `PackageChangeReport`, `PackagePartChange`, `PackagePartChangeKind`, `compare_docx_packages` |
| Core keyword properties | `read_core_keywords`, `remove_core_keyword_values`, `set_core_keywords` |
| Publication receipts | `PublishError`, `PublishReceipt` |
| Physical review geometry | `PhysicalCommentAnchor`, `PhysicalCommentRange`, `PhysicalCommentSpan`, `PhysicalParagraphGeometry`, `PhysicalReviewGeometry`, `PhysicalTextOrigin`, `PhysicalTextInsertion`, `physical_span_for_semantic_range`, `project_docx_review_geometry`, `add_physical_comment_bytes`, `add_physical_comment_regions_bytes` |
| Review markup inventory and batches | `inventory_review_markup`, `CommentRevisionAssociation`, `OperationReceipt`, `OperationStatus`, `ReviewBatchReceipt`, `ReviewCoverage`, `ReviewDiagnostic`, `ReviewMarkupInventory`, `ReviewCommand`, `ReviewTransactionError`, `apply_review_batch` |
| Revision mutations | `RevisionAuthor`, `RevisionMutationError`, `RevisionMutationResult`, `RevisionPosition`, `RevisionRange` |
| Selective revision disposition | `dispose_revisions_bytes`, `inspect_revision_dispositions`, `RevisionDecision`, `RevisionDisposition`, `RevisionDispositionCoverage`, `RevisionDispositionDiagnostic`, `RevisionDispositionInventory`, `RevisionDispositionReceipt`, `RevisionDispositionResult`, `RevisionDispositionTarget`, `RevisionSelectionError` |
| Revision inventory and bulk cleanup | `AcceptRevisionsError`, `RejectRevisionsError`, `Revision`, `RevisionInventory`, `RevisionInventoryCoverage`, `RevisionKind`, `RevisionOperation`, `RevisionOperationError`, `RevisionOperationReceipt`, `accept_all_revisions_bytes`, `inventory_revisions_bytes`, `reject_all_revisions_bytes`, `remove_all_comments_bytes` |
| DOCX admission | `DocumentPackageKind`, `DocxAdmissionInspection`, `inspect_docx_admission` |
| Package and combined transactions | `CombinedTransactionReceipt`, `apply_docx_transaction`, `PackageDisposition`, `PackageDispositionStatus`, `PackageMutation`, `PackageMutationError`, `PackageMutationKind`, `PackageTransactionReceipt`, `apply_package_transaction` |
| DOCX creation | `DocxCreationError`, `create_docx_from_paragraphs` |
| Hyperlink fields | `flatten_link_fields`, `instruction_is_strippable_hyperlink_field` |
| Metadata inspection and sanitization | `MetadataInspection`, `inspect_docx_metadata`, `sanitize_docx_metadata` |
| Publication marks and appendices | `BodyAppendix`, `DocumentMark`, `PublicationMarkError`, `append_body_appendix`, `has_body_appendix`, `has_document_mark`, `remove_body_appendix`, `stamp_document_mark`, `write_publication_bytes` |
| Review projection | `DocxReviewProjection`, `ReviewNoteProjection`, `ReviewParagraphProjection`, `project_docx_for_review` |
| Physical review rendering | `PhysicalReviewComment`, `PhysicalReviewEdit`, `PhysicalReviewer`, `PhysicalReviewNote`, `PhysicalReviewPlan`, `PhysicalReviewRenderError`, `render_physical_clean`, `render_physical_review` |

## Supported Formats
| Format | Read | Write | Notes |
| --- | --- | --- | --- |
| TXT/Markdown/text files | yes | same text format when known | Decodes UTF-8, UTF-16, CP1250, or Latin-1 input. |
| DOCX | yes | DOCX | High-fidelity editing powered by `python-docx`, a third-party dependency (not Python's standard library and not a Microsoft product) for the .docx / WordprocessingML format. Stable `container_id` + `paragraph_index`, whole-segment and offset-based partial replacements via `apply_replacements`; partial replacements update existing `w:t` text nodes without splitting `w:r` runs or creating new formatting boundaries. |
| PDF | text layer only | PDF | Layout-preserving redaction/overlays when possible; per-page rebuild next; full text reflow only if content no longer fits. OCR is not bundled. |

`load_document` identifies the document kind from bytes first, then consults
the MIME type and file extension as secondary signals.

## Installation

From GitHub:

```bash
python -m pip install git+https://github.com/mikolaj92/Docxtor.git@v0.19.1
```

With `uv`:

```bash
uv add git+https://github.com/mikolaj92/Docxtor.git@v0.19.1
```

Pin the current release tag `v0.19.1` (matching `pyproject.toml`). For the
feature history of earlier releases, see the git tags and `docs/releases/`.
Avoid `v0.4.1`: it shipped version `0.4.0`; `v0.4.4` corrected the pin.

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
output = document_to_bytes(document)

output.filename      # input.anonimizowany.docx
output.content_type  # application/vnd.openxmlformats-officedocument.wordprocessingml.document
output.data          # bytes
```

Each replacement targets a stable segment identifier. Offset-based partial
replacements are available when a whole segment should not be rewritten.

### DOCX creation

`create_docx_from_paragraphs(paragraphs)` accepts an `Iterable[str]`, creates
one paragraph per string, and returns DOCX bytes. Incomplete inventory coverage
of the created document raises `DocxCreationError`, a `ValueError` subclass,
instead of returning bytes.

### Hyperlink instruction predicate

`instruction_is_strippable_hyperlink_field(instruction: str) -> bool` checks only
instruction text. It returns `True` when the text contains a whole-word
`HYPERLINK` or `LINK` token and one of `mailto:`, `http://`, `https://`, `file:`,
or `ftp://`. Both checks ignore case and search anywhere in the text.
URI matching uses substrings, so `sftp://example.com` also matches `ftp://`.
It returns `False` for a full internal TOC instruction match of
`HYPERLINK \l _Toc[0-9]+` or `LINK \l _Toc[0-9]+`, allowing surrounding
whitespace and optional quotes around the bookmark. Empty instructions, missing
field tokens, and missing recognized URI prefixes also return `False`.

```python
from docxtor import instruction_is_strippable_hyperlink_field

assert instruction_is_strippable_hyperlink_field('HYPERLINK "https://example.com"')
assert not instruction_is_strippable_hyperlink_field(r'HYPERLINK \l "_Toc123"')
```

The predicate does not open or modify a DOCX package. A `True` result does not
confirm URI safety or the validity of the complete field. It is a text check,
not a substitute for the structural checks performed by `flatten_link_fields`.

### Metadata inspection and sanitization

```python
from docxtor import MetadataInspection, inspect_docx_metadata, sanitize_docx_metadata

report: MetadataInspection = inspect_docx_metadata("copy.docx")
sanitize_docx_metadata("copy.docx", review_author="Reviewer", review_initials="RV")
```

Both functions accept `str | Path`. Inspection does not write the file.
`violations` is a tuple of messages; `clean` means that tuple is empty.
Inspection checks non-whitespace element text in `docProps/core.xml` and `docProps/app.xml`, and flags other `docProps/` parts.
Messages omit values, but element names come from input. Some unreadable inputs produce violations; other errors, including some corrupt ZIP errors, can raise exceptions.
Inspection excludes authorship in `word/`, some attributes, and full OOXML conformance. `clean` does not establish whole-document safety.

Sanitization changes the file in place and returns `None`. It clears XML text and tails in core/app.
It removes other `docProps/` parts and matching root package relationships and content-type overrides.
In `word/` XML, it replaces existing `author` attributes on `comment`, `ins`, `del`, and `person`, and existing comment `initials`; it removes `person` `userId`/`providerId`.
Defaults are `review_author="Reviewer"` and `review_initials="RV"`; absent attributes are not added.
The complete ZIP is written to a sibling temporary file before atomic replacement.
Malformed XML during sanitization raises before replacement, preserving the input and removing the temporary file. This is not full anonymization.

### Complete package inventory

`DocxDocument.inventory()` enumerates every ZIP/OPC member, including orphan
parts that `python-docx` cannot reach through relationships. It exposes neutral
`DocumentSurface` records for XML text, XML attributes, and relationship targets.
Each surface has a stable ID, value hash, visibility, and mechanical capability.
Qualified context uses `element_qname` and `attribute_qname` (namespace-qualified
XML names), `ancestor_qnames` (root-to-parent names), and a mechanical `role`.
`locator_version` identifies the locator scheme (`docxtor-surface-v1`).

`inventory()`, `facts()`, and `apply_surface_replacements()` read the current
in-memory XML, including prior text replacements, while retaining orphan parts
and their content types. The original opening bytes remain the publication
source. Surface mutations return verified bytes in `result.data`; they do not
update the document instance.

`inventory.coverage` is `InventoryCoverage.COMPLETE` only when no parts are
unknown or unreadable; `unknown_parts` and `unreadable_parts` identify gaps.
Consumers must reject incomplete coverage when their policy requires it.
Docxtor reports physical facts, not PII, legal relevance, or review meaning.

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

`mc:AlternateContent` story traversal selects exactly one `Choice` whose
`Requires` namespaces are in Docxtor's declared support set, otherwise the
`Fallback` branch. It never concatenates mutually exclusive branches.
`document.alternate_content_coverage` counts selected regions and reports
unsupported or unbound `Requires` namespaces; a non-empty limitation means
coverage is incomplete. Read projection is for inventory only: mutation of a
paragraph inside `AlternateContent` writes the corresponding paragraph in every
branch, and fails closed when branch paragraph structures cannot be aligned.

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

### Package and combined transactions

`apply_package_transaction(data, mutations, *, validators=(),
require_complete_coverage=True)` transforms DOCX bytes in memory, without writing
files. `PackageMutation` requires a unique `operation_id`, `kind`, `target_id`,
and `expected_sha256`; optional fields are `value`, `cascade`, and `content_type`.
`content_type` is currently unused. `PackageMutationKind` supports
`REPLACE_SURFACE` (string value), `REMOVE_SURFACE`, `REMOVE_RELATIONSHIP`,
`REPLACE_PART` (bytes value), and `REMOVE_PART`. Surface/relationship targets use
`surface_id` and `value_sha256`; part targets use `name` and `sha256`.
Removing a part with incoming relationships requires `cascade=True`, which also
removes those links; part removal cleans its relationship part and type override.

Hashes must match the source inventory. Complete inventory coverage is required
before and after by default. `PackageMutationError` reports rejected mutations,
incomplete coverage, unexpected changes, or unconfirmed operations; package-read
and validator errors can propagate separately. No receipt or intermediate result
is returned on error. Validators receive the complete package receipt before return.
`PackageTransactionReceipt` contains `data`, `inventory_before`, `inventory_after`,
`operations`, `surfaces`, `parts`, and the derived `unexpected` dispositions.

`apply_docx_transaction(data, *, text_targets=(), package_mutations=(),
validators=(), require_complete_coverage=True)` applies strict text replacements
first. Its package stage checks hashes and inventories against the post-text state.
`CombinedTransactionReceipt` contains `data`, normalized `text_targets`, and
`package` (the package receipt). Neither function updates an existing handle.

This example requires `input.docx` with a nonempty creator property and a text
segment. The two variants independently use the same original bytes; changing
body text leaves the selected metadata surface unchanged.

```python
from docxtor import (
    DocxDocument, PackageMutation, PackageMutationKind, SegmentReplacement,
    apply_docx_transaction, apply_package_transaction,
)

document = DocxDocument.open("input.docx")
data = document.to_bytes()
inventory = document.inventory()
author = next(s for s in inventory.surfaces if s.part_name == "docProps/core.xml"
              and s.element_qname == "{http://purl.org/dc/elements/1.1/}creator")
mutation = PackageMutation(
    "replace-author", PackageMutationKind.REPLACE_SURFACE,
    author.surface_id, author.value_sha256, value=author.value + " (edited)",
)
receipt = apply_package_transaction(data, [mutation])
assert not receipt.unexpected
package_bytes = receipt.data

segment = next(s for s in document.segments if s.text)
receipt = apply_docx_transaction(
    data,
    text_targets=[SegmentReplacement(container_id=segment.container_id, text="Updated text")],
    package_mutations=[mutation],
)
combined_bytes = receipt.data
assert not receipt.package.unexpected
```

For atomic file publication, pass `read_package_entries(receipt.data)` to
`write_package_atomically(destination, entries)`. It validates a sibling temporary
file, runs an optional `validate(path)` callback, then atomically replaces the
destination; failure before replacement preserves the existing destination.

### Text search across runs

`DocxDocument.find_text()` finds overlapping exact matches across Word run
boundaries while returning source-character offsets in a stable paragraph
segment. Matches stay within one paragraph, and `container_id` can narrow the
search. Opt into `match="normalized"` to fold case, common typographic
characters, and whitespace for search only; returned text and offsets still
refer to the original document. `replace_match()` applies a result only while
its paragraph still matches the captured snapshot, then delegates the edit to
Docxtor's offset-based replacement machinery.

```python
from docxtor import DocxDocument

document = DocxDocument.open("input.docx")
match = document.find_text("thirty days", match="normalized")[0]
document.replace_match(match, "business days")
```

Search stays within paragraphs; callers choose matches and paragraph coordinates.

### Typed paragraph and run locators

Consumers that need paragraph/run identity and values can stay independent of
`python-docx`. `paragraph_resolutions` includes empty paragraphs in stable global
order across body, table, header, and footer stories. Typed lookup returns
`None` for an unknown paragraph or run coordinate.

```python
from docxtor import DocxDocument, ParagraphLocator, RunLocator

document = DocxDocument.open("input.docx")
for paragraph in document.paragraph_resolutions:
    print(paragraph.identity.container_id, paragraph.paragraph_index, paragraph.value)
    for run in paragraph.runs:
        print(run.identity.run_index, run.value)

body = ParagraphLocator("body:p:0")
paragraph = document.resolve_paragraph_locator(body)
run = document.resolve_run_locator(RunLocator(body, 0))
```

The resolution DTOs contain only typed identity/value data. Physical traversal
and `python-docx` objects remain internal to Docxtor.

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

## Sole mechanical DOCX layer

`SegmentReplacement` targets a `container_id` with optional `start_offset` and
`end_offset`, or a nested `span_id`. `apply_replacements(..., strict=True)` raises
on unknown targets or bad offsets. Consumers delegate physical OOXML edits here.

### Neutral Word review markup

Docxtor owns the physical OOXML mechanics for comments and tracked changes. It
stays domain-blind: the caller supplies exact locators, ranges, author identity,
and the accept/reject decision.

Both renderers take `source` as a DOCX path (`str` or `Path`) or `bytes`, take
`output` as the destination path, and return the published `Path`.
`render_physical_review` takes a `PhysicalReviewPlan` and an optional keyword-only
`PhysicalReviewer` (defaulting to `PhysicalReviewer()`) to create review markup;
`render_physical_clean` takes a tuple of `PhysicalReviewEdit` values and applies
inline edits without new tracked-change markup, skipping edits with `new_paragraph=True`.
Unknown locators, invalid inline ranges, and mismatched `expected_text` raise
`PhysicalReviewRenderError` before publication. Source reading and opening precede
exception handling, so not all input errors are wrapped; neither renderer guarantees
atomicity of the entire call.

Text replacement and deletion through `render_physical_review` or
`render_physical_clean` preserve structurally empty inline content controls
(`w:sdt` with an empty `w:sdtContent`). Controls inside the changed range retain
source order at the leading boundary of the replacement after acceptance;
rejection restores their original positions among the original runs. Controls
at range boundaries stay outside the edit. Nonempty controls and other opaque
objects remain protected, even when they have no visible text.

For a whole-paragraph insertion rendered by `render_physical_review`, the new
paragraph inherits the anchor's explicit paragraph style and direct paragraph
formatting, except that direct paragraph `w:numPr`, section properties, identity
attributes, and tracked property-change snapshots are not copied. Style-level
numbering can still apply through the inherited paragraph style. The inserted
text uses direct run properties from the first text run outside deleted or
moved-from revisions that has no enabled direct `w:vanish` (`w:val="0"`,
`"false"`, and `"off"` disable it; a missing value enables it). Visibility
inherited through paragraph or character styles is not resolved. If that run
has no effective properties after revision history is removed, paragraph-mark
run properties are used when present; otherwise the paragraph style's character
defaults apply. When the anchor has no explicit paragraph style, the inserted
paragraph leaves its style unset so the document's default paragraph style
applies.

```python
from docxtor import (
    CommentAuthor,
    CommentRange,
    DocxDocument,
    RevisionAuthor,
    RevisionRange,
    accept_all_revisions_bytes,
)

doc = DocxDocument.open("input.docx")
author = CommentAuthor("Reviewer", "RV", "2026-01-01T00:00:00Z")
reviewer = RevisionAuthor("Reviewer", "2026-01-01T00:00:00Z")
added = doc.add_comment(
    CommentRange("body:p:0", 0, 5, expected_text="Hello"),
    "Check this range",
    author,
)
doc.update_comment(added.receipt.created_ids[0], "Revised note")
doc.replace_revision(
    RevisionRange("body:p:0", 6, 11, expected_text="world"),
    "there",
    reviewer,
)
doc.remove_comments({added.receipt.created_ids[0]})
doc.publish()
accepted = accept_all_revisions_bytes(doc.to_bytes(), drop_comments=False)
```

`DocxDocument` is the open DOCX handle. Comment add/update/remove and tracked
insert/delete/replace (`w:ins` / `w:del`) run on that same file.
Repeated `add_comment` calls on the same sentence, including a span that already
contains native comment markers, each add another Word comment
(`w:commentRangeStart` / `w:commentRangeEnd` / comment part).
Tracked insert/delete/replace wrap the named characters when that span sits
in the middle of an existing comment, revision, bookmark, or hyperlink, and
refuse instead of reporting success for a range they did not wrap.
`publish()` with no arguments writes back to the path passed to
`DocxDocument.open`. If the handle was opened from bytes and has no path,
`publish()` fails closed. Docxtor does not invent a second review markup format.

`inventory_review_markup()` and `inventory_revisions_bytes()` distinguish an
empty document from incomplete coverage. Unsupported structural revisions fail
closed. `apply_review_batch()` returns no intermediate bytes if any command
fails.

`RevisionInventory.tracking_enabled` reports enabled `w:trackRevisions` settings
throughout the package, including orphan or relocated XML settings parts. It
recognizes the OOXML on/off values. Malformed, duplicated, misplaced or
unsupported tracking settings make coverage incomplete. Bulk acceptance and
rejection require complete revision inventories before and after transformation.

`accept_all_revisions_bytes(data, drop_comments=False, disable_tracking=True)`
accepts the revisions and disables future change tracking. Its receipt records
the tracking state before and after the operation. The default preserves the
tracking setting. Disabling tracking leaves other settings, personal metadata,
bookmarks and unresolved comments intact. Callers decide when accepting all
revisions is appropriate for their publication workflow.

#### Source-bound comment regions

`DocxDocument.review_geometry()` inventories body paragraphs, tables, headers,
footers, user notes, and paragraph-bearing text boxes. Comment bodies remain
separate context stories. Separator notes are excluded. The geometry carries
an exact source SHA-256 and ordered spans for each physical story.

`physical_span_for_semantic_range()` maps effective text through typed
`PhysicalTextOrigin` records. Inserted text participates in the effective view.
Deleted text keeps its raw source coordinates. A mapped range can include
intervening deleted text without removing or accepting that revision.

`DocxDocument.add_physical_comment_regions(regions, text, author,
document_sha256=source_sha256)` creates one native comment for each story region.
Each region is a tuple of `PhysicalCommentSpan` values. The operation validates
all regions before mutation and adopts only the complete candidate. Call
`publish()` to atomically write the resulting handle.

The comments store their shared `regionSet` identity and `regionSetSize` in a
Docxtor namespace inside DOCX. `AddressableComment.region_set_id` and
`region_set_size` expose these values after reload. Missing members and
conflicting sizes make review inventory incomplete. No sidecar file is used.

Word ignores `commentReference` inside headers, footers, text boxes, footnotes,
and endnotes. This behavior is documented in
[MS-OI29500 §17.13.4.5](https://learn.microsoft.com/en-nz/openspecs/office_standards/ms-oi29500/84a35ef3-9c12-4b59-ab4c-4a88f520d06c).
New comments on these stories use a visible native body display anchor and a
native hyperlink to a hidden bookmark at the exact logical target. Notes use
their body reference. Text boxes use their body host. Headers and footers use
their section reference. No body text or paragraph identity is invented.

`PhysicalCommentAnchor.spans` exposes the logical target. `display_spans` exposes
the native body display range. `display_visible` distinguishes a body display
anchor from incoming non-body native comments, which remain unchanged.
`PhysicalReviewGeometry.scope_bookmarks` exposes validated `CommentScopeBookmark`
ownership. Missing links, duplicate identities, incomplete groups, and targets
that cross stories fail closed. Multi-story groups encode their identity and
member count in native bookmark names, including their body member.

`AddressableComment.text` contains the actual Word comment text, including its
scope hyperlink label. `user_note_text` contains the remark without that label
only when the provider has validated the native bridge. Updating one member
updates the logical region set and retains each native scope link. Cleanup
removes only proved owned bookmarks and reports them in
`OperationReceipt.removed_scope_bookmarks`. Existing user bookmarks stay intact.
These structural guarantees do not claim a completed live Word rendering check.

A native reply without markers inherits its parent's proved physical scope.
The parent graph rejects unresolved identities, cycles, and conflicting scopes.
Partial markers cannot inherit a scope. `PhysicalCommentAnchor.anchor_evidence`
reports the proved explicit or inherited anchor. The older
`AddressableComment.anchor_evidence` reports marker or parent-graph provenance
without changing its legacy locator and anchor-text fields.

`remove_comments()` expands a selection to its complete native region sets and
all descendant replies. It removes their range markers and thread identities in
one transaction. `OperationReceipt.removed_ids` lists the complete selection.
Unselected comment bodies, thread records, and native revisions remain intact.

`DocxDocument.map_review_spans(source_sha256, spans)` follows native mutation
receipts on the same open handle. It returns current raw envelopes for captured
source spans. A tracked deletion retains its raw range. A tracked insertion
expands the envelope at its proved source position. `PhysicalTextInsertion`
records carry the actual physical insertion offset, including insertions next
to existing deleted text. Unknown mutations and broken snapshot chains raise.
`apply_review_batch()` returns its typed receipts and records the same mapping.
This history is transient execution state and is not a second document format.

#### Selective disposition of existing revisions

`inspect_revision_dispositions(bytes)` inventories every XML OPC member,
including comment stories, alternative XML branches and orphan parts. Each
immutable target identifies the package part, namespace, revision kind, Word ID,
physical XPath and subtree hash. Its neutral payload text, paragraph XPath and
paragraph text provide context without deciding review meaning. The inventory
also identifies the exact input bytes and its complete set of records with
SHA-256 hashes.

Pass that inventory and explicit `RevisionDecision` records to
`dispose_revisions_bytes`. The caller chooses every decision; Docxtor does not
derive authority from authorship, text or residual Word markup.

```python
from docxtor import (
    RevisionDecision,
    RevisionDisposition,
    dispose_revisions_bytes,
    inspect_revision_dispositions,
    read_package_entries,
    write_package_atomically,
)

inventory = inspect_revision_dispositions(input_bytes)
# target_ids are an explicit selection made by the caller from this inventory.
result = dispose_revisions_bytes(
    input_bytes,
    inventory,
    [RevisionDecision(target_id, RevisionDisposition.ACCEPT) for target_id in target_ids],
)
assert all(receipt.resolved for receipt in result.receipts)
write_package_atomically("resolved.docx", read_package_entries(result.data))
```

Supported selections are transitional `w:ins` and `w:del` wrappers directly
inside a paragraph or hyperlink, containing ordinary runs with text, tabs and
line breaks. Accepting an insertion or rejecting a deletion retains its runs;
rejecting an insertion or accepting a deletion removes its runs. Restoring a
deletion converts only that wrapper's `w:delText` to `w:t`. Formatting and all
surrounding markup are preserved.

Two directly adjacent opposite-kind wrappers form a conservative atomic inline
replacement group, recorded by `group_id` and `required_target_ids`. This is a
physical adjacency rule, without authorship or provenance inference. Both
targets must be selected with the same disposition; larger mixed chains and
pairs containing an unsupported member are preserve-only. Independent supported
targets can receive mixed accept/reject decisions in one call, even in one
paragraph.

Move/range, paragraph-mark, property, nested, alternative-branch and opaque-content
revisions are inventoried as preserve-only. Selecting one fails closed. An
unknown/duplicate target, absent/duplicate Word identity, stale or modified
inventory, partial/conflicting replacement group, or incomplete XML coverage
fails before output is returned. Valid range start/end markers may share a Word
ID, while repeated endpoints or repeated wrapper IDs are ambiguous.

The result contains typed complete `before`/`after` inventories, per-selected-target
receipts, output bytes and their SHA-256, and the changed part names. Postflight
reopens the output and verifies the complete planned XML transformation, every
unselected revision subtree in source order, original OPC members and metadata,
and byte-exact untouched part payloads. Unsupported unselected revisions remain
intact. No intermediate or partially resolved output is returned. The existing
`write_package_atomically` publication API validates a staged package before one
atomic destination replacement; its optional `validate` callback can enforce
the caller's publication policy before replacement.

`publish()` serializes in memory, preserves semantically unchanged source XML,
normalizes ZIP timestamps, validates the complete package, runs optional
validators, and performs one atomic replace. The returned `PublishReceipt`
identifies the exact published bytes. A failure before replacement leaves an
existing destination byte-for-byte unchanged.

#### Optional Open XML SDK validation

The Python library does not depend on .NET or the Open XML SDK. To independently
validate a generated DOCX, install the pinned .NET SDK and run the opt-in
validator from `validation/openxml/`:

```bash
dotnet run --project validation/openxml/Docxtor.OpenXmlValidation.csproj -- output.docx
```

The validator targets the explicitly configured Office 2019 file-format profile
and emits structured findings (error ID/type, description, part URI, XPath when
available), with a bounded error count. Package-open and validator execution
failures are distinct from conformance findings and return non-zero. This is an
independent OOXML conformance check, not a rendering, business-semantics, or
complete Microsoft Office behavior guarantee; it never modifies the document.


### Complete mechanical facts and structural diff

`DocxDocument.facts()` and `compare_docx(before, after)` provide complete,
typed facts snapshots for package structure, OPC relationships, container
coordinates, fields, bookmarks, and stories.

```python
from docxtor import compare_docx, docx_facts

before_facts = docx_facts("before.docx")
after_facts = docx_facts("after.docx")

comparison = compare_docx(before_facts, after_facts)
if not comparison.compliant:
    print(comparison.changes)
```

`compare_docx` detects lost containers, broken relationships, part drift, and
unintended surface modifications without parsing raw XML in consumer code.
