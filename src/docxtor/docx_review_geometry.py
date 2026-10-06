"""Source-bound physical paragraph and native Word comment geometry."""

from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
from typing import Any

from docx import Document as PyDocxDocument
from docx.oxml.ns import qn

from .docx_alternate_content import select_alternate_branch
from .docx_inline import _visible_text, paragraph_to_inline_segments
from .docx_ns import _TEXT_NODE_TAGS, MC_NS, W_P
from .docx_package import PackageEntry, parse_package_xml, read_package_entries
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import ReviewCoverage, ReviewDiagnostic
from .docx_stories import index_stories
from .docx_xml import _is_text_box_container


@dataclass(frozen=True)
class PhysicalCommentSpan:
    """One exact source-character slice, including zero-width paragraph points."""

    locator: str
    start_offset: int
    end_offset: int
    expected_text: str


@dataclass(frozen=True)
class PhysicalCommentRange:
    """A contiguous physical range bound to the exact immutable source bytes."""

    document_sha256: str
    spans: tuple[PhysicalCommentSpan, ...]

    def __post_init__(self) -> None:
        spans = tuple(self.spans)
        if any(not isinstance(span, PhysicalCommentSpan) for span in spans):
            raise TypeError("physical comment ranges require typed physical spans")
        object.__setattr__(self, "spans", spans)


@dataclass(frozen=True)
class PhysicalParagraphGeometry:
    locator: str
    text: str
    paragraph_index: int | None
    part_name: str
    story_id: str
    story_kind: str
    order_index: int
    story_order_index: int
    is_heading: bool
    opaque_ranges: tuple[tuple[int, int], ...] = ()
    has_revisions: bool = False
    addressable: bool = True

    @property
    def span(self) -> PhysicalCommentSpan:
        """Full raw source span; consumers need not compute character offsets."""
        return PhysicalCommentSpan(self.locator, 0, len(self.text), self.text)


@dataclass(frozen=True)
class PhysicalCommentAnchor:
    comment_id: str
    start_locator: str | None
    end_locator: str | None
    spans: tuple[PhysicalCommentSpan, ...]
    document_sha256: str
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...] = ()

    @property
    def addressable(self) -> bool:
        return self.coverage is ReviewCoverage.COMPLETE and bool(self.spans)

    @property
    def physical_spans(self) -> tuple[PhysicalCommentSpan, ...]:
        return self.spans


@dataclass(frozen=True)
class PhysicalReviewGeometry:
    paragraphs: tuple[PhysicalParagraphGeometry, ...]
    comment_anchors: tuple[PhysicalCommentAnchor, ...]
    document_sha256: str
    coverage: ReviewCoverage
    diagnostics: tuple[ReviewDiagnostic, ...] = ()


@dataclass
class _GeometryParagraph:
    geometry: PhysicalParagraphGeometry
    element: Any
    part_paragraph_index: int


@dataclass
class _GeometryState:
    projection: PhysicalReviewGeometry
    entries: tuple[PackageEntry, ...]
    roots: dict[str, Any]
    paragraphs: dict[str, _GeometryParagraph]
    main_part_name: str


@dataclass(frozen=True)
class _Marker:
    kind: str
    paragraph: _GeometryParagraph | None
    offset: int | None
    order: int
    part_name: str


_MARKERS = {
    qn("w:commentRangeStart"): "start",
    qn("w:commentRangeEnd"): "end",
    qn("w:commentReference"): "reference",
}


def project_docx_review_geometry(data: bytes) -> PhysicalReviewGeometry:
    """Project every indexed paragraph, including empties, and proved native anchors.

    Coordinates use Docxtor's raw visible-text profile: inserted and deleted text,
    tabs and breaks are retained. No text search or semantic normalization occurs.
    Unsupported/unindexed stories and malformed anchors produce incomplete coverage.
    """
    return _read_geometry_state(data).projection


def physical_span_for_semantic_range(
    paragraph: PhysicalParagraphGeometry,
    semantic_text: str,
    start_offset: int,
    end_offset: int,
    expected_text: str,
) -> PhysicalCommentSpan:
    """Map the strict strip-only semantic view of an unrevised paragraph to source.

    This deliberately supports no normalization beyond leading/trailing whitespace.
    Revision-bearing paragraphs require a separate proved origin projection.
    """
    if not paragraph.addressable or paragraph.has_revisions:
        raise ValueError("paragraph has no supported semantic-to-physical coordinate mapping")
    if paragraph.text.strip() != semantic_text:
        raise ValueError("semantic text does not match the supported strip-only source profile")
    if (
        type(start_offset) is not int
        or type(end_offset) is not int
        or not 0 <= start_offset <= end_offset <= len(semantic_text)
        or semantic_text[start_offset:end_offset] != expected_text
    ):
        raise ValueError("semantic range does not match its captured paragraph text")
    leading = len(paragraph.text) - len(paragraph.text.lstrip())
    start, end = leading + start_offset, leading + end_offset
    if paragraph.text[start:end] != expected_text:
        raise ValueError("semantic range cannot be proved against the physical source")
    return PhysicalCommentSpan(paragraph.locator, start, end, expected_text)


def _read_geometry_state(data: bytes) -> _GeometryState:
    digest = sha256(data).hexdigest()
    entries: tuple[PackageEntry, ...] = ()
    roots: dict[str, Any] = {}
    try:
        entries = read_package_entries(data)
        document = PyDocxDocument(BytesIO(data))
        stories = index_stories(document)
        roots = {
            entry.name: parse_package_xml(entry.data, part_name=entry.name)
            for entry in entries
            if entry.name.startswith("word/") and entry.name.endswith(".xml")
        }
        main_part_name = str(document.part.partname).lstrip("/")
        part_roots = {
            root: str(part.partname).lstrip("/")
            for part in document.part.package.parts
            if (root := getattr(part, "element", None)) is not None
        }
        for part, root in stories.note_parts.values():
            part_roots[root] = str(part.partname).lstrip("/")
        raw_paragraphs = {name: tuple(root.iter(W_P)) for name, root in roots.items()}
        indexed_root_paragraphs: dict[Any, dict[Any, int]] = {}
        stable_indices: dict[Any, int] = {}
        for index, paragraph in stories.paragraphs_by_index.items():
            stable_indices.setdefault(paragraph._p, index)
        paragraphs: dict[str, _GeometryParagraph] = {}
        diagnostics = list(inventory_review_markup(data).diagnostics)
        if not stories.alternate_content_coverage.complete:
            diagnostics.append(
                ReviewDiagnostic(
                    "physical_alternate_content_incomplete",
                    "The alternate-content story projection has unsupported requirements.",
                )
            )
        raw_by_paragraph: dict[Any, list[_GeometryParagraph]] = {}
        story_counts: dict[str, int] = {}
        for locator, paragraph in stories.paragraphs_by_container.items():
            root = paragraph._p.getroottree().getroot()
            part_name = part_roots.get(root)
            if part_name is None or part_name not in roots:
                diagnostics.append(
                    ReviewDiagnostic("paragraph_part_unresolved", locator, part_name)
                )
                continue
            source = raw_paragraphs[part_name]
            loaded = indexed_root_paragraphs.get(root)
            if loaded is None:
                loaded = {element: index for index, element in enumerate(root.iter(W_P))}
                indexed_root_paragraphs[root] = loaded
            ordinal = loaded.get(paragraph._p)
            if ordinal is None or len(loaded) != len(source):
                diagnostics.append(
                    ReviewDiagnostic("paragraph_identity_unresolved", locator, part_name)
                )
                continue
            original = source[ordinal]
            segments = paragraph_to_inline_segments(paragraph)
            text = _visible_text(segments)
            _, marker_text = _marker_coordinates(original)
            addressable = marker_text == text
            if not addressable:
                diagnostics.append(
                    ReviewDiagnostic("paragraph_coordinate_profile_unresolved", locator, part_name)
                )
            cursor = 0
            opaque = []
            for segment in segments:
                if segment.kind == "opaque":
                    opaque.append((cursor, cursor + len(segment.text)))
                cursor += len(segment.text)
            story_kind = locator.split(":", 1)[0]
            story_id = _story_id(locator, part_name)
            ppr = original.find(qn("w:pPr"))
            style = ppr.find(qn("w:pStyle")) if ppr is not None else None
            style_id = style.get(qn("w:val"), "") if style is not None else ""
            physical = PhysicalParagraphGeometry(
                locator=locator,
                text=text,
                paragraph_index=stable_indices.get(paragraph._p),
                part_name=part_name,
                story_id=story_id,
                story_kind=story_kind,
                order_index=len(paragraphs),
                story_order_index=story_counts.get(story_id, 0),
                is_heading=style_id.startswith("Heading") or style_id == "Title",
                opaque_ranges=tuple(opaque),
                has_revisions=_has_revisions(original),
                addressable=addressable,
            )
            story_counts[story_id] = physical.story_order_index + 1
            record = _GeometryParagraph(physical, original, ordinal)
            paragraphs[locator] = record
            raw_by_paragraph.setdefault(original, []).append(record)
        _validate_paragraph_coverage(roots, raw_by_paragraph, diagnostics)
        markers: dict[str, list[_Marker]] = {}
        bodies: dict[str, int] = {}
        for part_name, root in roots.items():
            for element in root.iter(qn("w:comment")):
                comment_id = element.get(qn("w:id"))
                if comment_id is not None:
                    bodies[comment_id] = bodies.get(comment_id, 0) + 1
                else:
                    diagnostics.append(
                        ReviewDiagnostic("comment_body_without_id", "Missing w:id", part_name)
                    )
            offsets: dict[Any, int] = {}
            for record in paragraphs.values():
                if record.geometry.part_name == part_name:
                    offsets.update(_marker_offsets(record.element))
            for order, element in enumerate(root.iter()):
                kind = _MARKERS.get(element.tag)
                if kind is None:
                    continue
                comment_id = element.get(qn("w:id"))
                if comment_id is None:
                    diagnostics.append(
                        ReviewDiagnostic("comment_marker_without_id", kind, part_name)
                    )
                    continue
                owner = next((node for node in element.iterancestors() if node.tag == W_P), None)
                candidates = raw_by_paragraph.get(owner, [])
                record = (
                    candidates[0]
                    if len(candidates) == 1 and _marker_placement_valid(element, kind)
                    else None
                )
                markers.setdefault(comment_id, []).append(
                    _Marker(kind, record, offsets.get(element), order, part_name)
                )
        anchors = []
        for comment_id in sorted(bodies.keys() | markers.keys()):
            anchor = _project_anchor(
                comment_id,
                bodies.get(comment_id, 0),
                markers.get(comment_id, []),
                paragraphs,
                digest,
            )
            anchors.append(anchor)
            diagnostics.extend(anchor.diagnostics)
        projection = PhysicalReviewGeometry(
            paragraphs=tuple(record.geometry for record in paragraphs.values()),
            comment_anchors=tuple(anchors),
            document_sha256=digest,
            coverage=ReviewCoverage.INCOMPLETE if diagnostics else ReviewCoverage.COMPLETE,
            diagnostics=tuple(diagnostics),
        )
        return _GeometryState(projection, entries, roots, paragraphs, main_part_name)
    except Exception as exc:
        # Unreadable or ambiguous XML is never projected as an empty, complete document.
        projection = PhysicalReviewGeometry(
            (),
            (),
            digest,
            ReviewCoverage.INCOMPLETE,
            (ReviewDiagnostic("physical_geometry_unreadable", str(exc)),),
        )
        return _GeometryState(projection, entries, roots, {}, "")


def _story_id(locator: str, part_name: str) -> str:
    kind = locator.split(":", 1)[0]
    if kind in {"txbx", "comment", "footnote", "endnote"}:
        return f"{part_name}#{locator.rsplit(':p:', 1)[0]}"
    return part_name


def _selected_paragraphs(root: Any):
    def walk(node: Any):
        if node.tag == f"{{{MC_NS}}}AlternateContent":
            yield from walk(select_alternate_branch(node).selected)
            return
        if node.tag == W_P:
            yield node
        for child in node:
            yield from walk(child)

    yield from walk(root)


def _validate_paragraph_coverage(
    roots: dict[str, Any],
    by_paragraph: dict[Any, list[_GeometryParagraph]],
    diagnostics: list[ReviewDiagnostic],
) -> None:
    story_roots = {
        qn(f"w:{name}") for name in ("document", "hdr", "ftr", "comments", "footnotes", "endnotes")
    }
    for part_name, root in roots.items():
        if root.tag not in story_roots:
            continue
        story_positions: dict[str, int] = {}
        for paragraph in _selected_paragraphs(root):
            # Word's separator notes are mechanical infrastructure, not user paragraphs.
            if any(
                node.tag in {qn("w:footnote"), qn("w:endnote")}
                and node.get(qn("w:type")) in {"separator", "continuationSeparator"}
                for node in paragraph.iterancestors()
            ):
                continue
            records = by_paragraph.get(paragraph, [])
            if len(records) != 1:
                diagnostics.append(
                    ReviewDiagnostic(
                        "physical_paragraph_unaddressed"
                        if not records
                        else "physical_paragraph_ambiguous",
                        "A physical paragraph has no unique canonical locator.",
                        part_name,
                    )
                )
            else:
                record = records[0]
                expected = story_positions.get(record.geometry.story_id, 0)
                if record.geometry.story_order_index != expected:
                    diagnostics.append(
                        ReviewDiagnostic(
                            "physical_paragraph_order_unresolved",
                            record.geometry.locator,
                            part_name,
                        )
                    )
                story_positions[record.geometry.story_id] = expected + 1


def _story_host(paragraph: Any) -> Any:
    for ancestor in paragraph.iterancestors():
        if ancestor.tag in {
            qn("w:txbxContent"),
            qn("w:comment"),
            qn("w:footnote"),
            qn("w:endnote"),
        }:
            return ancestor
    return paragraph.getroottree().getroot()


def _marker_coordinates(paragraph: Any) -> tuple[dict[Any, int], str]:
    offsets: dict[Any, int] = {}
    cursor = 0
    text = []

    def walk(node: Any) -> None:
        nonlocal cursor
        if node is not paragraph and _is_text_box_container(node.tag):
            return
        if node.tag in {qn("w:pPr"), qn("w:rPr")}:
            return
        if node.tag == f"{{{MC_NS}}}AlternateContent":
            walk(select_alternate_branch(node).selected)
            return
        if node.tag in _MARKERS:
            offsets[node] = cursor
        if node.tag in _TEXT_NODE_TAGS:
            value = node.text or ""
            cursor += len(value)
            text.append(value)
            return
        if node.tag in {qn("w:tab"), qn("w:br"), qn("w:cr")}:
            cursor += 1
            text.append("\t" if node.tag == qn("w:tab") else "\n")
            return
        for child in node:
            walk(child)

    walk(paragraph)
    return offsets, "".join(text)


def _marker_offsets(paragraph: Any) -> dict[Any, int]:
    return _marker_coordinates(paragraph)[0]


def _has_revisions(paragraph: Any) -> bool:
    tags = {
        qn(f"w:{name}")
        for name in (
            "ins",
            "del",
            "moveFrom",
            "moveTo",
            "pPrChange",
            "rPrChange",
        )
    }
    return any(node.tag in tags for node in paragraph.iter()) or any(
        node.tag in tags for node in paragraph.iterancestors()
    )


def _canonical_comment_id(comment_id: str) -> bool:
    try:
        value = int(comment_id)
    except ValueError:
        return False
    return str(value) == comment_id and -(2**31) <= value < 2**31


def _marker_placement_valid(element: Any, kind: str) -> bool:
    parent = element.getparent()
    if kind == "reference" and (parent is None or parent.tag != qn("w:r")):
        return False
    for ancestor in element.iterancestors():
        if ancestor.tag == W_P:
            return True
        if ancestor.tag in {qn("w:pPr"), qn("w:rPr")}:
            return False
        if kind != "reference" and ancestor.tag == qn("w:r"):
            return False
    return False


def _project_anchor(
    comment_id: str,
    body_count: int,
    markers: list[_Marker],
    paragraphs: dict[str, _GeometryParagraph],
    digest: str,
) -> PhysicalCommentAnchor:
    issues = []
    starts = [marker for marker in markers if marker.kind == "start"]
    ends = [marker for marker in markers if marker.kind == "end"]
    references = [marker for marker in markers if marker.kind == "reference"]
    if not _canonical_comment_id(comment_id):
        issues.append(
            ReviewDiagnostic(
                "comment_geometry_noncanonical_id",
                f"Comment {comment_id} has an unsupported decimal identity.",
            )
        )
    if body_count != 1 or any(len(items) != 1 for items in (starts, ends, references)):
        issues.append(
            ReviewDiagnostic(
                "comment_geometry_nonunique",
                f"Comment {comment_id} has no unique body/start/end/reference.",
            )
        )
    start_locator = end_locator = None
    spans: tuple[PhysicalCommentSpan, ...] = ()
    if not issues:
        start, end, reference = starts[0], ends[0], references[0]
        start_locator = start.paragraph.geometry.locator if start.paragraph is not None else None
        end_locator = end.paragraph.geometry.locator if end.paragraph is not None else None
        if any(marker.paragraph is None or marker.offset is None for marker in markers):
            issues.append(
                ReviewDiagnostic(
                    "comment_geometry_unaddressed",
                    f"Comment {comment_id} has an unaddressed marker.",
                )
            )
        else:
            assert start.paragraph is not None and end.paragraph is not None
            assert reference.paragraph is not None
            first, last = start.paragraph.geometry, end.paragraph.geometry
            if any(not marker.paragraph.geometry.addressable for marker in markers):
                issues.append(
                    ReviewDiagnostic(
                        "comment_geometry_coordinate_profile_unresolved",
                        f"Comment {comment_id} has unsupported paragraph coordinates.",
                    )
                )
            if len({marker.paragraph.geometry.story_id for marker in markers}) != 1:
                issues.append(
                    ReviewDiagnostic(
                        "comment_geometry_cross_story",
                        f"Comment {comment_id} crosses physical stories.",
                    )
                )
            elif not start.order < end.order < reference.order:
                issues.append(
                    ReviewDiagnostic(
                        "comment_geometry_reversed",
                        f"Comment {comment_id} markers are not in source order.",
                    )
                )
            elif first.story_order_index > last.story_order_index:
                issues.append(
                    ReviewDiagnostic(
                        "comment_geometry_reversed", f"Comment {comment_id} endpoints are reversed."
                    )
                )
            else:
                assert start.offset is not None and end.offset is not None
                selected = [
                    record.geometry
                    for record in paragraphs.values()
                    if record.geometry.story_id == first.story_id
                    and first.story_order_index
                    <= record.geometry.story_order_index
                    <= last.story_order_index
                ]
                if any(not paragraph.addressable for paragraph in selected):
                    issues.append(
                        ReviewDiagnostic(
                            "comment_geometry_coordinate_profile_unresolved",
                            f"Comment {comment_id} contains unsupported paragraph coordinates.",
                        )
                    )
                root = start.paragraph.element.getroottree().getroot()
                all_paragraphs = list(root.iter(W_P))
                host = _story_host(start.paragraph.element)
                physical_between = [
                    paragraph
                    for paragraph in _selected_paragraphs(root)
                    if _story_host(paragraph) is host
                    and start.paragraph.part_paragraph_index
                    <= all_paragraphs.index(paragraph)
                    <= end.paragraph.part_paragraph_index
                ]
                if [paragraphs[item.locator].element for item in selected] != physical_between:
                    issues.append(
                        ReviewDiagnostic(
                            "comment_geometry_incomplete_paragraphs",
                            f"Comment {comment_id} has no complete ordered paragraph geometry.",
                        )
                    )
                spans = tuple(
                    PhysicalCommentSpan(
                        paragraph.locator,
                        start.offset if paragraph.locator == first.locator else 0,
                        end.offset if paragraph.locator == last.locator else len(paragraph.text),
                        paragraph.text[
                            start.offset if paragraph.locator == first.locator else 0 : end.offset
                            if paragraph.locator == last.locator
                            else len(paragraph.text)
                        ],
                    )
                    for paragraph in selected
                )
                if not spans or any(
                    not 0
                    <= span.start_offset
                    <= span.end_offset
                    <= len(paragraphs[span.locator].geometry.text)
                    for span in spans
                ):
                    issues.append(
                        ReviewDiagnostic(
                            "comment_geometry_invalid_offset",
                            f"Comment {comment_id} offsets are invalid.",
                        )
                    )
    return PhysicalCommentAnchor(
        comment_id,
        start_locator,
        end_locator,
        spans if not issues else (),
        digest,
        ReviewCoverage.INCOMPLETE if issues else ReviewCoverage.COMPLETE,
        tuple(issues),
    )
