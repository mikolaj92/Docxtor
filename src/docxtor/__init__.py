from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from io import BytesIO
from zipfile import ZipFile

from lxml import etree

__version__ = "0.8.13"

from .common import (
    DOCX_MIME,
    MD_MIME,
    PDF_MIME,
    TXT_MIME,
    DocumentBytes,
    DocumentError,
    DocumentKind,
)
from .detection import DetectedDocumentType, detect_document_type
from .docx import (
    AddressableComment,
    AddressableSpan,
    DocxDocument,
    InlineSegment,
    InlineSegmentKind,
    SegmentReplacement,
    SpanRole,
    TextSegment,
    UnsupportedRevisionError,
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
from .docx_comment_mutations import (
    CommentAuthor,
    CommentMutationError,
    CommentMutationResult,
    CommentRange,
    add_comment,
    remove_comments,
)
from .docx_facts import (
    ChangeKind,
    ContainerCoordinate,
    DocxComparison,
    DocxFactsSnapshot,
    DocxStructureSnapshot,
    FactChange,
    FactDiagnostic,
    FactsCoverage,
    NamedFact,
    PageBreakFact,
    PageLayoutFacts,
    ParagraphFact,
    RelationshipFact,
    StoryFact,
    TransformPolicy,
    UnreadablePartFact,
    compare_docx,
    docx_facts,
)
from .docx_inventory import (
    DocumentSurface,
    DocxInventory,
    InventoryCoverage,
    PackagePart,
    SurfaceCapability,
    SurfaceKind,
    SurfaceVisibility,
    inventory_docx,
)
from .docx_mutations import (
    SurfaceDisposition,
    SurfaceDispositionStatus,
    SurfaceMutationError,
    SurfaceMutationResult,
    SurfaceReplacement,
    apply_surface_replacements,
)
from .docx_package import (
    DEFAULT_PACKAGE_LIMITS,
    PackageEntry,
    PackageError,
    PackageLimits,
    normalize_docx_timestamps,
    parse_package_xml,
    read_package_entries,
    restore_semantically_unchanged_xml_parts,
    write_package_atomically,
)
from .docx_properties import (
    read_core_keywords,
    remove_core_keyword_values,
    set_core_keywords,
)
from .docx_publish import PublishError, PublishReceipt, publish_docx
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import (
    CommentRevisionAssociation,
    OperationReceipt,
    OperationStatus,
    ReviewBatchReceipt,
    ReviewCoverage,
    ReviewDiagnostic,
    ReviewMarkupInventory,
)
from .docx_review_transaction import (
    ReviewCommand,
    ReviewTransactionError,
    apply_review_batch,
)
from .docx_revision_mutations import (
    RevisionAuthor,
    RevisionMutationError,
    RevisionMutationResult,
    RevisionPosition,
    RevisionRange,
    delete_revision,
    insert_revision,
    mark_paragraph_revision,
    replace_revision,
)
from .docx_revisions import (
    AcceptRevisionsError,
    RejectRevisionsError,
    Revision,
    RevisionInventory,
    RevisionInventoryCoverage,
    RevisionKind,
    RevisionOperation,
    RevisionOperationError,
    RevisionOperationReceipt,
    accept_all_revisions_bytes,
    inventory_revisions_bytes,
    reject_all_revisions_bytes,
)
from .loader import document_to_bytes, load_document
from .pdf import PdfDocument, PdfExtractionMode
from .text import PlainTextDocument

__all__ = [
    "__version__",
    "ChangeKind",
    "ContainerCoordinate",
    "DocxComparison",
    "FactChange",
    "FactDiagnostic",
    "FactsCoverage",
    "NamedFact",
    "PageBreakFact",
    "PageLayoutFacts",
    "ParagraphFact",
    "RelationshipFact",
    "StoryFact",
    "DocxFactsSnapshot",
    "DocxStructureSnapshot",
    "TransformPolicy",
    "UnreadablePartFact",
    "compare_docx",
    "docx_facts",
    "read_core_keywords",
    "remove_core_keyword_values",
    "set_core_keywords",
    "RevisionAuthor",
    "RevisionMutationError",
    "RevisionMutationResult",
    "RevisionPosition",
    "RevisionRange",
    "delete_revision",
    "insert_revision",
    "mark_paragraph_revision",
    "replace_revision",
    "CommentAuthor",
    "CommentMutationError",
    "CommentMutationResult",
    "CommentRange",
    "add_comment",
    "remove_comments",
    "PublishError",
    "PublishReceipt",
    "publish_docx",
    "AcceptRevisionsError",
    "RejectRevisionsError",
    "Revision",
    "RevisionInventory",
    "RevisionInventoryCoverage",
    "RevisionKind",
    "RevisionOperation",
    "RevisionOperationError",
    "RevisionOperationReceipt",
    "accept_all_revisions_bytes",
    "inventory_revisions_bytes",
    "reject_all_revisions_bytes",
    "inventory_review_markup",
    "CommentRevisionAssociation",
    "OperationReceipt",
    "OperationStatus",
    "ReviewBatchReceipt",
    "ReviewCoverage",
    "ReviewDiagnostic",
    "ReviewMarkupInventory",
    "ReviewCommand",
    "ReviewTransactionError",
    "apply_review_batch",
    "DOCX_MIME",
    "MD_MIME",
    "PDF_MIME",
    "TXT_MIME",
    "DocumentBytes",
    "DocumentError",
    "DocumentKind",
    "DetectedDocumentType",
    "AddressableComment",
    "AddressableSpan",
    "DocumentSurface",
    "DocxDocument",
    "DocxInventory",
    "InventoryCoverage",
    "DEFAULT_PACKAGE_LIMITS",
    "PackageEntry",
    "PackageError",
    "PackageLimits",
    "PackagePart",
    "SurfaceCapability",
    "SurfaceDisposition",
    "SurfaceDispositionStatus",
    "SurfaceKind",
    "SurfaceMutationError",
    "SurfaceMutationResult",
    "SurfaceReplacement",
    "SurfaceVisibility",
    "inventory_docx",
    "InlineSegment",
    "InlineSegmentKind",
    "PdfExtractionMode",
    "PdfDocument",
    "PlainTextDocument",
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
    "apply_surface_replacements",
    "detect_document_type",
    "document_to_bytes",
    "load_document",
    "normalize_docx_timestamps",
    "parse_package_xml",
    "read_package_entries",
    "restore_semantically_unchanged_xml_parts",
    "write_package_atomically",
    "CombinedTransactionReceipt",
    "PackageDisposition",
    "PackageDispositionStatus",
    "PackageMutation",
    "PackageMutationError",
    "PackageMutationKind",
    "PackageTransactionReceipt",
    "apply_docx_transaction",
    "apply_package_transaction",
    "PackageGraph",
    "PackageRelationship",
    "BodyAppendix",
    "DocumentMark",
    "PublicationMarkError",
    "append_body_appendix",
    "has_body_appendix",
    "has_document_mark",
    "remove_body_appendix",
    "stamp_document_mark",
    "write_publication_bytes",
    "SURFACE_LOCATOR_VERSION",
    "LinkFieldFlattenError",
    "LinkFieldFlattenReceipt",
    "flatten_link_fields",
]

from .docx_combined_transaction import CombinedTransactionReceipt, apply_docx_transaction
from .docx_inventory import SURFACE_LOCATOR_VERSION, PackageGraph, PackageRelationship
from .docx_package_transaction import (
    PackageDisposition,
    PackageDispositionStatus,
    PackageMutation,
    PackageMutationError,
    PackageMutationKind,
    PackageTransactionReceipt,
    apply_package_transaction,
)
from .docx_publication_marks import (
    BodyAppendix,
    DocumentMark,
    PublicationMarkError,
    append_body_appendix,
    has_body_appendix,
    has_document_mark,
    remove_body_appendix,
    stamp_document_mark,
    write_publication_bytes,
)


class LinkFieldFlattenError(PackageError):
    """A link-field package could not be flattened and verified safely."""


@dataclass(frozen=True)
class LinkFieldFlattenReceipt:
    """Verified result of :func:`flatten_link_fields`."""

    data: bytes
    flattened_fields: int
    flattened_hyperlinks: int
    removed_relationships: int
    cleared_custom_xml_values: int
    changed_parts: tuple[str, ...]

    @property
    def changed(self) -> bool:
        return bool(self.changed_parts)


_LINK_FIELD_TOKEN = re.compile(r"(?i)(?<![A-Z])(?:HYPERLINK|LINK)(?![A-Z])")
_LINK_RELATIONSHIP_SUFFIX = "/hyperlink"
_LINK_STORY_PART = re.compile(
    r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$"
)
_LINK_FIELD_INSTRUCTION_NAMES = frozenset({"instrText", "delInstrText"})
_LINK_FIELD_WRAPPERS = frozenset({"ins", "del"})
_LINK_FIELD_SKIPPED_CHILDREN = frozenset(
    {"pPr", "rPr", "sectPr", "tblPr", "trPr", "tcPr", "sdtPr", "sdtEndPr"}
)


def flatten_link_fields(
    data: bytes,
    value_predicate: Callable[[str], bool],
    *,
    include_custom_xml: bool = True,
) -> LinkFieldFlattenReceipt:
    """Flatten selected Word link fields without applying content policy.

    ``value_predicate`` receives field instructions, external hyperlink targets,
    and values carried by ``customXml/item*.xml``. Matching HYPERLINK/LINK fields
    retain their display runs; matching relationship-backed ``w:hyperlink``
    elements are unwrapped and their now-unused hyperlink relationships removed;
    matching custom XML text, tails, and attributes are cleared.

    The source is never modified. The returned bytes are made available only
    after the complete ZIP/XML result has been reopened and all requested
    removals have been verified.
    """
    try:
        source_entries = read_package_entries(data)
        entries = {entry.name: entry for entry in source_entries}
        changed_payloads: dict[str, bytes] = {}
        dropped_by_rels_part: dict[str, set[str]] = {}
        flattened_fields = 0
        flattened_hyperlinks = 0
        cleared_custom_xml_values = 0

        for entry in source_entries:
            if not _LINK_STORY_PART.fullmatch(entry.name):
                continue
            root = parse_package_xml(entry.data, part_name=entry.name)
            fields = _flatten_fields_in_element(root, value_predicate)
            rels_name = _link_relationship_part(entry.name)
            targets = _link_relationship_targets(entries, rels_name)
            dropped: set[str] = set()
            hyperlinks = _unwrap_selected_hyperlinks(root, targets, value_predicate, dropped)
            if fields or hyperlinks:
                changed_payloads[entry.name] = etree.tostring(
                    root, encoding="utf-8", xml_declaration=True
                )
                flattened_fields += fields
                flattened_hyperlinks += hyperlinks
            if dropped:
                dropped_by_rels_part.setdefault(rels_name, set()).update(dropped)

        for rels_name, relationship_ids in dropped_by_rels_part.items():
            entry = entries.get(rels_name)
            if entry is None:
                raise LinkFieldFlattenError(
                    f"story references hyperlinks but relationship part is missing: {rels_name}"
                )
            root = parse_package_xml(entry.data, part_name=entry.name)
            removed = _remove_link_relationships(root, relationship_ids)
            if removed != len(relationship_ids):
                raise LinkFieldFlattenError(
                    f"could not remove every selected hyperlink relationship from {rels_name}"
                )
            changed_payloads[rels_name] = etree.tostring(
                root, encoding="utf-8", xml_declaration=True
            )

        if include_custom_xml:
            for entry in source_entries:
                if not re.fullmatch(r"customXml/item\d*\.xml", entry.name):
                    continue
                root = parse_package_xml(entry.data, part_name=entry.name)
                cleared = _clear_selected_xml_values(root, value_predicate)
                if cleared:
                    changed_payloads[entry.name] = etree.tostring(
                        root, encoding="utf-8", xml_declaration=True
                    )
                    cleared_custom_xml_values += cleared

        if not changed_payloads:
            return LinkFieldFlattenReceipt(data, 0, 0, 0, 0, ())
        output = BytesIO()
        with ZipFile(output, "w") as archive:
            for entry in source_entries:
                payload = changed_payloads.get(entry.name, entry.data)
                archive.writestr(entry.zip_info(), payload)
        result = output.getvalue()
        read_package_entries(result)
        _verify_flattened_links(result, dropped_by_rels_part)
        return LinkFieldFlattenReceipt(
            result,
            flattened_fields,
            flattened_hyperlinks,
            sum(len(ids) for ids in dropped_by_rels_part.values()),
            cleared_custom_xml_values,
            tuple(sorted(changed_payloads)),
        )
    except LinkFieldFlattenError:
        raise
    except (OSError, PackageError, etree.LxmlError, ValueError) as exc:
        raise LinkFieldFlattenError(f"link-field flatten transaction failed: {exc}") from exc


def _link_local_name(name: object) -> str:
    return str(name).rsplit("}", 1)[-1]


def _link_attribute(element: etree._Element, local_name: str) -> str | None:
    return next(
        (value for name, value in element.attrib.items() if _link_local_name(name) == local_name),
        None,
    )


def _link_relationship_part(story_part: str) -> str:
    parent, filename = story_part.rsplit("/", 1)
    return f"{parent}/_rels/{filename}.rels"


def _link_relationship_targets(entries: dict[str, PackageEntry], rels_name: str) -> dict[str, str]:
    entry = entries.get(rels_name)
    if entry is None:
        return {}
    root = parse_package_xml(entry.data, part_name=entry.name)
    return {
        relationship_id: target
        for child in root
        if _link_local_name(child.tag) == "Relationship"
        and (_link_attribute(child, "Type") or "").endswith(_LINK_RELATIONSHIP_SUFFIX)
        and (relationship_id := _link_attribute(child, "Id")) is not None
        and (target := _link_attribute(child, "Target")) is not None
    }


def _unwrap_selected_hyperlinks(
    parent: etree._Element,
    targets: dict[str, str],
    predicate: Callable[[str], bool],
    dropped: set[str],
) -> int:
    count = 0
    for child in list(parent):
        count += _unwrap_selected_hyperlinks(child, targets, predicate, dropped)
        if _link_local_name(child.tag) != "hyperlink":
            continue
        relationship_id = _link_attribute(child, "id")
        target = targets.get(relationship_id or "")
        if relationship_id is None or target is None or not predicate(target):
            continue
        index = parent.index(child)
        children = list(child)
        for nested in children:
            child.remove(nested)
        if child.tail and children:
            children[-1].tail = (children[-1].tail or "") + child.tail
        parent[index : index + 1] = children
        dropped.add(relationship_id)
        count += 1
    return count


def _flatten_fields_in_element(parent: etree._Element, predicate: Callable[[str], bool]) -> int:
    count = 0
    for child in list(parent):
        count += _flatten_fields_in_element(child, predicate)
    for child in list(parent):
        if _link_local_name(child.tag) != "fldSimple":
            continue
        instruction = _link_attribute(child, "instr") or ""
        if not _LINK_FIELD_TOKEN.search(instruction) or not predicate(instruction):
            continue
        index = parent.index(child)
        children = list(child)
        for nested in children:
            child.remove(nested)
        if child.tail and children:
            children[-1].tail = (children[-1].tail or "") + child.tail
        parent[index : index + 1] = children
        count += 1

    slots = _link_field_slots(parent)
    index = 0
    while index < len(slots):
        if _link_field_marker(slots[index][0]) != "begin":
            index += 1
            continue
        end, nested, separated, instruction, drop = _collect_link_field(slots, index)
        if end is None:
            index += 1
            continue
        if (
            not nested
            and separated
            and _LINK_FIELD_TOKEN.search(instruction)
            and predicate(instruction)
        ):
            for element in drop:
                parent_element = element.getparent()
                if parent_element is not None:
                    parent_element.remove(element)
            count += 1
            slots = _link_field_slots(parent)
            index = 0
            continue
        index = end + 1
    return count


def _link_field_slots(parent: etree._Element) -> list[tuple[etree._Element, etree._Element]]:
    slots: list[tuple[etree._Element, etree._Element]] = []
    for child in parent:
        local = _link_local_name(child.tag)
        if local in _LINK_FIELD_SKIPPED_CHILDREN:
            continue
        if local in _LINK_FIELD_WRAPPERS:
            slots.extend(_link_field_slots(child))
        else:
            slots.append((child, parent))
    return slots


def _link_field_marker(slot: etree._Element) -> str | None:
    if _link_local_name(slot.tag) == "fldChar":
        return _link_attribute(slot, "fldCharType")
    for child in slot:
        if _link_local_name(child.tag) == "fldChar":
            return _link_attribute(child, "fldCharType")
    return None


def _collect_link_field(
    slots: list[tuple[etree._Element, etree._Element]], begin: int
) -> tuple[int | None, bool, bool, str, list[etree._Element]]:
    depth = 1
    nested = False
    separated = False
    instruction: list[str] = []
    drop = [slots[begin][0]]
    for index in range(begin + 1, len(slots)):
        element = slots[index][0]
        marker = _link_field_marker(element)
        if marker == "begin":
            depth += 1
            nested = True
            drop.append(element)
        elif marker == "end":
            depth -= 1
            drop.append(element)
            if depth == 0:
                return index, nested, separated, "".join(instruction), drop
        elif depth != 1:
            drop.append(element)
        elif marker == "separate":
            separated = True
            drop.append(element)
        elif not separated:
            instruction.extend(
                node.text
                for node in element.iter()
                if _link_local_name(node.tag) in _LINK_FIELD_INSTRUCTION_NAMES and node.text
            )
            drop.append(element)
    return None, nested, separated, "".join(instruction), drop


def _remove_link_relationships(root: etree._Element, relationship_ids: set[str]) -> int:
    removed = 0
    for child in list(root):
        if (
            _link_local_name(child.tag) == "Relationship"
            and _link_attribute(child, "Id") in relationship_ids
            and (_link_attribute(child, "Type") or "").endswith(_LINK_RELATIONSHIP_SUFFIX)
        ):
            root.remove(child)
            removed += 1
    return removed


def _clear_selected_xml_values(root: etree._Element, predicate: Callable[[str], bool]) -> int:
    cleared = 0
    for element in root.iter():
        if element.text and predicate(element.text):
            element.text = ""
            cleared += 1
        if element.tail and predicate(element.tail):
            element.tail = ""
            cleared += 1
        for name, value in list(element.attrib.items()):
            if value and predicate(value):
                element.set(name, "")
                cleared += 1
    return cleared


def _verify_flattened_links(data: bytes, dropped: dict[str, set[str]]) -> None:
    entries = {entry.name: entry for entry in read_package_entries(data)}
    for rels_name, relationship_ids in dropped.items():
        rels = entries.get(rels_name)
        if rels is None:
            raise LinkFieldFlattenError(f"relationship part disappeared: {rels_name}")
        root = parse_package_xml(rels.data, part_name=rels.name)
        remaining = {
            _link_attribute(child, "Id")
            for child in root
            if _link_local_name(child.tag) == "Relationship"
        }
        if relationship_ids & remaining:
            raise LinkFieldFlattenError(f"selected hyperlink relationship survived: {rels_name}")
        story_name = rels_name.replace("/_rels/", "/").removesuffix(".rels")
        story = entries.get(story_name)
        if story is None:
            raise LinkFieldFlattenError(f"story part disappeared: {story_name}")
        story_root = parse_package_xml(story.data, part_name=story.name)
        surviving_references = {
            relationship_id
            for element in story_root.iter()
            if _link_local_name(element.tag) == "hyperlink"
            and (relationship_id := _link_attribute(element, "id")) is not None
        }
        if relationship_ids & surviving_references:
            raise LinkFieldFlattenError(f"selected hyperlink span survived: {story_name}")
