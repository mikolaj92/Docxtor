from __future__ import annotations

from dataclasses import dataclass

__version__ = "0.5.2"

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
from .loader import document_to_bytes, load_document
from .pdf import PdfDocument, PdfExtractionMode
from .text import PlainTextDocument

# Public names deliberately describe facts rather than a consumer's policy.  Keep
# DocxInventory as a compatibility name while consumers migrate to the snapshot API.
DocxFactsSnapshot = DocxInventory
DocxStructureSnapshot = DocxInventory


@dataclass(frozen=True)
class DocxComparison:
    """Mechanical differences between two complete DOCX package snapshots."""

    before: DocxFactsSnapshot
    after: DocxFactsSnapshot
    added_parts: tuple[str, ...]
    removed_parts: tuple[str, ...]
    changed_parts: tuple[str, ...]
    added_relationships: tuple[str, ...]
    removed_relationships: tuple[str, ...]
    changed_relationships: tuple[str, ...]
    added_surfaces: tuple[str, ...]
    removed_surfaces: tuple[str, ...]
    changed_surfaces: tuple[str, ...]
    lost_containers: tuple[str, ...]


@dataclass(frozen=True)
class TransformPolicy:
    """Neutral allow-list contract for evaluating a :class:`DocxComparison`."""

    added_parts: frozenset[str] = frozenset()
    removed_parts: frozenset[str] = frozenset()
    changed_parts: frozenset[str] = frozenset()
    added_relationships: frozenset[str] = frozenset()
    removed_relationships: frozenset[str] = frozenset()
    changed_relationships: frozenset[str] = frozenset()
    added_surfaces: frozenset[str] = frozenset()
    removed_surfaces: frozenset[str] = frozenset()
    changed_surfaces: frozenset[str] = frozenset()
    lost_containers: frozenset[str] = frozenset()

    def violations(self, comparison: DocxComparison) -> tuple[str, ...]:
        """Return stable identifiers for every change not explicitly allowed."""
        violations: list[str] = []
        for category in (
            "added_parts",
            "removed_parts",
            "changed_parts",
            "added_relationships",
            "removed_relationships",
            "changed_relationships",
            "added_surfaces",
            "removed_surfaces",
            "changed_surfaces",
            "lost_containers",
        ):
            allowed = getattr(self, category)
            violations.extend(
                f"{category}:{identifier}"
                for identifier in getattr(comparison, category)
                if identifier not in allowed
            )
        return tuple(violations)

    def allows(self, comparison: DocxComparison) -> bool:
        return not self.violations(comparison)


def snapshot_docx(data: bytes) -> DocxFactsSnapshot:
    """Return a complete neutral inventory, including opaque and orphan package parts."""
    return inventory_docx(data)


def compare_docx(before: bytes, after: bytes) -> DocxComparison:
    """Compare package parts, relationships, surfaces, and addressable containers."""
    before_snapshot = snapshot_docx(before)
    after_snapshot = snapshot_docx(after)
    before_parts = {part.name: part for part in before_snapshot.parts}
    after_parts = {part.name: part for part in after_snapshot.parts}
    before_surfaces = {surface.surface_id: surface for surface in before_snapshot.surfaces}
    after_surfaces = {surface.surface_id: surface for surface in after_snapshot.surfaces}

    relationship_prefix = "relationship:"
    before_relationships = {
        key: value for key, value in before_surfaces.items() if key.startswith(relationship_prefix)
    }
    after_relationships = {
        key: value for key, value in after_surfaces.items() if key.startswith(relationship_prefix)
    }
    before_containers = {
        surface.container_id for surface in before_snapshot.surfaces if surface.container_id
    }
    after_containers = {
        surface.container_id for surface in after_snapshot.surfaces if surface.container_id
    }

    return DocxComparison(
        before=before_snapshot,
        after=after_snapshot,
        added_parts=_added(before_parts, after_parts),
        removed_parts=_added(after_parts, before_parts),
        changed_parts=_changed(before_parts, after_parts),
        added_relationships=_added(before_relationships, after_relationships),
        removed_relationships=_added(after_relationships, before_relationships),
        changed_relationships=_changed(before_relationships, after_relationships),
        added_surfaces=_added(before_surfaces, after_surfaces),
        removed_surfaces=_added(after_surfaces, before_surfaces),
        changed_surfaces=_changed(before_surfaces, after_surfaces),
        lost_containers=tuple(sorted(before_containers - after_containers)),
    )


def _added(before: dict[str, object], after: dict[str, object]) -> tuple[str, ...]:
    return tuple(sorted(after.keys() - before.keys()))


def _changed(before: dict[str, object], after: dict[str, object]) -> tuple[str, ...]:
    return tuple(sorted(key for key in before.keys() & after.keys() if before[key] != after[key]))


__all__ = [
    "__version__",
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
    "DocxComparison",
    "DocxDocument",
    "DocxFactsSnapshot",
    "DocxInventory",
    "DocxStructureSnapshot",
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
    "InlineSegment",
    "InlineSegmentKind",
    "PdfExtractionMode",
    "PdfDocument",
    "PlainTextDocument",
    "SegmentReplacement",
    "SpanRole",
    "TextSegment",
    "TransformPolicy",
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
    "compare_docx",
    "detect_document_type",
    "document_to_bytes",
    "load_document",
    "normalize_docx_timestamps",
    "parse_package_xml",
    "read_package_entries",
    "restore_semantically_unchanged_xml_parts",
    "snapshot_docx",
    "write_package_atomically",
]
