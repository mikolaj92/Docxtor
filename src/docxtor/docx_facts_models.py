"""Neutral, typed data model for DOCX mechanical facts and comparisons."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import TypeAlias

from .docx import DocxDocument
from .docx_inventory import DocumentSurface

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

Source: TypeAlias = str | Path | bytes | DocxDocument


class FactsCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class ChangeKind(StrEnum):
    CREATED = "created"
    REMOVED = "removed"
    CHANGED = "changed"


@dataclass(frozen=True)
class FactDiagnostic:
    code: str
    message: str
    part_name: str | None = None


@dataclass(frozen=True)
class ContentTypeFact:
    key: str
    content_type: str
    is_default: bool


@dataclass(frozen=True)
class PartFact:
    name: str
    content_type: str
    size: int
    sha256: str
    is_xml: bool
    reachable: bool
    understood: bool


@dataclass(frozen=True)
class RelationshipFact:
    source_part: str
    relationship_part: str
    relationship_id: str
    relationship_type: str
    target: str
    target_part: str | None
    external: bool

    @property
    def identity(self) -> tuple[str, str]:
        return (self.source_part, self.relationship_id)


@dataclass(frozen=True)
class ContainerCoordinate:
    container_id: str
    story_kind: str
    ordinal: int
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    paragraph_index: int | None = None


@dataclass(frozen=True)
class ParagraphFact:
    container_id: str
    paragraph_index: int | None
    story_kind: str
    text: str
    text_sha256: str
    style_id: str | None
    outline_level: int | None
    coordinate: ContainerCoordinate
    xml_path: str | None = None
    part_name: str | None = None


@dataclass(frozen=True)
class StoryFact:
    story_id: str
    story_kind: str
    paragraph_ids: tuple[str, ...]


@dataclass(frozen=True)
class UnreadablePartFact:
    part_name: str
    error: str


@dataclass(frozen=True)
class NamedFact:
    """A typed mechanical OOXML feature, without domain interpretation."""

    kind: str
    part_name: str
    fact_id: str
    container_id: str | None = None
    value: str | None = None
    target: str | None = None


@dataclass(frozen=True)
class PageBreakFact:
    kind: str
    part_name: str
    fact_id: str
    container_id: str | None = None


@dataclass(frozen=True)
class PageLayoutFacts:
    application_page_count: int | None
    breaks: tuple[PageBreakFact, ...]
    body_page_text: tuple[str, ...]

    @property
    def estimated_page_count(self) -> int | None:
        candidates = [len(self.body_page_text)] if self.body_page_text else []
        if self.application_page_count is not None:
            candidates.append(self.application_page_count)
        return max(candidates) if candidates else None


@dataclass(frozen=True)
class DocxStructureSnapshot:
    part_names: tuple[str, ...]
    relationship_identities: tuple[tuple[str, str], ...]
    story_ids: tuple[str, ...]
    container_ids: tuple[str, ...]
    field_ids: tuple[str, ...]
    bookmark_ids: tuple[str, ...]
    table_ids: tuple[str, ...] = ()
    body_block_ids: tuple[str, ...] = ()
    section_property_hashes: tuple[str, ...] = ()
    coverage: FactsCoverage = FactsCoverage.COMPLETE


@dataclass(frozen=True)
class DocxFactsSnapshot:
    coverage: FactsCoverage
    diagnostics: tuple[FactDiagnostic, ...]
    parts: tuple[PartFact, ...]
    content_types: tuple[ContentTypeFact, ...]
    relationships: tuple[RelationshipFact, ...]
    reachable_parts: tuple[str, ...]
    orphan_parts: tuple[str, ...]
    surfaces: tuple[DocumentSurface, ...]
    stories: tuple[StoryFact, ...]
    paragraphs: tuple[ParagraphFact, ...]
    fields: tuple[NamedFact, ...]
    bookmarks: tuple[NamedFact, ...]
    links: tuple[NamedFact, ...]
    hidden: tuple[NamedFact, ...]
    comments: tuple[NamedFact, ...]
    notes: tuple[NamedFact, ...]
    textboxes: tuple[NamedFact, ...]
    properties: tuple[NamedFact, ...]
    embedded_objects: tuple[NamedFact, ...]
    media: tuple[NamedFact, ...]
    unreadable_parts: tuple[UnreadablePartFact, ...]
    page_layout: PageLayoutFacts
    structure: DocxStructureSnapshot


@dataclass(frozen=True)
class FactChange:
    kind: ChangeKind
    category: str
    identity: str
    before_hash: str | None = None
    after_hash: str | None = None


@dataclass(frozen=True)
class TransformPolicy:
    """Neutral allow-list. Empty allow-lists mean that no such change is allowed."""

    allowed_part_changes: frozenset[ChangeKind] = field(default_factory=frozenset)
    allowed_relationship_changes: frozenset[ChangeKind] = field(default_factory=frozenset)
    allowed_surface_changes: frozenset[ChangeKind] = field(default_factory=frozenset)
    allowed_container_changes: frozenset[ChangeKind] = field(default_factory=frozenset)
    allowed_fact_categories: frozenset[str] = field(default_factory=frozenset)
    require_complete_coverage: bool = True

    @classmethod
    def allow_all(cls) -> TransformPolicy:
        all_changes = frozenset(ChangeKind)
        return cls(all_changes, all_changes, all_changes, all_changes, frozenset({"*"}))


@dataclass(frozen=True)
class DocxComparison:
    before: DocxFactsSnapshot
    after: DocxFactsSnapshot
    part_changes: tuple[FactChange, ...]
    relationship_changes: tuple[FactChange, ...]
    surface_changes: tuple[FactChange, ...]
    container_changes: tuple[FactChange, ...]
    fact_changes: tuple[FactChange, ...]
    violations: tuple[FactChange, ...]
    diagnostics: tuple[FactDiagnostic, ...]

    @property
    def allowed(self) -> bool:
        return not self.violations and not self.diagnostics

    @property
    def compliant(self) -> bool:
        return self.allowed

    @property
    def changes(self) -> tuple[FactChange, ...]:
        return (
            self.part_changes
            + self.relationship_changes
            + self.surface_changes
            + self.container_changes
            + self.fact_changes
        )
