"""Neutral, typed facts and before/after audit for DOCX packages.

This module deliberately builds on Docxtor's package reader, inventory, and
``DocxDocument`` story index.  It contains no consumer policy and does not use
ZIP or OOXML parser APIs directly.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from hashlib import sha256
from pathlib import Path, PurePosixPath
from posixpath import normpath
from typing import TypeAlias

from lxml import etree

from .docx import DocxDocument
from .docx_inventory import DocumentSurface, inventory_docx
from .docx_package import PackageError, parse_package_xml, read_package_entries

Source: TypeAlias = str | Path | bytes | DocxDocument
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XML_CT = "application/vnd.openxmlformats-package.relationships+xml"


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


@dataclass(frozen=True)
class StoryFact:
    story_id: str
    story_kind: str
    paragraph_ids: tuple[str, ...]


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
class DocxStructureSnapshot:
    part_names: tuple[str, ...]
    relationship_identities: tuple[tuple[str, str], ...]
    story_ids: tuple[str, ...]
    container_ids: tuple[str, ...]
    field_ids: tuple[str, ...]
    bookmark_ids: tuple[str, ...]
    table_ids: tuple[str, ...] = ()
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


def docx_facts(source: Source) -> DocxFactsSnapshot:
    """Return the complete mechanical snapshot, or raise on unreadable/malformed input."""
    payload = _payload(source)
    entries = read_package_entries(payload)
    entry_data = {entry.name: entry.data for entry in entries}
    if "[Content_Types].xml" not in entry_data:
        raise PackageError("DOCX package has no [Content_Types].xml")
    content_types = _content_types(entry_data["[Content_Types].xml"])
    ct_defaults, ct_overrides = _content_type_map(content_types)
    inventory = inventory_docx(payload)
    if inventory.unreadable_parts:
        raise PackageError(
            "DOCX inventory has unreadable parts: " + ", ".join(inventory.unreadable_parts)
        )

    relationships = _relationships(entry_data)
    reachable = _reachable(set(entry_data), relationships)
    diagnostics: list[FactDiagnostic] = [
        FactDiagnostic("unknown_part", "content type is not mechanically understood", name)
        for name in inventory.unknown_parts
    ]
    for rel in relationships:
        if not rel.external and rel.target_part not in entry_data:
            diagnostics.append(
                FactDiagnostic(
                    "missing_relationship_target",
                    f"missing target {rel.target_part}",
                    rel.relationship_part,
                )
            )
    missing_targets = [item for item in diagnostics if item.code == "missing_relationship_target"]
    if missing_targets:
        raise PackageError("DOCX relationship target is missing: " + missing_targets[0].message)

    inventory_parts = {part.name: part for part in inventory.parts}
    parts = tuple(
        PartFact(
            name,
            _content_type_for(name, ct_defaults, ct_overrides),
            len(data),
            sha256(data).hexdigest(),
            inventory_parts[name].is_xml,
            name in reachable,
            inventory_parts[name].understood,
        )
        for name, data in sorted(entry_data.items())
    )

    # DocxDocument is the canonical story/addressing implementation. It also
    # rejects packages that python-docx cannot interpret.
    document = source if isinstance(source, DocxDocument) else DocxDocument.open_bytes(payload)
    paragraphs = _paragraphs(document)
    stories = _stories(paragraphs)
    xml_roots = {
        name: parse_package_xml(data, part_name=name)
        for name, data in entry_data.items()
        if inventory_parts[name].is_xml
        and not name.endswith(".rels")
        and name != "[Content_Types].xml"
    }
    features = _features(xml_roots, paragraphs, relationships, set(entry_data))
    orphans = tuple(sorted(set(entry_data) - reachable - {"[Content_Types].xml", "_rels/.rels"}))
    coverage = FactsCoverage.COMPLETE if not diagnostics else FactsCoverage.INCOMPLETE
    structure = DocxStructureSnapshot(
        part_names=tuple(sorted(entry_data)),
        relationship_identities=tuple(rel.identity for rel in relationships),
        story_ids=tuple(story.story_id for story in stories),
        container_ids=tuple(p.container_id for p in paragraphs),
        field_ids=tuple(f.fact_id for f in features["fields"]),
        bookmark_ids=tuple(f.fact_id for f in features["bookmarks"]),
        table_ids=tuple(
            f"table:{index}"
            for index in sorted(
                {
                    paragraph.coordinate.table_index
                    for paragraph in paragraphs
                    if paragraph.coordinate.table_index is not None
                }
            )
        ),
        section_property_hashes=_section_property_hashes(xml_roots),
        coverage=coverage,
    )
    return DocxFactsSnapshot(
        FactsCoverage.COMPLETE if not diagnostics else FactsCoverage.INCOMPLETE,
        tuple(diagnostics),
        parts,
        content_types,
        relationships,
        tuple(sorted(reachable)),
        orphans,
        inventory.surfaces,
        stories,
        paragraphs,
        features["fields"],
        features["bookmarks"],
        features["links"],
        features["hidden"],
        features["comments"],
        features["notes"],
        features["textboxes"],
        features["properties"],
        features["embedded_objects"],
        features["media"],
        structure,
    )


# A discoverable noun/verb pair for callers that prefer ``snapshot_docx``.
snapshot_docx = docx_facts


def _section_property_hashes(
    xml_roots: dict[str, etree._Element],
) -> tuple[str, ...]:
    values: list[str] = []
    for part_name, root in sorted(xml_roots.items()):
        if not part_name.startswith("word/"):
            continue
        for section in root.iter(f"{{{_W_NS}}}sectPr"):
            payload = etree.tostring(section, method="c14n")
            values.append(sha256(payload).hexdigest())
    return tuple(values)


def compare_docx(
    before: Source, after: Source, policy: TransformPolicy | None = None
) -> DocxComparison:
    """Compare two fail-closed snapshots and apply an optional neutral allow-list."""
    b = docx_facts(before)
    a = docx_facts(after)
    selected = policy or TransformPolicy()
    part_changes = _diff(
        {x.name: x.sha256 for x in b.parts}, {x.name: x.sha256 for x in a.parts}, "part"
    )
    rel_changes = _diff(
        {_rel_key(x): _hash_repr(x) for x in b.relationships},
        {_rel_key(x): _hash_repr(x) for x in a.relationships},
        "relationship",
    )
    surface_changes = _diff(
        {x.surface_id: x.value_sha256 for x in b.surfaces},
        {x.surface_id: x.value_sha256 for x in a.surfaces},
        "surface",
    )
    container_changes = _diff(
        {x.container_id: x.text_sha256 for x in b.paragraphs},
        {x.container_id: x.text_sha256 for x in a.paragraphs},
        "container",
    )
    fact_changes: list[FactChange] = []
    for category in _FEATURE_CATEGORIES:
        left = getattr(b, category)
        right = getattr(a, category)
        fact_changes.extend(
            _diff(
                {x.fact_id: _hash_repr(x) for x in left},
                {x.fact_id: _hash_repr(x) for x in right},
                category,
            )
        )
    violations = [c for c in part_changes if c.kind not in selected.allowed_part_changes]
    violations += [c for c in rel_changes if c.kind not in selected.allowed_relationship_changes]
    violations += [c for c in surface_changes if c.kind not in selected.allowed_surface_changes]
    violations += [c for c in container_changes if c.kind not in selected.allowed_container_changes]
    violations += [
        c
        for c in fact_changes
        if "*" not in selected.allowed_fact_categories
        and c.category not in selected.allowed_fact_categories
    ]
    diagnostics: tuple[FactDiagnostic, ...] = ()
    if selected.require_complete_coverage and (
        b.coverage is not FactsCoverage.COMPLETE or a.coverage is not FactsCoverage.COMPLETE
    ):
        diagnostics = (
            FactDiagnostic("incomplete_coverage", "comparison requires complete facts coverage"),
        )
    return DocxComparison(
        b,
        a,
        part_changes,
        rel_changes,
        surface_changes,
        container_changes,
        tuple(fact_changes),
        tuple(violations),
        diagnostics,
    )


def _payload(source: Source) -> bytes:
    if isinstance(source, DocxDocument):
        raw = getattr(source, "_source_bytes", None)
        return raw if isinstance(raw, bytes) else source.to_bytes()
    if isinstance(source, bytes):
        return source
    try:
        return Path(source).read_bytes()
    except OSError as exc:
        raise PackageError(f"cannot read DOCX: {exc}") from exc


def _content_types(data: bytes) -> tuple[ContentTypeFact, ...]:
    root = parse_package_xml(data, part_name="[Content_Types].xml")
    if etree.QName(root).localname != "Types":
        raise PackageError("[Content_Types].xml has an invalid root")
    facts: list[ContentTypeFact] = []
    for element in root:
        local = etree.QName(element).localname
        key = element.get("Extension") if local == "Default" else element.get("PartName")
        value = element.get("ContentType")
        if local not in {"Default", "Override"} or not key or not value:
            raise PackageError("[Content_Types].xml contains an invalid declaration")
        facts.append(ContentTypeFact(key.lstrip("/"), value, local == "Default"))
    return tuple(sorted(facts, key=lambda x: (not x.is_default, x.key)))


def _content_type_map(
    facts: tuple[ContentTypeFact, ...],
) -> tuple[dict[str, str], dict[str, str]]:
    defaults = {x.key.lower(): x.content_type for x in facts if x.is_default}
    overrides = {x.key: x.content_type for x in facts if not x.is_default}
    overrides["[Content_Types].xml"] = "application/xml"
    return defaults, overrides


def _content_type_for(name: str, defaults: dict[str, str], overrides: dict[str, str]) -> str:
    return overrides.get(
        name,
        defaults.get(
            PurePosixPath(name).suffix.lstrip(".").lower(),
            "application/octet-stream",
        ),
    )


def _relationships(entries: dict[str, bytes]) -> tuple[RelationshipFact, ...]:
    facts: list[RelationshipFact] = []
    for name, data in sorted(entries.items()):
        if not name.endswith(".rels"):
            continue
        root = parse_package_xml(data, part_name=name)
        if etree.QName(root).localname != "Relationships":
            raise PackageError(f"relationship part {name} has an invalid root")
        source = _relationship_source(name)
        seen: set[str] = set()
        for element in root:
            if etree.QName(element).localname != "Relationship":
                raise PackageError(f"relationship part {name} contains an invalid element")
            rid, rtype, target = element.get("Id"), element.get("Type"), element.get("Target")
            if not rid or not rtype or target is None or rid in seen:
                raise PackageError(f"relationship part {name} contains an invalid relationship")
            seen.add(rid)
            external = element.get("TargetMode", "").lower() == "external"
            facts.append(
                RelationshipFact(
                    source,
                    name,
                    rid,
                    rtype,
                    target,
                    None if external else _resolve_target(source, target),
                    external,
                )
            )
    return tuple(sorted(facts, key=lambda x: x.identity))


def _relationship_source(name: str) -> str:
    if name == "_rels/.rels":
        return ""
    path = PurePosixPath(name)
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _resolve_target(source: str, target: str) -> str:
    base = str(PurePosixPath(source).parent) if source else ""
    resolved = normpath(f"{base}/{target}".lstrip("/"))
    if resolved == ".." or resolved.startswith("../"):
        raise PackageError(f"relationship target escapes package: {target}")
    return resolved


def _reachable(names: set[str], rels: tuple[RelationshipFact, ...]) -> set[str]:
    reachable = {"_rels/.rels"} if "_rels/.rels" in names else set()
    queue = [""]
    while queue:
        source = queue.pop(0)
        for rel in rels:
            if (
                rel.source_part == source
                and not rel.external
                and rel.target_part in names
                and rel.target_part not in reachable
            ):
                reachable.add(rel.target_part or "")
                queue.append(rel.target_part or "")
                rel_part = _rels_name(rel.target_part or "")
                if rel_part in names:
                    reachable.add(rel_part)
    return reachable


def _rels_name(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _paragraphs(document: DocxDocument) -> tuple[ParagraphFact, ...]:
    result: list[ParagraphFact] = []
    for index, container_id, paragraph in document.get_indexed_paragraphs():
        style_id = paragraph.style.style_id if paragraph.style is not None else None
        outline = None
        nodes = paragraph._p.xpath("./w:pPr/w:outlineLvl")
        if nodes:
            raw = nodes[0].get(f"{{{_W_NS}}}val")
            try:
                outline = int(raw) if raw is not None else None
            except ValueError:
                raise PackageError(f"invalid outline level in {container_id}") from None
        coord = _coordinate(container_id, index)
        result.append(
            ParagraphFact(
                container_id,
                index,
                coord.story_kind,
                paragraph.text,
                sha256(paragraph.text.encode()).hexdigest(),
                style_id,
                outline,
                coord,
            )
        )
    return tuple(result)


def _coordinate(cid: str, paragraph_index: int) -> ContainerCoordinate:
    bits = cid.split(":")

    def after(token: str) -> int | None:
        try:
            return int(bits[bits.index(token) + 1])
        except (ValueError, IndexError):
            return None

    kind = bits[0]
    ordinal = (
        after(kind) if kind in {"header", "footer", "txbx", "comment", "footnote", "endnote"} else 0
    )
    return ContainerCoordinate(
        cid, kind, ordinal or 0, after("table"), after("r"), after("c"), paragraph_index
    )


def _stories(paragraphs: tuple[ParagraphFact, ...]) -> tuple[StoryFact, ...]:
    grouped: dict[str, list[str]] = {}
    for paragraph in paragraphs:
        cid = paragraph.container_id
        story = cid.rsplit(":p:", 1)[0]
        if ":table:" in story:
            story = story.split(":table:", 1)[0] or "body"
        grouped.setdefault(story, []).append(cid)
    return tuple(
        StoryFact(key, key.split(":", 1)[0], tuple(values))
        for key, values in sorted(grouped.items())
    )


_FEATURE_CATEGORIES = (
    "fields",
    "bookmarks",
    "links",
    "hidden",
    "comments",
    "notes",
    "textboxes",
    "properties",
    "embedded_objects",
    "media",
)


def _features(
    roots: dict[str, etree._Element],
    paragraphs: tuple[ParagraphFact, ...],
    relationships: tuple[RelationshipFact, ...],
    part_names: set[str],
) -> dict[str, tuple[NamedFact, ...]]:
    singular = {
        "fields": "field",
        "bookmarks": "bookmark",
        "links": "link",
        "hidden": "hidden",
        "comments": "comment",
        "notes": "note",
        "textboxes": "textbox",
        "properties": "property",
        "embedded_objects": "embedded_object",
        "media": "media",
    }
    out: dict[str, list[NamedFact]] = {name: [] for name in singular.values()}
    containers = {p.text: p.container_id for p in paragraphs}
    for part, root in sorted(roots.items()):
        tree = root.getroottree()
        for element in root.iter():
            local = etree.QName(element).localname
            path = tree.getpath(element)
            fid = f"{part}:{path}"
            text = "".join(element.itertext()) or None
            container = containers.get(text or "")
            if local in {"fldSimple", "instrText", "fldChar"}:
                value = (
                    element.get(f"{{{_W_NS}}}instr")
                    or element.get(f"{{{_W_NS}}}fldCharType")
                    or text
                )
                out["field"].append(NamedFact("field", part, fid, container, value))
            if local in {"bookmarkStart", "bookmarkEnd"}:
                value = element.get(f"{{{_W_NS}}}name") or element.get(f"{{{_W_NS}}}id")
                out["bookmark"].append(NamedFact("bookmark", part, fid, container, value))
            if local == "hyperlink":
                out["link"].append(
                    NamedFact(
                        "link",
                        part,
                        fid,
                        container,
                        element.get(f"{{{_W_NS}}}anchor"),
                        element.get(f"{{{_R_NS}}}id"),
                    )
                )
            if local in {"vanish", "webHidden", "specVanish"}:
                out["hidden"].append(NamedFact("hidden", part, fid, container, local))
            if local.startswith("comment") or part.startswith("word/comments"):
                out["comment"].append(
                    NamedFact("comment", part, fid, container, element.get(f"{{{_W_NS}}}id"), text)
                )
            if local in {"footnote", "endnote", "footnoteReference", "endnoteReference"}:
                out["note"].append(
                    NamedFact(local, part, fid, container, element.get(f"{{{_W_NS}}}id"), text)
                )
            if local == "txbxContent":
                out["textbox"].append(NamedFact("textbox", part, fid, container, text))
            if part.startswith("docProps/") and element is not root:
                out["property"].append(NamedFact("property", part, fid, None, local, text))
            if local in {"object", "oleObject", "control"}:
                out["embedded_object"].append(
                    NamedFact(
                        "embedded_object",
                        part,
                        fid,
                        container,
                        local,
                        element.get(f"{{{_R_NS}}}id"),
                    )
                )
    # Binary package facts use relationship targets so identities stay stable.
    binary_parts = set(part_names)
    binary_parts.update(r.target_part for r in relationships if r.target_part)
    for part in sorted(binary_parts):
        if part.startswith("word/media/"):
            out["media"].append(NamedFact("media", part, part))
        elif part.startswith(("word/embeddings/", "word/activeX/")):
            out["embedded_object"].append(NamedFact("embedded_object", part, part))
    return {plural: tuple(out[singular[plural]]) for plural in _FEATURE_CATEGORIES}


def _hash_repr(value: object) -> str:
    return sha256(repr(value).encode()).hexdigest()


def _rel_key(rel: RelationshipFact) -> str:
    return f"{rel.source_part}#{rel.relationship_id}"


def _diff(before: dict[str, str], after: dict[str, str], category: str) -> tuple[FactChange, ...]:
    changes: list[FactChange] = []
    for key in sorted(before.keys() | after.keys()):
        if key not in before:
            changes.append(FactChange(ChangeKind.CREATED, category, key, None, after[key]))
        elif key not in after:
            changes.append(FactChange(ChangeKind.REMOVED, category, key, before[key], None))
        elif before[key] != after[key]:
            changes.append(FactChange(ChangeKind.CHANGED, category, key, before[key], after[key]))
    return tuple(changes)
