from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from hashlib import sha256
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from lxml import etree


class SurfaceKind(StrEnum):
    """Mechanical kind of a value carried by a DOCX package."""

    TEXT = "text"
    XML_TEXT = "xml_text"
    XML_ATTRIBUTE = "xml_attribute"
    RELATIONSHIP = "relationship"


class SurfaceVisibility(StrEnum):
    VISIBLE = "visible"
    HIDDEN = "hidden"


class SurfaceCapability(StrEnum):
    """What Docxtor can mechanically do with a surface."""

    RANGE_REPLACE = "range_replace"
    VALUE_REPLACE = "value_replace"
    REMOVE = "remove"
    PRESERVE_ONLY = "preserve_only"
    READ_ONLY = "read_only"
    UNSUPPORTED = "unsupported"


class InventoryCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


@dataclass(frozen=True)
class PackagePart:
    name: str
    content_type: str
    size: int
    sha256: str
    is_xml: bool
    understood: bool
    error: str | None = None


SURFACE_LOCATOR_VERSION = "docxtor-surface-v1"


@dataclass(frozen=True)
class DocumentSurface:
    """One neutral, fully qualified value carried by a DOCX package.

    Qualified names and ancestors let consumers apply their own policy without
    opening or reparsing XML. Docxtor reports physical facts and capabilities
    only; it never classifies a value as PII, legal content, or review policy.
    """

    surface_id: str
    kind: SurfaceKind
    part_name: str
    value: str
    value_sha256: str
    visibility: SurfaceVisibility
    capability: SurfaceCapability
    container_id: str | None = None
    relationship_id: str | None = None
    relationship_type: str | None = None
    external: bool = False
    xml_path: str | None = None
    xml_name: str | None = None
    element_qname: str | None = None
    attribute_qname: str | None = None
    ancestor_qnames: tuple[str, ...] = ()
    role: str | None = None
    locator_version: str = SURFACE_LOCATOR_VERSION


@dataclass(frozen=True)
class PackageRelationship:
    source_part: str
    relationship_part: str
    relationship_id: str
    relationship_type: str
    target: str
    target_part: str | None
    external: bool


@dataclass(frozen=True)
class PackageGraph:
    relationships: tuple[PackageRelationship, ...]
    reachable_parts: tuple[str, ...]
    orphan_parts: tuple[str, ...]


@dataclass(frozen=True)
class DocxInventory:
    parts: tuple[PackagePart, ...]
    surfaces: tuple[DocumentSurface, ...]
    coverage: InventoryCoverage
    unknown_parts: tuple[str, ...]
    unreadable_parts: tuple[str, ...]
    graph: PackageGraph = PackageGraph((), (), ())


_CONTENT_TYPES_PART = "[Content_Types].xml"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_CONTENT_TYPES = {
    "application/xml",
    "text/xml",
    "application/vnd.openxmlformats-package.relationships+xml",
}
_XML_SUFFIXES = ("+xml", "/xml")
_KNOWN_BINARY_PREFIXES = ("image/", "audio/", "video/")
_TEXT_PART_SUFFIXES = (".html", ".htm", ".xhtml", ".mht", ".txt", ".css")
_SKIP_XML_LOCAL_NAMES = {
    # Namespace compatibility declarations are syntax, not carried document values.
    "Ignorable",
}
_VISIBLE_TEXT_NAMES = {"t", "delText"}


def inventory_docx(data: bytes) -> DocxInventory:
    """Inventory every ZIP member, including orphan and unreachable OPC parts."""
    parts: list[PackagePart] = []
    surfaces: list[DocumentSurface] = []
    unknown_parts: list[str] = []
    unreadable_parts: list[str] = []

    try:
        archive = ZipFile(BytesIO(data))
    except BadZipFile:
        return DocxInventory(
            parts=(),
            surfaces=(),
            coverage=InventoryCoverage.INCOMPLETE,
            unknown_parts=(),
            unreadable_parts=("<package>",),
        )

    with archive:
        content_types, content_type_error = _content_types(archive)
        if content_type_error:
            unreadable_parts.append(_CONTENT_TYPES_PART)

        for name in sorted(item.filename for item in archive.infolist() if not item.is_dir()):
            payload = archive.read(name)
            content_type = _content_type_for(name, content_types)
            is_relationships = name.endswith(".rels")
            is_xml = is_relationships or _is_xml_part(content_type, payload)
            is_text_part = name.lower().endswith(_TEXT_PART_SUFFIXES)
            understood = is_xml or is_text_part or content_type.startswith(_KNOWN_BINARY_PREFIXES)
            error: str | None = None

            if is_xml:
                try:
                    root = etree.fromstring(payload, parser=_safe_xml_parser())
                except (etree.XMLSyntaxError, ValueError) as exc:
                    error = type(exc).__name__
                    unreadable_parts.append(name)
                else:
                    if is_relationships:
                        surfaces.extend(_relationship_surfaces(name, root))
                    elif name != _CONTENT_TYPES_PART:
                        surfaces.extend(_xml_surfaces(name, root))
            elif is_text_part:
                try:
                    value = payload.decode("utf-8")
                except UnicodeDecodeError as exc:
                    error = type(exc).__name__
                    unreadable_parts.append(name)
                else:
                    if name.lower().endswith((".html", ".htm", ".xhtml")):
                        try:
                            root = etree.fromstring(payload, parser=_safe_xml_parser())
                        except (etree.XMLSyntaxError, ValueError):
                            root = etree.HTML(value)
                        if root is None:
                            error = "HTMLParseError"
                            unreadable_parts.append(name)
                        else:
                            surfaces.extend(_xml_surfaces(name, root))
                    else:
                        surfaces.append(
                            _surface(
                                surface_id=f"text-part:{name}",
                                kind=SurfaceKind.XML_TEXT,
                                part_name=name,
                                value=value,
                                visibility=SurfaceVisibility.HIDDEN,
                                capability=SurfaceCapability.VALUE_REPLACE,
                                xml_path="/",
                                xml_name="content",
                                role="text_part_content",
                            )
                        )
            elif not understood:
                unknown_parts.append(name)

            parts.append(
                PackagePart(
                    name=name,
                    content_type=content_type,
                    size=len(payload),
                    sha256=sha256(payload).hexdigest(),
                    is_xml=is_xml,
                    understood=understood and error is None,
                    error=error,
                )
            )

    coverage = (
        InventoryCoverage.COMPLETE
        if not unknown_parts and not unreadable_parts
        else InventoryCoverage.INCOMPLETE
    )
    graph = _package_graph(tuple(part.name for part in parts), tuple(surfaces))
    return DocxInventory(
        parts=tuple(parts),
        surfaces=tuple(sorted(surfaces, key=lambda surface: surface.surface_id)),
        coverage=coverage,
        unknown_parts=tuple(sorted(set(unknown_parts))),
        unreadable_parts=tuple(sorted(set(unreadable_parts))),
        graph=graph,
    )


def _content_types(archive: ZipFile) -> tuple[dict[str, str], bool]:
    try:
        payload = archive.read(_CONTENT_TYPES_PART)
        root = etree.fromstring(payload, parser=_safe_xml_parser())
    except (KeyError, etree.XMLSyntaxError, ValueError):
        return {}, True

    values: dict[str, str] = {}
    for element in root:
        local = _local_name(element.tag)
        if local == "Default":
            extension = element.get("Extension")
            content_type = element.get("ContentType")
            if extension and content_type:
                values[f"*.{extension.lower()}"] = content_type
        elif local == "Override":
            part_name = element.get("PartName")
            content_type = element.get("ContentType")
            if part_name and content_type:
                values[part_name.lstrip("/")] = content_type
    return values, False


def _content_type_for(name: str, content_types: dict[str, str]) -> str:
    direct = content_types.get(name)
    if direct:
        return direct
    suffix = PurePosixPath(name).suffix.lower()
    return content_types.get(f"*{suffix}", "application/octet-stream")


def _xml_surfaces(part_name: str, root: etree._Element) -> list[DocumentSurface]:
    tree = root.getroottree()
    surfaces: list[DocumentSurface] = []
    for element in root.iter():
        path = tree.getpath(element)
        local = _local_name(element.tag)
        if element.text is not None or (local in _VISIBLE_TEXT_NAMES and len(element) == 0):
            surfaces.append(
                _surface(
                    surface_id=f"xml:{part_name}:{path}:text",
                    kind=(
                        SurfaceKind.TEXT if local in _VISIBLE_TEXT_NAMES else SurfaceKind.XML_TEXT
                    ),
                    part_name=part_name,
                    value=element.text or "",
                    visibility=(
                        SurfaceVisibility.VISIBLE
                        if local in _VISIBLE_TEXT_NAMES
                        else SurfaceVisibility.HIDDEN
                    ),
                    capability=SurfaceCapability.VALUE_REPLACE,
                    xml_path=path,
                    xml_name=local,
                    element_qname=str(element.tag),
                    ancestor_qnames=_ancestor_qnames(element),
                    role=_surface_role(part_name, element, None),
                )
            )
        for raw_name, value in sorted(element.attrib.items()):
            attr_name = _local_name(raw_name)
            if attr_name in _SKIP_XML_LOCAL_NAMES:
                continue
            surfaces.append(
                _surface(
                    surface_id=f"xml:{part_name}:{path}:attr:{raw_name}",
                    kind=SurfaceKind.XML_ATTRIBUTE,
                    part_name=part_name,
                    value=value,
                    visibility=SurfaceVisibility.HIDDEN,
                    capability=SurfaceCapability.VALUE_REPLACE,
                    xml_path=path,
                    xml_name=attr_name,
                    element_qname=str(element.tag),
                    attribute_qname=str(raw_name),
                    ancestor_qnames=_ancestor_qnames(element),
                    role=_surface_role(part_name, element, raw_name),
                )
            )
    return surfaces


def _relationship_surfaces(part_name: str, root: etree._Element) -> list[DocumentSurface]:
    surfaces: list[DocumentSurface] = []
    for element in root:
        if _local_name(element.tag) != "Relationship":
            continue
        rel_id = element.get("Id")
        target = element.get("Target")
        rel_type = element.get("Type")
        if not rel_id or target is None or rel_type is None:
            continue
        external = element.get("TargetMode", "").lower() == "external"
        surfaces.append(
            _surface(
                surface_id=f"relationship:{part_name}:{rel_id}:target",
                kind=SurfaceKind.RELATIONSHIP,
                part_name=part_name,
                value=target,
                visibility=SurfaceVisibility.HIDDEN,
                capability=(
                    SurfaceCapability.VALUE_REPLACE if external else SurfaceCapability.PRESERVE_ONLY
                ),
                relationship_id=rel_id,
                relationship_type=rel_type,
                external=external,
                xml_path=f"/*[local-name()='Relationships']/*[@Id='{rel_id}']",
                xml_name="Target",
                element_qname=str(element.tag),
                attribute_qname="Target",
                ancestor_qnames=_ancestor_qnames(element),
                role="external_relationship" if external else "internal_relationship",
            )
        )
    return surfaces


def _ancestor_qnames(element: etree._Element) -> tuple[str, ...]:
    ancestors: list[str] = []
    parent = element.getparent()
    while parent is not None:
        ancestors.append(str(parent.tag))
        parent = parent.getparent()
    return tuple(reversed(ancestors))


def _surface_role(part_name: str, element: etree._Element, attribute: object | None) -> str:
    local = _local_name(element.tag).lower()
    attr = _local_name(attribute).lower() if attribute is not None else "text"
    normalized = part_name.lower()
    if normalized.endswith(".rels"):
        return "relationship_target"
    if "customxml/" in normalized:
        return (
            "custom_xml_binding" if local in {"databinding", "schemaRef".lower()} else "custom_xml"
        )
    if "numbering" in normalized:
        return f"numbering_{local}_{attr}"
    if "drawings/" in normalized or normalized.endswith((".vml", ".svg")):
        return f"drawing_{local}_{attr}"
    if local in {"fldsimple", "instrtext"}:
        return "field_instruction"
    if local == "hyperlink":
        return "hyperlink"
    if local in {"sdt", "sdtpr", "databinding", "tag", "alias"} or any(
        _local_name(parent).lower() in {"sdt", "sdtpr"} for parent in element.iterancestors()
    ):
        return f"content_control_{local}_{attr}"
    return f"xml_{local}_{attr}"


def _relationship_source_part(relationship_part: str) -> str:
    if relationship_part == "_rels/.rels":
        return ""
    path = PurePosixPath(relationship_part)
    parent = path.parent.parent
    source_name = path.name.removesuffix(".rels")
    return str(parent / source_name) if str(parent) != "." else source_name


def _resolve_target(source_part: str, target: str) -> str:
    if not source_part:
        return str(PurePosixPath(target))
    parent = PurePosixPath(source_part).parent
    stack: list[str] = []
    for item in (parent / target).parts:
        if item in {"", "."}:
            continue
        if item == "..":
            if stack:
                stack.pop()
        else:
            stack.append(item)
    return "/".join(stack)


def _package_graph(
    part_names: tuple[str, ...], surfaces: tuple[DocumentSurface, ...]
) -> PackageGraph:
    names = set(part_names)
    relationships: list[PackageRelationship] = []
    for surface in surfaces:
        if surface.kind is not SurfaceKind.RELATIONSHIP:
            continue
        source = _relationship_source_part(surface.part_name)
        target_part = None if surface.external else _resolve_target(source, surface.value)
        relationships.append(
            PackageRelationship(
                source_part=source,
                relationship_part=surface.part_name,
                relationship_id=surface.relationship_id or "",
                relationship_type=surface.relationship_type or "",
                target=surface.value,
                target_part=target_part,
                external=surface.external,
            )
        )
    reachable = {"[Content_Types].xml", "_rels/.rels"} & names
    queue = [""]
    visited_sources: set[str] = set()
    while queue:
        source = queue.pop()
        if source in visited_sources:
            continue
        visited_sources.add(source)
        for relationship in relationships:
            if relationship.source_part != source or relationship.external:
                continue
            target = relationship.target_part
            if target and target in names and target not in reachable:
                reachable.add(target)
                queue.append(target)
                rel_part = str(
                    PurePosixPath(target).parent / "_rels" / (PurePosixPath(target).name + ".rels")
                )
                if rel_part in names:
                    reachable.add(rel_part)
    return PackageGraph(
        relationships=tuple(
            sorted(relationships, key=lambda item: (item.source_part, item.relationship_id))
        ),
        reachable_parts=tuple(sorted(reachable)),
        orphan_parts=tuple(sorted(names - reachable)),
    )


def _surface(*, value: str, **kwargs: object) -> DocumentSurface:
    return DocumentSurface(
        value=value,
        value_sha256=sha256(value.encode("utf-8")).hexdigest(),
        **kwargs,
    )


def _is_xml_part(content_type: str, payload: bytes) -> bool:
    return (
        content_type in _XML_CONTENT_TYPES
        or content_type.endswith(_XML_SUFFIXES)
        or payload.lstrip().startswith(b"<?xml")
    )


def _safe_xml_parser() -> etree.XMLParser:
    return etree.XMLParser(resolve_entities=False, no_network=True, recover=False)


def _local_name(name: object) -> str:
    if not isinstance(name, str):
        return ""
    return name.rsplit("}", 1)[-1]
