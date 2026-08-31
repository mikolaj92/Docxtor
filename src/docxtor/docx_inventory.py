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


@dataclass(frozen=True)
class DocumentSurface:
    """One neutral value carried by a DOCX package.

    Docxtor reports physical location and mutation capability only. Consumers
    decide whether a value is sensitive, legally relevant, or review content.
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


@dataclass(frozen=True)
class DocxInventory:
    parts: tuple[PackagePart, ...]
    surfaces: tuple[DocumentSurface, ...]
    coverage: InventoryCoverage
    unknown_parts: tuple[str, ...]
    unreadable_parts: tuple[str, ...]


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
            understood = is_xml or content_type.startswith(_KNOWN_BINARY_PREFIXES)
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
    return DocxInventory(
        parts=tuple(parts),
        surfaces=tuple(sorted(surfaces, key=lambda surface: surface.surface_id)),
        coverage=coverage,
        unknown_parts=tuple(sorted(set(unknown_parts))),
        unreadable_parts=tuple(sorted(set(unreadable_parts))),
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
        if element.text and element.text.strip():
            surfaces.append(
                _surface(
                    surface_id=f"xml:{part_name}:{path}:text",
                    kind=(
                        SurfaceKind.TEXT
                        if local in _VISIBLE_TEXT_NAMES
                        else SurfaceKind.XML_TEXT
                    ),
                    part_name=part_name,
                    value=element.text,
                    visibility=(
                        SurfaceVisibility.VISIBLE
                        if local in _VISIBLE_TEXT_NAMES
                        else SurfaceVisibility.HIDDEN
                    ),
                    capability=SurfaceCapability.VALUE_REPLACE,
                    xml_path=path,
                    xml_name=local,
                )
            )
        for raw_name, value in sorted(element.attrib.items()):
            attr_name = _local_name(raw_name)
            if not value or attr_name in _SKIP_XML_LOCAL_NAMES:
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
                    SurfaceCapability.VALUE_REPLACE
                    if external
                    else SurfaceCapability.PRESERVE_ONLY
                ),
                relationship_id=rel_id,
                relationship_type=rel_type,
                external=external,
                xml_path=f"/*[local-name()='Relationships']/*[@Id='{rel_id}']",
                xml_name="Target",
            )
        )
    return surfaces


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
