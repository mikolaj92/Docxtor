from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from hashlib import sha256
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .common import DocumentError
from .docx_inventory import (
    DocumentSurface,
    DocxInventory,
    SurfaceCapability,
    SurfaceKind,
    inventory_docx,
)


class SurfaceMutationError(DocumentError):
    """Raised before output when a surface mutation cannot be proven safe."""


class SurfaceDispositionStatus(StrEnum):
    UNCHANGED = "unchanged"
    REWRITTEN = "rewritten"
    REMOVED = "removed"
    PRESERVED = "preserved"
    UNSUPPORTED = "unsupported"
    DISAPPEARED = "disappeared"
    CREATED = "created"


@dataclass(frozen=True)
class SurfaceReplacement:
    surface_id: str
    value: str
    expected_value_sha256: str


@dataclass(frozen=True)
class SurfaceDisposition:
    surface_id: str
    status: SurfaceDispositionStatus
    before_value_sha256: str | None
    after_value_sha256: str | None


@dataclass(frozen=True)
class SurfaceMutationResult:
    data: bytes
    inventory: DocxInventory
    dispositions: tuple[SurfaceDisposition, ...]

    @property
    def unresolved(self) -> tuple[SurfaceDisposition, ...]:
        return tuple(
            disposition
            for disposition in self.dispositions
            if disposition.status is not SurfaceDispositionStatus.REWRITTEN
        )


def apply_surface_replacements(
    data: bytes,
    replacements: list[SurfaceReplacement],
) -> SurfaceMutationResult:
    """Apply exact neutral surface mutations and verify them after round-trip."""
    if not all(isinstance(item, SurfaceReplacement) for item in replacements):
        raise TypeError("replacements must contain only SurfaceReplacement instances")
    ids = [item.surface_id for item in replacements]
    if len(ids) != len(set(ids)):
        raise SurfaceMutationError("surface replacement IDs must be unique")

    before = inventory_docx(data)
    before_by_id = {surface.surface_id: surface for surface in before.surfaces}
    resolved: list[tuple[DocumentSurface, SurfaceReplacement]] = []
    for replacement in replacements:
        surface = before_by_id.get(replacement.surface_id)
        if surface is None:
            raise SurfaceMutationError(f"unknown surface: {replacement.surface_id}")
        if surface.capability is not SurfaceCapability.VALUE_REPLACE:
            raise SurfaceMutationError(
                f"surface is not value-replaceable: {replacement.surface_id} ({surface.capability})"
            )
        if replacement.expected_value_sha256 != surface.value_sha256:
            raise SurfaceMutationError(
                f"surface value changed before mutation: {replacement.surface_id}"
            )
        resolved.append((surface, replacement))

    output = _rewrite_parts(data, resolved)
    after = inventory_docx(output)
    after_by_id = {surface.surface_id: surface for surface in after.surfaces}
    dispositions: list[SurfaceDisposition] = []
    for surface, replacement in resolved:
        after_surface = after_by_id.get(surface.surface_id)
        expected_hash = sha256(replacement.value.encode("utf-8")).hexdigest()
        if after_surface is None:
            status = SurfaceDispositionStatus.DISAPPEARED
            after_hash = None
        elif after_surface.value_sha256 == expected_hash:
            status = SurfaceDispositionStatus.REWRITTEN
            after_hash = after_surface.value_sha256
        else:
            status = SurfaceDispositionStatus.UNCHANGED
            after_hash = after_surface.value_sha256
        dispositions.append(
            SurfaceDisposition(
                surface_id=surface.surface_id,
                status=status,
                before_value_sha256=surface.value_sha256,
                after_value_sha256=after_hash,
            )
        )

    result = SurfaceMutationResult(
        data=output,
        inventory=after,
        dispositions=tuple(dispositions),
    )
    if result.unresolved:
        unresolved = ", ".join(item.surface_id for item in result.unresolved)
        raise SurfaceMutationError(
            f"surface mutation was not confirmed after round-trip: {unresolved}"
        )
    return result


def _rewrite_parts(
    data: bytes,
    resolved: list[tuple[DocumentSurface, SurfaceReplacement]],
) -> bytes:
    by_part: dict[str, list[tuple[DocumentSurface, SurfaceReplacement]]] = {}
    for item in resolved:
        by_part.setdefault(item[0].part_name, []).append(item)

    source_buffer = BytesIO(data)
    output_buffer = BytesIO()
    with ZipFile(source_buffer) as source, ZipFile(output_buffer, "w", ZIP_DEFLATED) as output:
        for info in source.infolist():
            payload = source.read(info.filename)
            mutations = by_part.get(info.filename)
            if mutations:
                payload = _rewrite_xml_part(payload, mutations)
            output.writestr(info, payload)
    return output_buffer.getvalue()


def _rewrite_xml_part(
    payload: bytes,
    mutations: list[tuple[DocumentSurface, SurfaceReplacement]],
) -> bytes:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
    root = etree.fromstring(payload, parser=parser)
    tree = root.getroottree()
    for surface, replacement in mutations:
        if surface.kind is SurfaceKind.RELATIONSHIP:
            element = next(
                (candidate for candidate in root if candidate.get("Id") == surface.relationship_id),
                None,
            )
            if element is None:
                raise SurfaceMutationError(f"relationship disappeared: {surface.surface_id}")
            element.set("Target", replacement.value)
            continue

        if not surface.xml_path:
            raise SurfaceMutationError(f"surface has no XML locator: {surface.surface_id}")
        matches = tree.xpath(surface.xml_path, namespaces=_xpath_namespaces(root))
        if len(matches) != 1 or not isinstance(matches[0], etree._Element):
            raise SurfaceMutationError(f"surface XML locator drifted: {surface.surface_id}")
        element = matches[0]
        if surface.kind in {SurfaceKind.TEXT, SurfaceKind.XML_TEXT}:
            element.text = replacement.value
        elif surface.kind is SurfaceKind.XML_ATTRIBUTE and surface.xml_name:
            raw_name = next(
                (name for name in element.attrib if _local_name(name) == surface.xml_name),
                None,
            )
            if raw_name is None:
                raise SurfaceMutationError(f"surface attribute disappeared: {surface.surface_id}")
            element.set(raw_name, replacement.value)
        else:
            raise SurfaceMutationError(f"unsupported surface kind: {surface.surface_id}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _xpath_namespaces(root: etree._Element) -> dict[str, str]:
    return {prefix: uri for prefix, uri in root.nsmap.items() if prefix is not None}


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]
