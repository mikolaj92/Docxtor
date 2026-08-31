from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from enum import StrEnum
from hashlib import sha256
from pathlib import PurePosixPath

from lxml import etree

from .docx_inventory import (
    DocumentSurface,
    DocxInventory,
    InventoryCoverage,
    SurfaceKind,
    inventory_docx,
)
from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries


class PackageMutationError(PackageError):
    """The requested package transaction cannot be proven complete and safe."""


class PackageMutationKind(StrEnum):
    REPLACE_SURFACE = "replace_surface"
    REMOVE_SURFACE = "remove_surface"
    REMOVE_RELATIONSHIP = "remove_relationship"
    REPLACE_PART = "replace_part"
    REMOVE_PART = "remove_part"


class PackageDispositionStatus(StrEnum):
    REWRITTEN = "rewritten"
    REMOVED = "removed"
    PRESERVED = "preserved"
    CREATED = "created"
    UNEXPECTED = "unexpected"


@dataclass(frozen=True)
class PackageMutation:
    operation_id: str
    kind: PackageMutationKind
    target_id: str
    expected_sha256: str
    value: str | bytes | None = None
    cascade: bool = False
    content_type: str | None = None


@dataclass(frozen=True)
class PackageDisposition:
    identity: str
    status: PackageDispositionStatus
    before_sha256: str | None
    after_sha256: str | None
    operation_id: str | None = None


@dataclass(frozen=True)
class PackageTransactionReceipt:
    data: bytes
    inventory_before: DocxInventory
    inventory_after: DocxInventory
    operations: tuple[PackageDisposition, ...]
    surfaces: tuple[PackageDisposition, ...]
    parts: tuple[PackageDisposition, ...]

    @property
    def unexpected(self) -> tuple[PackageDisposition, ...]:
        return tuple(
            item
            for item in self.surfaces + self.parts
            if item.status is PackageDispositionStatus.UNEXPECTED
        )


PackageValidator = Callable[[PackageTransactionReceipt], None]


def apply_package_transaction(
    data: bytes,
    mutations: Sequence[PackageMutation],
    *,
    validators: Sequence[PackageValidator] = (),
    require_complete_coverage: bool = True,
) -> PackageTransactionReceipt:
    """Apply text-independent package operations as one verified transaction.

    No intermediate bytes are returned. Every source and result surface and part
    receives a disposition. Validators run over the complete receipt before the
    caller can publish ``receipt.data``.
    """
    if len({item.operation_id for item in mutations}) != len(mutations):
        raise PackageMutationError("package mutation operation IDs must be unique")
    before = inventory_docx(data)
    if require_complete_coverage and before.coverage is InventoryCoverage.INCOMPLETE:
        raise PackageMutationError("source package coverage is incomplete")
    entries = {entry.name: entry for entry in read_package_entries(data)}
    before_surfaces = {item.surface_id: item for item in before.surfaces}
    before_parts = {item.name: item for item in before.parts}
    operation_results: list[PackageDisposition] = []
    touched_surfaces: dict[str, str] = {}
    touched_parts: dict[str, str] = {}

    for mutation in mutations:
        if mutation.kind in {
            PackageMutationKind.REPLACE_SURFACE,
            PackageMutationKind.REMOVE_SURFACE,
        }:
            surface = before_surfaces.get(mutation.target_id)
            if surface is None:
                raise PackageMutationError(f"unknown surface: {mutation.target_id}")
            _expect_hash(mutation, surface.value_sha256)
            _mutate_surface(entries, surface, mutation)
            touched_surfaces[surface.surface_id] = mutation.operation_id
            operation_results.append(
                PackageDisposition(
                    surface.surface_id,
                    PackageDispositionStatus.REWRITTEN
                    if mutation.kind is PackageMutationKind.REPLACE_SURFACE
                    else PackageDispositionStatus.REMOVED,
                    surface.value_sha256,
                    _value_hash(mutation.value),
                    mutation.operation_id,
                )
            )
        elif mutation.kind is PackageMutationKind.REMOVE_RELATIONSHIP:
            surface = before_surfaces.get(mutation.target_id)
            if surface is None or surface.kind is not SurfaceKind.RELATIONSHIP:
                raise PackageMutationError(f"unknown relationship surface: {mutation.target_id}")
            _expect_hash(mutation, surface.value_sha256)
            _remove_relationship(entries, surface)
            touched_surfaces[surface.surface_id] = mutation.operation_id
            operation_results.append(
                PackageDisposition(
                    surface.surface_id,
                    PackageDispositionStatus.REMOVED,
                    surface.value_sha256,
                    None,
                    mutation.operation_id,
                )
            )
        elif mutation.kind is PackageMutationKind.REMOVE_PART:
            part = before_parts.get(mutation.target_id)
            if part is None:
                raise PackageMutationError(f"unknown package part: {mutation.target_id}")
            _expect_hash(mutation, part.sha256)
            incoming = _incoming_relationships(entries, mutation.target_id)
            removed = _remove_part(entries, mutation.target_id, cascade=mutation.cascade)
            for name in removed:
                touched_parts[name] = mutation.operation_id
            for surface in before.surfaces:
                if surface.part_name in removed or (
                    surface.kind is SurfaceKind.RELATIONSHIP
                    and (surface.part_name, surface.relationship_id or "") in incoming
                ):
                    touched_surfaces[surface.surface_id] = mutation.operation_id
            for relationship_part, _relationship_id in incoming:
                touched_parts[relationship_part] = mutation.operation_id
            touched_parts["[Content_Types].xml"] = mutation.operation_id
            operation_results.append(
                PackageDisposition(
                    mutation.target_id,
                    PackageDispositionStatus.REMOVED,
                    part.sha256,
                    None,
                    mutation.operation_id,
                )
            )
        elif mutation.kind is PackageMutationKind.REPLACE_PART:
            part = before_parts.get(mutation.target_id)
            if part is None:
                raise PackageMutationError(f"unknown package part: {mutation.target_id}")
            _expect_hash(mutation, part.sha256)
            if not isinstance(mutation.value, bytes):
                raise PackageMutationError("replace_part requires bytes")
            source = entries[mutation.target_id]
            entries[mutation.target_id] = PackageEntry(
                source.name,
                mutation.value,
                source.compress_type,
                source.external_attr,
                source.internal_attr,
                source.create_system,
            )
            touched_parts[mutation.target_id] = mutation.operation_id
            operation_results.append(
                PackageDisposition(
                    mutation.target_id,
                    PackageDispositionStatus.REWRITTEN,
                    part.sha256,
                    sha256(mutation.value).hexdigest(),
                    mutation.operation_id,
                )
            )
        else:
            raise PackageMutationError(f"unsupported package mutation: {mutation.kind}")

    output = _package_bytes(tuple(entries.values()))
    after = inventory_docx(output)
    if require_complete_coverage and after.coverage is InventoryCoverage.INCOMPLETE:
        raise PackageMutationError("result package coverage is incomplete")
    receipt = _receipt(
        output,
        before,
        after,
        tuple(operation_results),
        touched_surfaces,
        touched_parts,
    )
    if receipt.unexpected:
        identities = ", ".join(item.identity for item in receipt.unexpected[:5])
        raise PackageMutationError(f"unexpected package change: {identities}")
    _assert_operations(receipt)
    for validator in validators:
        validator(receipt)
    return receipt


def _expect_hash(mutation: PackageMutation, actual: str) -> None:
    if mutation.expected_sha256 != actual:
        raise PackageMutationError(f"target changed before mutation: {mutation.target_id}")


def _value_hash(value: str | bytes | None) -> str | None:
    if value is None:
        return None
    payload = value.encode() if isinstance(value, str) else value
    return sha256(payload).hexdigest()


def _mutate_surface(
    entries: dict[str, PackageEntry], surface: DocumentSurface, mutation: PackageMutation
) -> None:
    entry = entries[surface.part_name]
    root = parse_package_xml(entry.data, part_name=entry.name)
    tree = root.getroottree()
    if not surface.xml_path:
        raise PackageMutationError(f"surface has no XML locator: {surface.surface_id}")
    matches = tree.xpath(surface.xml_path, namespaces={k: v for k, v in root.nsmap.items() if k})
    if len(matches) != 1 or not isinstance(matches[0], etree._Element):
        raise PackageMutationError(f"surface locator drifted: {surface.surface_id}")
    element = matches[0]
    if mutation.kind is PackageMutationKind.REMOVE_SURFACE:
        parent = element.getparent()
        if surface.kind in {SurfaceKind.TEXT, SurfaceKind.XML_TEXT}:
            element.text = None
        elif surface.kind is SurfaceKind.XML_ATTRIBUTE and surface.attribute_qname:
            element.attrib.pop(surface.attribute_qname, None)
        elif parent is not None:
            parent.remove(element)
        else:
            raise PackageMutationError(f"surface cannot be removed: {surface.surface_id}")
    else:
        if not isinstance(mutation.value, str):
            raise PackageMutationError("surface replacement requires a string value")
        if surface.kind in {SurfaceKind.TEXT, SurfaceKind.XML_TEXT}:
            element.text = mutation.value
        elif surface.kind is SurfaceKind.XML_ATTRIBUTE and surface.attribute_qname:
            element.set(surface.attribute_qname, mutation.value)
        elif surface.kind is SurfaceKind.RELATIONSHIP:
            element.set("Target", mutation.value)
        else:
            raise PackageMutationError(f"surface cannot be replaced: {surface.surface_id}")
    _set_payload(
        entries,
        entry,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _remove_relationship(entries: dict[str, PackageEntry], surface: DocumentSurface) -> None:
    entry = entries[surface.part_name]
    root = parse_package_xml(entry.data, part_name=entry.name)
    match = next((item for item in root if item.get("Id") == surface.relationship_id), None)
    if match is None:
        raise PackageMutationError(f"relationship disappeared: {surface.surface_id}")
    root.remove(match)
    _set_payload(
        entries,
        entry,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _remove_part(entries: dict[str, PackageEntry], name: str, *, cascade: bool) -> tuple[str, ...]:
    if not cascade:
        incoming = _incoming_relationships(entries, name)
        if incoming:
            raise PackageMutationError(f"part has incoming relationships: {name}")
    removed = {name}
    rel_part = _relationship_part_for(name)
    if rel_part in entries:
        removed.add(rel_part)
    if cascade:
        for relationship_part, relationship_id in _incoming_relationships(entries, name):
            _remove_relationship_by_id(entries, relationship_part, relationship_id)
    for item in removed:
        entries.pop(item, None)
    _remove_content_type(entries, name)
    return tuple(sorted(removed))


def _incoming_relationships(
    entries: dict[str, PackageEntry], target_name: str
) -> list[tuple[str, str]]:
    incoming: list[tuple[str, str]] = []
    for name, entry in entries.items():
        if not name.endswith(".rels"):
            continue
        root = parse_package_xml(entry.data, part_name=name)
        source = _source_part_for_relationship(name)
        for relationship in root:
            if relationship.get("TargetMode", "").lower() == "external":
                continue
            target = relationship.get("Target")
            if target and _resolve_target(source, target) == target_name:
                incoming.append((name, relationship.get("Id", "")))
    return incoming


def _remove_relationship_by_id(
    entries: dict[str, PackageEntry], name: str, relationship_id: str
) -> None:
    entry = entries[name]
    root = parse_package_xml(entry.data, part_name=name)
    match = next((item for item in root if item.get("Id") == relationship_id), None)
    if match is not None:
        root.remove(match)
        _set_payload(
            entries,
            entry,
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
        )


def _remove_content_type(entries: dict[str, PackageEntry], part_name: str) -> None:
    entry = entries["[Content_Types].xml"]
    root = parse_package_xml(entry.data, part_name=entry.name)
    for item in list(root):
        if item.get("PartName", "").lstrip("/") == part_name:
            root.remove(item)
    _set_payload(
        entries,
        entry,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def _set_payload(entries: dict[str, PackageEntry], source: PackageEntry, payload: bytes) -> None:
    entries[source.name] = PackageEntry(
        source.name,
        payload,
        source.compress_type,
        source.external_attr,
        source.internal_attr,
        source.create_system,
    )


def _relationship_part_for(part_name: str) -> str:
    path = PurePosixPath(part_name)
    return str(path.parent / "_rels" / (path.name + ".rels"))


def _source_part_for_relationship(relationship_part: str) -> str:
    if relationship_part == "_rels/.rels":
        return ""
    path = PurePosixPath(relationship_part)
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _resolve_target(source: str, target: str) -> str:
    base = PurePosixPath(source).parent if source else PurePosixPath()
    stack: list[str] = []
    for piece in (base / target).parts:
        if piece in {"", "."}:
            continue
        if piece == "..":
            if stack:
                stack.pop()
        else:
            stack.append(piece)
    return "/".join(stack)


def _package_bytes(entries: tuple[PackageEntry, ...]) -> bytes:
    from io import BytesIO
    from zipfile import ZipFile

    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for entry in entries:
            archive.writestr(entry.zip_info(), entry.data)
    data = output.getvalue()
    read_package_entries(data)
    return data


def _receipt(
    data: bytes,
    before: DocxInventory,
    after: DocxInventory,
    operations: tuple[PackageDisposition, ...],
    touched_surfaces: dict[str, str],
    touched_parts: dict[str, str],
) -> PackageTransactionReceipt:
    before_surfaces = {item.surface_id: item for item in before.surfaces}
    after_surfaces = {item.surface_id: item for item in after.surfaces}
    surface_dispositions: list[PackageDisposition] = []
    for identity in sorted(before_surfaces.keys() | after_surfaces.keys()):
        left = before_surfaces.get(identity)
        right = after_surfaces.get(identity)
        operation_id = touched_surfaces.get(identity)
        if left and right and left.value_sha256 == right.value_sha256:
            status = PackageDispositionStatus.PRESERVED
        elif operation_id and right:
            status = PackageDispositionStatus.REWRITTEN
        elif operation_id and not right:
            status = PackageDispositionStatus.REMOVED
        elif not left and right:
            status = PackageDispositionStatus.CREATED
        else:
            status = PackageDispositionStatus.UNEXPECTED
        surface_dispositions.append(
            PackageDisposition(
                identity,
                status,
                left.value_sha256 if left else None,
                right.value_sha256 if right else None,
                operation_id,
            )
        )
    before_parts = {item.name: item for item in before.parts}
    after_parts = {item.name: item for item in after.parts}
    part_dispositions: list[PackageDisposition] = []
    for identity in sorted(before_parts.keys() | after_parts.keys()):
        left = before_parts.get(identity)
        right = after_parts.get(identity)
        operation_id = touched_parts.get(identity)
        if left and right and left.sha256 == right.sha256:
            status = PackageDispositionStatus.PRESERVED
        elif operation_id and right:
            status = PackageDispositionStatus.REWRITTEN
        elif operation_id and not right:
            status = PackageDispositionStatus.REMOVED
        elif not left and right:
            status = PackageDispositionStatus.CREATED
        elif _part_change_explained(identity, surface_dispositions):
            status = PackageDispositionStatus.REWRITTEN
        else:
            status = PackageDispositionStatus.UNEXPECTED
        part_dispositions.append(
            PackageDisposition(
                identity,
                status,
                left.sha256 if left else None,
                right.sha256 if right else None,
                operation_id,
            )
        )
    return PackageTransactionReceipt(
        data,
        before,
        after,
        operations,
        tuple(surface_dispositions),
        tuple(part_dispositions),
    )


def _part_change_explained(part_name: str, dispositions: list[PackageDisposition]) -> bool:
    return (
        any(
            item.operation_id is not None
            and item.identity.startswith((f"xml:{part_name}:", f"relationship:{part_name}:"))
            for item in dispositions
        )
        or part_name == "[Content_Types].xml"
    )


def _assert_operations(receipt: PackageTransactionReceipt) -> None:
    surface_by_id = {item.identity: item for item in receipt.surfaces}
    part_by_id = {item.identity: item for item in receipt.parts}
    for operation in receipt.operations:
        actual = surface_by_id.get(operation.identity) or part_by_id.get(operation.identity)
        if actual is None or actual.status not in {
            PackageDispositionStatus.REWRITTEN,
            PackageDispositionStatus.REMOVED,
        }:
            raise PackageMutationError(
                f"operation has no confirmed disposition: {operation.operation_id}"
            )
