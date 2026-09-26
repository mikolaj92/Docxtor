"""Deterministic byte and semantic comparison of OPC package entries."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from hashlib import sha256

from lxml import etree

from .docx_package import PackageError, parse_package_xml, read_package_entries


class PackagePartChangeKind(StrEnum):
    ADDED = "added"
    REMOVED = "removed"
    BYTE_IDENTICAL = "byte_identical"
    BYTE_CHANGED = "byte_changed"
    SEMANTIC_XML_CHANGED = "semantic_xml_changed"
    MALFORMED_XML = "malformed_xml"


@dataclass(frozen=True)
class PackagePartChange:
    part_name: str
    kind: PackagePartChangeKind
    before_sha256: str | None
    after_sha256: str | None


@dataclass(frozen=True)
class PackageChangeReport:
    parts: tuple[PackagePartChange, ...]

    @property
    def added(self) -> tuple[PackagePartChange, ...]:
        return tuple(item for item in self.parts if item.kind is PackagePartChangeKind.ADDED)

    @property
    def removed(self) -> tuple[PackagePartChange, ...]:
        return tuple(item for item in self.parts if item.kind is PackagePartChangeKind.REMOVED)

    @property
    def byte_identical(self) -> tuple[PackagePartChange, ...]:
        return tuple(
            item for item in self.parts if item.kind is PackagePartChangeKind.BYTE_IDENTICAL
        )

    @property
    def byte_changed(self) -> tuple[PackagePartChange, ...]:
        return tuple(item for item in self.parts if item.kind is PackagePartChangeKind.BYTE_CHANGED)

    @property
    def semantically_changed_xml(self) -> tuple[PackagePartChange, ...]:
        return tuple(
            item for item in self.parts if item.kind is PackagePartChangeKind.SEMANTIC_XML_CHANGED
        )

    @property
    def malformed_xml(self) -> tuple[PackagePartChange, ...]:
        return tuple(
            item for item in self.parts if item.kind is PackagePartChangeKind.MALFORMED_XML
        )


def compare_docx_packages(before: str | bytes, after: str | bytes) -> PackageChangeReport:
    """Compare input package bytes before publication/preservation is applied."""
    left = {entry.name: entry.data for entry in read_package_entries(before, validate_xml=False)}
    right = {entry.name: entry.data for entry in read_package_entries(after, validate_xml=False)}
    changes: list[PackagePartChange] = []
    for name in sorted(left.keys() | right.keys()):
        old, new = left.get(name), right.get(name)
        old_hash = sha256(old).hexdigest() if old is not None else None
        new_hash = sha256(new).hexdigest() if new is not None else None
        if old is None:
            kind = PackagePartChangeKind.ADDED
        elif new is None:
            kind = PackagePartChangeKind.REMOVED
        elif old == new:
            kind = PackagePartChangeKind.BYTE_IDENTICAL
        elif name.endswith((".xml", ".rels")):
            old_xml, new_xml = _canonical(old), _canonical(new)
            if old_xml is None or new_xml is None:
                kind = PackagePartChangeKind.MALFORMED_XML
            elif old_xml != new_xml:
                kind = PackagePartChangeKind.SEMANTIC_XML_CHANGED
            else:
                kind = PackagePartChangeKind.BYTE_CHANGED
        else:
            kind = PackagePartChangeKind.BYTE_CHANGED
        changes.append(PackagePartChange(name, kind, old_hash, new_hash))
    return PackageChangeReport(tuple(changes))


def _canonical(data: bytes) -> bytes | None:
    try:
        root = parse_package_xml(data)
        return etree.tostring(root, method="c14n2", with_comments=True)
    except (PackageError, etree.C14NError):
        return None
