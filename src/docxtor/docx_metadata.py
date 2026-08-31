"""DOCX package metadata policy for anonymized artifacts.

An anonymized DOCX is a public artifact, so its package properties must not
carry identity from the source document.  The policy is deliberately
lossy:

* every value in ``docProps/core.xml`` is cleared, including creator,
  last-modified-by, timestamps, revision, and descriptive properties;
* every value in ``docProps/app.xml`` is cleared, including company and
  manager; and
* every other ``docProps/*`` part is removed, which covers custom properties
  and thumbnails together with their package relationships.

Archive restore is different by contract: it is an explicit, vault-backed
operation that copies the original artifact and records a restore audit.  The
anonymized and reinjected artifacts always pass through this policy; metadata
is never restored implicitly during reinjection.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

_CORE_PART = "docProps/core.xml"
_APP_PART = "docProps/app.xml"
_METADATA_PREFIX = "docProps/"
_PACKAGE_RELATIONSHIPS = "_rels/.rels"
_CONTENT_TYPES = "[Content_Types].xml"


@dataclass(frozen=True, slots=True)
class MetadataInspection:
    """Safe-to-report metadata policy violations in a DOCX package."""

    violations: tuple[str, ...] = ()

    @property
    def clean(self) -> bool:
        return not self.violations


def sanitize_docx_metadata(
    path: str | Path, *, review_author: str = "Reviewer", review_initials: str = "RV"
) -> None:
    """Strip identifying and user-controlled metadata from a DOCX in place.

    The rewrite is performed through a sibling temporary file and atomically
    replaced only after the complete package has been written.  Malformed XML
    raises instead of leaving an unproven artifact in place.
    """
    source = Path(path)
    with ZipFile(source) as archive:
        members = [(info, archive.read(info.filename)) for info in archive.infolist()]

    removed_parts = {
        info.filename
        for info, _ in members
        if info.filename.startswith(_METADATA_PREFIX)
        and info.filename not in {_CORE_PART, _APP_PART}
    }
    removed_targets = {name.removeprefix("/") for name in removed_parts}

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=source.parent,
            prefix=f".{source.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)

        assert temporary_path is not None
        with ZipFile(temporary_path, mode="w", compression=ZIP_DEFLATED) as output:
            for info, data in members:
                name = info.filename
                if name in removed_parts:
                    continue
                if name in {_CORE_PART, _APP_PART}:
                    data = _clear_xml_values(data)
                elif name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name:
                    data = _sanitize_review_metadata(
                        data, review_author=review_author, review_initials=review_initials
                    )
                elif name == _PACKAGE_RELATIONSHIPS:
                    data = _remove_package_relationships(data, removed_targets)
                elif name == _CONTENT_TYPES:
                    data = _remove_content_type_overrides(data, removed_targets)
                rewritten_info = ZipInfo(name, date_time=info.date_time)
                rewritten_info.compress_type = ZIP_DEFLATED
                rewritten_info.external_attr = info.external_attr
                output.writestr(rewritten_info, data)
        os.replace(temporary_path, source)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def inspect_docx_metadata(path: str | Path) -> MetadataInspection:
    """Return metadata policy violations without exposing metadata values.

    The inspection intentionally reports only package part/element names.  A
    processing report must not become a second channel through which source PII
    is disclosed.
    """
    violations: list[str] = []
    try:
        with ZipFile(path) as archive:
            names = set(archive.namelist())
            for name in sorted(names):
                if not name.startswith(_METADATA_PREFIX):
                    continue
                if name not in {_CORE_PART, _APP_PART}:
                    # Do not echo an attacker-controlled part name: package
                    # member names are another possible PII channel.
                    violations.append("docProps:unexpected_part")
                    continue
                try:
                    root = ET.fromstring(archive.read(name))
                except (KeyError, ET.ParseError):
                    violations.append(f"{name}:unreadable")
                    continue
                for element in root.iter():
                    if isinstance(element.text, str) and element.text.strip():
                        local_name = _local_name(element.tag)
                        violations.append(f"{name}:{local_name}")
    except (OSError, ValueError):
        violations.append("DOCX package:unreadable")
    return MetadataInspection(violations=tuple(dict.fromkeys(violations)))


def _clear_xml_values(data: bytes) -> bytes:
    root = ET.fromstring(data)
    for element in root.iter():
        element.text = None
        element.tail = None
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sanitize_review_metadata(data: bytes, *, review_author: str, review_initials: str) -> bytes:
    root = ET.fromstring(data)
    changed = False
    for element in root.iter():
        local = _local_name(element.tag)
        if local in {"comment", "ins", "del"}:
            for attribute in list(element.attrib):
                if _local_name(attribute) == "author":
                    element.set(attribute, review_author)
                    changed = True
            if local == "comment":
                for attribute in list(element.attrib):
                    if _local_name(attribute) == "initials":
                        element.set(attribute, review_initials)
                        changed = True
        elif local == "person":
            for attribute in list(element.attrib):
                if _local_name(attribute) == "author":
                    element.set(attribute, review_author)
                    changed = True
                elif _local_name(attribute) in {"userId", "providerId"}:
                    del element.attrib[attribute]
                    changed = True
    return ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else data


def _remove_package_relationships(data: bytes, removed_parts: set[str]) -> bytes:
    root = ET.fromstring(data)
    for relationship in list(root):
        target = relationship.attrib.get("Target", "").lstrip("/")
        if target in removed_parts:
            root.remove(relationship)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_content_type_overrides(data: bytes, removed_parts: set[str]) -> bytes:
    root = ET.fromstring(data)
    for override in list(root):
        part_name = override.attrib.get("PartName", "").lstrip("/")
        if part_name in removed_parts:
            root.remove(override)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
