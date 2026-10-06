"""Package surgery helpers for revision operations."""

from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from .docx_package import PackageEntry, parse_package_xml, read_package_entries
from .docx_revision_models import _tag

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _strip_comment_anchors(root: etree._Element) -> None:
    for name in ("commentRangeStart", "commentRangeEnd"):
        for element in list(root.iter(_tag(name))):
            _remove(element)
    for element in list(root.iter(_tag("commentReference"))):
        parent = element.getparent()
        _remove(parent if parent is not None and parent.tag == _tag("r") else element)


def _strip_comment_relationships(root: etree._Element, original: bytes) -> bytes:
    changed = False
    for rel in list(root):
        target = rel.get("Target", "").removeprefix("/")
        if rel.tag == f"{{{_REL_NS}}}Relationship" and (
            "comment" in rel.get("Type", "").lower()
            or target.startswith(("comments", "word/comments", "people.xml", "word/people.xml"))
        ):
            root.remove(rel)
            changed = True
    return _serialize(root) if changed else original


def _strip_comment_content_types(root: etree._Element, original: bytes) -> bytes:
    changed = False
    for override in list(root):
        if override.tag == f"{{{_CT_NS}}}Override" and _is_comment_part(
            override.get("PartName", "").removeprefix("/")
        ):
            root.remove(override)
            changed = True
    return _serialize(root) if changed else original

def _remove(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    if element.tail:
        previous = element.getprevious()
        if previous is not None:
            previous.tail = (previous.tail or "") + element.tail
        else:
            parent.text = (parent.text or "") + element.tail
    parent.remove(element)

def _unwrap(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    children = list(element)
    for child in children:
        element.addprevious(child)
    if element.tail:
        target = children[-1] if children else element.getprevious()
        if target is not None:
            target.tail = (target.tail or "") + element.tail
        else:
            parent.text = (parent.text or "") + element.tail
    parent.remove(element)

def _has_comments(data: bytes) -> bool:
    entries = read_package_entries(data)
    if any(_is_comment_part(e.name) for e in entries):
        return True
    for entry in entries:
        if _is_word_xml(entry.name):
            root = parse_package_xml(entry.data, part_name=entry.name)
            if any(
                next(root.iter(_tag(n)), None) is not None
                for n in ("commentReference", "commentRangeStart", "commentRangeEnd")
            ):
                return True
    return False

def _is_comment_part(name: str) -> bool:
    normalized = name.removeprefix("/")
    if not normalized.startswith("word/"):
        return False
    base = normalized.rsplit("/", 1)[-1]
    if base == "people.xml" or (base.startswith("comments") and base.endswith(".xml")):
        return True
    if normalized.startswith("word/_rels/") and base.endswith(".xml.rels"):
        source = base.removesuffix(".rels")
        return source == "people.xml" or (source.startswith("comments") and source.endswith(".xml"))
    return False

def _is_word_xml(name: str) -> bool:
    return name.startswith("word/") and name.endswith(".xml") and not _is_comment_part(name)

def _replace_entry(entry: PackageEntry, data: bytes) -> PackageEntry:
    return PackageEntry(
        entry.name,
        data,
        entry.compress_type,
        entry.external_attr,
        entry.internal_attr,
        entry.create_system,
    )

def _write_bytes(entries: list[PackageEntry]) -> bytes:
    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for entry in entries:
            archive.writestr(entry.zip_info(), entry.data)
    result = output.getvalue()
    # Canonical reader is also the postcondition validator.
    read_package_entries(result)
    return result

def _serialize(root: etree._Element) -> bytes:
    return (_XML_DECLARATION + etree.tostring(root, encoding="unicode")).encode()
