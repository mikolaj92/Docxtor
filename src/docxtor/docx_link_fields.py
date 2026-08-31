"""Rewrite leftover hyperlink and customXml URIs on a working copy.

Word stores click targets in fields, relationship-backed spans, and custom XML.
This transaction flattens those package forms on a working copy.

The fail-closed gates stay: DrawingML ``hlinkClick`` / ``hlinkHover``, VML
``href``, orphaned external rels, linked images, attachedTemplate, nested
or unmatched fields, and unreadable XML still block the parser. Word TOC
``HYPERLINK \\l _Toc[0-9]+`` and TOC ``w:hyperlink`` anchors without an
``r:id`` are left for Docxtor. The caller chooses the working-copy path; malformed XML fails closed.
"""

from __future__ import annotations

import io
import os
import re
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

_HYPERLINK_FIELD_TOKEN = re.compile(r"\b(?:HYPERLINK|LINK)\b", re.IGNORECASE)
_INTERNAL_TOC_HYPERLINK_FIELD = re.compile(
    r'\s*(?:HYPERLINK|LINK)\s+\\l\s+"?_Toc[0-9]+"?\s*', re.IGNORECASE
)
_URL_IN_INSTRUCTION = re.compile(r"(?:mailto:|https?://|file:|ftp://)", re.IGNORECASE)


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _attribute(element: ET.Element, local: str) -> str | None:
    return next((value for key, value in element.attrib.items() if _local_name(key) == local), None)


def _value_has_leftover_uri(value: str) -> bool:
    return _URL_IN_INSTRUCTION.search(value) is not None


_XMLNS_NS = "http://www.w3.org/2000/xmlns/"
_CUSTOMXML_SCHEMA_ATTR_LOCALS = frozenset(
    {"schemalocation", "targetnamespace", "nonamespaceschemalocation"}
)
_CUSTOMXML_SCHEMA_URI_ATTR_LOCALS = frozenset({"uri", "namespace"})
_CUSTOMXML_SCHEMA_ELEMENT_LOCALS = frozenset(
    {"schemaref", "schema", "datastoreitem", "import", "include", "redefine"}
)


def _is_xmlns_attribute(key: str) -> bool:
    return key == "xmlns" or key.startswith("xmlns:") or key.startswith("{" + _XMLNS_NS + "}")


def _is_customxml_schema_identifier_attr(element: ET.Element, key: str) -> bool:
    if _is_xmlns_attribute(key):
        return True
    local = _local_name(key).lower()
    if local in _CUSTOMXML_SCHEMA_ATTR_LOCALS:
        return True
    return (
        local in _CUSTOMXML_SCHEMA_URI_ATTR_LOCALS
        and _local_name(element.tag).lower() in _CUSTOMXML_SCHEMA_ELEMENT_LOCALS
    )


def _is_supported_story_part(name: str) -> bool:
    return name == "word/document.xml" or (
        name.startswith(("word/header", "word/footer")) and name.endswith(".xml")
    )


def _is_customxml_part(name: str) -> bool:
    return name.startswith("customXml/") and name.endswith(".xml")


_FIELD_INSTRUCTION_LOCALS = frozenset({"instrText", "delInstrText"})
_FIELD_BEGIN = "begin"
_FIELD_SEPARATE = "separate"
_FIELD_END = "end"
_HYPERLINK_REL_SUFFIX = "/hyperlink"
_RESERVED_ET_PREFIX = re.compile(r"ns\d+$")
_REVISION_WRAPPERS = frozenset({"ins", "del"})
_SKIP_FIELD_SLOT_LOCALS = frozenset(
    {"pPr", "rPr", "sectPr", "tblPr", "trPr", "tcPr", "sdtPr", "sdtEndPr"}
)


def instruction_is_strippable_hyperlink_field(instruction: str) -> bool:
    """True for HYPERLINK/LINK fields that still hold a leftover URI.

    Word TOC ``HYPERLINK \\l _Toc[0-9]+`` stays for Docxtor. Free-form
    ``\\l`` anchors and opaque switches stay fail-closed at the parser
    gate. Only a HYPERLINK/LINK instruction with mailto / http(s) / file /
    ftp is flattened to its display run.
    """
    if _INTERNAL_TOC_HYPERLINK_FIELD.fullmatch(instruction):
        return False
    if _HYPERLINK_FIELD_TOKEN.search(instruction) is None:
        return False
    return _URL_IN_INSTRUCTION.search(instruction) is not None


def flatten_link_fields(path: str | Path) -> bool:
    """Flatten leftover hyperlink fields, ``w:hyperlink`` rels, and customXml URIs.

    Returns True when the package was rewritten. Malformed story XML raises
    instead of leaving an unproven working copy. Unreadable customXml is
    left for the fail-closed parser gate.
    """
    source = Path(path)
    with ZipFile(source) as archive:
        members = [(info, archive.read(info.filename)) for info in archive.infolist()]

    rewritten: list[tuple[ZipInfo, bytes]] = []
    changed = False
    dropped_rids: dict[str, set[str]] = {}
    for info, data in members:
        name = info.filename
        if _is_supported_story_part(name):
            new_data, part_changed, rids = _strip_story_part(data)
            if part_changed:
                data = new_data
                changed = True
            if rids:
                rels_name = _relationship_part_for_story(name)
                if rels_name is not None:
                    dropped_rids.setdefault(rels_name, set()).update(rids)
        rewritten.append((info, data))

    members = rewritten
    rewritten = []
    for info, data in members:
        name = info.filename
        if name in dropped_rids:
            new_data, part_changed = _drop_hyperlink_relationships(data, dropped_rids[name])
            if part_changed:
                data = new_data
                changed = True
        elif _is_customxml_part(name):
            new_data, part_changed = _blank_customxml_leftover_uris(data)
            if part_changed:
                data = new_data
                changed = True
        rewritten.append((info, data))

    if not changed:
        return False

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=source.parent,
            prefix=f".{source.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
        with ZipFile(temporary_path, "w") as output:
            for info, data in rewritten:
                rewritten_info = ZipInfo(info.filename, date_time=info.date_time)
                rewritten_info.compress_type = ZIP_DEFLATED
                rewritten_info.external_attr = info.external_attr
                output.writestr(rewritten_info, data)
        os.replace(temporary_path, source)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    return True


def _relationship_part_for_story(story: str) -> str | None:
    """Return the OPC rels part that holds hyperlink Targets for ``story``."""
    if "/_rels/" in story or not story.endswith(".xml"):
        return None
    if "/" not in story:
        return f"_rels/{story}.rels"
    parent, filename = story.rsplit("/", 1)
    return f"{parent}/_rels/{filename}.rels"


def _strip_story_part(data: bytes) -> tuple[bytes, bool, set[str]]:
    root = ET.fromstring(data)
    changed = _strip_element(root)
    rids: set[str] = set()
    if _unwrap_relationship_hyperlinks(root, rids):
        changed = True
    if not changed:
        return data, False, rids
    _register_source_namespaces(data)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True, rids


def _register_source_namespaces(data: bytes) -> None:
    """Keep the source prefixes so a rewrite does not emit ns0: markup.

    CPython reserves ``ns0``, ``ns1``, ... for generated prefixes. Word
    documents that already declare those names must not crash the rewrite
    (#5220). The empty default xmlns stays registered.
    """
    for _event, (prefix, uri) in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
        name = prefix or ""
        if _RESERVED_ET_PREFIX.fullmatch(name):
            continue
        ET.register_namespace(name, uri)


def _strip_element(elem: ET.Element) -> bool:
    changed = _strip_sibling_fields(elem)
    for child in list(elem):
        if _strip_element(child):
            changed = True
    return changed


def _unwrap_relationship_hyperlinks(parent: ET.Element, rids: set[str]) -> bool:
    """Replace relationship-backed ``w:hyperlink`` with its display children."""
    changed = False
    index = 0
    while index < len(parent):
        child = parent[index]
        if _unwrap_relationship_hyperlinks(child, rids):
            changed = True
        if _local_name(child.tag) == "hyperlink":
            rid = _attribute(child, "id")
            if rid:
                rids.add(rid)
                children = list(child)
                parent[index : index + 1] = children
                changed = True
                index += len(children)
                continue
        index += 1
    return changed


def _drop_hyperlink_relationships(data: bytes, rids: set[str]) -> tuple[bytes, bool]:
    """Drop hyperlink Relationships whose Id was unwrapped from story markup."""
    if not rids:
        return data, False
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data, False
    changed = False
    for child in list(root):
        if _local_name(child.tag) != "Relationship":
            continue
        rid = _attribute(child, "Id")
        rel_type = _attribute(child, "Type") or ""
        if rid in rids and rel_type.endswith(_HYPERLINK_REL_SUFFIX):
            root.remove(child)
            changed = True
    if not changed:
        return data, False
    _register_source_namespaces(data)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def _blank_customxml_leftover_uris(payload: bytes) -> tuple[bytes, bool]:
    """Blank leftover mailto / http / file in customXml item content.

    Schema identifiers stay. mailto / file on a schema identifier is left
    for the fail-closed parser. Unreadable XML is left untouched.
    """
    try:
        root = ET.fromstring(payload)
    except ET.ParseError:
        return payload, False
    changed = False
    for elem in root.iter():
        if elem.text and _value_has_leftover_uri(elem.text):
            elem.text = ""
            changed = True
        if elem.tail and _value_has_leftover_uri(elem.tail):
            elem.tail = ""
            changed = True
        for key, value in list(elem.attrib.items()):
            if not value:
                continue
            if _is_customxml_schema_identifier_attr(elem, key):
                continue
            if _value_has_leftover_uri(value):
                elem.set(key, "")
                changed = True
    if not changed:
        return payload, False
    _register_source_namespaces(payload)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), True


def _strip_sibling_fields(parent: ET.Element) -> bool:
    changed = False
    index = 0
    while index < len(parent):
        child = parent[index]
        if _local_name(child.tag) == "fldSimple":
            instruction = _attribute(child, "instr") or ""
            if instruction_is_strippable_hyperlink_field(instruction):
                display = list(child)
                parent[index : index + 1] = display
                changed = True
                index += len(display)
                continue
        index += 1

    slots = _field_slots(parent)
    slot_index = 0
    while slot_index < len(slots):
        container, child_index = slots[slot_index]
        child = container[child_index]
        if _fld_char_type(child) != _FIELD_BEGIN:
            slot_index += 1
            continue
        collected = _collect_field_slots(slots, slot_index)
        end_index, nested, saw_separate, instruction, drop = collected
        if end_index is None:
            slot_index += 1
            continue
        if not nested and saw_separate and instruction_is_strippable_hyperlink_field(instruction):
            _drop_field_slots(drop)
            changed = True
            slots = _field_slots(parent)
            slot_index = 0
            continue
        slot_index = end_index + 1
    return changed


def _field_slots(parent: ET.Element) -> list[tuple[ET.Element, int]]:
    """Document-order runs, including those wrapped in ``w:ins`` / ``w:del``."""
    slots: list[tuple[ET.Element, int]] = []
    for index, child in enumerate(parent):
        local = _local_name(child.tag)
        if local in _SKIP_FIELD_SLOT_LOCALS:
            continue
        if local in _REVISION_WRAPPERS:
            slots.extend(_field_slots(child))
            continue
        slots.append((parent, index))
    return slots


def _collect_field_slots(
    slots: list[tuple[ET.Element, int]], begin_index: int
) -> tuple[int | None, bool, bool, str, list[tuple[ET.Element, int]]]:
    """Return (end_index, nested, saw_separate, instruction, drop_slots)."""
    depth = 1
    nested = False
    saw_separate = False
    instruction_parts: list[str] = []
    drop: list[tuple[ET.Element, int]] = [slots[begin_index]]
    for index in range(begin_index + 1, len(slots)):
        container, child_index = slots[index]
        child = container[child_index]
        marker = _fld_char_type(child)
        if marker == _FIELD_BEGIN:
            depth += 1
            nested = True
            drop.append(slots[index])
            continue
        if marker == _FIELD_END:
            depth -= 1
            drop.append(slots[index])
            if depth == 0:
                return index, nested, saw_separate, "".join(instruction_parts), drop
            continue
        if depth != 1:
            drop.append(slots[index])
            continue
        if marker == _FIELD_SEPARATE:
            saw_separate = True
            drop.append(slots[index])
            continue
        if saw_separate:
            continue
        instruction_parts.extend(_instruction_texts(child))
        drop.append(slots[index])
    return None, nested, saw_separate, "".join(instruction_parts), drop


def _drop_field_slots(slots: list[tuple[ET.Element, int]]) -> None:
    grouped: dict[int, list[tuple[ET.Element, int]]] = {}
    for container, index in slots:
        grouped.setdefault(id(container), []).append((container, index))
    for items in grouped.values():
        container = items[0][0]
        for _, index in sorted(items, key=lambda item: item[1], reverse=True):
            del container[index]


def _fld_char_type(elem: ET.Element) -> str | None:
    """Return the marker on this run, not a nested ins/del field."""
    if _local_name(elem.tag) == "fldChar":
        return _attribute(elem, "fldCharType")
    for child in list(elem):
        if _local_name(child.tag) == "fldChar":
            return _attribute(child, "fldCharType")
    return None


def _instruction_texts(elem: ET.Element) -> list[str]:
    texts: list[str] = []
    for node in elem.iter():
        if _local_name(node.tag) in _FIELD_INSTRUCTION_LOCALS and node.text:
            texts.append(node.text)
    return texts
