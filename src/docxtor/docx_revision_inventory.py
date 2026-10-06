"""Neutral inventory of tracked-change XML across Word stories."""

from __future__ import annotations

from lxml import etree

from .docx_package import PackageError, parse_package_xml, read_package_entries
from .docx_revision_models import (
    _STRICT_W,
    _W,
    _W14,
    Revision,
    RevisionCoverageDiagnostic,
    RevisionInventory,
    RevisionInventoryCoverage,
    RevisionKind,
    RevisionOperationError,
    _tag,
)
from .docx_revision_package import _is_word_xml
from .docx_revision_paragraphs import _is_paragraph_mark

_REVISION_NAMES = (
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "tblPrChange",
    "trPrChange",
    "tcPrChange",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tblGridChange",
    "tblPrExChange",
    "numberingChange",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
    "conflictIns",
    "conflictDel",
    "customXmlConflictInsRangeStart",
    "customXmlConflictInsRangeEnd",
    "customXmlConflictDelRangeStart",
    "customXmlConflictDelRangeEnd",
)
_UNSUPPORTED_RANGES = (
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
)

_FORMATTING_PROPERTY_CHANGES = frozenset({"pPrChange", "rPrChange"})
_EMPTY_CUSTOM_XML_RANGE_MARKERS = frozenset(
    {
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
    }
)
_BLOCK_REVISION_CHILDREN = frozenset({"p", "tbl", "tr", "tc"})
_PROPERTY_CHANGE_SNAPSHOT = {"pPrChange": "pPr", "rPrChange": "rPr"}

def inventory_revisions_bytes(data: bytes) -> RevisionInventory:
    """Return a neutral inventory of tracked-change XML in every Word XML story."""
    try:
        entries = read_package_entries(data)
    except PackageError as exc:
        raise RevisionOperationError(str(exc)) from exc
    revisions: list[Revision] = []
    diagnostics: list[RevisionCoverageDiagnostic] = []
    for entry in entries:
        if not _is_word_xml(entry.name):
            continue
        root = parse_package_xml(entry.data, part_name=entry.name)
        tree = root.getroottree()
        for namespace in (_STRICT_W, _W14):
            for raw_kind in _present_kinds(root, namespace):
                diagnostics.append(
                    RevisionCoverageDiagnostic(
                        entry.name,
                        "unsupported_namespace",
                        f"{raw_kind} uses non-transitional WordprocessingML",
                        None,
                    )
                )
        for element in root.iter():
            tag = element.tag
            if not isinstance(tag, str):
                continue
            element_namespace, local = _split_tag(tag)
            if element_namespace != _W or local not in _REVISION_NAMES:
                continue
            locator = tree.getpath(element)
            unsafe_detail = _unsafe_revision_detail(element, local)
            if unsafe_detail:
                diagnostics.append(
                    RevisionCoverageDiagnostic(
                        entry.name, "unsafe_revision_shape", unsafe_detail, locator
                    )
                )
            if local in _UNSUPPORTED_RANGES:
                diagnostics.append(
                    RevisionCoverageDiagnostic(
                        entry.name, "unsupported_structural_revision", local, locator
                    )
                )
            revisions.append(
                Revision(
                    kind=_kind(local),
                    raw_kind=local,
                    part_name=entry.name,
                    revision_id=element.get(_tag("id")),
                    author=element.get(_tag("author")),
                    date=element.get(_tag("date")),
                    locator=locator,
                    paragraph_mark=_is_paragraph_mark(element),
                )
            )
    return RevisionInventory(
        revisions=tuple(revisions),
        coverage=(
            RevisionInventoryCoverage.COMPLETE
            if not diagnostics
            else RevisionInventoryCoverage.INCOMPLETE
        ),
        diagnostics=tuple(diagnostics),
    )

def _unsafe_revision_detail(element: etree._Element, kind: str) -> str | None:
    if kind in _FORMATTING_PROPERTY_CHANGES:
        expected = _PROPERTY_CHANGE_SNAPSHOT[kind]
        children = list(element)
        if len(children) != 1 or _split_tag(children[0].tag)[1] != expected:
            return f"malformed {kind} snapshot"
        for descendant in element.iterdescendants():
            local = _split_tag(descendant.tag)[1]
            if local in _BLOCK_REVISION_CHILDREN:
                return f"{kind} contains block content"
            if local in {"ins", "del"} and (list(descendant) or _direct_revision_text(descendant)):
                return f"{kind} contains text-bearing revision"
        if _direct_revision_text(element):
            return f"{kind} owns text"
        return None
    if kind in _EMPTY_CUSTOM_XML_RANGE_MARKERS:
        if list(element) or _direct_revision_text(element):
            return f"{kind} is not an empty range marker"
        return None
    if kind in {"ins", "del"}:
        for descendant in element.iterdescendants():
            if _split_tag(descendant.tag)[1] in _BLOCK_REVISION_CHILDREN:
                return f"{kind} contains block content"
        return None
    return f"unsupported revision kind {kind}"

def _direct_revision_text(element: etree._Element) -> str:
    parts: list[str] = []
    for node in element.iter():
        local = _split_tag(node.tag)[1]
        if local not in {"t", "delText", "tab", "br", "cr"}:
            continue
        parent = node.getparent()
        nested = False
        while parent is not None and parent is not element:
            if _split_tag(parent.tag)[1] in {"ins", "del"}:
                nested = True
                break
            parent = parent.getparent()
        if nested:
            continue
        parts.append(
            node.text or "" if local in {"t", "delText"} else "\t" if local == "tab" else "\n"
        )
    return "".join(parts)

def _split_tag(tag: str) -> tuple[str | None, str]:
    if tag.startswith("{"):
        namespace, local = tag[1:].split("}", 1)
        return namespace, local
    return None, tag

def _present_kinds(root: etree._Element, namespace: str) -> set[str]:
    return {
        name
        for name in _REVISION_NAMES
        if next(root.iter(f"{{{namespace}}}{name}"), None) is not None
    }

def _kind(raw: str) -> RevisionKind:
    try:
        return RevisionKind(raw)
    except ValueError:
        if "Range" in raw:
            return RevisionKind.RANGE_MARKER
        return RevisionKind.CONFLICT
