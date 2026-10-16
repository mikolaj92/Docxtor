from __future__ import annotations

import re
from dataclasses import dataclass
from posixpath import basename, dirname, join, normpath
from typing import Any
from uuid import uuid4

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .docx_models import CommentScopeBookmark
from .docx_ns import REL_NS, W_P
from .docx_package import parse_package_xml

_SCOPE_NAME = re.compile(r"_dxs_([0-9a-f]{16})_([1-9][0-9]*)_([1-9][0-9]*)\Z")
_LABELS = {
    "footnote": "Go to footnote text",
    "endnote": "Go to endnote text",
    "header": "Go to header text",
    "footer": "Go to footer text",
    "textbox": "Go to text box",
    "document": "Go to document text",
}


@dataclass(frozen=True)
class ScopeLink:
    name: str
    label: str
    paragraph: Any


def scope_link(comment: Any) -> ScopeLink | None:
    matches = [
        node
        for node in comment.iter(qn("w:hyperlink"))
        if (node.get(qn("w:anchor")) or "").startswith("_dxs_")
    ]
    if not matches:
        return None
    if len(matches) != 1:
        raise ValueError("native comment has duplicate logical scope links")
    link = matches[0]
    name = link.get(qn("w:anchor"), "")
    match = _SCOPE_NAME.fullmatch(name)
    paragraph = link.getparent()
    label = "".join(node.text or "" for node in link.iter(qn("w:t")))
    if (
        match is None
        or len(name) > 40
        or int(match[2]) > int(match[3])
        or paragraph is None
        or paragraph.tag != W_P
        or paragraph.getparent() is not comment
        or list(comment)[-1] is not paragraph
        or any(node is not link and node.tag != qn("w:pPr") for node in paragraph)
        or link.get(qn("r:id")) is not None
        or label not in _LABELS.values()
    ):
        raise ValueError("native comment logical scope link is malformed")
    return ScopeLink(name, label, paragraph)


def append_scope_link(comment: Any, name: str, kind: str) -> None:
    paragraph = OxmlElement("w:p")
    comment.append(paragraph)
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), name)
    link.set(qn("w:history"), "1")
    paragraph.append(link)
    run = OxmlElement("w:r")
    link.append(run)
    properties = OxmlElement("w:rPr")
    run.append(properties)
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), "Hyperlink")
    properties.append(style)
    text = OxmlElement("w:t")
    text.text = _LABELS[kind]
    run.append(text)


def new_scope_name(roots: dict[str, Any]) -> str:
    used = {
        node.get(qn("w:name"))
        for root in roots.values()
        for node in root.iter(qn("w:bookmarkStart"))
    }
    while True:
        name = f"_dxs_{uuid4().hex[:16]}_1_1"
        if name not in used:
            return name


def new_bookmark_id(roots: dict[str, Any]) -> str:
    used = {
        node.get(qn("w:id"))
        for root in roots.values()
        for tag in ("bookmarkStart", "bookmarkEnd")
        for node in root.iter(qn(f"w:{tag}"))
    }
    candidate = 0
    while str(candidate) in used:
        candidate += 1
    return str(candidate)


def display_target(state: Any, spans: tuple[Any, ...]) -> tuple[tuple[Any, ...], str | None]:
    first = state.paragraphs[spans[0].locator]
    if first.geometry.part_name == state.main_part_name and first.geometry.story_kind in {
        "body",
        "table",
    }:
        return spans, None
    body = [
        record
        for record in state.paragraphs.values()
        if record.geometry.part_name == state.main_part_name
        and record.geometry.story_kind in {"body", "table"}
    ]
    if not body:
        raise ValueError("logical comment scope has no existing body display paragraph")
    from .docx_review_geometry import PhysicalCommentSpan

    def point(record: Any, node: Any | None = None) -> tuple[Any, ...]:
        offset = source_offset(record.element, node) if node is not None else 0
        return (PhysicalCommentSpan(record.geometry.locator, offset, offset, ""),)

    story_root = state.roots[first.geometry.part_name]
    if story_root.tag in {qn("w:footnotes"), qn("w:endnotes")}:
        kind = "footnote" if story_root.tag == qn("w:footnotes") else "endnote"
        note = next(
            (node for node in first.element.iterancestors() if node.tag == qn(f"w:{kind}")), None
        )
        if note is None:
            raise ValueError("logical note scope has no native note identity")
        for record in body:
            for reference in record.element.iter(qn(f"w:{kind}Reference")):
                if (
                    reference.get(qn("w:id")) == note.get(qn("w:id"))
                    and next((node for node in reference.iterancestors() if node.tag == W_P), None)
                    is record.element
                ):
                    return point(record, reference), kind
        raise ValueError("logical note scope has no body note reference")
    if first.geometry.part_name == state.main_part_name:
        for parent in first.element.iterancestors():
            record = next((item for item in body if item.element is parent), None)
            if record is not None:
                return point(record, first.element), "textbox"
        raise ValueError("logical text box scope has no body host paragraph")
    kind = (
        "header"
        if story_root.tag == qn("w:hdr")
        else "footer"
        if story_root.tag == qn("w:ftr")
        else None
    )
    if kind is None:
        raise ValueError("logical comment story has no supported body display relationship")
    rels_name = join(
        dirname(state.main_part_name), "_rels", basename(state.main_part_name) + ".rels"
    )
    rels_entry = next((entry for entry in state.entries if entry.name == rels_name), None)
    if rels_entry is None:
        raise ValueError("logical header or footer has no document relationships")
    rels = parse_package_xml(rels_entry.data, part_name=rels_name)
    ids = {
        node.get("Id")
        for node in rels
        if node.tag == f"{{{REL_NS}}}Relationship"
        and node.get("TargetMode") != "External"
        and node.get("Target")
        and (
            node.get("Target").lstrip("/")
            if node.get("Target").startswith("/")
            else normpath(join(dirname(state.main_part_name), node.get("Target")))
        )
        == first.geometry.part_name
    }
    for reference in state.roots[state.main_part_name].iter(qn(f"w:{kind}Reference")):
        if reference.get(qn("r:id")) not in ids:
            continue
        parent = next((node for node in reference.iterancestors() if node.tag == W_P), None)
        record = next(
            (item for item in body if item.element is parent), body[-1] if parent is None else None
        )
        if record is not None:
            return point(record), kind
    raise ValueError("logical header or footer has no explicit section reference")


def source_offset(paragraph: Any, node: Any) -> int:
    from .docx_review_geometry import _marker_coordinates

    offset = 0
    child = node
    while child is not paragraph:
        parent = child.getparent()
        if parent is None:
            raise ValueError("display reference has no paragraph origin")
        offset += sum(
            len(_marker_coordinates(sibling)[1]) for sibling in list(parent)[: parent.index(child)]
        )
        child = parent
    return offset


def owned_scope_bookmarks(roots: dict[str, Any]) -> tuple[CommentScopeBookmark, ...]:
    result = []
    referenced: set[str] = set()
    for root in roots.values():
        for comment in root.iter(qn("w:comment")):
            link = scope_link(comment)
            if link is None:
                continue
            if link.name in referenced:
                raise ValueError("native logical scope bookmark has multiple comment owners")
            referenced.add(link.name)
            found = [
                (name, node)
                for name, tree in roots.items()
                for node in tree.iter(qn("w:bookmarkStart"))
                if node.get(qn("w:name")) == link.name
            ]
            if len(found) != 1 or comment.get(qn("w:id")) is None:
                raise ValueError("native logical scope bookmark identity is missing or duplicate")
            part, start = found[0]
            identity = start.get(qn("w:id"))
            starts = [
                node
                for tree in roots.values()
                for node in tree.iter(qn("w:bookmarkStart"))
                if node.get(qn("w:id")) == identity
            ]
            ends = [
                node
                for tree in roots.values()
                for node in tree.iter(qn("w:bookmarkEnd"))
                if node.get(qn("w:id")) == identity
            ]
            same_part = len(ends) == 1 and ends[0].getroottree().getroot() is roots[part]
            ordered = same_part and list(roots[part].iter()).index(start) < list(
                roots[part].iter()
            ).index(ends[0])
            if identity is None or len(starts) != 1 or len(ends) != 1 or not ordered:
                raise ValueError("native logical scope bookmark range is missing or duplicate")
            result.append(CommentScopeBookmark(part, identity, link.name, comment.get(qn("w:id"))))
    reserved = [
        node.get(qn("w:name"))
        for root in roots.values()
        for node in root.iter(qn("w:bookmarkStart"))
        if (node.get(qn("w:name")) or "").startswith("_dxs_")
    ]
    if len(reserved) != len(set(reserved)) or set(reserved) != referenced:
        raise ValueError("native logical scope bookmarks have missing or duplicate owners")
    groups: dict[str, list[tuple[int, int]]] = {}
    for scope in result:
        match = _SCOPE_NAME.fullmatch(scope.name)
        assert match is not None
        groups.setdefault(match[1], []).append((int(match[2]), int(match[3])))
    for identity, members in groups.items():
        counts = {count for _, count in members}
        if (
            len(counts) != 1
            or len(members) != members[0][1]
            or {index for index, _ in members} != set(range(1, len(members) + 1))
        ):
            raise ValueError(f"native logical scope group is incomplete: {identity}")
    return tuple(result)


def remove_scope_bookmark_markers(
    data: bytes, scopes: tuple[CommentScopeBookmark, ...]
) -> tuple[bytes, tuple[str, ...]]:
    from dataclasses import replace

    from .docx_package import read_package_entries
    from .docx_review_cleanup import _remove_ranges, _xml_spans
    from .docx_review_geometry_comments import _package_bytes

    entries = read_package_entries(data)
    output = []
    changed = []
    for entry in entries:
        selected = [scope for scope in scopes if scope.part_name == entry.name]
        if not selected:
            output.append(entry)
            continue
        root = parse_package_xml(entry.data, part_name=entry.name)
        nodes = [
            node
            for scope in selected
            for node in root.iter()
            if node.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
            and node.get(qn("w:id")) == scope.bookmark_id
        ]
        if len(nodes) != 2 * len(selected):
            raise ValueError("owned logical bookmark changed before cleanup")
        spans = _xml_spans(entry.data, root, entry.name)
        payload = _remove_ranges(
            entry.data, sorted((spans[node].start, spans[node].end) for node in nodes)
        )
        output.append(replace(entry, data=payload))
        changed.append(entry.name)
    return _package_bytes(tuple(output)), tuple(sorted(changed))


def expand_scope_bookmarks_for_insertion(
    source: bytes, paragraph: Any, wrapper: Any, locator: str, offset: int
) -> None:
    from .docx_review_geometry import project_docx_review_geometry
    from .docx_review_models import ReviewCoverage

    geometry = project_docx_review_geometry(source)
    if not geometry.scope_bookmarks:
        return
    if geometry.coverage is not ReviewCoverage.COMPLETE:
        raise ValueError("native scope insertion has no complete source proof")
    anchors = {anchor.comment_id: anchor for anchor in geometry.comment_anchors}
    for scope in geometry.scope_bookmarks:
        anchor = anchors[scope.comment_id]
        for kind, span, endpoint in (
            ("bookmarkStart", anchor.spans[0], "start"),
            ("bookmarkEnd", anchor.spans[-1], "end"),
        ):
            edge = span.start_offset if endpoint == "start" else span.end_offset
            if span.locator != locator or edge != offset:
                continue
            matches = [
                node
                for node in paragraph.iter(qn(f"w:{kind}"))
                if node.get(qn("w:id")) == scope.bookmark_id
            ]
            if len(matches) != 1:
                raise ValueError("native scope insertion lost its owned endpoint")
            marker = matches[0]
            order = list(paragraph.iter())
            if endpoint == "start" and order.index(marker) > order.index(wrapper):
                wrapper.addprevious(marker)
            elif endpoint == "end" and order.index(marker) < order.index(wrapper):
                wrapper.addnext(marker)
