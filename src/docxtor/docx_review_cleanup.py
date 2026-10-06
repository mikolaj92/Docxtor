"""Package-preserving removal of native Word comments, without resolving revisions."""

from __future__ import annotations

from dataclasses import dataclass, replace
from hashlib import sha256
from io import BytesIO
from posixpath import basename, dirname, join, normpath
from urllib.parse import urlsplit
from xml.parsers import expat
from zipfile import ZipFile

from lxml import etree

from .docx_comment_mutations import CommentMutationError, CommentMutationResult
from .docx_inventory import InventoryCoverage, inventory_docx
from .docx_ns import _THREAD_REL_BY_TARGET, CT_NS, REL_NS, W14_NS, W15_NS
from .docx_package import (
    PackageEntry,
    PackageError,
    _needs_xml_validation,
    parse_package_xml,
    read_package_entries,
)
from .docx_review_models import OperationReceipt, OperationStatus
from .docx_revision_inventory import _REVISION_NAMES

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
_CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
_COMMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_COMMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
_PART_ROOTS = {
    _COMMENT_TYPE: f"{{{_W}}}comments",
    "application/vnd.ms-word.commentsExtended+xml": f"{{{W15_NS}}}commentsEx",
    "application/vnd.ms-word.commentsIds+xml": f"{{{_CID}}}commentsIds",
    "application/vnd.ms-word.commentsExtensible+xml": f"{{{_CEX}}}commentsExtensible",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml": (
        f"{{{W15_NS}}}people"
    ),
}
_REL_ROOTS = {_COMMENT_REL: _PART_ROOTS[_COMMENT_TYPE]} | {
    relationship: _PART_ROOTS[content_type]
    for relationship, content_type in _THREAD_REL_BY_TARGET.values()
}
_METADATA_CHILDREN = {
    f"{{{W15_NS}}}commentsEx": f"{{{W15_NS}}}commentEx",
    f"{{{_CID}}}commentsIds": f"{{{_CID}}}commentId",
    f"{{{_CEX}}}commentsExtensible": f"{{{_CEX}}}commentExtensible",
    f"{{{W15_NS}}}people": f"{{{W15_NS}}}person",
}
_MARKERS = {
    f"{{{_W}}}{name}"
    for name in (
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
    )
}
_STORIES = {
    f"{{{_W}}}{name}"
    for name in (
        "body",
        "hdr",
        "ftr",
        "footnote",
        "endnote",
        "txbxContent",
        "docPartBody",
    )
}
_RANGE_PARENTS = {
    f"{{{_W}}}{name}"
    for name in (
        "bdo",
        "body",
        "customXml",
        "del",
        "dir",
        "docPartBody",
        "endnote",
        "fldSimple",
        "footnote",
        "ftr",
        "hdr",
        "hyperlink",
        "ins",
        "moveFrom",
        "moveTo",
        "p",
        "rt",
        "rubyBase",
        "sdtContent",
        "smartTag",
        "tbl",
        "tc",
        "tr",
    )
} | {
    f"{{http://schemas.openxmlformats.org/officeDocument/2006/math}}{name}"
    for name in (
        "deg",
        "den",
        "e",
        "fName",
        "lim",
        "num",
        "oMath",
        "sub",
        "sup",
    )
}


@dataclass
class _Span:
    tag: str
    start: int
    open_end: int
    end: int = 0


@dataclass
class _State:
    entries: tuple[PackageEntry, ...]
    roots: dict[str, etree._Element]
    removed_parts: set[str]
    removals: dict[str, list[etree._Element]]
    comment_ids: set[str]
    marker_count: int
    metadata_count: int


def remove_all_comments_bytes(data: bytes) -> CommentMutationResult:
    """Remove all native Word comments without accepting or rejecting any revision.

    The complete package is checked before and after the in-memory operation.
    Only exact comment elements and their OPC records are removed; retained XML
    bytes, including revision wrappers, text, properties and orphan parts, are
    copied verbatim. A marker inside a revision is the sole permitted change to
    that revision. Revisions inside a part that would be deleted are refused.
    No-comment input returns the original archive bytes. Publication remains the
    caller's responsibility through ``write_package_atomically``.
    """
    try:
        state = _read_state(data)
        if not (state.comment_ids or state.marker_count or state.metadata_count):
            return _result(data, data, ())
        replacements: dict[str, bytes] = {}
        for entry in state.entries:
            root = state.roots.get(entry.name)
            if entry.name in state.removed_parts:
                if root is not None and any(_is_revision(node) for node in root.iter()):
                    raise CommentMutationError(
                        f"comment cleanup would delete revision payloads in {entry.name}"
                    )
                continue
            nodes = state.removals.get(entry.name, [])
            if nodes:
                assert root is not None
                spans = _xml_spans(entry.data, root, entry.name)
                ranges = sorted((spans[node].start, spans[node].end) for node in nodes)
                replacements[entry.name] = _remove_ranges(entry.data, ranges)
        final_entries = tuple(
            replace(entry, data=replacements[entry.name]) if entry.name in replacements else entry
            for entry in state.entries
            if entry.name not in state.removed_parts
        )
        output = BytesIO()
        with ZipFile(output, "w") as archive:
            for entry in final_entries:
                archive.writestr(entry.zip_info(), entry.data)
        payload = output.getvalue()
        after = _read_state(payload)
        if after.comment_ids or after.marker_count or after.metadata_count:
            raise CommentMutationError("comment cleanup was not confirmed by complete postflight")
        actual = {entry.name: entry.data for entry in after.entries}
        expected = {entry.name: entry.data for entry in final_entries}
        if actual != expected:
            raise CommentMutationError("comment cleanup changed retained package entry bytes")
        if _revision_payloads(state) != _revision_payloads(after):
            raise CommentMutationError("comment cleanup changed existing revision payloads")
        affected = tuple(sorted(state.removed_parts | set(replacements)))
        return _result(data, payload, affected)
    except PackageError as exc:
        raise CommentMutationError(str(exc)) from exc


def _result(before: bytes, after: bytes, affected: tuple[str, ...]) -> CommentMutationResult:
    return CommentMutationResult(
        data=after,
        receipt=OperationReceipt(
            operation="remove_all_comments",
            status=OperationStatus.APPLIED if affected else OperationStatus.NOOP,
            affected_parts=affected,
            before_sha256=sha256(before).hexdigest(),
            after_sha256=sha256(after).hexdigest(),
        ),
        comments=(),
    )


def _read_state(data: bytes) -> _State:
    entries = read_package_entries(data)
    inventory = inventory_docx(data)
    if inventory.coverage is not InventoryCoverage.COMPLETE:
        raise CommentMutationError("comment cleanup requires complete package inventory")
    parts = {part.name: part for part in inventory.parts}
    roots = {
        entry.name: parse_package_xml(entry.data, part_name=entry.name)
        for entry in entries
        if parts[entry.name].is_xml or _needs_xml_validation(entry.name, entry.data)
    }
    removed: set[str] = set()
    for name, root in roots.items():
        content_type = parts[name].content_type
        expected = _PART_ROOTS.get(content_type)
        if expected is not None and root.tag != expected:
            raise CommentMutationError(f"unsupported native comment part root in {name}")
        if expected is not None or root.tag in _PART_ROOTS.values():
            removed.add(name)
        elif _word_comment_type(content_type):
            raise CommentMutationError(f"unsupported native comment content type in {name}")
    state = _State(entries, roots, removed, {}, set(), 0, 0)
    _read_comments(state)
    _read_markers(state)
    _read_relationships(state)
    _read_content_types(state)
    return state


def _read_comments(state: _State) -> None:
    para_ids: set[str] = set()
    durable_ids: set[str] = set()
    thread_parents: dict[str, str | None] = {}
    mapped_paragraphs: set[str] = set()
    for name in sorted(state.removed_parts):
        root = state.roots[name]
        expected_child = (
            f"{{{_W}}}comment" if root.tag == f"{{{_W}}}comments" else _METADATA_CHILDREN[root.tag]
        )
        if root.text and root.text.strip():
            raise CommentMutationError(f"unsupported comment part text in {name}")
        for node in root:
            if not isinstance(node.tag, str):
                continue
            if node.tag != expected_child or (node.tail and node.tail.strip()):
                raise CommentMutationError(f"unsupported native comment definition in {name}")
            if node.tag == f"{{{_W}}}comment":
                comment_id = _comment_id(node)
                if comment_id in state.comment_ids:
                    raise CommentMutationError(f"duplicate native comment ID: {comment_id}")
                state.comment_ids.add(comment_id)
                for paragraph in node.iter(f"{{{_W}}}p"):
                    para_id = paragraph.get(f"{{{W14_NS}}}paraId")
                    if para_id is not None:
                        if para_id in para_ids:
                            raise CommentMutationError("duplicate comment paragraph identity")
                        para_ids.add(para_id)
            else:
                state.metadata_count += 1
                if node.tag != f"{{{W15_NS}}}person":
                    _empty_record(node, name)
                elif any(
                    child.tag != f"{{{W15_NS}}}presenceInfo"
                    or len(child)
                    or (child.text and child.text.strip())
                    for child in node
                    if isinstance(child.tag, str)
                ):
                    raise CommentMutationError(f"unsupported comment person metadata in {name}")
                if node.tag == f"{{{_CID}}}commentId":
                    durable_id = node.get(f"{{{_CID}}}durableId")
                    if not durable_id or durable_id in durable_ids:
                        raise CommentMutationError("invalid comment durable identity")
                    durable_ids.add(durable_id)
    for name in sorted(state.removed_parts):
        root = state.roots[name]
        for node in root:
            if node.tag == f"{{{W15_NS}}}commentEx":
                values = [node.get(f"{{{W15_NS}}}paraId")]
                parent = node.get(f"{{{W15_NS}}}paraIdParent")
                if parent is not None:
                    values.append(parent)
                if any(value not in para_ids for value in values):
                    raise CommentMutationError("dangling comment thread paragraph identity")
                para_id = values[0]
                assert para_id is not None
                if para_id in thread_parents:
                    raise CommentMutationError("duplicate comment thread paragraph identity")
                thread_parents[para_id] = parent
            elif node.tag == f"{{{_CID}}}commentId":
                para_id = node.get(f"{{{_CID}}}paraId")
                if para_id not in para_ids or para_id in mapped_paragraphs:
                    raise CommentMutationError("dangling comment paragraph identity")
                assert para_id is not None
                mapped_paragraphs.add(para_id)
            elif node.tag == f"{{{_CEX}}}commentExtensible":
                if node.get(f"{{{_CEX}}}durableId") not in durable_ids:
                    raise CommentMutationError("dangling extensible comment identity")
    if state.metadata_count and not state.comment_ids:
        raise CommentMutationError("native comment metadata has no comment definitions")
    for para_id in thread_parents:
        visited: set[str] = set()
        current: str | None = para_id
        while current is not None:
            if current in visited:
                raise CommentMutationError("cyclic native comment thread metadata")
            visited.add(current)
            current = thread_parents.get(current)


def _read_markers(state: _State) -> None:
    anchors: dict[str, list[tuple[str, tuple[str, str], int]]] = {}
    for name, root in state.roots.items():
        for index, node in enumerate(root.iter()):
            if not isinstance(node.tag, str):
                continue
            qname = etree.QName(node)
            if _word_namespace(qname.namespace) and "comment" in qname.localname.lower():
                allowed_definition = (
                    name in state.removed_parts
                    and node.tag in {f"{{{_W}}}comments", f"{{{_W}}}comment"}
                    and (node is root or node.getparent() is root)
                )
                allowed_metadata = (
                    name in state.removed_parts
                    and root.tag in _METADATA_CHILDREN
                    and (
                        node is root
                        or (node.getparent() is root and node.tag == _METADATA_CHILDREN[root.tag])
                    )
                )
                if node.tag not in _MARKERS and not allowed_definition and not allowed_metadata:
                    raise CommentMutationError(f"unsupported Word comment element in {name}")
            for attribute in node.attrib:
                attr = etree.QName(attribute)
                if _word_namespace(attr.namespace) and "comment" in attr.localname.lower():
                    raise CommentMutationError(f"unsupported Word comment attribute in {name}")
            if node.tag not in _MARKERS:
                continue
            if name in state.removed_parts or len(node) or (node.text and node.text.strip()):
                raise CommentMutationError(f"unsupported native comment marker in {name}")
            allowed = {f"{{{_W}}}id"}
            if node.tag != f"{{{_W}}}commentReference":
                allowed.add(f"{{{_W}}}displacedByCustomXml")
            if set(node.attrib) - allowed:
                raise CommentMutationError(
                    f"unsupported native comment marker attributes in {name}"
                )
            value = node.get(f"{{{_W}}}displacedByCustomXml")
            if value is not None and value not in {"prev", "next"}:
                raise CommentMutationError(f"malformed native comment marker flag in {name}")
            comment_id = _comment_id(node)
            parent = node.getparent()
            allowed_parents = (
                {f"{{{_W}}}r"} if node.tag == f"{{{_W}}}commentReference" else _RANGE_PARENTS
            )
            if parent is None or parent.tag not in allowed_parents:
                raise CommentMutationError(f"unsupported native comment marker parent in {name}")
            ancestors = list(node.iterancestors())
            story = next((item for item in ancestors if item.tag in _STORIES), root)
            identity = (name, root.getroottree().getpath(story))
            anchors.setdefault(comment_id, []).append((node.tag, identity, index))
            state.removals.setdefault(name, []).append(node)
            state.marker_count += 1
    for comment_id, rows in anchors.items():
        if comment_id not in state.comment_ids or len({row[1] for row in rows}) != 1:
            raise CommentMutationError(f"dangling or cross-story comment markers: {comment_id}")
        positions = {tag: [row[2] for row in rows if row[0] == tag] for tag in _MARKERS}
        starts = positions[f"{{{_W}}}commentRangeStart"]
        ends = positions[f"{{{_W}}}commentRangeEnd"]
        refs = positions[f"{{{_W}}}commentReference"]
        # Reference-only point comments are supported. A one-sided range has
        # conflicting conformance rules in the native Word specification.
        if len(refs) != 1 or not (
            (not starts and not ends)
            or (len(starts) == len(ends) == 1 and starts[0] < ends[0] < refs[0])
        ):
            raise CommentMutationError(f"malformed native comment markers: {comment_id}")


def _read_relationships(state: _State) -> None:
    names = {entry.name for entry in state.entries}
    comment_parts = set(state.removed_parts)
    owned_rels = {join(dirname(name), "_rels", basename(name) + ".rels") for name in comment_parts}
    state.removed_parts.update(owned_rels & names)
    for name, root in state.roots.items():
        if not name.endswith(".rels"):
            if root.tag == f"{{{REL_NS}}}Relationships":
                raise CommentMutationError(f"unsupported relationships part locator: {name}")
            continue
        if root.tag != f"{{{REL_NS}}}Relationships":
            raise CommentMutationError(f"invalid relationships part: {name}")
        seen: set[str] = set()
        for node in root:
            if not isinstance(node.tag, str):
                continue
            rid, kind, target = node.get("Id"), node.get("Type", ""), node.get("Target")
            mode = node.get("TargetMode", "Internal")
            if (
                node.tag != f"{{{REL_NS}}}Relationship"
                or not rid
                or rid in seen
                or not kind
                or target is None
                or mode not in {"Internal", "External"}
            ):
                raise CommentMutationError(f"invalid relationship record in {name}")
            seen.add(rid)
            expected_root = _REL_ROOTS.get(kind)
            if expected_root is None and _word_comment_relationship(kind):
                raise CommentMutationError(f"unsupported native comment relationship in {name}")
            if expected_root is not None and mode == "External":
                raise CommentMutationError(f"external native comment relationship in {name}")
            resolved = _resolve_target(name, target) if mode == "Internal" else None
            if expected_root is not None:
                if resolved not in comment_parts or state.roots[resolved].tag != expected_root:
                    raise CommentMutationError(f"dangling native comment relationship in {name}")
                _empty_record(node, name)
                if set(node.attrib) - {"Id", "Type", "Target", "TargetMode"}:
                    raise CommentMutationError(f"unsupported native comment relationship in {name}")
                state.removals.setdefault(name, []).append(node)
            elif resolved in state.removed_parts and name not in state.removed_parts:
                raise CommentMutationError(f"unrelated relationship targets comment part in {name}")


def _read_content_types(state: _State) -> None:
    root = state.roots["[Content_Types].xml"]
    if root.tag != f"{{{CT_NS}}}Types":
        raise CommentMutationError("invalid content-types root")
    if any(
        node.tag == f"{{{CT_NS}}}Types"
        for name, node in state.roots.items()
        if name != "[Content_Types].xml"
    ):
        raise CommentMutationError("unsupported extra content-types part")
    names = {entry.name for entry in state.entries}
    for node in root:
        if not isinstance(node.tag, str):
            continue
        if node.tag not in {f"{{{CT_NS}}}Override", f"{{{CT_NS}}}Default"}:
            raise CommentMutationError("unsupported content-type record shape")
        content_type = node.get("ContentType", "")
        if _word_comment_type(content_type) and content_type not in _PART_ROOTS:
            raise CommentMutationError("unsupported native comment content-type record")
        if node.tag == f"{{{CT_NS}}}Override":
            part = node.get("PartName", "").removeprefix("/")
            if content_type in _PART_ROOTS and part not in state.removed_parts:
                raise CommentMutationError("dangling native comment content-type record")
            if part in state.removed_parts:
                _empty_record(node, "[Content_Types].xml")
                if set(node.attrib) != {"PartName", "ContentType"}:
                    raise CommentMutationError("unsupported native comment override attributes")
                state.removals.setdefault("[Content_Types].xml", []).append(node)
        elif node.tag == f"{{{CT_NS}}}Default" and content_type in _PART_ROOTS:
            extension = node.get("Extension", "").lower()
            matching = {name for name in names if name.rsplit(".", 1)[-1].lower() == extension}
            if not matching or not matching <= state.removed_parts:
                raise CommentMutationError("native comment default content type is ambiguous")
            _empty_record(node, "[Content_Types].xml")
            if set(node.attrib) != {"Extension", "ContentType"}:
                raise CommentMutationError("unsupported native comment default attributes")
            state.removals.setdefault("[Content_Types].xml", []).append(node)


def _empty_record(node: etree._Element, name: str) -> None:
    if len(node) or (node.text and node.text.strip()):
        raise CommentMutationError(f"unsupported nonempty native comment OPC record in {name}")


def _comment_id(node: etree._Element) -> str:
    value = node.get(f"{{{_W}}}id", "")
    digits = value.removeprefix("-")
    if not digits or not digits.isascii() or not digits.isdecimal():
        raise CommentMutationError("native comment identity is missing or malformed")
    try:
        return str(int(value))
    except ValueError as exc:
        raise CommentMutationError("native comment identity exceeds decimal bounds") from exc


def _word_namespace(namespace: str | None) -> bool:
    return namespace in {_W, _STRICT_W} or bool(
        namespace and namespace.startswith("http://schemas.microsoft.com/office/word/")
    )


def _word_comment_type(value: str) -> bool:
    return "comment" in value.lower() and value.startswith(
        (
            "application/vnd.ms-word.",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.",
        )
    )


def _word_comment_relationship(value: str) -> bool:
    return "comment" in value.lower() and value.startswith(
        (
            "http://schemas.microsoft.com/office/",
            "http://schemas.openxmlformats.org/officeDocument/",
            "http://purl.oclc.org/ooxml/officeDocument/",
        )
    )


def _is_revision(node: etree._Element) -> bool:
    if not isinstance(node.tag, str):
        return False
    qname = etree.QName(node)
    return _word_namespace(qname.namespace) and qname.localname in (
        set(_REVISION_NAMES) | {"delText", "delInstrText"}
    )


def _revision_payloads(state: _State) -> dict[str, tuple[bytes, ...]]:
    result: dict[str, tuple[bytes, ...]] = {}
    for entry in state.entries:
        root = state.roots.get(entry.name)
        if root is None:
            continue
        revisions = [node for node in root.iter() if _is_revision(node)]
        if not revisions:
            continue
        spans = _xml_spans(entry.data, root, entry.name)
        markers = [spans[node] for node in root.iter() if node.tag in _MARKERS]
        payloads = []
        for node in revisions:
            span = spans[node]
            removals = sorted(
                (marker.start - span.start, marker.end - span.start)
                for marker in markers
                if span.start <= marker.start < marker.end <= span.end
            )
            payloads.append(_remove_ranges(entry.data[span.start : span.end], removals))
        result[entry.name] = tuple(payloads)
    return result


def _resolve_target(relationship_part: str, target: str) -> str:
    try:
        parsed = urlsplit(target)
    except ValueError as exc:
        raise CommentMutationError("malformed internal relationship target") from exc
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment or "\\" in target:
        raise CommentMutationError("ambiguous internal relationship target")
    if relationship_part == "_rels/.rels":
        source = ""
    elif dirname(relationship_part).endswith("_rels"):
        source = join(dirname(dirname(relationship_part)), basename(relationship_part)[:-5])
    else:
        raise CommentMutationError("invalid relationship part locator")
    resolved = normpath(target[1:] if target.startswith("/") else join(dirname(source), target))
    if not target or resolved == ".." or resolved.startswith("../"):
        raise CommentMutationError("internal relationship target escapes the package")
    return resolved


def _xml_spans(
    data: bytes,
    root: etree._Element,
    name: str,
) -> dict[etree._Element, _Span]:
    """Locate exact element bytes; XML is validated separately before this scan."""
    encoding = root.getroottree().docinfo.encoding.upper().replace("_", "-")
    if encoding.startswith("UTF-16"):
        little = (
            encoding.endswith("LE") or data.startswith(b"\xff\xfe") or data.startswith(b"<\x00")
        )
        codec, width = ("utf-16-le" if little else "utf-16-be"), 2
    elif encoding.startswith("UTF-32"):
        raise CommentMutationError(f"unsupported comment mutation XML encoding in {name}")
    else:
        codec, width = "ascii", 1
    quotes = {token.encode(codec) for token in ("'", '"')}
    close, slash = ">".encode(codec), "/".encode(codec)

    def tag_end(start: int) -> int:
        quote = None
        for offset in range(start, len(data), width):
            token = data[offset : offset + width]
            if quote is not None:
                if token == quote:
                    quote = None
            elif token in quotes:
                quote = token
            elif token == close:
                return offset + width
        raise CommentMutationError(f"unreadable comment mutation XML markup in {name}")

    parser = expat.ParserCreate(namespace_separator="\x1f")
    spans: list[_Span] = []
    stack: list[_Span] = []

    def start(tag: str, _attributes: dict[str, str]) -> None:
        expanded = "{" + tag.replace("\x1f", "}", 1) if "\x1f" in tag else tag
        span = _Span(expanded, parser.CurrentByteIndex, tag_end(parser.CurrentByteIndex))
        spans.append(span)
        stack.append(span)

    def end(_tag: str) -> None:
        span = stack.pop()
        empty = data[span.open_end - 2 * width : span.open_end - width] == slash
        span.end = span.open_end if empty else tag_end(parser.CurrentByteIndex)

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except (expat.ExpatError, ValueError, LookupError, UnicodeError) as exc:
        raise CommentMutationError(f"unreadable comment mutation XML in {name}") from exc
    elements = [node for node in root.iter() if isinstance(node.tag, str)]
    if len(elements) != len(spans) or any(
        node.tag != span.tag for node, span in zip(elements, spans, strict=True)
    ):
        raise CommentMutationError(f"ambiguous comment mutation XML element spans in {name}")
    return dict(zip(elements, spans, strict=True))


def _remove_ranges(data: bytes, ranges: list[tuple[int, int]]) -> bytes:
    chunks: list[bytes] = []
    offset = 0
    for start, end in ranges:
        if not offset <= start < end <= len(data):
            raise CommentMutationError("overlapping or invalid comment removal byte ranges")
        chunks.append(data[offset:start])
        offset = end
    chunks.append(data[offset:])
    return b"".join(chunks)
