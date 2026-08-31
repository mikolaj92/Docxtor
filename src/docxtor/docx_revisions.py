from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from enum import StrEnum
from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

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
_RANGE_PAIRS = (
    ("customXmlDelRangeStart", "customXmlDelRangeEnd"),
    ("customXmlInsRangeStart", "customXmlInsRangeEnd"),
)
_REJECT_UNSUPPORTED = (
    "cellDel",
    "cellIns",
    "cellMerge",
    "sectPrChange",
    "tblPrChange",
    "trPrChange",
    "tcPrChange",
    "tblPrExChange",
    "tblGridChange",
    "numberingChange",
)
_NON_TEXT = ("tab", "br", "cr", "drawing", "object", "pict", "fldChar", "sym")


class RevisionKind(StrEnum):
    INSERTION = "ins"
    DELETION = "del"
    MOVE_FROM = "moveFrom"
    MOVE_TO = "moveTo"
    RUN_PROPERTIES = "rPrChange"
    PARAGRAPH_PROPERTIES = "pPrChange"
    SECTION_PROPERTIES = "sectPrChange"
    TABLE_PROPERTIES = "tblPrChange"
    ROW_PROPERTIES = "trPrChange"
    CELL_PROPERTIES = "tcPrChange"
    CELL_INSERTION = "cellIns"
    CELL_DELETION = "cellDel"
    CELL_MERGE = "cellMerge"
    TABLE_GRID = "tblGridChange"
    TABLE_EXCEPTION_PROPERTIES = "tblPrExChange"
    NUMBERING = "numberingChange"
    RANGE_MARKER = "range_marker"
    CONFLICT = "conflict"


class RevisionInventoryCoverage(StrEnum):
    COMPLETE = "complete"
    INCOMPLETE = "incomplete"


class RevisionOperation(StrEnum):
    ACCEPT_ALL = "accept_all"
    REJECT_ALL = "reject_all"


@dataclass(frozen=True)
class RevisionCoverageDiagnostic:
    part_name: str
    code: str
    detail: str
    locator: str | None = None


@dataclass(frozen=True)
class Revision:
    kind: RevisionKind
    raw_kind: str
    part_name: str
    revision_id: str | None
    author: str | None
    date: str | None
    locator: str
    paragraph_mark: bool


@dataclass(frozen=True)
class RevisionInventory:
    revisions: tuple[Revision, ...]
    coverage: RevisionInventoryCoverage
    diagnostics: tuple[RevisionCoverageDiagnostic, ...]

    @property
    def count(self) -> int:
        return len(self.revisions)


@dataclass(frozen=True)
class RevisionOperationReceipt:
    operation: RevisionOperation
    output_bytes: bytes
    before: RevisionInventory
    after: RevisionInventory
    comments_dropped: bool


class RevisionOperationError(PackageError):
    """The package cannot be flattened without guessing."""


# Compatibility-specific names make callers' exception handling intention clear.
class AcceptRevisionsError(RevisionOperationError):
    pass


class RejectRevisionsError(RevisionOperationError):
    pass


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


def accept_all_revisions_bytes(
    data: bytes, *, drop_comments: bool = True
) -> RevisionOperationReceipt:
    return _operate(data, RevisionOperation.ACCEPT_ALL, drop_comments=drop_comments)


def reject_all_revisions_bytes(
    data: bytes, *, drop_comments: bool = True
) -> RevisionOperationReceipt:
    return _operate(data, RevisionOperation.REJECT_ALL, drop_comments=drop_comments)


def _operate(
    data: bytes, operation: RevisionOperation, *, drop_comments: bool
) -> RevisionOperationReceipt:
    error_type = (
        AcceptRevisionsError if operation is RevisionOperation.ACCEPT_ALL else RejectRevisionsError
    )
    try:
        entries = read_package_entries(data)
        before = inventory_revisions_bytes(data)
        parsed = _preflight(entries, operation, drop_comments=drop_comments, error_type=error_type)
        transformed: list[PackageEntry] = []
        for entry in entries:
            if drop_comments and _is_comment_part(entry.name):
                continue
            payload = entry.data
            root = parsed.get(entry.name)
            if root is not None:
                if _is_word_xml(entry.name):
                    semantic_before = etree.tostring(root, method="c14n")
                    if operation is RevisionOperation.ACCEPT_ALL:
                        _accept_tree(root, entry.name, error_type)
                    else:
                        _reject_tree(root, entry.name, drop_comments, error_type)
                    if drop_comments:
                        _strip_comment_anchors(root)
                    if etree.tostring(root, method="c14n") != semantic_before:
                        payload = _serialize(root)
                elif entry.name.endswith(".rels") and drop_comments:
                    payload = _strip_comment_relationships(root, entry.data)
                elif entry.name == "[Content_Types].xml" and drop_comments:
                    payload = _strip_comment_content_types(root, entry.data)
            transformed.append(_replace_entry(entry, payload))
        output = _write_bytes(transformed)
        after = inventory_revisions_bytes(output)
    except RevisionOperationError:
        raise
    except PackageError as exc:
        raise error_type(str(exc)) from exc
    if after.revisions:
        raise error_type("revision markup survived the operation")
    if drop_comments and _has_comments(output):
        raise error_type("comment markup survived the operation")
    return RevisionOperationReceipt(operation, output, before, after, drop_comments)


def _preflight(
    entries: tuple[PackageEntry, ...],
    operation: RevisionOperation,
    *,
    drop_comments: bool,
    error_type: type[RevisionOperationError],
) -> dict[str, etree._Element]:
    parsed: dict[str, etree._Element] = {}
    for entry in entries:
        relevant = (
            _is_word_xml(entry.name)
            or entry.name.endswith(".rels")
            or entry.name == "[Content_Types].xml"
        )
        if not relevant:
            continue
        root = parse_package_xml(entry.data, part_name=entry.name)
        parsed[entry.name] = root
        if not _is_word_xml(entry.name):
            continue
        if _present_kinds(root, _STRICT_W) or _present_kinds(root, _W14):
            raise error_type(f"{entry.name}: non-transitional review markup is unsupported")
        unsupported = (
            ("cellDel",) if operation is RevisionOperation.ACCEPT_ALL else _REJECT_UNSUPPORTED
        )
        for name in unsupported:
            if next(root.iter(_tag(name)), None) is not None:
                raise error_type(f"{entry.name}: {operation.value} {name} is unsupported")
        problem = _range_problem(root)
        if problem:
            raise error_type(f"{entry.name}: {problem}")
        if operation is RevisionOperation.ACCEPT_ALL:
            for element in root.iter(_tag("del"), _tag("moveFrom")):
                if (
                    _is_paragraph_mark(element)
                    and _next_paragraph_for_mark(element) is None
                    and not _is_content_control_paragraph(element)
                ):
                    raise error_type(
                        f"{entry.name}: paragraph-mark deletion has no following paragraph"
                    )
        else:
            for element in root.iter(_tag("ins"), _tag("moveTo")):
                if _is_paragraph_mark(element):
                    _validate_rejected_inserted_mark(element, entry.name, drop_comments, error_type)
    return parsed


def _accept_tree(root: etree._Element, part: str, error_type: type[RevisionOperationError]) -> None:
    _drop_range_markers(root)
    for element in list(root.iter(_tag("del"), _tag("moveFrom"))):
        if (
            _is_paragraph_mark(element)
            and element.getparent() is not None
            and not _merge_paragraph_into_next(element)
        ):
            if _is_content_control_paragraph(element):
                _remove(element)
            else:
                raise error_type(f"{part}: cannot merge paragraph")
    for element in list(root.iter(_tag("del"), _tag("moveFrom"))):
        _remove(element)
    for element in list(root.iter(_tag("ins"), _tag("moveTo"))):
        if element.getparent() is None:
            continue
        _remove(element) if _is_paragraph_mark(element) else _unwrap(element)
    for name in (
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "tblPrExChange",
        "tblGridChange",
        "numberingChange",
        "cellIns",
        "cellMerge",
    ):
        for element in list(root.iter(_tag(name))):
            _remove(element)


def _reject_tree(
    root: etree._Element, part: str, drop_comments: bool, error_type: type[RevisionOperationError]
) -> None:
    _drop_range_markers(root)
    for element in list(root.iter(_tag("ins"), _tag("moveTo"))):
        if _is_paragraph_mark(element) and element.getparent() is not None:
            _reject_inserted_mark(element, part, drop_comments, error_type)
    leftovers: list[etree._Element] = []
    seen: set[int] = set()
    for element in list(root.iter(_tag("ins"), _tag("moveTo"))):
        if element.getparent() is not None:
            paragraph = next((a for a in element.iterancestors() if a.tag == _tag("p")), None)
            if paragraph is not None and id(paragraph) not in seen:
                leftovers.append(paragraph)
                seen.add(id(paragraph))
            _remove(element)
    for element in list(root.iter(_tag("del"), _tag("moveFrom"))):
        if element.getparent() is None:
            continue
        _remove(element) if _is_paragraph_mark(element) else _unwrap(element)
    for element in root.iter(_tag("delText")):
        element.tag = _tag("t")
    for change_name, property_name in (("rPrChange", "rPr"), ("pPrChange", "pPr")):
        for change in list(root.iter(_tag(change_name))):
            parent, snapshot = change.getparent(), change.find(_tag(property_name))
            if parent is None or snapshot is None:
                raise error_type(f"{part}: malformed {change_name}")
            for child in list(parent):
                parent.remove(child)
            for child in snapshot:
                parent.append(deepcopy(child))
    for paragraph in leftovers:
        _drop_numbering_leftover(paragraph, drop_comments)


def _validate_rejected_inserted_mark(
    mark: etree._Element, part: str, drop_comments: bool, error_type: type[RevisionOperationError]
) -> None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        raise error_type(f"{part}: malformed paragraph-mark insertion")
    if _paragraph_has_original_content(paragraph):
        if _next_paragraph_for_mark(mark) is None and not _is_content_control_paragraph(mark):
            raise error_type(f"{part}: tracked paragraph-mark insertion has no following paragraph")
        return
    allowed = {
        _tag("pPr"),
        _tag("ins"),
        _tag("moveTo"),
        _tag("sdt"),
        _tag("bookmarkStart"),
        _tag("bookmarkEnd"),
        _tag("permStart"),
        _tag("permEnd"),
        _tag("proofErr"),
    }
    for child in paragraph:
        if child.tag in allowed or _is_empty_text_container(child):
            continue
        if drop_comments and (
            child.tag in {_tag("commentRangeStart"), _tag("commentRangeEnd")}
            or _is_comment_reference_run(child)
        ):
            continue
        raise error_type(f"{part}: inserted paragraph contains unsupported structural children")


def _reject_inserted_mark(
    mark: etree._Element, part: str, drop_comments: bool, error_type: type[RevisionOperationError]
) -> None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        raise error_type(f"{part}: malformed paragraph-mark insertion")
    if _paragraph_has_original_content(paragraph):
        if not _merge_paragraph_into_next(mark):
            if _is_content_control_paragraph(mark):
                _remove(mark)
            else:
                raise error_type(f"{part}: cannot merge inserted paragraph")
        return
    for child in list(paragraph):
        if child.tag in {
            _tag("bookmarkStart"),
            _tag("bookmarkEnd"),
            _tag("permStart"),
            _tag("permEnd"),
            _tag("proofErr"),
        }:
            paragraph.addprevious(child)
    parent = _paragraph_block(paragraph).getparent()
    if parent is None:
        raise error_type(f"{part}: tracked paragraph has no parent")
    parent.remove(_paragraph_block(paragraph))


def _range_problem(root: etree._Element) -> str | None:
    for name in _UNSUPPORTED_RANGES:
        if next(root.iter(_tag(name)), None) is not None:
            return f"{name} is unsupported"
    starts = {_tag(start): start for start, _ in _RANGE_PAIRS}
    ends = {_tag(end): start for start, end in _RANGE_PAIRS}
    stack: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for element in (e for e in root.iter() if e.tag in starts or e.tag in ends):
        name = starts.get(element.tag) or ends[element.tag]
        marker_id = element.get(_tag("id"))
        if marker_id is None:
            return f"{name} is malformed"
        key = (name, marker_id)
        if element.tag in starts:
            if key in seen:
                return f"{name} is malformed"
            seen.add(key)
            stack.append(key)
        elif not stack or stack.pop() != key:
            return f"{name} is malformed"
    return f"{stack[-1][0]} is malformed" if stack else None


def _drop_range_markers(root: etree._Element) -> None:
    names = {name for pair in _RANGE_PAIRS for name in pair}
    for name in names:
        for element in list(root.iter(_tag(name))):
            _remove(element)


def _paragraph_for_mark(mark: etree._Element) -> etree._Element | None:
    rpr = mark.getparent()
    ppr = rpr.getparent() if rpr is not None else None
    paragraph = ppr.getparent() if ppr is not None else None
    return paragraph if paragraph is not None and paragraph.tag == _tag("p") else None


def _is_paragraph_mark(element: etree._Element) -> bool:
    return _paragraph_for_mark(element) is not None and element.getparent().tag == _tag("rPr")


def _paragraph_block(paragraph: etree._Element) -> etree._Element:
    content = paragraph.getparent()
    control = content.getparent() if content is not None else None
    if (
        content is not None
        and content.tag == _tag("sdtContent")
        and control is not None
        and control.tag == _tag("sdt")
        and len(content) == 1
    ):
        return control
    return paragraph


def _next_paragraph_for_mark(mark: etree._Element) -> etree._Element | None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        return None
    candidate = _paragraph_block(paragraph).getnext()
    if candidate is None:
        return None
    if candidate.tag == _tag("p"):
        return candidate
    if candidate.tag != _tag("sdt"):
        return None
    content = candidate.find(_tag("sdtContent"))
    return (
        content[0]
        if content is not None and len(content) == 1 and content[0].tag == _tag("p")
        else None
    )


def _merge_paragraph_into_next(mark: etree._Element) -> bool:
    paragraph = _paragraph_for_mark(mark)
    next_paragraph = _next_paragraph_for_mark(mark)
    if paragraph is None or next_paragraph is None:
        return False
    at = 1 if len(next_paragraph) and next_paragraph[0].tag == _tag("pPr") else 0
    for child in [node for node in paragraph if node.tag != _tag("pPr")]:
        next_paragraph.insert(at, child)
        at += 1
    block = _paragraph_block(paragraph)
    parent = block.getparent()
    if parent is None:
        return False
    parent.remove(block)
    return True


def _is_content_control_paragraph(mark: etree._Element) -> bool:
    paragraph = _paragraph_for_mark(mark)
    return paragraph is not None and _paragraph_block(paragraph) is not paragraph


def _paragraph_has_original_content(paragraph: etree._Element) -> bool:
    tags = [_tag("t"), _tag("delText"), *(_tag(n) for n in _NON_TEXT)]
    for node in paragraph.iter(*tags):
        if node.tag in {_tag("t"), _tag("delText")} and not (node.text or "").strip():
            continue
        if not any(a.tag in {_tag("ins"), _tag("moveTo")} for a in node.iterancestors()):
            return True
    return False


def _is_comment_reference_run(element: etree._Element) -> bool:
    return element.tag == _tag("r") and element.find(f".//{_tag('commentReference')}") is not None


def _is_empty_text_container(element: etree._Element) -> bool:
    if element.tag not in {_tag("r"), _tag("del"), _tag("moveFrom")}:
        return False
    for node in element.iter(_tag("t"), _tag("delText"), *(_tag(n) for n in _NON_TEXT)):
        if node.tag not in {_tag("t"), _tag("delText")} or (node.text or "").strip():
            return False
    return True


def _drop_numbering_leftover(paragraph: etree._Element, drop_comments: bool) -> None:
    if paragraph.getparent() is None:
        return
    text = "".join(n.text or "" for n in paragraph.iter(_tag("t"), _tag("delText"))).strip()
    ppr = paragraph.find(_tag("pPr"))
    numbered = ppr is not None and ppr.find(_tag("numPr")) is not None
    token = text.rstrip(".)").isdigit() and bool(text)
    if not ((numbered and not text) or token):
        return
    if any(next(paragraph.iter(_tag(n)), None) is not None for n in _NON_TEXT):
        return
    if not drop_comments and any(
        next(paragraph.iter(_tag(n)), None) is not None
        for n in ("commentRangeStart", "commentRangeEnd", "commentReference")
    ):
        return
    block = _paragraph_block(paragraph)
    parent = block.getparent()
    if parent is not None:
        parent.remove(block)


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


def _tag(name: str) -> str:
    return f"{{{_W}}}{name}"


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
