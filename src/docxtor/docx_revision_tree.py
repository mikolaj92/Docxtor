"""Accept/reject tree transforms for tracked revisions."""

from __future__ import annotations

from copy import deepcopy

from lxml import etree

from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries
from .docx_revision_inventory import (
    _UNSUPPORTED_RANGES,
    _present_kinds,
    inventory_revisions_bytes,
)
from .docx_revision_models import (
    _STRICT_W,
    _W14,
    AcceptRevisionsError,
    RejectRevisionsError,
    RevisionOperation,
    RevisionOperationError,
    RevisionOperationReceipt,
    _tag,
)
from .docx_revision_package import (
    _has_comments,
    _is_comment_part,
    _is_word_xml,
    _remove,
    _replace_entry,
    _serialize,
    _strip_comment_anchors,
    _strip_comment_content_types,
    _strip_comment_relationships,
    _unwrap,
    _write_bytes,
)
from .docx_revision_paragraphs import (
    _drop_numbering_leftover,
    _is_content_control_paragraph,
    _is_paragraph_mark,
    _merge_paragraph_into_next,
    _next_paragraph_for_mark,
    _reject_inserted_mark,
    _validate_rejected_inserted_mark,
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
