"""Package-wide physical discovery for selective revision disposition."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, replace
from hashlib import sha256
from pathlib import PurePosixPath

from lxml import etree

from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries
from .docx_revision_selection_models import (
    RevisionDispositionCoverage,
    RevisionDispositionDiagnostic,
    RevisionDispositionInventory,
    RevisionDispositionTarget,
    RevisionSelectionError,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_REVISION_NAMES = frozenset(
    {
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
    }
)
_XML_ENCODINGS = (
    "utf-8-sig",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "utf-32",
    "utf-32-le",
    "utf-32-be",
)
_XML_BOMS = (b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff", b"\x00\x00\xfe\xff")


@dataclass
class _Inspection:
    entries: tuple[PackageEntry, ...]
    roots: dict[str, etree._Element]
    nodes: dict[str, etree._Element]
    inventory: RevisionDispositionInventory


def inspect_revision_dispositions(data: bytes) -> RevisionDispositionInventory:
    """Inspect all XML package members, including comment and orphan stories.

    Unsupported forms are explicit preserve-only targets. Incomplete package/XML
    coverage or ambiguous identities are diagnostics, never an empty inventory.
    """
    try:
        return _inspect(data).inventory
    except RevisionSelectionError:
        raise
    except (PackageError, etree.LxmlError, ValueError, TypeError) as exc:
        raise RevisionSelectionError(f"revision inventory failed: {exc}") from exc


def _inspect(data: bytes) -> _Inspection:
    if not isinstance(data, bytes):
        raise RevisionSelectionError("revision disposition requires immutable bytes")
    try:
        entries = read_package_entries(data)
    except PackageError as exc:
        raise RevisionSelectionError(str(exc)) from exc
    diagnostics: list[RevisionDispositionDiagnostic] = []
    types = _content_types(entries, diagnostics)
    roots: dict[str, etree._Element] = {}
    nodes: dict[str, etree._Element] = {}
    targets: list[RevisionDispositionTarget] = []
    for entry in entries:
        content_type = types.get(entry.name)
        if entry.name != "[Content_Types].xml" and content_type is None:
            diagnostics.append(
                RevisionDispositionDiagnostic(
                    "part_type_unknown", "part has no unambiguous content type", entry.name
                )
            )
        if not _xml_member(entry, content_type):
            continue
        try:
            root = parse_package_xml(entry.data, part_name=entry.name)
        except PackageError as exc:
            diagnostics.append(
                RevisionDispositionDiagnostic("xml_unreadable", str(exc), entry.name)
            )
            continue
        roots[entry.name] = root
        tree = root.getroottree()
        identities: dict[tuple[str, str], list[str]] = {}
        part_targets: list[RevisionDispositionTarget] = []
        for element in root.iter():
            if not _is_revision(element):
                continue
            namespace, kind = _name(element)
            locator = tree.getpath(element)
            attribute_namespace = _STRICT_W if namespace == _STRICT_W else _W
            revision_id = element.get(f"{{{attribute_namespace}}}id")
            node_hash = _node_sha256(element)
            target_id = _digest((entry.name, namespace, kind, revision_id, locator, node_hash))
            paragraph = next((a for a in element.iterancestors() if a.tag == _tag("p")), None)
            reason = _unsupported_reason(element)
            target = RevisionDispositionTarget(
                target_id=target_id,
                part_name=entry.name,
                namespace=namespace,
                kind=kind,
                revision_id=revision_id,
                locator=locator,
                node_sha256=node_hash,
                author=element.get(f"{{{attribute_namespace}}}author"),
                date=element.get(f"{{{attribute_namespace}}}date"),
                text=_physical_text(element),
                paragraph_locator=tree.getpath(paragraph) if paragraph is not None else None,
                paragraph_text=_physical_text(paragraph) if paragraph is not None else None,
                supported=reason is None,
                unsupported_reason=reason,
            )
            part_targets.append(target)
            nodes[target_id] = element
            if revision_id is None or not revision_id.strip():
                diagnostics.append(
                    RevisionDispositionDiagnostic(
                        "revision_identity_missing", "revision has no id", entry.name, locator
                    )
                )
            elif re.fullmatch(r"[+-]?[0-9]+", revision_id) is None:
                diagnostics.append(
                    RevisionDispositionDiagnostic(
                        "revision_identity_invalid",
                        "revision id is not a decimal integer",
                        entry.name,
                        locator,
                    )
                )
            else:
                # A range's start and end intentionally share an id; duplicate
                # wrappers or duplicate endpoints of the same kind are ambiguous.
                family = kind if kind.endswith(("RangeStart", "RangeEnd")) else "wrapper"
                digits = revision_id.lstrip("+-").lstrip("0") or "0"
                identity = f"-{digits}" if revision_id.startswith("-") and digits != "0" else digits
                identities.setdefault((family, identity), []).append(locator)
        for (_, revision_id), locators in identities.items():
            if len(locators) > 1:
                diagnostics.append(
                    RevisionDispositionDiagnostic(
                        "revision_identity_ambiguous",
                        f"revision id {revision_id!r} is repeated",
                        entry.name,
                        locators[0],
                    )
                )
        for element in root.iter():
            namespace, kind = _name(element)
            if namespace not in {_W, _STRICT_W} or kind not in {"delText", "delInstrText"}:
                continue
            deleted = any(
                _name(a)[1] in {"del", "moveFrom", "conflictDel"} and _is_revision(a)
                for a in element.iterancestors()
            )
            if not deleted:
                diagnostics.append(
                    RevisionDispositionDiagnostic(
                        "orphan_deleted_payload",
                        f"{kind} has no deletion ancestor",
                        entry.name,
                        tree.getpath(element),
                    )
                )
        targets.extend(_replacement_groups(part_targets, nodes))
    inventory = RevisionDispositionInventory(
        profile_id="docxtor-revision-disposition-v1",
        input_sha256=sha256(data).hexdigest(),
        inventory_sha256="",
        revisions=tuple(targets),
        coverage=(
            RevisionDispositionCoverage.INCOMPLETE
            if diagnostics
            else RevisionDispositionCoverage.COMPLETE
        ),
        diagnostics=tuple(diagnostics),
        scanned_parts=tuple(roots),
    )
    inventory = replace(inventory, inventory_sha256=_inventory_sha256(inventory))
    return _Inspection(entries, roots, nodes, inventory)


def _content_types(
    entries: tuple[PackageEntry, ...], diagnostics: list[RevisionDispositionDiagnostic]
) -> dict[str, str]:
    entry = next((e for e in entries if e.name == "[Content_Types].xml"), None)
    if entry is None:
        diagnostics.append(
            RevisionDispositionDiagnostic(
                "content_types_missing", "[Content_Types].xml is required for complete coverage"
            )
        )
        return {}
    root = parse_package_xml(entry.data, part_name=entry.name)
    if root.tag != f"{{{_CT}}}Types":
        diagnostics.append(
            RevisionDispositionDiagnostic(
                "content_types_invalid", "unexpected content types root", entry.name
            )
        )
        return {}
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if not isinstance(child.tag, str):
            continue
        if child.tag == f"{{{_CT}}}Default":
            key = (child.get("Extension") or "").lower()
            table = defaults
        elif child.tag == f"{{{_CT}}}Override":
            key = (child.get("PartName") or "").removeprefix("/")
            table = overrides
        else:
            diagnostics.append(
                RevisionDispositionDiagnostic(
                    "content_types_unknown", "unknown content type declaration", entry.name
                )
            )
            continue
        value = child.get("ContentType") or ""
        if not key or not value or key in table:
            diagnostics.append(
                RevisionDispositionDiagnostic(
                    "content_types_ambiguous",
                    "missing or repeated content type declaration",
                    entry.name,
                )
            )
            continue
        table[key] = value
    return {
        item.name: value
        for item in entries
        if (
            value := overrides.get(item.name)
            or defaults.get(PurePosixPath(item.name).suffix.removeprefix(".").lower())
        )
    }


def _xml_member(entry: PackageEntry, content_type: str | None) -> bool:
    name = entry.name.casefold()
    if (
        name.endswith((".xml", ".rels"))
        or name.startswith("customxml/")
        or (
            content_type
            and (content_type.endswith("+xml") or content_type in {"application/xml", "text/xml"})
        )
    ):
        return True
    # Unicode XML can have arbitrarily long leading whitespace, including zero
    # bytes between ASCII codepoints. Do not hide it behind the bounded probe.
    if entry.data.lstrip(b" \t\r\n\x00").startswith(b"<") or entry.data.startswith(_XML_BOMS):
        return True
    probe = entry.data[:65536]
    return any(
        probe.decode(encoding, errors="ignore").lstrip().startswith("<")
        for encoding in _XML_ENCODINGS
    )


def _is_revision(element: etree._Element) -> bool:
    namespace, kind = _name(element)
    word_namespace = namespace in {_W, _STRICT_W} or (
        namespace.startswith("http://schemas.microsoft.com/office/word/")
    )
    return word_namespace and (
        kind in _REVISION_NAMES
        or kind.endswith("PrChange")
        or ("Range" in kind and any(token in kind for token in ("Ins", "Del", "Move", "Conflict")))
    )


def _unsupported_reason(element: etree._Element) -> str | None:
    namespace, kind = _name(element)
    if namespace != _W:
        return "only transitional WordprocessingML revisions can be selected"
    if kind not in {"ins", "del"}:
        return "move, range, paragraph and property revision groups are preserve-only"
    ancestors = tuple(element.iterancestors())
    if any(_is_revision(a) for a in ancestors) or any(
        _is_revision(d) for d in element.iterdescendants()
    ):
        return "nested revision groups are preserve-only"
    if any(a.tag == f"{{{_MC}}}AlternateContent" for a in ancestors):
        return "revisions in alternative XML branches are preserve-only"
    parent = element.getparent()
    if parent is None or parent.tag not in {_tag("p"), _tag("hyperlink")}:
        return "revision is not an ordinary inline paragraph or hyperlink wrapper"
    paragraph = next((a for a in ancestors if a.tag == _tag("p")), None)
    if paragraph is None:
        return "revision has no paragraph"
    if any(_name(a)[1] in {"pPr", "rPr", "trPr", "numPr", "ctrlPr"} for a in ancestors):
        return "revision is inside a property or math control context"
    for node in paragraph.iter():
        if _is_revision(node) and any(
            a.tag in {_tag("pPr"), _tag("rPr"), _tag("numPr")}
            for a in node.iterancestors()
            if a is not paragraph
        ):
            return "paragraph or run property revision context is preserve-only"
        if _is_revision(node) and "Range" in _name(node)[1]:
            return "tracked range revision context is preserve-only"
    if not len(element) or (element.text or "").strip():
        return "revision must contain ordinary runs only"
    text_kind = "t" if kind == "ins" else "delText"
    for run in element:
        if run.tag != _tag("r") or (run.text or "").strip() or (run.tail or "").strip():
            return "revision contains an unsupported inline child"
        children = tuple(run)
        properties = [node for node in children if node.tag == _tag("rPr")]
        if len(properties) > 1 or (properties and children[0] is not properties[0]):
            return "run has malformed properties"
        content = [node for node in children if node.tag != _tag("rPr")]
        if not content:
            return "revision contains an empty run"
        for node in content:
            if node.tag not in {_tag(text_kind), _tag("tab"), _tag("br"), _tag("cr")}:
                return "revision contains field, drawing, math, marker or opaque run content"
            if len(node) or (node.tail or "").strip():
                return "revision contains malformed text content"
            if node.tag != _tag(text_kind) and (node.text or "").strip():
                return "revision contains malformed control content"
    return None


def _replacement_groups(
    targets: list[RevisionDispositionTarget], nodes: dict[str, etree._Element]
) -> list[RevisionDispositionTarget]:
    by_node = {nodes[target.target_id]: target for target in targets}
    updated: dict[str, RevisionDispositionTarget] = {}
    seen: set[str] = set()
    for target in targets:
        node = nodes[target.target_id]
        if target.target_id in seen or target.kind not in {"ins", "del"}:
            continue
        sequence = [target]
        following = node.getnext()
        while (candidate := by_node.get(following)) is not None and candidate.kind in {
            "ins",
            "del",
        }:
            sequence.append(candidate)
            following = following.getnext()
        seen.update(item.target_id for item in sequence)
        if len({item.kind for item in sequence}) < 2:
            continue
        if len(sequence) != 2 or not all(item.supported for item in sequence):
            for item in sequence:
                updated[item.target_id] = replace(
                    item,
                    supported=False,
                    unsupported_reason="adjacent mixed group is not a safe two-member pair",
                )
            continue
        group_ids = tuple(item.target_id for item in sequence)
        group_id = _digest(("adjacent_inline_replacement", group_ids))
        for item in sequence:
            updated[item.target_id] = replace(
                item, group_id=group_id, required_target_ids=group_ids
            )
    return [updated.get(target.target_id, target) for target in targets]


def _physical_text(element: etree._Element) -> str:
    parts: list[str] = []
    for node in element.iter():
        namespace, kind = _name(node)
        if namespace not in {_W, _STRICT_W}:
            continue
        if kind in {"t", "delText", "instrText", "delInstrText"}:
            parts.append(node.text or "")
        elif kind == "tab":
            parts.append("\t")
        elif kind in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts)


def _name(element: etree._Element) -> tuple[str, str]:
    if not isinstance(element.tag, str):
        return "", ""
    name = etree.QName(element)
    return name.namespace or "", name.localname


def _tag(kind: str) -> str:
    return f"{{{_W}}}{kind}"


def _node_sha256(element: etree._Element) -> str:
    return sha256(
        etree.tostring(element, method="c14n", with_comments=True, with_tail=False)
    ).hexdigest()


def _digest(value: object) -> str:
    return sha256(
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()


def _inventory_sha256(inventory: RevisionDispositionInventory) -> str:
    payload = asdict(inventory)
    payload.pop("inventory_sha256")
    return _digest(payload)
