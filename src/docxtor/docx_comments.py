from __future__ import annotations

from io import BytesIO
from typing import Any
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from docx.document import Document as DocxDocumentType
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .docx_models import AddressableComment
from .docx_ns import (
    _COMMENTS_EXTENDED_PART,
    _CONTENT_TYPES_PART,
    _DOCUMENT_RELS_PART,
    _THREAD_PARTS,
    _THREAD_REL_BY_TARGET,
    _UNSUPPORTED_MOVE_TAGS,
    CT_NS,
    REL_NS,
    W14_NS,
    W15_NS,
    W_T,
)
from .docx_units import _paragraph_units
from .docx_xml import _clark, _local_tag, _w_get


def _unsupported_revision_reason(doc: DocxDocumentType) -> str | None:
    roots: list[Any] = [doc.element]
    for section in doc.sections:
        for story in (section.header, section.footer):
            if not story._has_definition:
                continue
            roots.append(story._definition.element)
    comments_part = _existing_comments_part(doc)
    if comments_part is not None:
        roots.append(comments_part.element)
    for root in roots:
        for element in root.iter():
            local = _local_tag(element.tag)
            if local in _UNSUPPORTED_MOVE_TAGS:
                return local
            if local in {"ins", "del"}:
                for child in element:
                    if _local_tag(child.tag) in {"p", "tbl", "tr", "tc"}:
                        return f"block-{local}"
    return None


def _existing_comments_part(doc: DocxDocumentType) -> Any | None:
    try:
        return doc.part.part_related_by(RT.COMMENTS)
    except KeyError:
        return None


def _package_part_name(partname: Any) -> str:
    text = str(partname)
    return text[1:] if text.startswith("/") else text


def _is_thread_part_name(name: str) -> bool:
    return (
        name in _THREAD_PARTS
        or name.startswith("word/_rels/comments")
        or name.startswith("word/_rels/people")
    )


def _capture_thread_parts(package: Any) -> dict[str, tuple[bytes, str]]:
    captured: dict[str, tuple[bytes, str]] = {}
    for part in package.iter_parts():
        name = _package_part_name(part.partname)
        if _is_thread_part_name(name):
            captured[name] = (part.blob, part.content_type)
    return captured


def _ensure_thread_parts(package: Any, captured: dict[str, tuple[bytes, str]]) -> None:
    if not captured:
        return
    document_part = package.main_document_part
    existing = {_package_part_name(part.partname): part for part in package.iter_parts()}
    for name, (blob, content_type) in captured.items():
        if name.startswith("word/_rels/"):
            continue
        part = existing.get(name)
        if part is None:
            part = Part(PackURI(f"/{name}"), content_type, blob, package)
            target = name.rsplit("/", 1)[-1]
            reltype, _ctype = _THREAD_REL_BY_TARGET.get(target, (None, None))
            if reltype is None:
                continue
            document_part.relate_to(part, reltype)
            existing[name] = part
        else:
            part._blob = blob


def _restore_thread_sidecars(data: bytes, captured: dict[str, tuple[bytes, str]]) -> bytes:
    if not captured:
        return data
    with ZipFile(BytesIO(data)) as bundle:
        names = set(bundle.namelist())
        entries = [(info, bundle.read(info.filename)) for info in bundle.infolist()]
    missing = {name: blob for name, (blob, _ctype) in captured.items() if name not in names}
    if not missing:
        return data

    restored: list[tuple[Any, bytes]] = []
    for info, blob in entries:
        if info.filename == _CONTENT_TYPES_PART:
            blob = _merge_content_type_overrides(blob, missing, captured)
        elif info.filename == _DOCUMENT_RELS_PART:
            blob = _merge_document_relationships(blob, missing)
        restored.append((info, blob))
    for name, blob in sorted(missing.items()):
        restored.append((name, blob))

    output = BytesIO()
    with ZipFile(output, mode="w", compression=ZIP_DEFLATED) as bundle:
        for info, blob in restored:
            if isinstance(info, str):
                bundle.writestr(info, blob)
            else:
                bundle.writestr(info, blob)
    return output.getvalue()


def _merge_content_type_overrides(
    rendered: bytes,
    missing: dict[str, bytes],
    captured: dict[str, tuple[bytes, str]],
) -> bytes:
    root = ElementTree.fromstring(rendered)
    have = {override.get("PartName") for override in root.findall(_clark(CT_NS, "Override"))}
    changed = False
    for name in missing:
        part_name = f"/{name}"
        if part_name in have:
            continue
        _blob, content_type = captured[name]
        override = ElementTree.SubElement(root, _clark(CT_NS, "Override"))
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        changed = True
    if not changed:
        return rendered
    return ElementTree.tostring(root, encoding="UTF-8", xml_declaration=True)


def _merge_document_relationships(rendered: bytes, missing: dict[str, bytes]) -> bytes:
    root = ElementTree.fromstring(rendered)
    have_targets = {rel.get("Target") for rel in root.findall(_clark(REL_NS, "Relationship"))}
    used_ids = {rel.get("Id") for rel in root.findall(_clark(REL_NS, "Relationship"))}
    next_id = 1
    changed = False
    for name in missing:
        target = name.rsplit("/", 1)[-1]
        reltype, _ctype = _THREAD_REL_BY_TARGET.get(target, (None, None))
        if reltype is None or target in have_targets:
            continue
        while f"rId{next_id}" in used_ids:
            next_id += 1
        rel = ElementTree.SubElement(root, _clark(REL_NS, "Relationship"))
        rel.set("Id", f"rId{next_id}")
        rel.set("Type", reltype)
        rel.set("Target", target)
        used_ids.add(f"rId{next_id}")
        have_targets.add(target)
        next_id += 1
        changed = True
    if not changed:
        return rendered
    return ElementTree.tostring(root, encoding="UTF-8", xml_declaration=True)


def _comment_date_attr(element: Any) -> str | None:
    return _w_get(element, "date")


def _comment_parent_ids(comments_part: Any, package: Any) -> dict[str, str]:
    parts = {_package_part_name(part.partname): part for part in package.iter_parts()}
    extended = parts.get(_COMMENTS_EXTENDED_PART)
    if comments_part is None or extended is None:
        return {}
    para_to_comment: dict[str, str] = {}
    for comment in comments_part.element.findall(qn("w:comment")):
        comment_id = _w_get(comment, "id")
        if comment_id is None:
            continue
        for paragraph in comment.iter(qn("w:p")):
            para_id = paragraph.get(f"{{{W14_NS}}}paraId")
            if para_id:
                para_to_comment[para_id] = comment_id
                break
    if not para_to_comment:
        return {}
    try:
        root = ElementTree.fromstring(extended.blob)
    except ElementTree.ParseError:
        return {}
    parents: dict[str, str] = {}
    for element in root.iter():
        if element.tag != _clark(W15_NS, "commentEx"):
            continue
        para_id = element.get(_clark(W15_NS, "paraId"))
        parent_para = element.get(_clark(W15_NS, "paraIdParent"))
        if not para_id or not parent_para:
            continue
        comment_id = para_to_comment.get(para_id)
        parent_id = para_to_comment.get(parent_para)
        if comment_id and parent_id and comment_id != parent_id:
            parents[comment_id] = parent_id
    return parents


def _comment_anchors(
    paragraphs_by_container: dict[str, Paragraph],
) -> dict[str, tuple[str, str]]:
    open_ids: dict[str, list[str]] = {}
    started_at: dict[str, str] = {}
    finished: dict[str, tuple[str, str]] = {}
    reference_at: dict[str, str] = {}
    for locator, paragraph in paragraphs_by_container.items():
        if locator.startswith("comment:"):
            continue
        p_element = paragraph._p
        for node in p_element.iter():
            local = _local_tag(node.tag)
            if local == "commentRangeStart":
                comment_id = _w_get(node, "id")
                if comment_id is None:
                    continue
                open_ids.setdefault(comment_id, [])
                started_at.setdefault(comment_id, locator)
            elif local == "commentRangeEnd":
                comment_id = _w_get(node, "id")
                if comment_id is None or comment_id not in open_ids:
                    continue
                finished[comment_id] = (
                    started_at.get(comment_id, locator),
                    "".join(open_ids.pop(comment_id)),
                )
                started_at.pop(comment_id, None)
            elif local == "commentReference":
                comment_id = _w_get(node, "id")
                if comment_id is not None:
                    reference_at.setdefault(comment_id, locator)
            elif open_ids and node.tag == W_T and node.text:
                for buffer in open_ids.values():
                    buffer.append(node.text)
            elif open_ids and node.tag == qn("w:tab"):
                for buffer in open_ids.values():
                    buffer.append("\t")
            elif open_ids and node.tag in (qn("w:br"), qn("w:cr")):
                for buffer in open_ids.values():
                    buffer.append("\n")
    for comment_id, buffer in open_ids.items():
        finished[comment_id] = (started_at.get(comment_id, ""), "".join(buffer))
    for comment_id, locator in reference_at.items():
        current = finished.get(comment_id)
        if current is None or not current[0]:
            finished[comment_id] = (locator, current[1] if current is not None else "")
    return finished


def _collect_comments(
    comments_part: Any | None,
    paragraphs_by_container: dict[str, Paragraph],
    package: Any,
) -> list[AddressableComment]:
    if comments_part is None:
        return []
    anchors = _comment_anchors(paragraphs_by_container)
    parents = _comment_parent_ids(comments_part, package)
    comments: list[AddressableComment] = []
    for comment_elm in comments_part.element.findall(qn("w:comment")):
        comment_id = _w_get(comment_elm, "id")
        if comment_id is None:
            continue
        prefix = f"comment:{comment_id}"
        texts: list[str] = []
        first_container: str | None = None
        for local_idx, paragraph in enumerate(comment_elm.iter(qn("w:p"))):
            text = "".join(unit.text for unit in _paragraph_units(paragraph))
            if not text:
                continue
            if first_container is None:
                first_container = f"{prefix}:p:{local_idx}"
            texts.append(text)
        if not texts:
            continue
        locator, anchor_text = anchors.get(comment_id, (None, ""))
        comments.append(
            AddressableComment(
                comment_id=comment_id,
                container_id=first_container or f"{prefix}:p:0",
                text="\n".join(texts),
                author=_w_get(comment_elm, "author") or "",
                initials=_w_get(comment_elm, "initials"),
                locator=locator or None,
                anchor_text=anchor_text,
                parent_id=parents.get(comment_id),
                date=_comment_date_attr(comment_elm),
            )
        )
    return comments
