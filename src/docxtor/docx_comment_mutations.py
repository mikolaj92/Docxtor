from __future__ import annotations

from dataclasses import dataclass, replace
from hashlib import sha256
from io import BytesIO
from typing import Any

from docx import Document as PyDocxDocument
from docx.oxml.ns import qn

from .common import DocumentError
from .docx_comments import (
    _ensure_thread_parts,
    _existing_comments_part,
    _restore_thread_sidecars,
)
from .docx_models import AddressableComment
from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries
from .docx_package_transaction import (
    PackageDisposition,
    PackageDispositionStatus,
    _package_bytes,
    _resolve_target,
    _source_part_for_relationship,
)
from .docx_review_models import OperationReceipt, OperationStatus, ReviewCoverage
from .docx_stories import index_stories
from .docx_units import _paragraph_visible_text, _replace_plain_range


class CommentMutationError(DocumentError):
    """A comment mutation could not be proven complete."""


@dataclass(frozen=True)
class CommentRange:
    locator: str
    start_offset: int
    end_offset: int
    expected_text: str | None = None


@dataclass(frozen=True)
class CommentAuthor:
    author: str
    initials: str | None = None
    date: str | None = None


@dataclass(frozen=True)
class CommentMutationResult:
    data: bytes
    receipt: OperationReceipt
    comments: tuple[AddressableComment, ...]


@dataclass(frozen=True)
class CommentReplyResult:
    data: bytes
    reply: AddressableComment
    parent_id: str
    receipt: OperationReceipt
    coverage: ReviewCoverage
    parts: tuple[PackageDisposition, ...]


def append_comment_reply(
    data: bytes,
    parent_comment_id: str,
    text: str,
    author: CommentAuthor,
) -> CommentReplyResult:
    """Append a native reply to an identity in these exact input bytes."""
    if not isinstance(data, bytes) or not isinstance(parent_comment_id, str):
        raise CommentMutationError("reply requires bytes and an exact string comment ID")
    if not isinstance(text, str) or not text.strip():
        raise CommentMutationError("comment text must not be empty")
    if not isinstance(author, CommentAuthor) or not isinstance(author.author, str):
        raise CommentMutationError("reply requires CommentAuthor")
    try:
        return _append_comment_reply(data, parent_comment_id, text, author)
    except (PackageError, ValueError, TypeError, KeyError) as exc:
        raise CommentMutationError(f"cannot append comment reply: {exc}") from exc


def _reply_graph(data: bytes) -> tuple[dict[str, PackageEntry], dict[str, Any]]:
    from .docx_comment_threads import comment_parent_graph
    from .docx_ns import _THREAD_REL_BY_TARGET, CT_NS, REL_NS, W14_NS, W15_NS
    from .docx_review_inventory import inventory_review_markup

    entries = {entry.name: entry for entry in read_package_entries(data)}
    roots = {
        name: parse_package_xml(entry.data, part_name=name)
        for name, entry in entries.items()
        if name.endswith((".xml", ".rels"))
    }
    comments = roots.get("word/comments.xml")
    if comments is None or comments.tag != qn("w:comments"):
        raise CommentMutationError("native comments part is unavailable")
    expected_roots = {
        "word/commentsExtended.xml": f"{{{W15_NS}}}commentsEx",
        "word/commentsIds.xml": "{http://schemas.microsoft.com/office/word/2016/wordml/cid}commentsIds",
        "word/people.xml": f"{{{W15_NS}}}people",
    }
    for name, expected in expected_roots.items():
        if name in roots and roots[name].tag != expected:
            raise CommentMutationError(f"unsupported comment sidecar root: {name}")
    for name, root in roots.items():
        if name != "word/comments.xml" and any(root.iter(qn("w:comment"))):
            raise CommentMutationError("unsupported relocated comment body")
    ids: set[str] = set()
    para_ids: set[str] = set()
    for comment in comments.iter(qn("w:comment")):
        identity = comment.get(qn("w:id"), "")
        if (
            comment.getparent() is not comments
            or not identity.isascii()
            or not identity.isdecimal()
            or str(int(identity)) != identity
            or identity in ids
            or not list(comment.iter(qn("w:p")))
        ):
            raise CommentMutationError("duplicate, missing or unsupported comment identity")
        ids.add(identity)
    for root in roots.values():
        for paragraph in root.iter(qn("w:p")):
            identity = paragraph.get(f"{{{W14_NS}}}paraId")
            if identity is None:
                continue
            if (
                len(identity) != 8
                or any(c not in "0123456789ABCDEFabcdef" for c in identity)
                or identity.upper() in para_ids
            ):
                raise CommentMutationError("duplicate or unsupported paragraph identity")
            para_ids.add(identity.upper())
    _, issues = comment_parent_graph(roots)
    if issues:
        raise CommentMutationError("; ".join(item.message for item in issues))
    for node in roots.get("word/commentsExtended.xml", ()):
        if node.tag == f"{{{W15_NS}}}commentEx" and node.get(f"{{{W15_NS}}}done") not in {
            None,
            "0",
            "1",
            "true",
            "false",
        }:
            raise CommentMutationError("unsupported comment completion value")
    # These newer records require additional identity/state semantics, not a blind copy.
    if any(
        name.startswith("word/comments")
        and name.endswith(".xml")
        and name
        not in {
            "word/comments.xml",
            "word/commentsExtended.xml",
            "word/commentsIds.xml",
        }
        for name in entries
    ):
        raise CommentMutationError("unsupported comment extension representation")
    known = dict(_THREAD_REL_BY_TARGET)
    known["comments.xml"] = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    )
    incoming: dict[str, list[tuple[str, Any]]] = {}
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        if root.tag != f"{{{REL_NS}}}Relationships":
            raise CommentMutationError("invalid relationship root")
        seen: set[str] = set()
        for node in root:
            identity = node.get("Id")
            if not identity or identity in seen:
                raise CommentMutationError("duplicate or missing relationship identity")
            seen.add(identity)
            target = _resolve_target(_source_part_for_relationship(name), node.get("Target", ""))
            if target in {"word/" + target for target in known} or node.get("Type") in {
                item[0] for item in known.values()
            }:
                incoming.setdefault(target, []).append((name, node))
                if (
                    node.get("TargetMode") is not None
                    or target not in entries
                    or target not in {"word/" + item for item in known}
                ):
                    raise CommentMutationError("orphan, relocated or external comment relationship")
    for name, root in roots.items():
        for node in root.iter(f"{{{W15_NS}}}commentEx"):
            if name != "word/commentsExtended.xml" or node.getparent() is not root:
                raise CommentMutationError("unsupported comment thread representation")
    types = roots["[Content_Types].xml"]
    declared: set[tuple[str, str | None]] = set()
    for node in types:
        key = (
            node.tag,
            node.get("PartName") if node.tag == f"{{{CT_NS}}}Override" else node.get("Extension"),
        )
        if key in declared:
            raise CommentMutationError("duplicate content type declaration")
        declared.add(key)
    for target, (rel_type, content_type) in known.items():
        name = "word/" + target
        if name not in entries:
            continue
        links = incoming.get(name, [])
        overrides = [node for node in types if node.get("PartName") == "/" + name]
        if (
            len(links) != 1
            or links[0][0] != "word/_rels/document.xml.rels"
            or links[0][1].get("Type") != rel_type
            or len(overrides) != 1
            or overrides[0].get("ContentType") != content_type
        ):
            raise CommentMutationError("ambiguous comment relationship or content type")
    cid = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
    owners = {
        paragraph.get(f"{{{W14_NS}}}paraId"): comment.get(qn("w:id"))
        for comment in comments
        for paragraph in comment.iter(qn("w:p"))
        if paragraph.get(f"{{{W14_NS}}}paraId") is not None
    }
    seen_paras: set[str] = set()
    durable_ids: set[str] = set()
    seen_cid_owners: set[str] = set()
    for node in roots.get("word/commentsIds.xml", ()):
        if node.tag != cid + "commentId":
            continue
        para = node.get(cid + "paraId")
        durable = node.get(cid + "durableId", "")
        if (
            para not in owners
            or para in seen_paras
            or owners.get(para) in seen_cid_owners
            or durable.upper() in durable_ids
            or len(durable) != 8
            or any(c not in "0123456789abcdefABCDEF" for c in durable)
        ):
            raise CommentMutationError("ambiguous commentsIds identity")
        seen_paras.add(para)
        seen_cid_owners.add(owners[para])
        durable_ids.add(durable.upper())
    inventory = inventory_review_markup(data)
    if inventory.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment reply source coverage is incomplete")
    return entries, roots


def _append_comment_reply(
    data: bytes,
    parent_id: str,
    text: str,
    author: CommentAuthor,
) -> CommentReplyResult:
    from lxml import etree

    from .docx_comment_threads import comment_parent_graph
    from .docx_ns import _THREAD_REL_BY_TARGET, CT_NS, REL_NS, W14_NS, W15_NS
    from .docx_review_inventory import inventory_review_markup

    entries, roots = _reply_graph(data)
    before = inventory_review_markup(data)
    graph_before, _ = comment_parent_graph(roots)
    comments = roots["word/comments.xml"]
    parent = next((node for node in comments if node.get(qn("w:id")) == parent_id), None)
    if parent is None:
        raise CommentMutationError(f"unknown comment ID: {parent_id}")
    # Snapshots verify old subtrees after adding only the parent's required identity.
    para_attr = f"{{{W14_NS}}}paraId"
    used = {
        value.upper()
        for root in roots.values()
        for node in root.iter()
        for attr, value in node.attrib.items()
        if attr.endswith(("}paraId", "}durableId"))
    }

    def allocate() -> str:
        value = 1
        while f"{value:08X}" in used:
            value += 1
        if value > 0x7FFFFFFF:
            raise CommentMutationError("paragraph identity space exhausted")
        identity = f"{value:08X}"
        used.add(identity)
        return identity

    ex_name = "word/commentsExtended.xml"
    ex = roots.get(ex_name)
    if ex is None:
        ex = etree.Element(f"{{{W15_NS}}}commentsEx", nsmap={"w15": W15_NS})
        roots[ex_name] = ex
    parent_paras = list(parent.iter(qn("w:p")))
    owned = {node.get(para_attr) for node in parent_paras if node.get(para_attr)}
    records = [node for node in ex if node.get(f"{{{W15_NS}}}paraId") in owned]
    if records:
        parent_para = records[0].get(f"{{{W15_NS}}}paraId")
    else:
        paragraph = next((node for node in parent_paras if node.get(para_attr)), parent_paras[0])
        parent_para = paragraph.get(para_attr)
        if parent_para is None:
            parent_para = allocate()
            paragraph.set(para_attr, parent_para)
        etree.SubElement(ex, f"{{{W15_NS}}}commentEx", {f"{{{W15_NS}}}paraId": parent_para})
    old_comments = [etree.tostring(node) for node in comments]
    old_ex = [etree.tostring(node) for node in ex]
    reply_id = str(max(int(node.get(qn("w:id"))) for node in comments.iter(qn("w:comment"))) + 1)
    reply_para = allocate()
    reply = etree.SubElement(
        comments,
        qn("w:comment"),
        {
            qn("w:id"): reply_id,
            qn("w:author"): author.author,
        },
    )
    for key, value in (("initials", author.initials), ("date", author.date)):
        if value is not None:
            reply.set(qn("w:" + key), value)
    paragraph = etree.SubElement(reply, qn("w:p"), {para_attr: reply_para})
    run = etree.SubElement(paragraph, qn("w:r"))
    for index, line in enumerate(text.split("\n")):
        if index:
            etree.SubElement(run, qn("w:br"))
        node = etree.SubElement(run, qn("w:t"), {qn("xml:space"): "preserve"})
        node.text = line
    etree.SubElement(
        ex,
        f"{{{W15_NS}}}commentEx",
        {
            f"{{{W15_NS}}}paraId": reply_para,
            f"{{{W15_NS}}}paraIdParent": parent_para,
        },
    )
    changed = {"word/comments.xml", ex_name}
    cid_name = "word/commentsIds.xml"
    old_cid = None
    if cid_name in roots:
        cid = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
        old_cid = [etree.tostring(node) for node in roots[cid_name]]
        etree.SubElement(
            roots[cid_name],
            cid + "commentId",
            {
                cid + "paraId": reply_para,
                cid + "durableId": allocate(),
            },
        )
        changed.add(cid_name)
    if ex_name not in entries:
        rel_type, content_type = _THREAD_REL_BY_TARGET["commentsExtended.xml"]
        rels = roots["word/_rels/document.xml.rels"]
        rel_ids = {node.get("Id") for node in rels}
        index = 1
        while f"rId{index}" in rel_ids:
            index += 1
        etree.SubElement(
            rels,
            f"{{{REL_NS}}}Relationship",
            {
                "Id": f"rId{index}",
                "Type": rel_type,
                "Target": "commentsExtended.xml",
            },
        )
        etree.SubElement(
            roots["[Content_Types].xml"],
            f"{{{CT_NS}}}Override",
            {
                "PartName": "/" + ex_name,
                "ContentType": content_type,
            },
        )
        changed.update({"word/_rels/document.xml.rels", "[Content_Types].xml"})
    output = dict(entries)
    for name in changed:
        payload = etree.tostring(roots[name], encoding="UTF-8", xml_declaration=True)
        output[name] = (
            replace(entries[name], data=payload)
            if name in entries
            else PackageEntry(
                name,
                payload,
            )
        )
    candidate = _package_bytes(tuple(output.values()))
    actual_entries, actual_roots = _reply_graph(candidate)
    after = inventory_review_markup(candidate)
    confirmed = next((item for item in after.comments if item.comment_id == reply_id), None)
    graph_after, _ = comment_parent_graph(actual_roots)
    if (
        confirmed is None
        or confirmed.parent_id != parent_id
        or confirmed.text != text
        or confirmed.author != author.author
        or confirmed.initials != author.initials
        or confirmed.date != author.date
        or after.comments[:-1] != before.comments
        or graph_after != {**graph_before, reply_id: parent_id}
        or [etree.tostring(node) for node in actual_roots["word/comments.xml"]][:-1] != old_comments
        or [etree.tostring(node) for node in actual_roots[ex_name]][:-1] != old_ex
        or (
            old_cid is not None
            and [etree.tostring(node) for node in actual_roots[cid_name]][:-1] != old_cid
        )
    ):
        raise CommentMutationError("comment reply preservation was not confirmed")
    parts: list[PackageDisposition] = []
    affected: list[str] = []
    for name, entry in actual_entries.items():
        previous = entries.get(name)
        if previous is None:
            status = PackageDispositionStatus.CREATED
        elif previous.data == entry.data:
            status = PackageDispositionStatus.PRESERVED
        elif name in changed:
            status = PackageDispositionStatus.REWRITTEN
        else:
            raise CommentMutationError(f"unexpected payload change: {name}")
        if status is not PackageDispositionStatus.PRESERVED:
            affected.append(name)
        parts.append(
            PackageDisposition(
                name,
                status,
                sha256(previous.data).hexdigest() if previous else None,
                sha256(entry.data).hexdigest(),
                "append_comment_reply" if name in changed else None,
            )
        )
    return CommentReplyResult(
        candidate,
        confirmed,
        parent_id,
        OperationReceipt(
            "append_comment_reply",
            OperationStatus.APPLIED,
            tuple(sorted(affected)),
            created_ids=(reply_id,),
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(candidate).hexdigest(),
        ),
        after.coverage,
        tuple(parts),
    )


def add_comment(
    data: bytes,
    target: CommentRange,
    text: str,
    author: CommentAuthor,
) -> CommentMutationResult:
    from .docx_review_geometry import (
        PhysicalCommentRange,
        PhysicalCommentSpan,
        project_docx_review_geometry,
    )
    from .docx_review_geometry_comments import add_physical_comment_bytes

    geometry = project_docx_review_geometry(data)
    paragraph = next((item for item in geometry.paragraphs if item.locator == target.locator), None)
    if paragraph is None:
        raise CommentMutationError(f"unknown comment locator: {target.locator}")
    if not 0 <= target.start_offset < target.end_offset <= len(paragraph.text):
        raise CommentMutationError(f"invalid comment range for {target.locator}")
    selected = paragraph.text[target.start_offset : target.end_offset]
    if target.expected_text is not None and selected != target.expected_text:
        raise CommentMutationError(f"comment range text changed at {target.locator}")
    return add_physical_comment_bytes(
        data,
        PhysicalCommentRange(
            geometry.document_sha256,
            (
                PhysicalCommentSpan(
                    target.locator, target.start_offset, target.end_offset, selected
                ),
            ),
        ),
        text,
        author,
    )


def update_comment(
    data: bytes,
    comment_id: str,
    text: str,
    *,
    expected_text: str | None = None,
    author: CommentAuthor | None = None,
) -> CommentMutationResult:
    from .docx_review_inventory import inventory_review_markup
    from .docx_review_models import ReviewCoverage

    before = inventory_review_markup(data)
    if before.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment update source coverage is incomplete")
    selected = next((item for item in before.comments if item.comment_id == comment_id), None)
    if selected is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    ids = tuple(
        item.comment_id
        for item in before.comments
        if item.comment_id == comment_id
        or (selected.region_set_id is not None and item.region_set_id == selected.region_set_id)
    )
    candidate = data
    affected: set[str] = set()
    for identity in ids:
        result = _update_comment(
            candidate, identity, text, expected_text=expected_text, author=author
        )
        candidate = result.data
        affected.update(result.receipt.affected_parts)
    after = inventory_review_markup(candidate)
    if after.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment update result coverage is incomplete")
    return CommentMutationResult(
        candidate,
        OperationReceipt(
            operation="update_comment",
            status=OperationStatus.APPLIED,
            affected_parts=tuple(sorted(affected)),
            locator=selected.locator,
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(candidate).hexdigest(),
        ),
        after.comments,
    )


def _update_comment(
    data: bytes,
    comment_id: str,
    text: str,
    *,
    expected_text: str | None = None,
    author: CommentAuthor | None = None,
) -> CommentMutationResult:
    if not text:
        raise CommentMutationError("comment text must not be empty")
    document = PyDocxDocument(BytesIO(data))
    stories = index_stories(document)
    existing = next(
        (comment for comment in stories.comments if comment.comment_id == comment_id),
        None,
    )
    if existing is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    canonical = existing.user_note_text if existing.user_note_text is not None else existing.text
    if expected_text is not None and canonical != expected_text:
        raise CommentMutationError(f"comment text changed at {comment_id}")
    from .docx_comment_scopes import scope_link

    comments_part = _existing_comments_part(document)
    if comments_part is None:
        raise CommentMutationError("comment body part is unavailable for update")
    comment_element = next(
        node
        for node in comments_part.element.iter(qn("w:comment"))
        if node.get(qn("w:id")) == comment_id
    )
    bridge = scope_link(comment_element)
    extra = [
        segment
        for segment in stories.segments
        if (segment.container_id or "").startswith(f"comment:{comment_id}:")
        and segment.container_id != existing.container_id
        and not (
            bridge is not None
            and segment.container_id in stories.paragraphs_by_container
            and stories.paragraphs_by_container[segment.container_id]._p is bridge.paragraph
        )
    ]
    if extra:
        raise CommentMutationError(f"comment {comment_id} spans multiple paragraphs")
    paragraph = stories.paragraphs_by_container.get(existing.container_id)
    if paragraph is None:
        raise CommentMutationError(f"unknown comment locator: {existing.container_id}")
    visible = _paragraph_visible_text(paragraph)
    if visible:
        _replace_plain_range(paragraph._p, 0, len(visible), text)
    elif paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)
    if author is not None:
        _apply_comment_author(document, comment_id, author)
    payload = _serialize(document, thread_parts=stories.thread_parts)
    after = index_stories(PyDocxDocument(BytesIO(payload)))
    updated = next((item for item in after.comments if item.comment_id == comment_id), None)
    if (
        updated is None
        or updated.text != text + ("\n" + bridge.label if bridge is not None else "")
        or (bridge is not None and updated.user_note_text != text)
        or updated.locator != existing.locator
        or updated.anchor_text != existing.anchor_text
        or updated.parent_id != existing.parent_id
    ):
        raise CommentMutationError("comment update was not confirmed after round-trip")
    if author is None and (
        updated.author != existing.author
        or updated.initials != existing.initials
        or updated.date != existing.date
    ):
        raise CommentMutationError("comment update was not confirmed after round-trip")
    return CommentMutationResult(
        data=payload,
        receipt=OperationReceipt(
            operation="update_comment",
            status=OperationStatus.APPLIED,
            affected_parts=("word/comments.xml",),
            created_ids=(),
            locator=existing.locator,
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(payload).hexdigest(),
        ),
        comments=tuple(after.comments),
    )


def remove_comments(data: bytes, comment_ids: set[str] | None = None) -> CommentMutationResult:
    """Remove complete native region sets and descendants in one transaction."""
    from .docx_comment_removal import remove_comment_sets

    return remove_comment_sets(data, comment_ids)


def _apply_comment_author(document: Any, comment_id: str, author: CommentAuthor) -> None:
    comments_part = _existing_comments_part(document)
    if comments_part is None:
        raise CommentMutationError(f"unknown comment ID: {comment_id}")
    for comment in comments_part.element.findall(qn("w:comment")):
        if comment.get(qn("w:id")) != comment_id:
            continue
        comment.set(qn("w:author"), author.author)
        if author.initials is not None:
            comment.set(qn("w:initials"), author.initials)
        if author.date is not None:
            comment.set(qn("w:date"), author.date)
        return
    raise CommentMutationError(f"unknown comment ID: {comment_id}")


def _serialize(doc: Any, thread_parts: dict[str, tuple[bytes, str]] | None = None) -> bytes:
    if thread_parts:
        _ensure_thread_parts(doc.part.package, thread_parts)
    buffer = BytesIO()
    doc.save(buffer)
    payload = buffer.getvalue()
    if thread_parts:
        return _restore_thread_sidecars(payload, thread_parts)
    return payload
