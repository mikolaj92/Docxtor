from __future__ import annotations

from dataclasses import dataclass, replace
from hashlib import sha256
from posixpath import basename, dirname, join, normpath
from uuid import uuid4

from docx.oxml.ns import qn
from lxml import etree

from .docx_comment_mutations import CommentAuthor, CommentMutationError, CommentMutationResult
from .docx_comment_threads import comment_parent_graph, region_metadata
from .docx_ns import CT_NS, REL_NS, W14_NS, W15_NS
from .docx_package import PackageEntry, parse_package_xml
from .docx_review_geometry import _read_geometry_state
from .docx_review_geometry_comments import (
    _append_comment,
    _comments_part,
    _package_bytes,
    _xml_bytes,
)
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import OperationReceipt, OperationStatus, ReviewCoverage

_CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
_CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
_EXT_REL = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
_EXT_TYPE = "application/vnd.ms-word.commentsExtended+xml"


@dataclass(frozen=True)
class CommentReplyTarget:
    document_sha256: str
    parent_id: str


@dataclass(frozen=True)
class CommentReplyResult(CommentMutationResult):
    reply_id: str
    parent_id: str
    coverage: ReviewCoverage


def append_comment_reply_bytes(
    data: bytes, target: CommentReplyTarget, text: str, author: CommentAuthor
) -> CommentReplyResult:
    if sha256(data).hexdigest() != target.document_sha256:
        raise CommentMutationError("comment reply target belongs to different source bytes")
    if not text:
        raise CommentMutationError("comment reply text must not be empty")
    state = _read_geometry_state(data)
    if state.projection.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment reply source geometry is incomplete")
    parts = {entry.name: entry for entry in state.entries}
    changed = {}
    comments_name, comments = _comments_part(state, parts, changed)
    if changed:
        raise CommentMutationError("comment reply requires an existing comments part")
    bodies = list(comments)
    ids = [node.get(qn("w:id")) for node in bodies]
    if (
        any(node.tag != qn("w:comment") for node in bodies)
        or any(value is None or not value.isdecimal() or str(int(value)) != value for value in ids)
        or len(ids) != len(set(ids))
        or target.parent_id not in ids
    ):
        raise CommentMutationError("unknown, duplicate or unsupported native comment identity")
    parent = bodies[ids.index(target.parent_id)]
    if region_metadata(parent)[0] is not None:
        raise CommentMutationError("reply to a multi-story region set is unsupported")
    parents, issues = comment_parent_graph(state.roots)
    if issues:
        raise CommentMutationError("comment reply parent graph is incomplete")
    if any(root.tag == f"{{{_CEX}}}commentsExtensible" for root in state.roots.values()):
        raise CommentMutationError("extensible comment identities are unsupported for replies")
    rels_name = join(
        dirname(state.main_part_name), "_rels", basename(state.main_part_name) + ".rels"
    )
    rels = parse_package_xml(parts[rels_name].data, part_name=rels_name)
    extended = [
        (name, root) for name, root in state.roots.items() if root.tag == f"{{{W15_NS}}}commentsEx"
    ]
    if len(extended) > 1:
        raise CommentMutationError("multiple native extended comment parts")
    types_name = "[Content_Types].xml"
    types = parse_package_xml(parts[types_name].data, part_name=types_name)
    if extended:
        extended_name, extended_root = extended[0]
        links = [node for node in rels if node.get("Type") == _EXT_REL]
        if (
            len(links) != 1
            or links[0].get("TargetMode") == "External"
            or normpath(join(dirname(state.main_part_name), links[0].get("Target", ""))).lstrip("/")
            != extended_name
        ):
            raise CommentMutationError("extended comment relationship is ambiguous")
        declarations = [node for node in types if node.get("PartName") == f"/{extended_name}"]
        if len(declarations) != 1 or declarations[0].get("ContentType") != _EXT_TYPE:
            raise CommentMutationError("extended comment content type is ambiguous")
    else:
        if parents or any(node.get("Type") == _EXT_REL for node in rels):
            raise CommentMutationError("extended comment part is missing")
        extended_name = join(dirname(comments_name), "commentsExtended.xml")
        if extended_name in parts or any(
            node.get("PartName") == f"/{extended_name}" for node in types
        ):
            raise CommentMutationError("extended comment part identity is occupied")
        extended_root = etree.Element(f"{{{W15_NS}}}commentsEx", nsmap={"w15": W15_NS})
        used_rels = {node.get("Id") for node in rels}
        rid = 1
        while f"rId{rid}" in used_rels:
            rid += 1
        etree.SubElement(
            rels,
            f"{{{REL_NS}}}Relationship",
            Id=f"rId{rid}",
            Type=_EXT_REL,
            Target=basename(extended_name),
        )
        etree.SubElement(
            types, f"{{{CT_NS}}}Override", PartName=f"/{extended_name}", ContentType=_EXT_TYPE
        )
        changed[rels_name] = rels
        changed[types_name] = types
    paragraph_ids = {
        node.get(f"{{{W14_NS}}}paraId")
        for root in state.roots.values()
        for node in root.iter(qn("w:p"))
    }

    def new_paragraph_id() -> str:
        while True:
            value = f"{int(uuid4().hex[:8], 16) & 0x7FFFFFFF:08X}"
            if value != "00000000" and value not in paragraph_ids:
                paragraph_ids.add(value)
                return value

    records = {}
    for node in extended_root.findall(f"{{{W15_NS}}}commentEx"):
        para_id = node.get(f"{{{W15_NS}}}paraId")
        if para_id in records or node.get(f"{{{W15_NS}}}done") not in {
            None,
            "0",
            "1",
            "false",
            "true",
            "on",
            "off",
        }:
            raise CommentMutationError("duplicate or unsupported comment completion record")
        records[para_id] = node
    for body in bodies:
        paragraphs = list(body.iter(qn("w:p")))
        if not paragraphs:
            raise CommentMutationError("native comment has no paragraph identity")
        last = paragraphs[-1]
        para_id = last.get(f"{{{W14_NS}}}paraId")
        if extended:
            if para_id not in records:
                raise CommentMutationError(
                    "native thread must identify each comment's last paragraph"
                )
        else:
            if para_id is None:
                para_id = new_paragraph_id()
                last.set(f"{{{W14_NS}}}paraId", para_id)
                changed[comments_name] = comments
            etree.SubElement(
                extended_root, f"{{{W15_NS}}}commentEx", {f"{{{W15_NS}}}paraId": para_id}
            )
    parent_para = list(parent.iter(qn("w:p")))[-1].get(f"{{{W14_NS}}}paraId")
    candidate = max((int(value) for value in ids), default=-1) + 1
    if candidate >= 2**31:
        raise CommentMutationError("native comment identity space is exhausted")
    reply_id = str(candidate)
    reply_para = new_paragraph_id()
    _append_comment(comments, reply_id, text, author)
    comments[-1].find(qn("w:p")).set(f"{{{W14_NS}}}paraId", reply_para)
    etree.SubElement(
        extended_root,
        f"{{{W15_NS}}}commentEx",
        {
            f"{{{W15_NS}}}paraId": reply_para,
            f"{{{W15_NS}}}paraIdParent": parent_para,
            f"{{{W15_NS}}}done": "0",
        },
    )
    changed[comments_name] = comments
    changed[extended_name] = extended_root
    cid_parts = [
        (name, root) for name, root in state.roots.items() if root.tag == f"{{{_CID}}}commentsIds"
    ]
    if cid_parts:
        raise CommentMutationError("durable comment identities are unsupported for replies")
    entries = tuple(
        replace(entry, data=_xml_bytes(changed[entry.name])) if entry.name in changed else entry
        for entry in state.entries
    ) + tuple(
        PackageEntry(name=name, data=_xml_bytes(root))
        for name, root in changed.items()
        if name not in parts
    )
    payload = _package_bytes(entries)
    after = _read_geometry_state(payload)
    actual_entries = {entry.name: entry for entry in after.entries}
    if set(actual_entries) != {entry.name for entry in entries} or any(
        entry.name not in actual_entries or actual_entries[entry.name].data != entry.data
        for entry in entries
    ):
        raise CommentMutationError("comment reply package payloads were not preserved")
    for name, expected in changed.items():
        if name in after.roots and etree.tostring(after.roots[name]) != etree.tostring(expected):
            raise CommentMutationError("comment reply XML transformation was not preserved")
    actual_parents, issues = comment_parent_graph(after.roots)
    inventory = inventory_review_markup(payload)
    created = next((item for item in inventory.comments if item.comment_id == reply_id), None)
    if (
        after.projection.coverage is not ReviewCoverage.COMPLETE
        or inventory.coverage is not ReviewCoverage.COMPLETE
        or issues
        or actual_parents != {**parents, reply_id: target.parent_id}
        or created is None
        or created.parent_id != target.parent_id
        or created.text != text
        or created.author != author.author
        or (author.initials is not None and created.initials != author.initials)
        or (author.date is not None and created.date != author.date)
    ):
        raise CommentMutationError("native comment reply was not confirmed after round-trip")
    for old in state.projection.comment_anchors:
        kept = next(
            (
                item
                for item in after.projection.comment_anchors
                if item.comment_id == old.comment_id
            ),
            None,
        )
        if kept is None or replace(kept, document_sha256=old.document_sha256) != old:
            raise CommentMutationError("existing comment scope was not preserved")
    return CommentReplyResult(
        payload,
        OperationReceipt(
            operation="append_comment_reply",
            status=OperationStatus.APPLIED,
            affected_parts=tuple(sorted(changed)),
            created_ids=(reply_id,),
            before_sha256=target.document_sha256,
            after_sha256=sha256(payload).hexdigest(),
        ),
        inventory.comments,
        reply_id,
        target.parent_id,
        ReviewCoverage.COMPLETE,
    )
