"""Atomic removal of native comment sets, replies, and their thread identities."""

from __future__ import annotations

from dataclasses import replace
from hashlib import sha256
from typing import Any

from docx.oxml.ns import qn

from .docx_comment_mutations import CommentMutationError, CommentMutationResult
from .docx_comment_threads import comment_parent_graph, region_metadata
from .docx_ns import W14_NS, W15_NS
from .docx_review_geometry import _read_geometry_state
from .docx_review_geometry_comments import _package_bytes, _revision_identity, _xml_bytes
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import OperationReceipt, OperationStatus, ReviewCoverage

_CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
_CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"


def remove_comment_sets(data: bytes, comment_ids: set[str] | None) -> CommentMutationResult:
    state = _read_geometry_state(data)
    if state.projection.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment removal source geometry is incomplete")
    bodies = {
        node.get(qn("w:id")): node
        for root in state.roots.values()
        for node in root.iter(qn("w:comment"))
    }
    existing = set(bodies)
    selected = set(existing) if comment_ids is None else set(comment_ids)
    if selected - existing:
        raise CommentMutationError(f"unknown comment IDs: {sorted(selected - existing)}")
    parents, issues = comment_parent_graph(state.roots)
    if issues:
        raise CommentMutationError("comment removal source parent graph is incomplete")
    groups = {identity: region_metadata(node)[0] for identity, node in bodies.items()}
    while True:
        group_ids = {groups[identity] for identity in selected if groups[identity] is not None}
        expanded = selected | {
            identity
            for identity, group in groups.items()
            if group is not None and group in group_ids
        }
        expanded.update(identity for identity, parent in parents.items() if parent in selected)
        if expanded == selected:
            break
        selected = expanded
    if selected == existing and selected:
        from .docx_review_cleanup import remove_all_comments_bytes

        result = remove_all_comments_bytes(data)
        return replace(
            result,
            receipt=replace(
                result.receipt, operation="remove_comments", removed_ids=tuple(sorted(selected))
            ),
        )
    if not selected:
        return CommentMutationResult(
            data,
            OperationReceipt(
                operation="remove_comments",
                status=OperationStatus.NOOP,
                affected_parts=(),
                before_sha256=sha256(data).hexdigest(),
                after_sha256=sha256(data).hexdigest(),
            ),
            inventory_review_markup(data).comments,
        )
    paragraph_ids = {
        paragraph.get(f"{{{W14_NS}}}paraId")
        for identity in selected
        for paragraph in bodies[identity].iter(qn("w:p"))
    }
    paragraph_ids.discard(None)
    durable_ids: set[str] = set()
    for root in state.roots.values():
        for node in root.iter(f"{{{_CID}}}commentId"):
            if node.get(f"{{{_CID}}}paraId") in paragraph_ids:
                identity = node.get(f"{{{_CID}}}durableId")
                if identity is not None:
                    durable_ids.add(identity)
    removed_scopes = tuple(
        scope for scope in state.projection.scope_bookmarks if scope.comment_id in selected
    )
    changed: dict[str, Any] = {}
    markers = {
        qn(f"w:{name}") for name in ("commentRangeStart", "commentRangeEnd", "commentReference")
    }
    for name, root in state.roots.items():
        for node in list(root.iter()):
            remove = (node.tag == qn("w:comment") or node.tag in markers) and node.get(
                qn("w:id")
            ) in selected
            remove |= (
                node.tag == f"{{{W15_NS}}}commentEx"
                and node.get(f"{{{W15_NS}}}paraId") in paragraph_ids
            )
            remove |= (
                node.tag == f"{{{_CID}}}commentId"
                and node.get(f"{{{_CID}}}paraId") in paragraph_ids
            )
            remove |= (
                node.tag == f"{{{_CEX}}}commentExtensible"
                and node.get(f"{{{_CEX}}}durableId") in durable_ids
            )
            remove |= node.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")} and any(
                scope.part_name == name and scope.bookmark_id == node.get(qn("w:id"))
                for scope in removed_scopes
            )
            if remove and (parent := node.getparent()) is not None:
                parent.remove(node)
                changed[name] = root
    entries = tuple(
        replace(entry, data=_xml_bytes(changed[entry.name])) if entry.name in changed else entry
        for entry in state.entries
    )
    payload = _package_bytes(entries)
    after = _read_geometry_state(payload)
    inventory = inventory_review_markup(payload)
    if after.projection.coverage is not ReviewCoverage.COMPLETE:
        raise CommentMutationError("comment removal result geometry is incomplete")
    if _revision_identity(state) != _revision_identity(after):
        raise CommentMutationError("comment removal changed native revisions")
    remaining = {comment.comment_id: comment for comment in inventory.comments}
    before = {comment.comment_id: comment for comment in inventory_review_markup(data).comments}
    if selected & remaining.keys() or any(
        remaining.get(identity) != comment
        for identity, comment in before.items()
        if identity not in selected
    ):
        raise CommentMutationError("comment removal result differs from the selected logical sets")
    return CommentMutationResult(
        payload,
        OperationReceipt(
            operation="remove_comments",
            status=OperationStatus.APPLIED,
            affected_parts=tuple(sorted(changed)),
            removed_ids=tuple(sorted(selected)),
            removed_scope_bookmarks=removed_scopes,
            before_sha256=sha256(data).hexdigest(),
            after_sha256=sha256(payload).hexdigest(),
        ),
        inventory.comments,
    )
