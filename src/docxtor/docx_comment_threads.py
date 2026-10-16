from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn

from .docx_ns import W14_NS, W15_NS
from .docx_review_models import ReviewDiagnostic

REGION_NS = "urn:docxtor:comment-regions:v1"


def comment_parent_graph(
    roots: dict[str, Any],
) -> tuple[dict[str, str], tuple[ReviewDiagnostic, ...]]:
    owners: dict[str, str] = {}
    diagnostics: list[ReviewDiagnostic] = []
    for name, root in roots.items():
        for comment in root.iter(qn("w:comment")):
            identity = comment.get(qn("w:id"))
            for paragraph in comment.iter(qn("w:p")):
                para_id = paragraph.get(f"{{{W14_NS}}}paraId")
                if para_id is None:
                    continue
                if identity is None or para_id in owners:
                    diagnostics.append(
                        ReviewDiagnostic(
                            "invalid_comment_thread",
                            "Comment paragraph identity is duplicate or missing its owner.",
                            name,
                        )
                    )
                else:
                    owners[para_id] = identity
    parents: dict[str, str] = {}
    seen: set[str] = set()
    seen_owners: set[str] = set()
    for name, root in roots.items():
        for node in root.iter(f"{{{W15_NS}}}commentEx"):
            para_id = node.get(f"{{{W15_NS}}}paraId")
            parent_para = node.get(f"{{{W15_NS}}}paraIdParent")
            identity = owners.get(para_id or "")
            parent = owners.get(parent_para or "")
            if (
                identity is None
                or para_id is None
                or para_id in seen
                or identity in seen_owners
                or (parent_para is not None and parent is None)
            ):
                diagnostics.append(
                    ReviewDiagnostic(
                        "invalid_comment_thread",
                        "Comment thread contains a duplicate or unresolved relationship.",
                        name,
                    )
                )
                continue
            seen.add(para_id)
            seen_owners.add(identity)
            if parent is not None:
                if identity in parents or identity == parent:
                    diagnostics.append(
                        ReviewDiagnostic(
                            "invalid_comment_thread",
                            "Comment thread contains conflicting or self-parent relationships.",
                            name,
                        )
                    )
                else:
                    parents[identity] = parent
    for identity in parents:
        visited: set[str] = set()
        cursor = identity
        while cursor in parents:
            if cursor in visited:
                diagnostics.append(ReviewDiagnostic("cyclic_comment_thread", identity))
                break
            visited.add(cursor)
            cursor = parents[cursor]
    return parents, tuple(diagnostics)


def region_metadata(comment: Any) -> tuple[str | None, int | None]:
    from .docx_comment_scopes import _SCOPE_NAME, scope_link

    link = scope_link(comment)
    native = _SCOPE_NAME.fullmatch(link.name) if link is not None else None
    identity = comment.get(f"{{{REGION_NS}}}regionSet")
    size = comment.get(f"{{{REGION_NS}}}regionSetSize")
    if identity is None and size is None:
        return (native[1], int(native[3])) if native is not None else (None, None)
    if not identity or size is None or not size.isdecimal() or int(size) < 1:
        raise ValueError("invalid native comment region-set metadata")
    if native is not None and (identity != native[1] or int(size) != int(native[3])):
        raise ValueError("native scope group conflicts with comment group metadata")
    return identity, int(size)


def validate_region_sets(roots: dict[str, Any]) -> tuple[ReviewDiagnostic, ...]:
    groups: dict[str, list[tuple[int, str | None]]] = {}
    diagnostics: list[ReviewDiagnostic] = []
    for name, root in roots.items():
        for comment in root.iter(qn("w:comment")):
            try:
                identity, size = region_metadata(comment)
            except ValueError as exc:
                diagnostics.append(ReviewDiagnostic("invalid_comment_region_set", str(exc), name))
                continue
            if identity is not None and size is not None:
                groups.setdefault(identity, []).append((size, comment.get(qn("w:id"))))
    for identity, members in groups.items():
        sizes = {size for size, _ in members}
        ids = {comment_id for _, comment_id in members}
        if len(sizes) != 1 or len(members) != members[0][0] or len(ids) != len(members):
            diagnostics.append(ReviewDiagnostic("incomplete_comment_region_set", identity))
    return tuple(diagnostics)
