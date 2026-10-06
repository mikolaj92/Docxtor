"""Selective acceptance/rejection of inventoried existing DOCX revisions."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import replace
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from .docx_package import PackageEntry, PackageError, read_package_entries
from .docx_revision_selection_inventory import (
    _inspect,
    _Inspection,
    _node_sha256,
    _tag,
)
from .docx_revision_selection_models import (
    RevisionDecision,
    RevisionDisposition,
    RevisionDispositionCoverage,
    RevisionDispositionInventory,
    RevisionDispositionReceipt,
    RevisionDispositionResult,
    RevisionDispositionTarget,
    RevisionSelectionError,
)


def dispose_revisions_bytes(
    data: bytes,
    inventory: RevisionDispositionInventory,
    decisions: Iterable[RevisionDecision],
) -> RevisionDispositionResult:
    """Resolve explicit existing targets, preserving every unselected revision.

    The exact source and the entire immutable inventory must still match. A pair
    of adjacent opposite-kind inline wrappers is an atomic replacement group:
    both targets must be selected with the same disposition. This conservative
    physical grammar does not infer an author's intent or a caller's authority.
    No intermediate bytes are returned if selection or postflight fails.
    """
    try:
        inspection = _inspect(data)
        before = inspection.inventory
        if not isinstance(inventory, RevisionDispositionInventory) or inventory != before:
            raise RevisionSelectionError("input or revision inventory is stale or modified")
        if before.coverage is not RevisionDispositionCoverage.COMPLETE:
            raise RevisionSelectionError("complete, unambiguous revision inventory is required")
        selected = _select(before, decisions)
        if not selected:
            return RevisionDispositionResult(data, sha256(data).hexdigest(), before, before, (), ())
        changed_parts = tuple(
            name
            for name in inspection.roots
            if any(target.part_name == name for target, _ in selected)
        )
        preserved = _preserved_fingerprints(before, {target.target_id for target, _ in selected})
        receipts: list[RevisionDispositionReceipt] = []
        for target, disposition in selected:
            node = inspection.nodes[target.target_id]
            if _node_sha256(node) != target.node_sha256 or node.getparent() is None:
                raise RevisionSelectionError(f"revision target changed: {target.target_id}")
            retain = (target.kind == "ins" and disposition is RevisionDisposition.ACCEPT) or (
                target.kind == "del" and disposition is RevisionDisposition.REJECT
            )
            _resolve(node, retain=retain, restore_deletion=retain and target.kind == "del")
            if node.getparent() is not None:
                raise RevisionSelectionError(f"revision was not resolved: {target.target_id}")
            receipts.append(
                RevisionDispositionReceipt(
                    target=target,
                    disposition=disposition,
                    coverage=RevisionDispositionCoverage.COMPLETE,
                    resolved=True,
                    retained_text=target.text if retain else "",
                )
            )
        planned = {name: _tree_bytes(inspection.roots[name]) for name in changed_parts}
        entries = tuple(
            replace(entry, data=_serialize(inspection.roots[entry.name]))
            if entry.name in changed_parts
            else entry
            for entry in inspection.entries
        )
        output = _write_bytes(entries)
        postflight = _inspect(output)
        _verify_postflight(inspection, postflight, changed_parts, planned, preserved)
        return RevisionDispositionResult(
            data=output,
            output_sha256=sha256(output).hexdigest(),
            before=before,
            after=postflight.inventory,
            receipts=tuple(receipts),
            changed_parts=changed_parts,
        )
    except RevisionSelectionError:
        raise
    except (PackageError, etree.LxmlError, ValueError, TypeError) as exc:
        raise RevisionSelectionError(f"revision disposition failed: {exc}") from exc


def _select(
    inventory: RevisionDispositionInventory, decisions: Iterable[RevisionDecision]
) -> tuple[tuple[RevisionDispositionTarget, RevisionDisposition], ...]:
    catalog = {target.target_id: target for target in inventory.revisions}
    selected: dict[str, RevisionDisposition] = {}
    for decision in decisions:
        if not isinstance(decision, RevisionDecision):
            raise RevisionSelectionError("every selection must be a RevisionDecision")
        if decision.target_id in selected:
            raise RevisionSelectionError(f"duplicate selected target: {decision.target_id}")
        target = catalog.get(decision.target_id)
        if target is None:
            raise RevisionSelectionError(f"unknown selected target: {decision.target_id}")
        if not target.supported:
            raise RevisionSelectionError(
                f"unsupported selected target: {target.unsupported_reason}"
            )
        try:
            disposition = RevisionDisposition(decision.disposition)
        except (TypeError, ValueError) as exc:
            raise RevisionSelectionError("disposition must be accept or reject") from exc
        selected[decision.target_id] = disposition
    for target_id, disposition in selected.items():
        target = catalog[target_id]
        if any(
            selected.get(required) is not disposition for required in target.required_target_ids
        ):
            raise RevisionSelectionError(
                "replacement group requires both targets and one disposition"
            )
    # Source order makes physical mutation and receipt order independent of caller ordering.
    return tuple(
        (target, selected[target.target_id])
        for target in inventory.revisions
        if target.target_id in selected
    )


def _resolve(node: etree._Element, *, retain: bool, restore_deletion: bool) -> None:
    parent = node.getparent()
    if parent is None:
        raise RevisionSelectionError("selected revision has no parent")
    if retain:
        if restore_deletion:
            for text in node.iter(_tag("delText")):
                text.tag = _tag("t")
        leading = node.text or ""
        previous = node.getprevious()
        if leading:
            if previous is None:
                parent.text = (parent.text or "") + leading
            else:
                previous.tail = (previous.tail or "") + leading
        for child in list(node):
            node.addprevious(child)
    # The wrapper's XML formatting tail belongs outside the selected content.
    if node.tail:
        previous = node.getprevious()
        if previous is None:
            parent.text = (parent.text or "") + node.tail
        else:
            previous.tail = (previous.tail or "") + node.tail
    parent.remove(node)


def _fingerprint(target: RevisionDispositionTarget) -> tuple[str, str, str, str | None, str]:
    return (target.part_name, target.namespace, target.kind, target.revision_id, target.node_sha256)


def _preserved_fingerprints(
    inventory: RevisionDispositionInventory, selected_ids: set[str]
) -> tuple[tuple[str, str, str, str | None, str], ...]:
    return tuple(
        _fingerprint(target)
        for target in inventory.revisions
        if target.target_id not in selected_ids
    )


def _verify_postflight(
    before: _Inspection,
    after: _Inspection,
    changed_parts: tuple[str, ...],
    planned: dict[str, bytes],
    preserved: tuple[tuple[str, str, str, str | None, str], ...],
) -> None:
    if after.inventory.coverage is not RevisionDispositionCoverage.COMPLETE:
        raise RevisionSelectionError("output revision inventory is incomplete")
    if tuple(_fingerprint(target) for target in after.inventory.revisions) != preserved:
        raise RevisionSelectionError("selected or unselected revision verification failed")
    if tuple(entry.name for entry in before.entries) != tuple(
        entry.name for entry in after.entries
    ):
        raise RevisionSelectionError("OPC members changed during revision disposition")
    for original, rendered in zip(before.entries, after.entries, strict=True):
        if original.name not in changed_parts:
            if original != rendered:
                raise RevisionSelectionError(f"untouched OPC part changed: {original.name}")
        elif replace(original, data=rendered.data) != rendered:
            raise RevisionSelectionError(f"OPC member metadata changed: {original.name}")
    for name, expected in planned.items():
        if _tree_bytes(after.roots[name]) != expected:
            raise RevisionSelectionError(f"planned XML transformation was not preserved: {name}")


def _tree_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root.getroottree(), method="c14n", with_comments=True)


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root.getroottree(), xml_declaration=True, encoding="UTF-8")


def _write_bytes(entries: tuple[PackageEntry, ...]) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w") as archive:
        for entry in entries:
            archive.writestr(entry.zip_info(), entry.data)
    output = buffer.getvalue()
    read_package_entries(output)
    return output
