"""Fail-closed comparison of two DOCX facts snapshots."""

from __future__ import annotations

from hashlib import sha256

from .docx_facts_features import _FEATURE_CATEGORIES
from .docx_facts_models import (
    ChangeKind,
    DocxComparison,
    FactChange,
    FactDiagnostic,
    FactsCoverage,
    RelationshipFact,
    Source,
    TransformPolicy,
)
from .docx_facts_scan import docx_facts


def compare_docx(
    before: Source, after: Source, policy: TransformPolicy | None = None
) -> DocxComparison:
    """Compare two fail-closed snapshots and apply an optional neutral allow-list."""
    b = docx_facts(before)
    a = docx_facts(after)
    selected = policy or TransformPolicy()
    part_changes = _diff(
        {x.name: x.sha256 for x in b.parts}, {x.name: x.sha256 for x in a.parts}, "part"
    )
    rel_changes = _diff(
        {_rel_key(x): _hash_repr(x) for x in b.relationships},
        {_rel_key(x): _hash_repr(x) for x in a.relationships},
        "relationship",
    )
    surface_changes = _diff(
        {x.surface_id: x.value_sha256 for x in b.surfaces},
        {x.surface_id: x.value_sha256 for x in a.surfaces},
        "surface",
    )
    container_changes = _diff(
        {x.container_id: x.text_sha256 for x in b.paragraphs},
        {x.container_id: x.text_sha256 for x in a.paragraphs},
        "container",
    )
    fact_changes: list[FactChange] = []
    for category in _FEATURE_CATEGORIES:
        left = getattr(b, category)
        right = getattr(a, category)
        fact_changes.extend(
            _diff(
                {x.fact_id: _hash_repr(x) for x in left},
                {x.fact_id: _hash_repr(x) for x in right},
                category,
            )
        )
    violations = [c for c in part_changes if c.kind not in selected.allowed_part_changes]
    violations += [c for c in rel_changes if c.kind not in selected.allowed_relationship_changes]
    violations += [c for c in surface_changes if c.kind not in selected.allowed_surface_changes]
    violations += [c for c in container_changes if c.kind not in selected.allowed_container_changes]
    violations += [
        c
        for c in fact_changes
        if "*" not in selected.allowed_fact_categories
        and c.category not in selected.allowed_fact_categories
    ]
    diagnostics: tuple[FactDiagnostic, ...] = ()
    if selected.require_complete_coverage and (
        b.coverage is not FactsCoverage.COMPLETE or a.coverage is not FactsCoverage.COMPLETE
    ):
        diagnostics = (
            FactDiagnostic("incomplete_coverage", "comparison requires complete facts coverage"),
        )
    return DocxComparison(
        b,
        a,
        part_changes,
        rel_changes,
        surface_changes,
        container_changes,
        tuple(fact_changes),
        tuple(violations),
        diagnostics,
    )

def _hash_repr(value: object) -> str:
    return sha256(repr(value).encode()).hexdigest()

def _rel_key(rel: RelationshipFact) -> str:
    return f"{rel.source_part}#{rel.relationship_id}"

def _diff(before: dict[str, str], after: dict[str, str], category: str) -> tuple[FactChange, ...]:
    changes: list[FactChange] = []
    for key in sorted(before.keys() | after.keys()):
        if key not in before:
            changes.append(FactChange(ChangeKind.CREATED, category, key, None, after[key]))
        elif key not in after:
            changes.append(FactChange(ChangeKind.REMOVED, category, key, before[key], None))
        elif before[key] != after[key]:
            changes.append(FactChange(ChangeKind.CHANGED, category, key, before[key], after[key]))
    return tuple(changes)
