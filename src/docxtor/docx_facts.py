"""Public facts surface: data model, scanner, and comparison.

The mechanical scanner lives in ``docx_facts_scan``; this module keeps the
historical ``docxtor.docx_facts`` import path stable for callers and tests.
"""

from __future__ import annotations

from .docx_facts_compare import compare_docx
from .docx_facts_models import (
    ChangeKind,
    ContainerCoordinate,
    ContentTypeFact,
    DocxComparison,
    DocxFactsSnapshot,
    DocxStructureSnapshot,
    FactChange,
    FactDiagnostic,
    FactsCoverage,
    NamedFact,
    PageBreakFact,
    PageLayoutFacts,
    ParagraphFact,
    PartFact,
    RelationshipFact,
    Source,
    StoryFact,
    TransformPolicy,
    UnreadablePartFact,
)
from .docx_facts_scan import docx_facts, snapshot_docx

__all__ = [
    "ChangeKind",
    "ContainerCoordinate",
    "ContentTypeFact",
    "DocxComparison",
    "DocxFactsSnapshot",
    "DocxStructureSnapshot",
    "FactChange",
    "FactDiagnostic",
    "FactsCoverage",
    "NamedFact",
    "PageBreakFact",
    "PartFact",
    "PageLayoutFacts",
    "ParagraphFact",
    "RelationshipFact",
    "Source",
    "StoryFact",
    "TransformPolicy",
    "UnreadablePartFact",
    "compare_docx",
    "docx_facts",
    "snapshot_docx",
]
