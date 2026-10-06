"""Public revisions surface: inventory and accept/reject operations.

The implementation lives in ``docx_revision_inventory`` and
``docx_revision_tree``; this module keeps the historical
``docxtor.docx_revisions`` import path stable for callers and tests.
"""

from __future__ import annotations

from .docx_revision_inventory import inventory_revisions_bytes
from .docx_revision_models import (
    AcceptRevisionsError,
    RejectRevisionsError,
    Revision,
    RevisionCoverageDiagnostic,
    RevisionInventory,
    RevisionInventoryCoverage,
    RevisionKind,
    RevisionOperation,
    RevisionOperationError,
    RevisionOperationReceipt,
)
from .docx_revision_tree import accept_all_revisions_bytes, reject_all_revisions_bytes

__all__ = [
    "AcceptRevisionsError",
    "RejectRevisionsError",
    "Revision",
    "RevisionCoverageDiagnostic",
    "RevisionInventory",
    "RevisionInventoryCoverage",
    "RevisionKind",
    "RevisionOperation",
    "RevisionOperationError",
    "RevisionOperationReceipt",
    "accept_all_revisions_bytes",
    "inventory_revisions_bytes",
    "reject_all_revisions_bytes",
]
