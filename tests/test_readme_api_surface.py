"""README must not misstate the public API surface (issue #169)."""

from __future__ import annotations

import re
from pathlib import Path

import docxtor

ROOT = Path(__file__).resolve().parents[1]


def _readme() -> str:
    return (ROOT / "README.md").read_text(encoding="utf-8")


def test_status_states_true_export_count() -> None:
    match = re.search(r"`docxtor.__all__` exports (\d+) names", _readme())
    assert match, "README Status must state the docxtor.__all__ export count"
    assert int(match.group(1)) == len(docxtor.__all__)


def test_status_does_not_claim_a_small_api() -> None:
    assert "public API is small" not in _readme()


def test_readme_lists_every_exported_name() -> None:
    missing = [name for name in docxtor.__all__ if f"`{name}`" not in _readme()]
    assert not missing, f"README does not list exports: {missing}"
