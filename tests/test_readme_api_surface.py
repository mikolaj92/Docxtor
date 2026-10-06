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


def _status_listed_names() -> set[str]:
    """Backticked names in the Status API-group bullets (prose backticks excluded)."""
    status = _readme().split("## Status", 1)[1].split("\n## ", 1)[0]
    names: set[str] = set()
    for line in status.splitlines():
        if line.startswith("- "):
            names.update(re.findall(r"`([^`]+)`", line))
    return names


def test_status_lists_only_exported_names() -> None:
    """README Status must not advertise names docxtor no longer exports."""
    stale = sorted(_status_listed_names() - set(docxtor.__all__))
    assert not stale, f"README Status lists names not in docxtor.__all__: {stale}"
