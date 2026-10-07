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


_SMALL_API_PHRASES = (
    "small api",
    "small interface",
    "small surface",
    "small set",
    "small public",
)


def test_status_does_not_sell_a_small_api_in_any_wording() -> None:
    """Issue #169: Status must stop describing the API as small, in any wording."""
    status = _readme().split("## Status", 1)[1].split("\n## ", 1)[0].lower()
    hits = [phrase for phrase in _SMALL_API_PHRASES if phrase in status]
    assert not hits, f"README Status still sells a small API (phrases: {hits})"


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


def test_status_prose_names_only_exports() -> None:
    """Backticked identifiers in Status prose (count sentence) must be exports."""
    status = _readme().split("## Status", 1)[1].split("\n## ", 1)[0]
    prose = "\n".join(line for line in status.splitlines() if not line.startswith("- "))
    non_exports = sorted(
        name
        for name in re.findall(r"`(\w+)`", prose)
        if name not in docxtor.__all__ and name not in {"docxtor", "__all__"}
    )
    assert not non_exports, f"README Status prose names non-exports: {non_exports}"


def test_status_group_tallies_match_listed_names() -> None:
    """Each Status group bullet's (N) tally must equal the names it lists."""
    status = _readme().split("## Status", 1)[1].split("\n## ", 1)[0]
    mismatches: list[str] = []
    for line in status.splitlines():
        if not line.startswith("- "):
            continue
        tally = re.search(r"\((\d+)\)", line)
        assert tally, f"Status group bullet without a count: {line[:60]}"
        listed = re.findall(r"`([^`]+)`", line.split(":", 1)[1])
        if int(tally.group(1)) != len(listed):
            mismatches.append(
                f"{line[:60]}...: claims {tally.group(1)}, lists {len(listed)}"
            )
    assert not mismatches, "Status group tallies wrong: " + "; ".join(mismatches)


def test_status_bullets_cover_exactly_the_exports() -> None:
    """Status group bullets must list every export except __version__ (prose only)."""
    expected = set(docxtor.__all__) - {"__version__"}
    listed = _status_listed_names()
    missing = sorted(expected - listed)
    extra = sorted(listed - expected)
    assert not missing and not extra, (
        f"Status bullets drift from docxtor.__all__: missing={missing} extra={extra}"
    )
