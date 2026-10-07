"""README must not misstate the public API surface (issue #169)."""

from __future__ import annotations

import re
from pathlib import Path

import docxtor

ROOT = Path(__file__).resolve().parents[1]


def _readme() -> str:
    return (ROOT / "README.md").read_text(encoding="utf-8")


def test_status_states_true_export_count() -> None:
    matches = re.findall(r"`docxtor\.__all__` exports (\d+) names", _readme())
    assert matches, "README Status must state the docxtor.__all__ export count"
    wrong = [count for count in matches if int(count) != len(docxtor.__all__)]
    assert not wrong, (
        f"README states wrong export count(s) {wrong}; actual {len(docxtor.__all__)}"
    )


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


def test_all_exported_names_are_importable() -> None:
    """The Status claim 'importable from docxtor' must hold for every export."""
    unimportable = sorted(
        name for name in docxtor.__all__ if not hasattr(docxtor, name)
    )
    assert not unimportable, f"__all__ names not importable from docxtor: {unimportable}"


def test_readme_import_examples_reference_exports() -> None:
    """Every name README tells consumers to import from docxtor must be exported."""
    readme = _readme()
    names: set[str] = set()
    for match in re.findall(r"from docxtor import ([^(\n][^\n]*)", readme):
        names.update(part.strip() for part in match.split(",") if part.strip())
    for match in re.findall(r"from docxtor import \(([^)]*)\)", readme):
        names.update(part.strip().rstrip(",") for part in match.split("\n") if part.strip())
    stale = sorted(names - set(docxtor.__all__))
    assert not stale, f"README import examples use non-exported names: {stale}"


_DOCXDOCUMENT_MEMBERS_USED_IN_EXAMPLES = (
    "open",
    "apply_replacements",
    "apply_surface_replacements",
    "add_comment",
    "update_comment",
    "replace_revision",
    "remove_comments",
    "publish",
    "save_docx",
    "to_bytes",
    "inventory",
    "resolve_paragraph_locator",
    "resolve_run_locator",
    "segments",
    "spans",
    "comments",
    "paragraph_resolutions",
    "alternate_content_coverage",
)


def test_readme_example_apis_exist() -> None:
    """DocxDocument members used in README examples must exist (no copy-paste AttributeError)."""
    missing = sorted(
        name
        for name in _DOCXDOCUMENT_MEMBERS_USED_IN_EXAMPLES
        if not hasattr(docxtor.DocxDocument, name)
    )
    assert not missing, (
        f"README examples use DocxDocument members that do not exist: {missing}; "
        "update the examples or the member list"
    )


_DTO_FIELDS_USED_IN_EXAMPLES = (
    ("DocumentBytes", "filename"),
    ("DocumentBytes", "content_type"),
    ("DocumentBytes", "data"),
    ("DocxInventory", "coverage"),
    ("DocxInventory", "unknown_parts"),
    ("DocxInventory", "unreadable_parts"),
    ("DocxInventory", "surfaces"),
    ("DocumentSurface", "surface_id"),
    ("DocumentSurface", "kind"),
    ("DocumentSurface", "capability"),
    ("DocumentSurface", "value_sha256"),
    ("DocumentSurface", "external"),
    ("SurfaceMutationResult", "unresolved"),
    ("SurfaceMutationResult", "data"),
    ("SegmentReplacement", "container_id"),
    ("SegmentReplacement", "text"),
    ("SegmentReplacement", "start_offset"),
    ("SegmentReplacement", "end_offset"),
    ("SegmentReplacement", "span_id"),
    ("AddressableSpan", "span_id"),
    ("AddressableSpan", "role"),
    ("AddressableSpan", "text"),
    ("TextSegment", "container_id"),
    ("TextSegment", "text"),
    ("ParagraphResolution", "identity"),
    ("ParagraphResolution", "paragraph_index"),
    ("ParagraphResolution", "value"),
    ("ParagraphResolution", "runs"),
    ("RevisionDispositionReceipt", "resolved"),
)


def _dto_member_exists(cls: type, name: str) -> bool:
    return (
        name in getattr(cls, "__dataclass_fields__", {})
        or name in getattr(cls, "__annotations__", {})
        or hasattr(cls, name)
    )


def test_readme_example_dto_fields_exist() -> None:
    """DTO fields README examples read must exist on the exported classes."""
    missing = sorted(
        f"{cls_name}.{field}"
        for cls_name, field in _DTO_FIELDS_USED_IN_EXAMPLES
        if not _dto_member_exists(getattr(docxtor, cls_name), field)
    )
    assert not missing, (
        f"README examples read DTO fields that do not exist: {missing}; "
        "update the examples or the field list"
    )


def test_readme_module_calls_are_exports() -> None:
    """README must not call docxtor.<name> for a name that is not exported."""
    phantom = sorted(
        set(re.findall(r"docxtor\.(\w+)", _readme())) - set(docxtor.__all__) - {"__all__"}
    )
    assert not phantom, (
        f"README calls docxtor attributes that are not exported: {phantom}; "
        "use a real export or import-from syntax"
    )


def _block_imported_names(block: str) -> set[str]:
    names: set[str] = set()
    multi = re.search(r"from docxtor import \(([^)]*)\)", block, re.S)
    if multi:
        names |= {part.strip() for part in multi.group(1).split(",")}
    for single in re.findall(r"from docxtor import ([\w, ]+)", block):
        names |= {part.strip() for part in single.split(",")}
    return names


def test_readme_code_blocks_import_what_they_use() -> None:
    """Each README python block must import the exported names it uses (copy-paste works)."""
    unimported: list[str] = []
    for index, block in enumerate(re.findall(r"```python\n(.*?)```", _readme(), re.S)):
        imported = _block_imported_names(block)
        if "import docxtor" in block:
            continue
        for name in docxtor.__all__:
            if name == "__version__":
                continue
            used_bare = re.search(rf"(?<![\w.]){re.escape(name)}\b", block)
            if used_bare and name not in imported:
                unimported.append(f"block {index}: {name}")
    assert not unimported, (
        f"README code blocks use exported names without importing them: {unimported}; "
        "add the missing names to the block's docxtor import"
    )
