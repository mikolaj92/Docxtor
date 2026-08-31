from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "src" / "docxtor"


def _line_count(name: str) -> int:
    return len((ROOT / name).read_text(encoding="utf-8").splitlines())


def test_docx_godfile_is_split_into_small_modules() -> None:
    """#41: mechanical DOCX work lives in small modules, not one kitchen sink."""
    expected = {
        "docx_ns.py",
        "docx_xml.py",
        "docx_models.py",
        "docx_inline.py",
        "docx_units.py",
        "docx_comments.py",
        "docx_stories.py",
        "docx.py",
    }
    present = {path.name for path in ROOT.glob("docx*.py")}
    assert expected <= present
    for name in expected:
        count = _line_count(name)
        limit = 550 if name == "docx.py" else 450
        assert count <= limit, f"{name} grew to {count} lines (limit {limit})"
