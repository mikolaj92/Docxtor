from __future__ import annotations

import tomllib
from pathlib import Path

from docxtor import __version__

ROOT = Path(__file__).resolve().parents[1]


def _pyproject_version() -> str:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    version = data["project"]["version"]
    if not isinstance(version, str):
        raise TypeError(f"project.version must be a string, got {type(version).__name__}")
    return version


def test_package_version_matches_pyproject() -> None:
    assert __version__ == _pyproject_version()


def test_package_version_is_ahead_of_mismatched_tag() -> None:
    """v0.4.1 still ships dist 0.4.0; current metadata must not repeat that pin."""
    assert __version__ != "0.4.0"
    assert tuple(int(part) for part in __version__.split(".")) >= (0, 4, 2)
