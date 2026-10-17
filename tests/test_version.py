from __future__ import annotations

import re
import tomllib
from pathlib import Path

from docxtor import __all__, __version__

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
    assert tuple(int(part) for part in __version__.split(".")) >= (0, 4, 4)


def test_readme_installation_pins_match_project_version() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    status = readme.split("\n## Status\n", 1)[1].split("\n## ", 1)[0]
    counts = re.findall(r"The public API exports ([0-9]+) names\.", status)
    assert counts == [str(len(__all__))]
    assert "small" not in status.lower()
    pin = f"git+https://github.com/mikolaj92/Docxtor.git@v{_pyproject_version()}"
    installation = readme.split("\n## Installation\n", 1)[1].split("\n## ", 1)[0]
    commands = installation.splitlines()
    assert f"python -m pip install {pin}" in commands
    assert f"uv add {pin}" in commands


def test_pypdf_is_not_a_runtime_dependency() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    deps = data["project"]["dependencies"]
    assert all("pypdf" not in dep for dep in deps)
    lock = (ROOT / "uv.lock").read_text(encoding="utf-8")
    assert 'name = "pypdf"' not in lock


def test_reportlab_is_a_dev_only_dependency() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    runtime = data["project"]["dependencies"]
    dev = data["project"]["optional-dependencies"]["dev"]
    assert all("reportlab" not in dep for dep in runtime)
    assert any(dep.startswith("reportlab") for dep in dev)
