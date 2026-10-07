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
    assert tuple(int(part) for part in __version__.split(".")) >= (0, 4, 4)


def test_readme_documents_v041_pin_mismatch() -> None:
    """Consumers must be told not to pin the lying v0.4.1 tag."""
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "v0.4.1" in readme
    assert "0.4.0" in readme
    assert "v0.4.4" in readme


def test_readme_install_pin_matches_pyproject_version() -> None:
    """The README install pins must track pyproject (the v0.4.1 failure mode)."""
    version = _pyproject_version()
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"@v{version}" in readme, (
        f"README install pin does not match pyproject version {version}"
    )


def test_readme_license_matches_pyproject() -> None:
    """The README License section must state the pyproject license text."""
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    license_payload = data["project"]["license"]
    license_text = (
        license_payload["text"] if isinstance(license_payload, dict) else license_payload
    )
    if not isinstance(license_text, str):
        raise TypeError(f"unsupported pyproject license payload: {license_payload!r}")
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"## License\n\n{license_text}" in readme, (
        f"README License section does not state pyproject license {license_text!r}"
    )


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
