from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo

import pytest

from docxtor import (
    PackageEntry,
    PackageError,
    normalize_docx_timestamps,
    parse_package_xml,
    read_package_entries,
    restore_semantically_unchanged_xml_parts,
    write_package_atomically,
)


def _write(path: Path, entries: list[tuple[str, bytes]], *, compress: int = ZIP_DEFLATED) -> None:
    with ZipFile(path, "w", compress) as bundle:
        for name, data in entries:
            bundle.writestr(name, data)


def test_reads_entries_and_preserves_zip_attributes(tmp_path: Path) -> None:
    path = tmp_path / "package.docx"
    with ZipFile(path, "w") as bundle:
        info = ZipInfo("word/document.xml")
        info.compress_type = ZIP_STORED
        info.external_attr = 0o600 << 16
        bundle.writestr(info, b"<document/>")

    entries = read_package_entries(path)

    assert len(entries) == 1
    assert entries[0].name == "word/document.xml"
    assert entries[0].data == b"<document/>"
    assert entries[0].compress_type == ZIP_STORED
    assert entries[0].external_attr == 0o600 << 16


@pytest.mark.parametrize("name", ["../evil.xml", "/root.xml", "word\\evil.xml", "word/%2e.xml"])
def test_rejects_invalid_member_paths(tmp_path: Path, name: str) -> None:
    path = tmp_path / "invalid.docx"
    _write(path, [(name, b"<x/>")])

    with pytest.raises(PackageError, match="invalid package member path"):
        read_package_entries(path)


def test_rejects_case_insensitive_duplicate_names(tmp_path: Path) -> None:
    path = tmp_path / "duplicate.docx"
    _write(path, [("word/a.xml", b"<a/>"), ("WORD/A.XML", b"<a/>")])

    with pytest.raises(PackageError, match="duplicate"):
        read_package_entries(path)


def test_rejects_malformed_xml_and_doctype(tmp_path: Path) -> None:
    malformed = tmp_path / "malformed.docx"
    _write(malformed, [("word/document.xml", b"<document>")])
    with pytest.raises(PackageError, match="invalid"):
        read_package_entries(malformed)

    doctype = tmp_path / "doctype.docx"
    _write(doctype, [("word/document.xml", b'<!DOCTYPE x [<!ENTITY y "z">]><x>&y;</x>')])
    with pytest.raises(PackageError, match="DOCTYPE"):
        read_package_entries(doctype)


def test_atomic_writer_does_not_publish_when_validator_fails(tmp_path: Path) -> None:
    path = tmp_path / "output.docx"
    path.write_bytes(b"original")

    def reject(_path: Path) -> None:
        raise RuntimeError("reject output")

    with pytest.raises(RuntimeError, match="reject output"):
        write_package_atomically(
            path,
            [PackageEntry("word/document.xml", b"<document/>")],
            validate=reject,
        )

    assert path.read_bytes() == b"original"
    assert not list(tmp_path.glob(".output.docx.*.tmp"))


def test_normalization_is_deterministic_and_idempotent(tmp_path: Path) -> None:
    path = tmp_path / "normalized.docx"
    _write(
        path,
        [("word/document.xml", b"<document/>"), ("word/media/image.bin", b"image")],
    )

    normalize_docx_timestamps(path)
    first = path.read_bytes()
    normalize_docx_timestamps(path)

    assert path.read_bytes() == first
    with ZipFile(path) as bundle:
        assert {info.date_time for info in bundle.infolist()} == {(1980, 1, 1, 0, 0, 0)}


def test_restores_semantically_unchanged_xml_bytes(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    rendered = tmp_path / "rendered.docx"
    original = b'<?xml version="1.0"?><root><value a="1">x</value></root>'
    rewritten = b'<root><value a="1">x</value></root>'
    _write(source, [("word/document.xml", original)])
    _write(rendered, [("word/document.xml", rewritten)])

    restore_semantically_unchanged_xml_parts(source, rendered)

    assert read_package_entries(rendered)[0].data == original


def test_parse_package_xml_accepts_safe_xml() -> None:
    root = parse_package_xml(b"<root><child/></root>")
    assert root.tag == "root"
