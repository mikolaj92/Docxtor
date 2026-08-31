from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile, ZipInfo

from lxml import etree

from .common import DocumentError

MAX_PACKAGE_ENTRIES = 4096
MAX_ENTRY_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
MAX_TOTAL_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_COMPRESSION_RATIO = 1000
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)
_XML_BOMS = (b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff", b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")
_XML_PROBE_BYTES = 64 * 1024
_XML_ENCODINGS = (
    "utf-8-sig",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "utf-32",
    "utf-32-le",
    "utf-32-be",
)
_OPC_PART_NAME_CHARS = frozenset("!$&'()*+,-.:;=@_~")
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")
_ASCII_IUNRESERVED = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")


@dataclass(frozen=True)
class PackageLimits:
    max_entries: int = MAX_PACKAGE_ENTRIES
    max_entry_uncompressed_bytes: int = MAX_ENTRY_UNCOMPRESSED_BYTES
    max_total_uncompressed_bytes: int = MAX_TOTAL_UNCOMPRESSED_BYTES
    max_compression_ratio: int = MAX_COMPRESSION_RATIO


DEFAULT_PACKAGE_LIMITS = PackageLimits()


class PackageError(DocumentError):
    """A DOCX package could not be read or written without ambiguity."""


@dataclass(frozen=True)
class PackageEntry:
    name: str
    data: bytes
    compress_type: int = ZIP_DEFLATED
    external_attr: int = 0
    internal_attr: int = 0
    create_system: int = 3

    @classmethod
    def from_zip(cls, info: ZipInfo, data: bytes) -> PackageEntry:
        return cls(
            name=info.filename,
            data=data,
            compress_type=info.compress_type,
            external_attr=info.external_attr,
            internal_attr=info.internal_attr,
            create_system=info.create_system,
        )

    def zip_info(self, *, deterministic: bool = True) -> ZipInfo:
        date_time = _ZIP_EPOCH if deterministic else (1980, 1, 1, 0, 0, 0)
        info = ZipInfo(self.name, date_time=date_time)
        info.compress_type = self.compress_type
        info.external_attr = self.external_attr
        info.internal_attr = self.internal_attr
        info.create_system = self.create_system
        return info


def read_package_entries(
    source: str | Path | bytes,
    *,
    limits: PackageLimits = DEFAULT_PACKAGE_LIMITS,
    validate_xml: bool = True,
) -> tuple[PackageEntry, ...]:
    """Read every DOCX entry with bounds and optional source-only XML validation."""
    opener: str | Path | BytesIO
    opener = BytesIO(source) if isinstance(source, bytes) else source
    try:
        archive = ZipFile(opener)
    except (OSError, BadZipFile) as exc:
        raise PackageError(f"cannot open DOCX package: {exc}") from exc
    with archive:
        infos = archive.infolist()
        _validate_infos(infos, limits)
        entries: list[PackageEntry] = []
        for info in infos:
            try:
                data = archive.read(info)
            except (OSError, BadZipFile, RuntimeError) as exc:
                raise PackageError(f"cannot read DOCX entry {info.filename}: {exc}") from exc
            if validate_xml and _needs_xml_validation(info.filename, data):
                parse_package_xml(data, part_name=info.filename)
            entries.append(PackageEntry.from_zip(info, data))
        return tuple(entries)


def write_package_atomically(
    destination: str | Path,
    entries: Iterable[PackageEntry],
    *,
    validate: Callable[[Path], None] | None = None,
    limits: PackageLimits = DEFAULT_PACKAGE_LIMITS,
) -> None:
    """Write a deterministic DOCX beside destination and publish it atomically."""
    target = Path(destination)
    materialized = tuple(entries)
    _validate_entry_records(materialized, limits)
    target.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        dir=target.parent,
        prefix=f".{target.name}.",
        suffix=".tmp",
        delete=False,
    ) as handle:
        temporary = Path(handle.name)
    try:
        _write_package(temporary, materialized)
        # Re-read with the same fail-closed contract before caller validation/publication.
        read_package_entries(temporary, limits=limits)
        if validate is not None:
            validate(temporary)
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)


def normalize_docx_timestamps(path: str | Path) -> None:
    """Atomically stamp every ZIP entry with the deterministic ZIP epoch."""
    entries = read_package_entries(path)
    write_package_atomically(path, entries)


def restore_semantically_unchanged_xml_parts(
    source_path: str | Path,
    rendered_path: str | Path,
) -> None:
    """Restore original bytes when rendered XML has identical canonical meaning."""
    original = {
        entry.name: entry.data
        for entry in read_package_entries(source_path, validate_xml=False)
        if entry.name.endswith((".xml", ".rels"))
    }
    rendered = read_package_entries(rendered_path)
    changed = False
    restored: list[PackageEntry] = []
    for entry in rendered:
        source_data = original.get(entry.name)
        data = entry.data
        if source_data is not None and data != source_data:
            canonical_source = _canonical_xml(source_data)
            if canonical_source is not None and canonical_source == _canonical_xml(data):
                data = source_data
                changed = True
        restored.append(
            PackageEntry(
                name=entry.name,
                data=data,
                compress_type=entry.compress_type,
                external_attr=entry.external_attr,
                internal_attr=entry.internal_attr,
                create_system=entry.create_system,
            )
        )
    if changed:
        write_package_atomically(rendered_path, restored)


def parse_package_xml(data: bytes, *, part_name: str = "<xml>") -> etree._Element:
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        recover=False,
    )
    try:
        root = etree.fromstring(data, parser=parser)
    except (etree.XMLSyntaxError, ValueError) as exc:
        raise PackageError(f"DOCX XML part {part_name} is invalid: {exc}") from exc
    if root.getroottree().docinfo.doctype:
        raise PackageError(f"DOCX XML part {part_name} must not contain a DOCTYPE")
    return root


def _validate_infos(infos: list[ZipInfo], limits: PackageLimits) -> None:
    if len(infos) > limits.max_entries:
        raise PackageError(f"DOCX has {len(infos)} entries; limit is {limits.max_entries}")
    names = [info.filename for info in infos]
    for name in names:
        if not _is_valid_package_member_name(name):
            raise PackageError(f"DOCX contains invalid package member path: {name}")
    equivalent = [name.casefold() for name in names]
    if len(equivalent) != len(set(equivalent)):
        raise PackageError("DOCX contains duplicate package member names")
    total = sum(info.file_size for info in infos)
    if total > limits.max_total_uncompressed_bytes:
        raise PackageError(
            f"DOCX uncompressed size {total} exceeds {limits.max_total_uncompressed_bytes}"
        )
    for info in infos:
        if info.file_size > limits.max_entry_uncompressed_bytes:
            raise PackageError(f"DOCX entry {info.filename} uncompressed size exceeds limit")
        ratio = info.file_size / max(info.compress_size, 1)
        if ratio > limits.max_compression_ratio:
            raise PackageError(f"DOCX entry {info.filename} compression ratio exceeds limit")


def _validate_entry_records(entries: tuple[PackageEntry, ...], limits: PackageLimits) -> None:
    infos = []
    for entry in entries:
        info = entry.zip_info()
        info.file_size = len(entry.data)
        info.compress_size = max(len(entry.data), 1)
        infos.append(info)
        if _needs_xml_validation(entry.name, entry.data):
            parse_package_xml(entry.data, part_name=entry.name)
    _validate_infos(infos, limits)


def _write_package(path: Path, entries: tuple[PackageEntry, ...]) -> None:
    with ZipFile(path, "w") as output:
        for entry in entries:
            output.writestr(entry.zip_info(), entry.data)


def _canonical_xml(data: bytes) -> bytes | None:
    try:
        root = parse_package_xml(data)
        return etree.tostring(root, method="c14n2", with_comments=True)
    except (PackageError, etree.C14NError):
        return None


def _is_valid_package_member_name(name: str) -> bool:
    if name == "[Content_Types].xml":
        return True
    normalized = name.removeprefix("/")
    return bool(
        normalized
        and not name.startswith(("/", "\\"))
        and "\\" not in name
        and all(_is_valid_package_segment(segment) for segment in normalized.split("/"))
    )


def _is_valid_package_segment(segment: str) -> bool:
    if not segment or segment in {".", ".."} or segment.endswith("."):
        return False
    offset = 0
    while offset < len(segment):
        character = segment[offset]
        if character == "%":
            if (
                offset + 2 >= len(segment)
                or segment[offset + 1] not in _HEX_DIGITS
                or segment[offset + 2] not in _HEX_DIGITS
            ):
                return False
            decoded = chr(int(segment[offset + 1 : offset + 3], 16))
            if decoded in "/\\" or decoded in _ASCII_IUNRESERVED:
                return False
            offset += 3
            continue
        if not (
            ord(character) < 128 and (character.isalnum() or character in _OPC_PART_NAME_CHARS)
        ):
            return False
        offset += 1
    return True


def _needs_xml_validation(name: str, data: bytes) -> bool:
    normalized = name.removeprefix("/").casefold()
    if normalized.startswith("customxml/") or normalized.endswith((".xml", ".rels")):
        return True
    candidate = data[:_XML_PROBE_BYTES].lstrip(b" \t\r\n")
    if candidate.startswith(b"<") or any(candidate.startswith(bom) for bom in _XML_BOMS):
        return True
    return any(
        data[:_XML_PROBE_BYTES].decode(encoding, errors="ignore").lstrip().startswith("<")
        for encoding in _XML_ENCODINGS
    )
