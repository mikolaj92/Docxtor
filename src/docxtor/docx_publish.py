from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from lxml import etree

from .docx_package import (
    PackageEntry,
    PackageError,
    parse_package_xml,
    read_package_entries,
    write_package_atomically,
)


class PublishError(PackageError):
    """A DOCX could not be validated and atomically published."""


@dataclass(frozen=True)
class PublishReceipt:
    destination: Path
    sha256: str
    size: int
    preserved_parts: tuple[str, ...]
    normalized_parts: tuple[str, ...]
    validators_run: int


def publish_docx(
    rendered: bytes,
    destination: str | Path,
    *,
    source: bytes | None = None,
    validators: Iterable[Callable[[Path], None]] = (),
) -> PublishReceipt:
    """Preserve, normalize, validate, then atomically replace ``destination``.

    All transformation and validation happens before the final ``Path.replace`` in
    :func:`write_package_atomically`; a failure therefore leaves an existing target
    byte-for-byte unchanged.
    """
    rendered_entries = read_package_entries(rendered)
    source_by_name = (
        {entry.name: entry for entry in read_package_entries(source, validate_xml=False)}
        if source is not None
        else {}
    )
    preserved: list[str] = []
    final_entries: list[PackageEntry] = []
    for entry in rendered_entries:
        original = source_by_name.get(entry.name)
        data = entry.data
        if (
            original is not None
            and original.data != data
            and _same_xml_meaning(original.data, data)
        ):
            data = original.data
            preserved.append(entry.name)
        final_entries.append(
            PackageEntry(
                name=entry.name,
                data=data,
                compress_type=entry.compress_type,
                external_attr=entry.external_attr,
                internal_attr=entry.internal_attr,
                create_system=entry.create_system,
            )
        )

    validator_list = tuple(validators)

    def validate(path: Path) -> None:
        read_package_entries(path)
        from .docx_review_inventory import inventory_review_markup
        from .docx_review_models import ReviewCoverage

        review_inventory = inventory_review_markup(path.read_bytes())
        if review_inventory.coverage is ReviewCoverage.INCOMPLETE:
            diagnostics = ", ".join(item.code for item in review_inventory.diagnostics)
            raise PublishError(f"review markup validation failed: {diagnostics}")
        for validator in validator_list:
            validator(path)

    target = Path(destination)
    try:
        write_package_atomically(target, final_entries, validate=validate)
    except (OSError, PackageError, ValueError) as exc:
        if isinstance(exc, PublishError):
            raise
        raise PublishError(f"DOCX publication failed: {exc}") from exc
    payload = target.read_bytes()
    return PublishReceipt(
        destination=target,
        sha256=sha256(payload).hexdigest(),
        size=len(payload),
        preserved_parts=tuple(sorted(preserved)),
        normalized_parts=tuple(entry.name for entry in final_entries),
        validators_run=len(validator_list),
    )


def _same_xml_meaning(left: bytes, right: bytes) -> bool:
    try:
        left_root = parse_package_xml(left)
        right_root = parse_package_xml(right)
        return etree.tostring(left_root, method="c14n2", with_comments=True) == etree.tostring(
            right_root, method="c14n2", with_comments=True
        )
    except (PackageError, etree.C14NError):
        return False
