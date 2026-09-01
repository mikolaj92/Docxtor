"""Typed, domain-blind admission facts for a prospective DOCX package."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_inventory import inventory_docx

_OLE_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class DocumentPackageKind(StrEnum):
    VALID = "valid"
    ENCRYPTED = "encrypted"
    INVALID = "invalid"


@dataclass(frozen=True)
class DocxAdmissionInspection:
    package_kind: DocumentPackageKind
    has_main_document: bool
    has_macros: bool
    page_count: int | None
    diagnostics: tuple[str, ...] = ()


def inspect_docx_admission(content: bytes) -> DocxAdmissionInspection:
    if content.startswith(_OLE_CFB_MAGIC):
        return DocxAdmissionInspection(
            DocumentPackageKind.ENCRYPTED, False, False, None, ("ole_encrypted_package",)
        )
    inventory = inventory_docx(content)
    names = {part.name.replace("\\", "/").lstrip("/").lower() for part in inventory.parts}
    encrypted = bool({"encryptedpackage", "encryptioninfo"} & names)
    has_main = "word/document.xml" in names
    macros = any(name == "vbaproject.bin" or name.endswith("/vbaproject.bin") for name in names)
    kind = (
        DocumentPackageKind.ENCRYPTED
        if encrypted
        else DocumentPackageKind.VALID
        if has_main
        else DocumentPackageKind.INVALID
    )
    page = next(
        (
            item.value
            for item in inventory.surfaces
            if item.part_name == "docProps/app.xml"
            and (item.xml_name or "").casefold() == "pages"
            and item.attribute_qname is None
        ),
        None,
    )
    try:
        parsed = int(page) if page is not None else 0
    except ValueError:
        parsed = 0
    pages = parsed if parsed > 0 else None
    diagnostics = tuple(inventory.unreadable_parts + inventory.unknown_parts)
    return DocxAdmissionInspection(kind, has_main, macros, pages, diagnostics)
