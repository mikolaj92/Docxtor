"""Bounded checks for OPC content-type declarations in DOCX packages."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import PurePosixPath

from lxml import etree

from .docx_package import PackageError, parse_package_xml, read_package_entries

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PROFILE_ID = "docxtor-opc-content-types-v1"


class ContentTypeSeverity(StrEnum):
    ERROR = "error"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class ContentTypeDiagnostic:
    code: str
    severity: ContentTypeSeverity
    constraint_id: str
    part_name: str | None = None
    declaration: str | None = None
    expected: str | None = None
    actual: str | None = None
    message: str = ""


@dataclass(frozen=True)
class ContentTypeValidation:
    profile_id: str
    diagnostics: tuple[ContentTypeDiagnostic, ...]

    @property
    def complete(self) -> bool:
        return not any(d.severity is ContentTypeSeverity.UNKNOWN for d in self.diagnostics)

    @property
    def valid(self) -> bool:
        return not any(d.severity is ContentTypeSeverity.ERROR for d in self.diagnostics)


_KNOWN_TYPES = {
    "[Content_Types].xml": "application/xml",
    "_rels/.rels": "application/vnd.openxmlformats-package.relationships+xml",
    "word/document.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    ),
    "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "word/numbering.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    ),
    "word/settings.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    ),
    "word/webSettings.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
    ),
    "word/fontTable.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    ),
    "word/footnotes.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    ),
    "word/endnotes.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
    ),
    "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
    "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    "docProps/custom.xml": "application/vnd.openxmlformats-officedocument.custom-properties+xml",
}
_KNOWN_DEFAULTS = {
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
}


def validate_docx_content_types(source: str | bytes) -> ContentTypeValidation:
    """Inspect declarations and effective types without changing or rejecting input."""
    entries = {entry.name: entry.data for entry in read_package_entries(source, validate_xml=False)}
    diagnostics: list[ContentTypeDiagnostic] = []
    payload = entries.get("[Content_Types].xml")
    defaults: dict[str, tuple[str, str]] = {}
    overrides: dict[str, tuple[str, str]] = {}
    if payload is None:
        diagnostics.append(
            _diag(
                "content_types_part_missing",
                "content_types_part_required",
                declaration="[Content_Types].xml",
            )
        )
    else:
        try:
            root = parse_package_xml(payload, part_name="[Content_Types].xml")
        except PackageError as exc:
            diagnostics.append(
                _diag(
                    "content_types_xml_malformed",
                    "content_types_xml_well_formed",
                    declaration="[Content_Types].xml",
                    message=str(exc),
                )
            )
            root = None
        if root is not None:
            if root.tag != f"{{{CONTENT_TYPES_NS}}}Types":
                diagnostics.append(
                    _diag(
                        "content_types_root_invalid",
                        "types_root",
                        declaration="[Content_Types].xml",
                    )
                )
            else:
                for node in root:
                    qname = etree.QName(node)
                    if qname.namespace != CONTENT_TYPES_NS:
                        diagnostics.append(
                            _diag(
                                "unknown_content_types_element",
                                "declaration_coverage",
                                declaration=qname.localname,
                                severity=ContentTypeSeverity.UNKNOWN,
                            )
                        )
                        continue
                    if qname.localname == "Default":
                        key, category = (node.get("Extension") or "").lower(), "Default"
                        table = defaults
                    elif qname.localname == "Override":
                        raw = node.get("PartName") or ""
                        key, category = raw.lstrip("/"), "Override"
                        table = overrides
                    else:
                        diagnostics.append(
                            _diag(
                                "unknown_content_types_element",
                                "declaration_coverage",
                                declaration=qname.localname,
                                severity=ContentTypeSeverity.UNKNOWN,
                            )
                        )
                        continue
                    value = node.get("ContentType") or ""
                    if not key or not value:
                        diagnostics.append(
                            _diag(
                                "content_type_declaration_incomplete",
                                "declaration_required_attributes",
                                declaration=f"{category}:{key}",
                            )
                        )
                        continue
                    if category == "Override" and (
                        not raw.startswith("/") or _unsafe_part_name(key)
                    ):
                        diagnostics.append(
                            _diag(
                                "content_type_override_part_name_invalid",
                                "override_part_name",
                                part_name=key,
                                declaration=f"Override:{raw}",
                            )
                        )
                    if key in table:
                        old_value, _ = table[key]
                        diagnostics.append(
                            _diag(
                                "content_type_declaration_duplicate"
                                if old_value == value
                                else "content_type_declaration_conflict",
                                f"{category.lower()}_unique_key",
                                part_name=key if category == "Override" else None,
                                declaration=f"{category}:{key}",
                                expected=old_value,
                                actual=value,
                            )
                        )
                    else:
                        table[key] = (value, f"{category}:{key}")
    for key, (value, declaration) in sorted(overrides.items()):
        if key not in entries:
            diagnostics.append(
                _diag(
                    "content_type_declaration_unresolved",
                    "override_targets_existing_part",
                    key,
                    declaration,
                    actual=value,
                )
            )
    for name in sorted(entries):
        if name == "[Content_Types].xml":
            continue
        extension = PurePosixPath(name).suffix.removeprefix(".").lower()
        override = overrides.get(name)
        default = defaults.get(extension)
        effective = override[0] if override else default[0] if default else None
        declaration = override[1] if override else default[1] if default else None
        if effective is None:
            diagnostics.append(
                _diag("content_type_declaration_missing", "part_has_effective_content_type", name)
            )
            if name not in _KNOWN_TYPES:
                diagnostics.append(
                    _diag(
                        "unknown_part_content_type",
                        "part_type_coverage",
                        name,
                        severity=ContentTypeSeverity.UNKNOWN,
                    )
                )
            continue
        expected = _KNOWN_TYPES.get(name)
        if expected is None and extension not in _KNOWN_DEFAULTS and not name.endswith(".rels"):
            diagnostics.append(
                _diag(
                    "unknown_part_content_type",
                    "part_type_coverage",
                    name,
                    declaration=declaration,
                    actual=effective,
                    severity=ContentTypeSeverity.UNKNOWN,
                )
            )
        elif expected is not None and effective != expected:
            diagnostics.append(
                _diag(
                    "known_part_content_type_mismatch",
                    "known_part_expected_type",
                    name,
                    declaration,
                    expected,
                    effective,
                )
            )
    ordered = tuple(
        sorted(
            set(diagnostics),
            key=lambda d: (d.part_name or "", d.constraint_id, d.code, d.declaration or ""),
        )
    )
    return ContentTypeValidation(PROFILE_ID, ordered)


def _unsafe_part_name(name: str) -> bool:
    return (
        not name or "\\" in name or any(segment in {"", ".", ".."} for segment in name.split("/"))
    )


def _diag(
    code: str,
    constraint: str,
    part_name: str | None = None,
    declaration: str | None = None,
    expected: str | None = None,
    actual: str | None = None,
    severity: ContentTypeSeverity = ContentTypeSeverity.ERROR,
    message: str = "",
) -> ContentTypeDiagnostic:
    return ContentTypeDiagnostic(
        code,
        severity,
        f"{PROFILE_ID}:{constraint}",
        part_name,
        declaration,
        expected,
        actual,
        message,
    )
