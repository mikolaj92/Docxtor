"""Bounded, versioned checks for required DOCX package relationships."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

from .docx_package import PackageError, parse_package_xml, read_package_entries

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
PROFILE_ID = "docxtor-docx-relationships-v1"


class RelationshipSeverity(StrEnum):
    ERROR = "error"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class RelationshipDiagnostic:
    code: str
    severity: RelationshipSeverity
    constraint_id: str
    source_part: str
    relationship_type: str | None = None
    relationship_id: str | None = None
    target: str | None = None
    message: str = ""


@dataclass(frozen=True)
class RelationshipValidation:
    profile_id: str
    diagnostics: tuple[RelationshipDiagnostic, ...]

    @property
    def complete(self) -> bool:
        return not any(d.severity is RelationshipSeverity.UNKNOWN for d in self.diagnostics)

    @property
    def valid(self) -> bool:
        return not any(d.severity is RelationshipSeverity.ERROR for d in self.diagnostics)


# Rel type -> source part, cardinality; only constraints asserted by this profile.
_SINGLETONS = {
    "officeDocument": ("", "exactly_one"),
    "core-properties": ("", "at_most_one"),
    "extended-properties": ("", "at_most_one"),
    "custom-properties": ("", "at_most_one"),
    "styles": ("word/document.xml", "at_most_one"),
    "numbering": ("word/document.xml", "at_most_one"),
    "settings": ("word/document.xml", "at_most_one"),
    "webSettings": ("word/document.xml", "at_most_one"),
    "fontTable": ("word/document.xml", "at_most_one"),
    "theme": ("word/document.xml", "at_most_one"),
    "footnotes": ("word/document.xml", "at_most_one"),
    "endnotes": ("word/document.xml", "at_most_one"),
}


def validate_docx_relationships(source: str | bytes) -> RelationshipValidation:
    """Diagnose constraints in the deliberately limited ``PROFILE_ID`` profile."""
    entries = {entry.name: entry.data for entry in read_package_entries(source, validate_xml=False)}
    diagnostics: list[RelationshipDiagnostic] = []
    rel_files = {name for name in entries if name.endswith(".rels")}
    parsed: list[tuple[str, str, str, str, str, bool]] = []
    for name in sorted(rel_files):
        source_part = _source_part(name)
        try:
            root = parse_package_xml(entries[name], part_name=name)
        except PackageError as exc:
            diagnostics.append(
                _diag(
                    "malformed_relationship_part", "relationship_xml", source_part, message=str(exc)
                )
            )
            continue
        if root.tag != f"{{{REL_NS}}}Relationships":
            diagnostics.append(
                _diag(
                    "invalid_relationship_root",
                    "relationship_root",
                    source_part,
                    message="unexpected root element",
                )
            )
            continue
        ids: set[str] = set()
        for element in root:
            if element.tag != f"{{{REL_NS}}}Relationship":
                diagnostics.append(
                    _diag(
                        "unknown_relationship_element",
                        "relationship_element",
                        source_part,
                        message=str(element.tag),
                        severity=RelationshipSeverity.UNKNOWN,
                    )
                )
                continue
            rid, rtype, target = element.get("Id"), element.get("Type"), element.get("Target")
            mode = element.get("TargetMode", "Internal")
            if not rid or not rtype or target is None:
                diagnostics.append(
                    _diag(
                        "incomplete_relationship",
                        "relationship_required_attributes",
                        source_part,
                        rid,
                        rtype,
                        target,
                    )
                )
                continue
            if rid in ids:
                diagnostics.append(
                    _diag(
                        "duplicate_relationship_id",
                        "relationship_id_unique_per_source",
                        source_part,
                        rid,
                        rtype,
                        target,
                    )
                )
            ids.add(rid)
            external = mode == "External"
            if mode not in {"External", "Internal"}:
                diagnostics.append(
                    _diag(
                        "unknown_target_mode",
                        "target_mode",
                        source_part,
                        rid,
                        rtype,
                        target,
                        RelationshipSeverity.UNKNOWN,
                    )
                )
            try:
                resolved_target = _resolve(source_part, target) if not external else ""
            except ValueError:
                diagnostics.append(
                    _diag(
                        "invalid_relationship_target",
                        "internal_target_is_package_relative",
                        source_part,
                        rid,
                        rtype,
                        target,
                    )
                )
                resolved_target = ""
            parsed.append((source_part, rid, rtype, target, resolved_target, external))
            short = rtype.removeprefix(OFFICE_REL)
            if (
                (rtype.startswith(OFFICE_REL) and short not in _SINGLETONS)
                or not rtype.startswith(OFFICE_REL)
            ) and short not in {
                "hyperlink",
                "image",
                "oleObject",
                "package",
                "chart",
                "slideLayout",
                "slideMaster",
                "notesMaster",
                "comments",
                "commentAuthors",
                "customXml",
                "aFChunk",
                "attachedTemplate",
                "control",
                "diagramData",
                "diagramLayout",
                "diagramColors",
                "diagramQuickStyle",
                "printerSettings",
                "vbaProject",
                "glossaryDocument",
                "subDocument",
                "themeOverride",
                "font",
            }:
                diagnostics.append(
                    _diag(
                        "unprofiled_relationship_type",
                        "relationship_type_coverage",
                        source_part,
                        rid,
                        rtype,
                        target,
                        RelationshipSeverity.UNKNOWN,
                    )
                )
    office_document = OFFICE_REL + "officeDocument"
    count = sum(r[2] == office_document and r[0] == "" for r in parsed)
    if count != 1:
        diagnostics.append(
            _diag(
                "required_relationship_cardinality",
                "root_office_document_exactly_one",
                "",
                rtype=office_document,
                message=f"expected exactly one, found {count}",
            )
        )
    for source_part, rid, rtype, target, resolved, external in parsed:
        short = rtype.removeprefix(OFFICE_REL)
        constraint = _SINGLETONS.get(short)
        if constraint:
            expected_source, cardinality = constraint
            if source_part != expected_source and (expected_source or source_part):
                diagnostics.append(
                    _diag(
                        "disallowed_relationship_source",
                        f"{PROFILE_ID}:{short}:source",
                        source_part,
                        rid,
                        rtype,
                        target,
                    )
                )
            if (
                cardinality == "at_most_one"
                and sum(x[0] == source_part and x[2] == rtype for x in parsed) > 1
            ):
                diagnostics.append(
                    _diag(
                        "relationship_cardinality_exceeded",
                        f"{PROFILE_ID}:{short}:at_most-one",
                        source_part,
                        rid,
                        rtype,
                        target,
                    )
                )
        if not external and resolved not in entries:
            diagnostics.append(
                _diag(
                    "missing_relationship_target",
                    "internal_target_exists",
                    source_part,
                    rid,
                    rtype,
                    target,
                )
            )
    if "word/document.xml" not in entries:
        diagnostics.append(
            _diag(
                "required_part_missing",
                "word_document_part_required",
                "",
                target="word/document.xml",
            )
        )
    # Stable ordering independent of ZIP member order.
    ordered = tuple(
        sorted(
            set(diagnostics),
            key=lambda d: (
                d.source_part,
                d.constraint_id,
                d.code,
                d.relationship_id or "",
                d.target or "",
            ),
        )
    )
    return RelationshipValidation(PROFILE_ID, ordered)


def _source_part(name: str) -> str:
    if name == "_rels/.rels":
        return ""
    from pathlib import PurePosixPath

    path = PurePosixPath(name)
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _resolve(source: str, target: str) -> str:
    from pathlib import PurePosixPath
    from posixpath import normpath

    path = normpath(f"{str(PurePosixPath(source).parent) if source else ''}/{target}".lstrip("/"))
    if path == ".." or path.startswith("../") or path.startswith("/"):
        raise ValueError("relationship target escapes package")
    return path


def _diag(
    code: str,
    constraint: str,
    source: str,
    rid: str | None = None,
    rtype: str | None = None,
    target: str | None = None,
    severity: RelationshipSeverity = RelationshipSeverity.ERROR,
    message: str = "",
) -> RelationshipDiagnostic:
    return RelationshipDiagnostic(
        code, severity, f"{PROFILE_ID}:{constraint}", source, rtype, rid, target, message
    )
