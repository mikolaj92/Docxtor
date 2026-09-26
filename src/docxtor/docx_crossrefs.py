"""Policy-free diagnostics for a bounded set of internal DOCX references.

Implemented forms: w:hyperlink/@w:anchor and REF/PAGEREF/NOTEREF field
instructions resolve against w:bookmarkStart/@w:name. Bookmark start/end IDs
are also checked for one-to-one pairing. Relationship-backed hyperlinks and
other field instructions are outside this pass; unsupported forms are surfaced
as incomplete coverage, not assumed valid.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import TypeAlias

from lxml import etree

from .docx_package import PackageError, parse_package_xml, read_package_entries

Source: TypeAlias = str | Path | bytes
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
PROFILE_ID = "docxtor-cross-references-v1"
_FIELD_REF = re.compile(r"^\s*(REF|PAGEREF|NOTEREF)\s+([^\s\\]+)", re.IGNORECASE)


class CrossReferenceSeverity(StrEnum):
    ERROR = "error"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class CrossReferenceDiagnostic:
    code: str
    severity: CrossReferenceSeverity
    part_name: str
    locator: str | None
    reference_kind: str
    referenced_value: str | None
    message: str = ""


@dataclass(frozen=True)
class CrossReferenceReport:
    profile_id: str
    diagnostics: tuple[CrossReferenceDiagnostic, ...]
    supported_forms: tuple[str, ...]
    unsupported_forms: tuple[str, ...]

    @property
    def complete(self) -> bool:
        return not self.unsupported_forms and not any(
            item.severity is CrossReferenceSeverity.UNKNOWN for item in self.diagnostics
        )

    @property
    def valid(self) -> bool:
        return not any(item.severity is CrossReferenceSeverity.ERROR for item in self.diagnostics)


def diagnose_docx_cross_references(source: Source) -> CrossReferenceReport:
    """Check internal bookmark references without making consumer-policy claims."""
    entries = read_package_entries(source, validate_xml=False)
    roots: dict[str, etree._Element] = {}
    diagnostics: list[CrossReferenceDiagnostic] = []
    unsupported: set[str] = set()
    for entry in entries:
        if not entry.name.startswith("word/") or not entry.name.endswith(".xml"):
            continue
        try:
            root = parse_package_xml(entry.data, part_name=entry.name)
        except PackageError:
            diagnostics.append(
                _diagnostic(
                    "crossref_xml_unreadable",
                    CrossReferenceSeverity.UNKNOWN,
                    entry.name,
                    None,
                    "xml",
                    None,
                )
            )
            unsupported.add("unreadable WordprocessingML part")
            continue
        roots[entry.name] = root
    for part, root in sorted(roots.items()):
        tree = root.getroottree()
        starts: dict[str, list[tuple[str, str | None]]] = {}
        ends: dict[str, list[str]] = {}
        bookmark_names: set[str] = set()
        for node in root.iter():
            if node.tag == W + "bookmarkStart":
                bid = node.get(W + "id")
                name = node.get(W + "name")
                locator = tree.getpath(node)
                if bid is None or name is None:
                    diagnostics.append(
                        _diagnostic(
                            "bookmark_start_incomplete",
                            CrossReferenceSeverity.ERROR,
                            part,
                            locator,
                            "bookmark",
                            bid or name,
                        )
                    )
                else:
                    starts.setdefault(bid, []).append((locator, name))
                    bookmark_names.add(name)
            elif node.tag == W + "bookmarkEnd":
                bid = node.get(W + "id")
                if bid is None:
                    diagnostics.append(
                        _diagnostic(
                            "bookmark_end_incomplete",
                            CrossReferenceSeverity.ERROR,
                            part,
                            tree.getpath(node),
                            "bookmark",
                            None,
                        )
                    )
                else:
                    ends.setdefault(bid, []).append(tree.getpath(node))
        for bid in sorted(starts.keys() | ends.keys()):
            if len(starts.get(bid, ())) != 1 or len(ends.get(bid, ())) != 1:
                for locator, _ in starts.get(bid, ()):
                    diagnostics.append(
                        _diagnostic(
                            "bookmark_pair_mismatch",
                            CrossReferenceSeverity.ERROR,
                            part,
                            locator,
                            "bookmark_id",
                            bid,
                        )
                    )
                for locator in ends.get(bid, ()):
                    diagnostics.append(
                        _diagnostic(
                            "bookmark_pair_mismatch",
                            CrossReferenceSeverity.ERROR,
                            part,
                            locator,
                            "bookmark_id",
                            bid,
                        )
                    )
        # Preserve field instructions split across multiple w:instrText runs.
        _check_complex_fields(root, part, tree, bookmark_names, diagnostics, unsupported)
        for node in root.iter():
            locator = tree.getpath(node)
            if node.tag == W + "hyperlink":
                anchor = node.get(W + "anchor")
                rel_id = node.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                )
                if anchor and anchor not in bookmark_names:
                    diagnostics.append(
                        _diagnostic(
                            "hyperlink_bookmark_missing",
                            CrossReferenceSeverity.ERROR,
                            part,
                            locator,
                            "bookmark_anchor",
                            anchor,
                        )
                    )
                if rel_id:
                    unsupported.add("relationship-backed hyperlink (w:hyperlink/@r:id)")
                    diagnostics.append(
                        _diagnostic(
                            "crossref_form_unsupported",
                            CrossReferenceSeverity.UNKNOWN,
                            part,
                            locator,
                            "relationship_hyperlink",
                            rel_id,
                        )
                    )
            if node.tag == W + "fldSimple":
                instruction = node.get(W + "instr") or ""
                _check_field(instruction, part, locator, bookmark_names, diagnostics, unsupported)
            if node.tag == W + "instrText":
                parent = node.getparent()
                if parent is None or parent.tag != W + "r":
                    unsupported.add("field instruction outside a simple w:r")
                    diagnostics.append(
                        _diagnostic(
                            "crossref_form_unsupported",
                            CrossReferenceSeverity.UNKNOWN,
                            part,
                            locator,
                            "field_instruction",
                            "instrText",
                        )
                    )
            if node.tag == W + "dataBinding":
                unsupported.add("content-control data binding (w:dataBinding)")
                diagnostics.append(
                    _diagnostic(
                        "crossref_form_unsupported",
                        CrossReferenceSeverity.UNKNOWN,
                        part,
                        locator,
                        "content_control_binding",
                        node.get(W + "storeItemID"),
                    )
                )
    ordered = tuple(
        sorted(
            set(diagnostics),
            key=lambda d: (d.part_name, d.locator or "", d.code, d.referenced_value or ""),
        )
    )
    return CrossReferenceReport(
        PROFILE_ID,
        ordered,
        ("bookmark start/end IDs", "w:hyperlink/@w:anchor", "REF/PAGEREF/NOTEREF field target"),
        tuple(sorted(unsupported)),
    )


def _check_complex_fields(
    root: etree._Element,
    part: str,
    tree: etree._ElementTree,
    bookmarks: set[str],
    diagnostics: list[CrossReferenceDiagnostic],
    unsupported: set[str],
) -> None:
    """Collect ECMA field-char sequences and inspect complete instruction text."""
    stack: list[tuple[str, list[str], str]] = []
    for node in root.iter():
        if node.tag == W + "fldChar":
            kind = (node.get(W + "fldCharType") or "").lower()
            if kind == "begin":
                stack.append((tree.getpath(node), [], ""))
            elif kind == "separate" and stack:
                locator, text, _ = stack[-1]
                stack[-1] = (locator, text, "separated")
            elif kind == "end" and stack:
                locator, text, _ = stack.pop()
                _check_field("".join(text), part, locator, bookmarks, diagnostics, unsupported)
        elif node.tag == W + "instrText" and stack:
            # Only the innermost open field directly owns this instruction run.
            stack[-1][1].append(node.text or "")
    for locator, text, _ in stack:
        unsupported.add("unterminated complex field sequence")
        diagnostics.append(
            _diagnostic(
                "complex_field_unterminated",
                CrossReferenceSeverity.UNKNOWN,
                part,
                locator,
                "field_instruction",
                "".join(text) or None,
            )
        )


def _check_field(
    instruction: str,
    part: str,
    locator: str,
    bookmarks: set[str],
    diagnostics: list[CrossReferenceDiagnostic],
    unsupported: set[str],
) -> None:
    value = instruction.strip()
    match = _FIELD_REF.match(value)
    if match:
        kind, target = match.groups()
        if target not in bookmarks:
            diagnostics.append(
                _diagnostic(
                    "field_bookmark_missing",
                    CrossReferenceSeverity.ERROR,
                    part,
                    locator,
                    kind.upper(),
                    target,
                )
            )
    elif value and not value.startswith("="):
        # Many field instructions are not reference-bearing; expose them without
        # guessing semantics or declaring coverage complete.
        if value.split(maxsplit=1)[0].upper() in {"REF", "PAGEREF", "NOTEREF"}:
            diagnostics.append(
                _diagnostic(
                    "field_reference_unparsed",
                    CrossReferenceSeverity.UNKNOWN,
                    part,
                    locator,
                    "field_instruction",
                    value,
                )
            )
            unsupported.add("malformed REF/PAGEREF/NOTEREF instruction")


def _diagnostic(
    code: str,
    severity: CrossReferenceSeverity,
    part: str,
    locator: str | None,
    kind: str,
    value: str | None,
) -> CrossReferenceDiagnostic:
    return CrossReferenceDiagnostic(code, severity, part, locator, kind, value)
