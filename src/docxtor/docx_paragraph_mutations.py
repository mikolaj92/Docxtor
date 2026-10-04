from __future__ import annotations

from collections import Counter
from typing import Any

from docx.oxml.ns import qn

from .common import DocumentError
from .docx_package import (
    _is_valid_package_member_name,
    _needs_xml_validation,
    parse_package_xml,
    read_package_entries,
)

_COMMENT_MARKERS = tuple(
    qn(f"w:{name}") for name in ("commentRangeStart", "commentRangeEnd", "commentReference")
)


def _is_comment_annotation(element: Any) -> bool:
    if element.tag in _COMMENT_MARKERS[:2]:
        return True
    return (
        element.tag == qn("w:r")
        and len(element) == 1
        and element[0].tag == _COMMENT_MARKERS[2]
        and not element.attrib
        and not (element.text or "").strip()
    )


def _require_local_comment_ranges(
    document: Any, paragraph: Any, *, source_bytes: bytes | None
) -> None:
    markers = [node for node in paragraph._p.iter() if node.tag in _COMMENT_MARKERS]
    if not markers:
        return
    ordered: dict[str, list[str]] = {}
    for marker in markers:
        comment_id = marker.get(qn("w:id"))
        if (
            comment_id is None
            or not comment_id.isdecimal()
            or set(marker.attrib) != {qn("w:id")}
            or len(marker)
            or (marker.text or "").strip()
            or (marker.tail or "").strip()
        ):
            raise DocumentError("inserted paragraph carries malformed comment annotations")
        ordered.setdefault(comment_id, []).append(marker.tag)
    if any(tags != list(_COMMENT_MARKERS) for tags in ordered.values()):
        raise DocumentError("inserted paragraph comment range is not balanced and local")
    local = Counter((node.get(qn("w:id")), node.tag) for node in markers)
    global_markers: Counter[tuple[str | None, str]] = Counter()
    # Facts import the public document class; defer these mechanical helpers
    # until the handle is fully initialized to avoid the module import cycle.
    from .docx_facts import _content_type_for, _content_type_map, _content_types
    from .docx_inventory import _is_xml_part

    # The serializer omits unlinked source parts. Retain those for validation,
    # including their source-only OPC type declarations. Live payload and type
    # metadata replace stale source versions of the same logical part.
    packages = []
    if source_bytes is not None:
        packages.append(read_package_entries(source_bytes))
    packages.append(read_package_entries(document.to_bytes()))
    entries = {}
    declared_xml = set()
    for package in packages:
        content_types = next(
            (entry.data for entry in package if entry.name == "[Content_Types].xml"), None
        )
        if content_types is None:
            raise DocumentError("DOCX package has no [Content_Types].xml")
        facts = _content_types(content_types)
        keys = set()
        package_names = {entry.name.casefold() for entry in package}
        for fact in facts:
            # Supported member names are ASCII; defaults use the owner's lower()
            # extension matching. Keep Override and Default namespaces separate.
            key = (fact.is_default, fact.key.lower())
            if key in keys:
                raise DocumentError("DOCX package has ambiguous content type declarations")
            keys.add(key)
            if not fact.is_default and (
                not _is_valid_package_member_name(fact.key)
                or fact.key.casefold() not in package_names
            ):
                raise DocumentError("DOCX content type Override does not resolve to a valid part")
        defaults, overrides = _content_type_map(facts)
        overrides = {name.casefold(): value for name, value in overrides.items()}
        for entry in package:
            name = entry.name.casefold()
            entries[name] = entry
            content_type = _content_type_for(name, defaults, overrides)
            if _is_xml_part(content_type.partition(";")[0].strip().casefold(), b""):
                declared_xml.add(name)
            else:
                declared_xml.discard(name)
    for name, entry in entries.items():
        if name not in declared_xml and not _needs_xml_validation(entry.name, entry.data):
            continue
        root = parse_package_xml(entry.data, part_name=entry.name)
        global_markers.update(
            (node.get(qn("w:id")), node.tag)
            for node in root.iter()
            if node.tag in _COMMENT_MARKERS and node.get(qn("w:id")) in ordered
        )
    if global_markers != local:
        raise DocumentError("inserted paragraph comment annotations extend outside the paragraph")


class DocxParagraphMutationOperations:
    """Physical paragraph mutations shared by the public DocxDocument handle.

    Reuse the handle's canonical physical projections and story indexing.
    This private mixin adds no separate public import surface.
    """

    filename: str
    _doc: Any
    _source_bytes: bytes | None

    def remove_inserted_paragraph(self, container_id: str, *, expected_text: str) -> None:
        """Remove an entirely inserted paragraph after exact physical text validation.

        Balanced paragraph-local comment annotations may accompany insertion wrappers.
        Mixed source content and other opaque payload outside wrappers fail closed.
        Reindex surviving paragraphs, spans and stories after removal.
        """
        paragraph = self.resolve_paragraph(container_id)
        if paragraph is None:
            raise DocumentError("inserted paragraph locator does not resolve")
        parent = paragraph._p.getparent()
        if parent is not self._doc.element.body:
            raise DocumentError("inserted paragraph is not attached to a supported body story")
        if paragraph._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            raise DocumentError("inserted paragraph carries protected section properties")
        current = self._from_pydocx(self._doc, filename=self.filename)
        indexed = current.resolve_paragraph(container_id)
        if indexed is None or indexed._p is not paragraph._p:
            raise DocumentError(
                "inserted paragraph locator no longer identifies the same paragraph"
            )
        spans = [span for span in current.spans if span.container_id == container_id]
        inserted = "".join(span.text for span in spans if span.role == "insertion")
        segments = current.get_inline_segments(container_id)
        if (
            not inserted
            or inserted != expected_text
            or any(span.role != "insertion" and span.text for span in spans)
            or any(
                element.tag not in {qn("w:pPr"), qn("w:ins")}
                and not _is_comment_annotation(element)
                for element in paragraph._p
            )
            or any(
                segment.kind != "opaque"
                or segment.element is None
                or (
                    segment.element.tag != qn("w:ins")
                    and not _is_comment_annotation(segment.element)
                )
                for segment in segments
            )
        ):
            raise DocumentError("paragraph is not the expected entirely inserted content")
        _require_local_comment_ranges(current, paragraph, source_bytes=self._source_bytes)
        self._require_supported_revisions()
        parent.remove(paragraph._p)
        replacement = self._from_pydocx(self._doc, filename=self.filename)
        replacement._source_bytes = self._source_bytes
        self.__dict__.update(replacement.__dict__)
