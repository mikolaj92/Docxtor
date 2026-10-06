"""Feature extraction for typed mechanical OOXML facts."""

from __future__ import annotations

from collections.abc import Mapping

from lxml import etree

from .docx_facts_models import _R_NS, _W_NS, NamedFact, ParagraphFact, RelationshipFact

_FEATURE_CATEGORIES = (
    "fields",
    "bookmarks",
    "links",
    "hidden",
    "comments",
    "notes",
    "textboxes",
    "properties",
    "embedded_objects",
    "media",
)

def _container_for_element(
    part_name: str,
    element: etree._Element,
    containers: Mapping[tuple[str, str], str],
) -> str | None:
    current: etree._Element | None = element
    tree = element.getroottree()
    while current is not None:
        if etree.QName(current).localname == "p":
            return containers.get((part_name, tree.getpath(current)))
        current = current.getparent()
    return None

def _features(
    roots: dict[str, etree._Element],
    paragraphs: tuple[ParagraphFact, ...],
    relationships: tuple[RelationshipFact, ...],
    part_names: set[str],
) -> dict[str, tuple[NamedFact, ...]]:
    singular = {
        "fields": "field",
        "bookmarks": "bookmark",
        "links": "link",
        "hidden": "hidden",
        "comments": "comment",
        "notes": "note",
        "textboxes": "textbox",
        "properties": "property",
        "embedded_objects": "embedded_object",
        "media": "media",
    }
    out: dict[str, list[NamedFact]] = {name: [] for name in singular.values()}
    containers_by_part_path = {
        (paragraph.part_name, paragraph.xml_path): paragraph.container_id
        for paragraph in paragraphs
        if paragraph.part_name and paragraph.xml_path
    }
    relationship_targets = {
        (relationship.source_part, relationship.relationship_id): relationship.target
        for relationship in relationships
    }
    for part, root in sorted(roots.items()):
        tree = root.getroottree()
        for element in root.iter():
            local = etree.QName(element).localname
            path = tree.getpath(element)
            fid = f"{part}:{path}"
            text = "".join(element.itertext()) or None
            container = _container_for_element(part, element, containers_by_part_path)
            if local in {"fldSimple", "instrText", "fldChar"}:
                value = (
                    element.get(f"{{{_W_NS}}}instr")
                    or element.get(f"{{{_W_NS}}}fldCharType")
                    or text
                )
                display = next(
                    (
                        paragraph.text
                        for paragraph in paragraphs
                        if paragraph.container_id == container
                    ),
                    None,
                )
                out["field"].append(NamedFact("field", part, fid, container, value, display))
            if local in {"bookmarkStart", "bookmarkEnd"}:
                value = element.get(f"{{{_W_NS}}}name") or element.get(f"{{{_W_NS}}}id")
                out["bookmark"].append(NamedFact("bookmark", part, fid, container, value))
            if local == "hyperlink":
                out["link"].append(
                    NamedFact(
                        "link",
                        part,
                        fid,
                        container,
                        element.get(f"{{{_W_NS}}}anchor"),
                        relationship_targets.get(
                            (part, element.get(f"{{{_R_NS}}}id") or ""),
                            element.get(f"{{{_R_NS}}}id"),
                        ),
                    )
                )
            if local in {"vanish", "webHidden", "specVanish"} and (
                element.get(f"{{{_W_NS}}}val") or "true"
            ).casefold() not in {"0", "false", "off", "no"}:
                run = element.getparent()
                while run is not None and etree.QName(run).localname != "r":
                    run = run.getparent()
                run_text = "" if run is None else "".join(run.itertext())
                out["hidden"].append(NamedFact("hidden", part, fid, container, local, run_text))
            if local.startswith("comment") or part.startswith("word/comments"):
                out["comment"].append(
                    NamedFact("comment", part, fid, container, element.get(f"{{{_W_NS}}}id"), text)
                )
            if local in {"footnote", "endnote"} and part == "word/settings.xml":
                # footnotePr/endnotePr list Word's default note IDs as settings
                # metadata; they are not user-authored note definitions.
                continue
            if local in {"footnote", "endnote", "footnoteReference", "endnoteReference"}:
                note_text = "".join(
                    (node.text or "")
                    if etree.QName(node).localname in {"t", "delText"}
                    else "\t"
                    if etree.QName(node).localname == "tab"
                    else "\n"
                    if etree.QName(node).localname in {"br", "cr"}
                    else ""
                    for node in element.iter()
                )
                out["note"].append(
                    NamedFact(
                        (
                            f"{local}_{element.get(f'{{{_W_NS}}}type') or 'user'}"
                            if local in {"footnote", "endnote"}
                            else local
                        ),
                        part,
                        fid,
                        container,
                        element.get(f"{{{_W_NS}}}id"),
                        note_text,
                    )
                )
            if local == "txbxContent":
                out["textbox"].append(NamedFact("textbox", part, fid, container, text))
            if part.startswith("docProps/") and element is not root:
                out["property"].append(NamedFact("property", part, fid, None, local, text))
            if local in {"object", "oleObject", "control"}:
                out["embedded_object"].append(
                    NamedFact(
                        "embedded_object",
                        part,
                        fid,
                        container,
                        " ".join(value for value in (local, *element.attrib.values()) if value),
                        element.get(f"{{{_R_NS}}}id"),
                    )
                )
    # Binary package facts use relationship targets so identities stay stable.
    binary_parts = set(part_names)
    binary_parts.update(r.target_part for r in relationships if r.target_part)
    for part in sorted(binary_parts):
        if part.startswith("word/embeddings/"):
            out["embedded_object"].append(NamedFact("embedded_object", part, part))
        elif part.startswith("word/media/"):
            out["media"].append(NamedFact("media", part, part))
        elif part.startswith("word/activeX/"):
            out["embedded_object"].append(NamedFact("embedded_object", part, part))
    return {plural: tuple(out[singular[plural]]) for plural in _FEATURE_CATEGORIES}
