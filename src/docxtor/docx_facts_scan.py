"""Scanner building the mechanical DOCX facts snapshot."""

from __future__ import annotations

from hashlib import sha256

from docx.text.paragraph import Paragraph
from lxml import etree

from .docx import DocxDocument
from .docx_facts_features import _features
from .docx_facts_models import (
    _W_NS,
    ContainerCoordinate,
    DocxFactsSnapshot,
    DocxStructureSnapshot,
    FactDiagnostic,
    FactsCoverage,
    PageBreakFact,
    PageLayoutFacts,
    ParagraphFact,
    PartFact,
    Source,
    StoryFact,
    UnreadablePartFact,
)
from .docx_facts_package import (
    _content_type_for,
    _content_type_map,
    _content_types,
    _payload,
    _reachable,
    _relationships,
)
from .docx_inventory import inventory_docx
from .docx_package import PackageError, parse_package_xml, read_package_entries


def docx_facts(source: Source) -> DocxFactsSnapshot:
    """Return the complete mechanical snapshot, or raise on unreadable/malformed input."""
    payload = _payload(source)
    entries = read_package_entries(payload, validate_xml=False)
    entry_data = {entry.name: entry.data for entry in entries}
    if "[Content_Types].xml" not in entry_data:
        raise PackageError("DOCX package has no [Content_Types].xml")
    content_types = _content_types(entry_data["[Content_Types].xml"])
    ct_defaults, ct_overrides = _content_type_map(content_types)
    inventory = inventory_docx(payload)
    fatal_unreadable = tuple(
        name
        for name in inventory.unreadable_parts
        if name.startswith("word/") or name in {"[Content_Types].xml", "_rels/.rels"}
    )
    if fatal_unreadable:
        raise PackageError("DOCX inventory has unreadable parts: " + ", ".join(fatal_unreadable))

    relationships = _relationships(entry_data)
    reachable = _reachable(set(entry_data), relationships)
    diagnostics: list[FactDiagnostic] = [
        FactDiagnostic("unknown_part", "content type is not mechanically understood", name)
        for name in inventory.unknown_parts
    ]
    diagnostics.extend(
        FactDiagnostic("unreadable_part", "XML part is mechanically unreadable", name)
        for name in inventory.unreadable_parts
    )
    for rel in relationships:
        if not rel.external and rel.target_part not in entry_data:
            diagnostics.append(
                FactDiagnostic(
                    "missing_relationship_target",
                    f"missing target {rel.target_part}",
                    rel.relationship_part,
                )
            )
    missing_targets = [item for item in diagnostics if item.code == "missing_relationship_target"]
    if missing_targets:
        raise PackageError("DOCX relationship target is missing: " + missing_targets[0].message)

    inventory_parts = {part.name: part for part in inventory.parts}
    parts = tuple(
        PartFact(
            name,
            _content_type_for(name, ct_defaults, ct_overrides),
            len(data),
            sha256(data).hexdigest(),
            inventory_parts[name].is_xml,
            name in reachable,
            inventory_parts[name].understood,
        )
        for name, data in sorted(entry_data.items())
    )

    # DocxDocument is the canonical story/addressing implementation. It also
    # rejects packages that python-docx cannot interpret.
    document = source if isinstance(source, DocxDocument) else DocxDocument.open_bytes(payload)
    xml_roots = {
        name: parse_package_xml(data, part_name=name)
        for name, data in entry_data.items()
        if inventory_parts[name].is_xml
        and name not in inventory.unreadable_parts
        and not name.endswith(".rels")
        and name != "[Content_Types].xml"
    }
    paragraphs = _paragraphs(document, xml_roots)
    stories = _stories(paragraphs)
    features = _features(xml_roots, paragraphs, relationships, set(entry_data))
    orphans = tuple(sorted(set(entry_data) - reachable - {"[Content_Types].xml", "_rels/.rels"}))
    page_layout = _page_layout(xml_roots, paragraphs)
    coverage = FactsCoverage.COMPLETE if not diagnostics else FactsCoverage.INCOMPLETE
    structure = DocxStructureSnapshot(
        part_names=tuple(sorted(entry_data)),
        relationship_identities=tuple(rel.identity for rel in relationships),
        story_ids=tuple(story.story_id for story in stories),
        container_ids=tuple(p.container_id for p in paragraphs),
        field_ids=tuple(f.fact_id for f in features["fields"]),
        bookmark_ids=tuple(f.fact_id for f in features["bookmarks"]),
        table_ids=tuple(
            f"table:{index}"
            for index in sorted(
                {
                    paragraph.coordinate.table_index
                    for paragraph in paragraphs
                    if paragraph.coordinate.table_index is not None
                }
            )
        ),
        body_block_ids=_body_block_ids(paragraphs),
        section_property_hashes=_section_property_hashes(xml_roots),
        coverage=coverage,
    )
    return DocxFactsSnapshot(
        FactsCoverage.COMPLETE if not diagnostics else FactsCoverage.INCOMPLETE,
        tuple(diagnostics),
        parts,
        content_types,
        relationships,
        tuple(sorted(reachable)),
        orphans,
        inventory.surfaces,
        stories,
        paragraphs,
        features["fields"],
        features["bookmarks"],
        features["links"],
        features["hidden"],
        features["comments"],
        features["notes"],
        features["textboxes"],
        features["properties"],
        features["embedded_objects"],
        features["media"],
        tuple(
            UnreadablePartFact(name, inventory_parts[name].error or "unreadable_xml")
            for name in inventory.unreadable_parts
        ),
        page_layout,
        structure,
    )


# A discoverable noun/verb pair for callers that prefer ``snapshot_docx``.
snapshot_docx = docx_facts

def _body_block_ids(paragraphs: tuple[ParagraphFact, ...]) -> tuple[str, ...]:
    result: list[str] = []
    for paragraph in paragraphs:
        if paragraph.part_name != "word/document.xml":
            continue
        table_index = paragraph.coordinate.table_index
        if table_index is None and paragraph.story_kind != "body":
            continue
        block_id = f"table:{table_index}" if table_index is not None else paragraph.container_id
        if not result or result[-1] != block_id:
            result.append(block_id)
    return tuple(result)

def _section_property_hashes(
    xml_roots: dict[str, etree._Element],
) -> tuple[str, ...]:
    values: list[str] = []
    for part_name, root in sorted(xml_roots.items()):
        if not part_name.startswith("word/"):
            continue
        for section in root.iter(f"{{{_W_NS}}}sectPr"):
            payload = etree.tostring(section, method="c14n")
            values.append(sha256(payload).hexdigest())
    return tuple(values)

def _page_layout(
    roots: dict[str, etree._Element], paragraphs: tuple[ParagraphFact, ...]
) -> PageLayoutFacts:
    app_pages: int | None = None
    app_root = roots.get("docProps/app.xml")
    if app_root is not None:
        for node in app_root.iter():
            if etree.QName(node).localname != "Pages":
                continue
            raw = (node.text or "").strip()
            if raw.isdigit() and int(raw) > 0:
                app_pages = int(raw)
            break
    breaks: list[PageBreakFact] = []
    chunks: list[list[str]] = [[]]
    body = roots.get("word/document.xml")
    paragraph_paths = {
        paragraph.xml_path: paragraph.container_id
        for paragraph in paragraphs
        if paragraph.part_name == "word/document.xml" and paragraph.xml_path
    }
    if body is not None:
        tree = body.getroottree()
        for node in body.iter():
            local = etree.QName(node).localname
            kind = None
            if local == "lastRenderedPageBreak":
                kind = "rendered"
            elif local == "br" and node.get(f"{{{_W_NS}}}type") == "page":
                kind = "explicit"
            if kind is not None:
                current = node
                container_id = None
                while current is not None:
                    if etree.QName(current).localname == "p":
                        container_id = paragraph_paths.get(tree.getpath(current))
                        break
                    current = current.getparent()
                breaks.append(
                    PageBreakFact(
                        kind,
                        "word/document.xml",
                        f"word/document.xml:{tree.getpath(node)}",
                        container_id,
                    )
                )
                chunks.append([])
            elif local in {"t", "delText"} and node.text:
                chunks[-1].append(node.text)
    return PageLayoutFacts(app_pages, tuple(breaks), tuple("".join(chunk) for chunk in chunks))

def _paragraphs(
    document: DocxDocument, roots: dict[str, etree._Element]
) -> tuple[ParagraphFact, ...]:
    result: list[ParagraphFact] = []
    seen: set[str] = set()
    candidates: list[tuple[int | None, str, Paragraph]] = [
        (index, container_id, paragraph)
        for index, container_id, paragraph in document.get_indexed_paragraphs()
    ]
    for segment in document.segments:
        container_id = segment.container_id
        if container_id is None or container_id in seen:
            continue
        paragraph = document.resolve_paragraph(container_id)
        if paragraph is not None and segment.paragraph_index is None:
            candidates.append((None, container_id, paragraph))
    for index, container_id, paragraph in candidates:
        if container_id in seen:
            continue
        seen.add(container_id)
        part_name = _part_for_container_id(container_id)
        path = paragraph._p.getroottree().getpath(paragraph._p)
        root = roots.get(part_name)
        namespaces = (
            {}
            if root is None
            else {key: value for key, value in root.nsmap.items() if key is not None}
        )
        source_paragraph = None if root is None else root.xpath(path, namespaces=namespaces)
        element: etree._Element = source_paragraph[0] if source_paragraph else paragraph._p
        style_nodes = element.findall(f"./{{{_W_NS}}}pPr/{{{_W_NS}}}pStyle")
        style_id = style_nodes[0].get(f"{{{_W_NS}}}val") if style_nodes else "Normal"
        outline = None
        nodes = element.findall(f"./{{{_W_NS}}}pPr/{{{_W_NS}}}outlineLvl")
        if nodes:
            raw = nodes[0].get(f"{{{_W_NS}}}val")
            try:
                outline = int(raw) if raw is not None else None
            except ValueError:
                raise PackageError(f"invalid outline level in {container_id}") from None
        coord = _coordinate(container_id, index)
        result.append(
            ParagraphFact(
                container_id,
                index,
                coord.story_kind,
                paragraph.text,
                sha256(paragraph.text.encode()).hexdigest(),
                style_id,
                outline,
                coord,
                path,
                part_name,
            )
        )
    for story in ("footnote", "endnote"):
        part_name = f"word/{story}s.xml"
        root = roots.get(part_name)
        if root is None:
            continue
        tree = root.getroottree()
        for note in root.findall(f"{{{_W_NS}}}{story}"):
            note_type = note.get(f"{{{_W_NS}}}type")
            note_id = note.get(f"{{{_W_NS}}}id")
            if note_id is None or note_type in {"separator", "continuationSeparator"}:
                continue
            for local_index, paragraph in enumerate(note.findall(f"{{{_W_NS}}}p")):
                container_id = f"{story}:{note_id}:p:{local_index}"
                if container_id in seen:
                    continue
                seen.add(container_id)
                text = "".join(
                    node.text or ""
                    for node in paragraph.iter()
                    if etree.QName(node).localname in {"t", "delText"}
                )
                style_nodes = paragraph.findall(f"./{{{_W_NS}}}pPr/{{{_W_NS}}}pStyle")
                style_id = style_nodes[0].get(f"{{{_W_NS}}}val") if style_nodes else "Normal"
                coord = _coordinate(container_id, None)
                result.append(
                    ParagraphFact(
                        container_id,
                        None,
                        story,
                        text,
                        sha256(text.encode()).hexdigest(),
                        style_id,
                        None,
                        coord,
                        tree.getpath(paragraph),
                        part_name,
                    )
                )
    return tuple(result)

def _part_for_container_id(container_id: str) -> str:
    if container_id.startswith(("body:", "table:", "txbx:")):
        return "word/document.xml"
    bits = container_id.split(":")
    if bits[0].startswith(("header", "footer")) and len(bits) > 1:
        base = bits[0].split("-", 1)[0]
        return f"word/{base}{int(bits[1]) + 1}.xml"
    if bits[0] in {"footnote", "endnote"}:
        return f"word/{bits[0]}s.xml"
    if bits[0] == "comment":
        return "word/comments.xml"
    return "word/document.xml"

def _coordinate(cid: str, paragraph_index: int | None) -> ContainerCoordinate:
    bits = cid.split(":")

    def after(token: str) -> int | None:
        try:
            return int(bits[bits.index(token) + 1])
        except (ValueError, IndexError):
            return None

    kind = bits[0]
    ordinal = (
        after(kind) if kind in {"header", "footer", "txbx", "comment", "footnote", "endnote"} else 0
    )
    return ContainerCoordinate(
        cid, kind, ordinal or 0, after("table"), after("r"), after("c"), paragraph_index
    )

def _stories(paragraphs: tuple[ParagraphFact, ...]) -> tuple[StoryFact, ...]:
    grouped: dict[str, list[str]] = {}
    for paragraph in paragraphs:
        cid = paragraph.container_id
        story = cid.rsplit(":p:", 1)[0]
        if ":table:" in story:
            story = story.split(":table:", 1)[0] or "body"
        grouped.setdefault(story, []).append(cid)
    return tuple(
        StoryFact(key, key.split(":", 1)[0], tuple(values))
        for key, values in sorted(grouped.items())
    )
