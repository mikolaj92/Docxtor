"""Package-level fact helpers: payload reading, content types, reachability."""

from __future__ import annotations

from pathlib import Path, PurePosixPath
from posixpath import normpath

from lxml import etree

from .docx import DocxDocument
from .docx_facts_models import ContentTypeFact, RelationshipFact, Source
from .docx_package import PackageError, parse_package_xml


def _payload(source: Source) -> bytes:
    if isinstance(source, DocxDocument):
        raw = getattr(source, "_source_bytes", None)
        return raw if isinstance(raw, bytes) else source.to_bytes()
    if isinstance(source, bytes):
        return source
    try:
        return Path(source).read_bytes()
    except OSError as exc:
        raise PackageError(f"cannot read DOCX: {exc}") from exc


def _content_types(data: bytes) -> tuple[ContentTypeFact, ...]:
    root = parse_package_xml(data, part_name="[Content_Types].xml")
    if etree.QName(root).localname != "Types":
        raise PackageError("[Content_Types].xml has an invalid root")
    facts: list[ContentTypeFact] = []
    for element in root:
        local = etree.QName(element).localname
        key = element.get("Extension") if local == "Default" else element.get("PartName")
        value = element.get("ContentType")
        if local not in {"Default", "Override"} or not key or not value:
            raise PackageError("[Content_Types].xml contains an invalid declaration")
        facts.append(ContentTypeFact(key.lstrip("/"), value, local == "Default"))
    return tuple(sorted(facts, key=lambda x: (not x.is_default, x.key)))


def _content_type_map(
    facts: tuple[ContentTypeFact, ...],
) -> tuple[dict[str, str], dict[str, str]]:
    defaults = {x.key.lower(): x.content_type for x in facts if x.is_default}
    overrides = {x.key: x.content_type for x in facts if not x.is_default}
    overrides["[Content_Types].xml"] = "application/xml"
    return defaults, overrides


def _content_type_for(name: str, defaults: dict[str, str], overrides: dict[str, str]) -> str:
    # OPC uses the text after the last dot, including the root ".rels" part.
    _, separator, extension = PurePosixPath(name).name.rpartition(".")
    return overrides.get(
        name,
        defaults.get(
            extension.lower() if separator else "",
            "application/octet-stream",
        ),
    )


def _relationships(entries: dict[str, bytes]) -> tuple[RelationshipFact, ...]:
    facts: list[RelationshipFact] = []
    for name, data in sorted(entries.items()):
        if not name.endswith(".rels"):
            continue
        root = parse_package_xml(data, part_name=name)
        if etree.QName(root).localname != "Relationships":
            raise PackageError(f"relationship part {name} has an invalid root")
        source = _relationship_source(name)
        seen: set[str] = set()
        for element in root:
            if etree.QName(element).localname != "Relationship":
                raise PackageError(f"relationship part {name} contains an invalid element")
            rid, rtype, target = element.get("Id"), element.get("Type"), element.get("Target")
            if not rid or not rtype or target is None or rid in seen:
                raise PackageError(f"relationship part {name} contains an invalid relationship")
            seen.add(rid)
            external = element.get("TargetMode", "").lower() == "external"
            facts.append(
                RelationshipFact(
                    source,
                    name,
                    rid,
                    rtype,
                    target,
                    None if external else _resolve_target(source, target),
                    external,
                )
            )
    return tuple(sorted(facts, key=lambda x: x.identity))


def _relationship_source(name: str) -> str:
    if name == "_rels/.rels":
        return ""
    path = PurePosixPath(name)
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _resolve_target(source: str, target: str) -> str:
    base = str(PurePosixPath(source).parent) if source else ""
    resolved = normpath(f"{base}/{target}".lstrip("/"))
    if resolved == ".." or resolved.startswith("../"):
        raise PackageError(f"relationship target escapes package: {target}")
    return resolved


def _reachable(names: set[str], rels: tuple[RelationshipFact, ...]) -> set[str]:
    reachable = {"_rels/.rels"} if "_rels/.rels" in names else set()
    queue = [""]
    while queue:
        source = queue.pop(0)
        for rel in rels:
            if (
                rel.source_part == source
                and not rel.external
                and rel.target_part in names
                and rel.target_part not in reachable
            ):
                reachable.add(rel.target_part or "")
                queue.append(rel.target_part or "")
                rel_part = _rels_name(rel.target_part or "")
                if rel_part in names:
                    reachable.add(rel_part)
    return reachable


def _rels_name(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / f"{path.name}.rels")
