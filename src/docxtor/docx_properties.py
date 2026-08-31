from __future__ import annotations

from hashlib import sha256

from lxml import etree

from .docx_inventory import SurfaceKind, inventory_docx
from .docx_package import PackageError, parse_package_xml, read_package_entries
from .docx_package_transaction import (
    PackageMutation,
    PackageMutationKind,
    apply_package_transaction,
)

_CORE = "docProps/core.xml"
_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"


def read_core_keywords(data: bytes) -> str:
    inventory = inventory_docx(data)
    surface = _keyword_surface(inventory)
    return "" if surface is None else surface.value


def set_core_keywords(data: bytes, value: str) -> bytes:
    inventory = inventory_docx(data)
    surface = _keyword_surface(inventory)
    if surface is not None:
        if surface.value == value:
            return data
        mutation = PackageMutation(
            "set-core-keywords",
            PackageMutationKind.REPLACE_SURFACE,
            surface.surface_id,
            surface.value_sha256,
            value,
        )
        return apply_package_transaction(data, [mutation]).data
    entries = read_package_entries(data)
    entry = next((item for item in entries if item.name == _CORE), None)
    if entry is None:
        raise PackageError("DOCX package has no docProps/core.xml")
    root = parse_package_xml(entry.data, part_name=_CORE)
    node = etree.SubElement(root, f"{{{_CP}}}keywords")
    node.text = value
    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    mutation = PackageMutation(
        "create-core-keywords",
        PackageMutationKind.REPLACE_PART,
        _CORE,
        sha256(entry.data).hexdigest(),
        payload,
    )
    return apply_package_transaction(data, [mutation]).data


def remove_core_keyword_values(data: bytes, *, prefix: str) -> bytes:
    current = read_core_keywords(data)
    kept = [value for value in current.split(";") if value and not value.startswith(prefix)]
    return set_core_keywords(data, ";".join(kept))


def _keyword_surface(inventory):
    return next(
        (
            surface
            for surface in inventory.surfaces
            if surface.part_name == _CORE
            and surface.kind is SurfaceKind.XML_TEXT
            and surface.element_qname
            and surface.element_qname.rsplit("}", 1)[-1] == "keywords"
        ),
        None,
    )
