from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn

from .docx_ns import V_TEXTBOX, W_TXBX_CONTENT, WPS_TXBX


def _local_tag(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _is_text_box_container(tag: str) -> bool:
    """True for Word/VML/DrawingML text-box wrappers.

    ``w:txbxContent`` is the actual paragraph host. VML ``v:textbox`` and
    DrawingML ``wps:txbx`` wrap that host; walk into them so a box nested
    inside a drawing is not silence.
    """
    if tag in {W_TXBX_CONTENT, V_TEXTBOX, WPS_TXBX}:
        return True
    return _local_tag(tag) in {"txbxContent", "textbox", "txbx"}


def _w_get(element: Any, name: str) -> str | None:
    return element.get(qn(f"w:{name}"))


def _clark(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"
