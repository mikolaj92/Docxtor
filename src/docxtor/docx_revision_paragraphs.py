"""Paragraph-mark surgery for accept/reject revision operations."""

from __future__ import annotations

from lxml import etree

from .docx_revision_models import RevisionOperationError, _tag
from .docx_revision_package import _remove

_NON_TEXT = ("tab", "br", "cr", "drawing", "object", "pict", "fldChar", "sym")

def _validate_rejected_inserted_mark(
    mark: etree._Element, part: str, drop_comments: bool, error_type: type[RevisionOperationError]
) -> None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        raise error_type(f"{part}: malformed paragraph-mark insertion")
    if _paragraph_has_original_content(paragraph):
        if _next_paragraph_for_mark(mark) is None and not _is_content_control_paragraph(mark):
            raise error_type(f"{part}: tracked paragraph-mark insertion has no following paragraph")
        return
    allowed = {
        _tag("pPr"),
        _tag("ins"),
        _tag("moveTo"),
        _tag("sdt"),
        _tag("bookmarkStart"),
        _tag("bookmarkEnd"),
        _tag("permStart"),
        _tag("permEnd"),
        _tag("proofErr"),
    }
    for child in paragraph:
        if child.tag in allowed or _is_empty_text_container(child):
            continue
        if drop_comments and (
            child.tag in {_tag("commentRangeStart"), _tag("commentRangeEnd")}
            or _is_comment_reference_run(child)
        ):
            continue
        raise error_type(f"{part}: inserted paragraph contains unsupported structural children")

def _reject_inserted_mark(
    mark: etree._Element, part: str, drop_comments: bool, error_type: type[RevisionOperationError]
) -> None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        raise error_type(f"{part}: malformed paragraph-mark insertion")
    if _paragraph_has_original_content(paragraph):
        if not _merge_paragraph_into_next(mark):
            if _is_content_control_paragraph(mark):
                _remove(mark)
            else:
                raise error_type(f"{part}: cannot merge inserted paragraph")
        return
    for child in list(paragraph):
        if child.tag in {
            _tag("bookmarkStart"),
            _tag("bookmarkEnd"),
            _tag("permStart"),
            _tag("permEnd"),
            _tag("proofErr"),
        }:
            paragraph.addprevious(child)
    parent = _paragraph_block(paragraph).getparent()
    if parent is None:
        raise error_type(f"{part}: tracked paragraph has no parent")
    parent.remove(_paragraph_block(paragraph))

def _paragraph_for_mark(mark: etree._Element) -> etree._Element | None:
    rpr = mark.getparent()
    ppr = rpr.getparent() if rpr is not None else None
    paragraph = ppr.getparent() if ppr is not None else None
    return paragraph if paragraph is not None and paragraph.tag == _tag("p") else None


def _is_paragraph_mark(element: etree._Element) -> bool:
    return _paragraph_for_mark(element) is not None and element.getparent().tag == _tag("rPr")


def _paragraph_block(paragraph: etree._Element) -> etree._Element:
    content = paragraph.getparent()
    control = content.getparent() if content is not None else None
    if (
        content is not None
        and content.tag == _tag("sdtContent")
        and control is not None
        and control.tag == _tag("sdt")
        and len(content) == 1
    ):
        return control
    return paragraph


def _next_paragraph_for_mark(mark: etree._Element) -> etree._Element | None:
    paragraph = _paragraph_for_mark(mark)
    if paragraph is None:
        return None
    candidate = _paragraph_block(paragraph).getnext()
    if candidate is None:
        return None
    if candidate.tag == _tag("p"):
        return candidate
    if candidate.tag != _tag("sdt"):
        return None
    content = candidate.find(_tag("sdtContent"))
    return (
        content[0]
        if content is not None and len(content) == 1 and content[0].tag == _tag("p")
        else None
    )


def _merge_paragraph_into_next(mark: etree._Element) -> bool:
    paragraph = _paragraph_for_mark(mark)
    next_paragraph = _next_paragraph_for_mark(mark)
    if paragraph is None or next_paragraph is None:
        return False
    at = 1 if len(next_paragraph) and next_paragraph[0].tag == _tag("pPr") else 0
    for child in [node for node in paragraph if node.tag != _tag("pPr")]:
        next_paragraph.insert(at, child)
        at += 1
    block = _paragraph_block(paragraph)
    parent = block.getparent()
    if parent is None:
        return False
    next_block = _paragraph_block(next_paragraph)
    if next_block is not next_paragraph and next_block.getparent() is parent:
        next_block.addprevious(next_paragraph)
        parent.remove(next_block)
    parent.remove(block)
    return True


def _is_content_control_paragraph(mark: etree._Element) -> bool:
    paragraph = _paragraph_for_mark(mark)
    return paragraph is not None and _paragraph_block(paragraph) is not paragraph


def _paragraph_has_original_content(paragraph: etree._Element) -> bool:
    tags = [_tag("t"), _tag("delText"), *(_tag(n) for n in _NON_TEXT)]
    for node in paragraph.iter(*tags):
        if node.tag in {_tag("t"), _tag("delText")} and not (node.text or "").strip():
            continue
        if not any(a.tag in {_tag("ins"), _tag("moveTo")} for a in node.iterancestors()):
            return True
    return False


def _is_comment_reference_run(element: etree._Element) -> bool:
    return element.tag == _tag("r") and element.find(f".//{_tag('commentReference')}") is not None


def _is_empty_text_container(element: etree._Element) -> bool:
    if element.tag not in {_tag("r"), _tag("del"), _tag("moveFrom")}:
        return False
    for node in element.iter(_tag("t"), _tag("delText"), *(_tag(n) for n in _NON_TEXT)):
        if node.tag not in {_tag("t"), _tag("delText")} or (node.text or "").strip():
            return False
    return True


def _drop_numbering_leftover(paragraph: etree._Element, drop_comments: bool) -> None:
    if paragraph.getparent() is None:
        return
    text = "".join(n.text or "" for n in paragraph.iter(_tag("t"), _tag("delText"))).strip()
    ppr = paragraph.find(_tag("pPr"))
    numbered = ppr is not None and ppr.find(_tag("numPr")) is not None
    token = text.rstrip(".)").isdigit() and bool(text)
    if not ((numbered and not text) or token):
        return
    if any(next(paragraph.iter(_tag(n)), None) is not None for n in _NON_TEXT):
        return
    if not drop_comments and any(
        next(paragraph.iter(_tag(n)), None) is not None
        for n in ("commentRangeStart", "commentRangeEnd", "commentReference")
    ):
        return
    block = _paragraph_block(paragraph)
    parent = block.getparent()
    if parent is not None:
        parent.remove(block)
