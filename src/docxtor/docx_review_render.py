"""Combined physical DOCX review rendering; contains no review policy."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, Literal

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .docx import DocxDocument
from .docx_inline import (
    _visible_text,
    paragraph_to_inline_segments,
)
from .docx_models import InlineSegment
from .docx_package import normalize_docx_timestamps
from .docx_publish import publish_docx
from .docx_revisions import inventory_revisions_bytes


@dataclass(frozen=True)
class PhysicalReviewEdit:
    action_id: str
    locator: str
    operation: Literal["insert", "delete", "replace"]
    start_offset: int
    end_offset: int
    replacement_text: str = ""
    expected_text: str | None = None
    comment_text: str | None = None
    new_paragraph: bool = False
    insert_after: bool = True


@dataclass(frozen=True)
class PhysicalReviewComment:
    locator: str
    text: str
    start_offset: int | None = None
    end_offset: int | None = None
    anchor_text: str | None = None


@dataclass(frozen=True)
class PhysicalReviewPlan:
    edits: tuple[PhysicalReviewEdit, ...] = ()
    comments: tuple[PhysicalReviewComment, ...] = ()


@dataclass(frozen=True)
class PhysicalReviewer:
    author: str = "Reviewer"
    initials: str = "RV"
    date: str | None = None


class PhysicalReviewRenderError(ValueError):
    pass


@dataclass
class _Piece:
    kind: Literal["text", "opaque", "ins", "del"]
    text: str
    rpr: Any | None = None
    element: Any | None = None
    action_id: str | None = None
    revision_id: int | None = None
    start_comments: list[int] = field(default_factory=list)
    end_comments: list[int] = field(default_factory=list)


def render_physical_review(
    source: str | Path | bytes,
    output: str | Path,
    plan: PhysicalReviewPlan,
    *,
    reviewer: PhysicalReviewer | None = None,
) -> Path:
    reviewer = reviewer or PhysicalReviewer()
    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    addressable = DocxDocument.open_bytes(data)
    doc = addressable._doc
    next_revision = (
        max(
            (
                int(r.revision_id)
                for r in inventory_revisions_bytes(data).revisions
                if r.revision_id and r.revision_id.isdigit()
            ),
            default=0,
        )
        + 1
    )
    by_locator: dict[str, list[PhysicalReviewEdit]] = {}
    for edit in plan.edits:
        by_locator.setdefault(edit.locator, []).append(edit)
    deferred = []
    try:
        for locator, edits in by_locator.items():
            paragraph = addressable.resolve_paragraph(locator)
            if paragraph is None:
                raise PhysicalReviewRenderError(f"unknown review locator: {locator}")
            blocks = [e for e in edits if e.new_paragraph]
            inline = [e for e in edits if not e.new_paragraph]
            pieces = [
                _Piece(s.kind, s.text, s.rpr, s.element)
                for s in paragraph_to_inline_segments(paragraph)
            ]
            for edit in inline:
                pieces, next_revision = _apply_edit(pieces, edit, next_revision)
            related = [c for c in plan.comments if c.locator == locator]
            for edit in inline:
                if edit.comment_text:
                    _mark_action_comment(doc, pieces, edit.action_id, edit.comment_text, reviewer)
            for comment in related:
                _mark_comment(doc, pieces, comment, reviewer)
            _rebuild(paragraph, pieces, reviewer)
            deferred.extend((paragraph, e) for e in blocks)
        # comments on paragraphs without edits
        for comment in plan.comments:
            if comment.locator in by_locator:
                continue
            paragraph = addressable.resolve_paragraph(comment.locator)
            if paragraph is None:
                raise PhysicalReviewRenderError(f"unknown comment locator: {comment.locator}")
            pieces = [
                _Piece(s.kind, s.text, s.rpr, s.element)
                for s in paragraph_to_inline_segments(paragraph)
            ]
            _mark_comment(doc, pieces, comment, reviewer)
            _rebuild(paragraph, pieces, reviewer)
        for anchor, edit in deferred:
            next_revision = _insert_block(doc, anchor, edit, reviewer, next_revision)
        out = BytesIO()
        doc.save(out)
        rendered = out.getvalue()
        receipt = publish_docx(rendered, output, source=data)
        normalize_docx_timestamps(receipt.destination)
        return receipt.destination
    except PhysicalReviewRenderError:
        raise
    except Exception as exc:
        raise PhysicalReviewRenderError(str(exc)) from exc


def render_physical_clean(
    source: str | Path | bytes, output: str | Path, edits: tuple[PhysicalReviewEdit, ...]
) -> Path:
    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    addressable = DocxDocument.open_bytes(data)
    doc = addressable._doc
    by_locator: dict[str, list[PhysicalReviewEdit]] = {}
    for edit in edits:
        by_locator.setdefault(edit.locator, []).append(edit)
    try:
        for locator, items in by_locator.items():
            paragraph = addressable.resolve_paragraph(locator)
            if paragraph is None:
                raise PhysicalReviewRenderError(f"unknown review locator: {locator}")
            pieces = [
                _Piece(s.kind, s.text, s.rpr, s.element)
                for s in paragraph_to_inline_segments(paragraph)
            ]
            for edit in items:
                if edit.new_paragraph:
                    continue
                pieces, _ = _apply_edit(pieces, edit, 1, tracked=False)
            _rebuild(paragraph, pieces, PhysicalReviewer(), tracked=False)
        out = BytesIO()
        doc.save(out)
        receipt = publish_docx(out.getvalue(), output, source=data)
        normalize_docx_timestamps(receipt.destination)
        return receipt.destination
    except PhysicalReviewRenderError:
        raise
    except Exception as exc:
        raise PhysicalReviewRenderError(str(exc)) from exc


def _apply_edit(
    pieces: list[_Piece], edit: PhysicalReviewEdit, next_id: int, tracked: bool = True
) -> tuple[list[_Piece], int]:
    base = [
        InlineSegment(
            "text" if p.kind in {"text", "ins", "del"} else "opaque", p.text, p.rpr, p.element
        )
        for p in pieces
    ]
    visible = _visible_text(base)
    if not 0 <= edit.start_offset <= edit.end_offset <= len(visible):
        raise PhysicalReviewRenderError(f"invalid review range at {edit.locator}")
    if (
        edit.expected_text is not None
        and visible[edit.start_offset : edit.end_offset] != edit.expected_text
    ):
        raise PhysicalReviewRenderError(f"review range text changed at {edit.locator}")
    pieces = _split_pieces(pieces, edit.end_offset)
    pieces = _split_pieces(pieces, edit.start_offset)
    start = _piece_index(pieces, edit.start_offset)
    end = _piece_index(pieces, edit.end_offset)
    chosen = pieces[start:end]
    if edit.end_offset > edit.start_offset and (
        not chosen or any(p.kind != "text" for p in chosen)
    ):
        raise PhysicalReviewRenderError(f"review range crosses opaque content at {edit.locator}")
    repl = []
    if edit.operation in {"delete", "replace"} and tracked:
        repl.extend(
            _Piece(
                "del",
                p.text,
                deepcopy(p.rpr),
                action_id=edit.action_id,
                revision_id=next_id,
            )
            for p in chosen
        )
        if chosen:
            next_id += 1
    if edit.operation in {"insert", "replace"} and edit.replacement_text:
        repl.append(
            _Piece(
                "ins" if tracked else "text",
                edit.replacement_text,
                deepcopy(_rpr(pieces, edit.start_offset)),
                action_id=edit.action_id,
                revision_id=next_id if tracked else None,
            )
        )
        next_id += 1 if tracked else 0
    if edit.operation == "insert":
        pieces[start:start] = repl
    else:
        pieces[start:end] = repl
    return pieces, next_id


def _split_pieces(pieces: list[_Piece], offset: int) -> list[_Piece]:
    pos = 0
    out = []
    for piece in pieces:
        length = len(piece.text) if piece.kind != "del" else len(piece.text)
        if piece.kind == "opaque":
            length = len(piece.text)
        if 0 < offset - pos < length and piece.kind == "text":
            cut = offset - pos
            out.extend(
                [
                    _Piece("text", piece.text[:cut], deepcopy(piece.rpr)),
                    _Piece("text", piece.text[cut:], deepcopy(piece.rpr)),
                ]
            )
        else:
            out.append(piece)
        pos += length
    return out


def _piece_index(pieces: list[_Piece], offset: int) -> int:
    pos = 0
    for i, p in enumerate(pieces):
        if pos >= offset:
            return i
        pos += len(p.text)
    return len(pieces)


def _rpr(pieces: list[_Piece], offset: int) -> Any | None:
    i = _piece_index(pieces, offset)
    candidates = pieces[i : i + 1] or pieces[-1:]
    return candidates[0].rpr if candidates else None


def _create_comment(doc: Any, text: str, reviewer: PhysicalReviewer) -> int:
    c = doc.comments.add_comment(text=text, author=reviewer.author, initials=reviewer.initials)
    if reviewer.date:
        c._comment_elm.set(qn("w:date"), reviewer.date)
    return int(c.comment_id)


def _mark_action_comment(
    doc: Any, pieces: list[_Piece], action_id: str, text: str, reviewer: PhysicalReviewer
) -> None:
    ids = [i for i, p in enumerate(pieces) if p.action_id == action_id]
    if not ids:
        return
    cid = _create_comment(doc, text, reviewer)
    pieces[min(ids)].start_comments.append(cid)
    pieces[max(ids)].end_comments.append(cid)


def _mark_comment(
    doc: Any, pieces: list[_Piece], comment: PhysicalReviewComment, reviewer: PhysicalReviewer
) -> None:
    if not pieces:
        return
    start, end = 0, len(pieces) - 1
    if comment.start_offset is not None and comment.end_offset is not None:
        pieces[:] = _split_pieces(_split_pieces(pieces, comment.end_offset), comment.start_offset)
        start = _piece_index(pieces, comment.start_offset)
        end = max(start, _piece_index(pieces, comment.end_offset) - 1)
    elif comment.anchor_text:
        text = "".join(p.text for p in pieces)
        at = text.find(comment.anchor_text)
        if at >= 0:
            pieces[:] = _split_pieces(_split_pieces(pieces, at + len(comment.anchor_text)), at)
            start = _piece_index(pieces, at)
            end = max(start, _piece_index(pieces, at + len(comment.anchor_text)) - 1)
    cid = _create_comment(doc, comment.text, reviewer)
    pieces[start].start_comments.append(cid)
    pieces[end].end_comments.append(cid)


def _rebuild(
    paragraph: Any, pieces: list[_Piece], reviewer: PhysicalReviewer, tracked: bool = True
) -> None:
    parent = paragraph._p
    for child in list(parent):
        if child.tag != qn("w:pPr"):
            parent.remove(child)
    for p in pieces:
        for cid in p.start_comments:
            parent.append(_marker("w:commentRangeStart", cid))
        if p.kind == "opaque" and p.element is not None:
            parent.append(deepcopy(p.element))
        elif p.kind in {"ins", "del"} and tracked:
            parent.append(_revision(p, reviewer))
        elif p.text:
            parent.append(_run(p.text, p.rpr, "w:t"))
        for cid in reversed(p.end_comments):
            parent.append(_marker("w:commentRangeEnd", cid))
            r = OxmlElement("w:r")
            r.append(_marker("w:commentReference", cid))
            parent.append(r)


def _run(text: str, rpr: Any | None, tag: str) -> Any:
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(deepcopy(rpr))
    node = OxmlElement(tag)
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run.append(node)
    return run


def _revision(piece: _Piece, reviewer: PhysicalReviewer) -> Any:
    e = OxmlElement(f"w:{piece.kind}")
    e.set(qn("w:id"), str(piece.revision_id or 0))
    e.set(qn("w:author"), reviewer.author)
    if reviewer.date:
        e.set(qn("w:date"), reviewer.date)
    e.append(_run(piece.text, piece.rpr, "w:t" if piece.kind == "ins" else "w:delText"))
    return e


def _marker(tag: str, cid: int) -> Any:
    e = OxmlElement(tag)
    e.set(qn("w:id"), str(cid))
    return e


def _insert_block(
    doc: Any, anchor: Any, edit: PhysicalReviewEdit, reviewer: PhysicalReviewer, next_id: int
) -> int:
    lines = edit.replacement_text.splitlines() or [edit.replacement_text]
    elements = []
    for line in lines:
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        rpr = OxmlElement("w:rPr")
        rpr.append(_revision(_Piece("ins", "", revision_id=next_id), reviewer))
        ppr.append(rpr)
        p.append(ppr)
        p.append(_revision(_Piece("ins", line, revision_id=next_id), reviewer))
        elements.append(p)
        next_id += 1
    ref = anchor._p
    for element in elements:
        if edit.insert_after:
            ref.addnext(element)
            ref = element
        else:
            ref.addprevious(element)
    return next_id
