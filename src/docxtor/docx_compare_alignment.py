"""Bounded story alignment and accepted-text span diffing."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from difflib import SequenceMatcher
from hashlib import sha256
from math import inf

from .docx_compare_models import (
    BlockPair,
    ComparisonDiagnostic,
    DocumentBlock,
    DocumentView,
    TextAnchor,
    TextChange,
    TextChangeKind,
)

_MAX_ALIGNMENT_PRODUCT = 200_000
_MAX_TOKEN_PRODUCT = 100_000
_TOKEN_PATTERN = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)


@dataclass(frozen=True)
class AlignmentResult:
    left: DocumentView
    right: DocumentView
    pairs: tuple[BlockPair, ...]
    diagnostics: tuple[ComparisonDiagnostic, ...]


def align_blocks(left: DocumentView, right: DocumentView, comparison_id: str) -> AlignmentResult:
    left_groups = _group_blocks(left.blocks)
    right_groups = _group_blocks(right.blocks)
    group_ids = sorted(
        set(left_groups) | set(right_groups),
        key=lambda key: _group_order(key, left_groups, right_groups),
    )
    ordered_pairs: list[tuple[float, int, BlockPair]] = []
    pair_sequence = 0
    diagnostics = _shifted_table_ordinal_diagnostics(left, right)

    def add_pair(
        before_block: DocumentBlock | None,
        after_block: DocumentBlock | None,
        position: float,
    ) -> None:
        nonlocal pair_sequence
        ordered_pairs.append(
            (position, pair_sequence, _pair(before_block, after_block, comparison_id))
        )
        pair_sequence += 1

    for group_id in group_ids:
        before = left_groups.get(group_id, [])
        after = right_groups.get(group_id, [])
        before_keys = [_match_key(block) for block in before]
        after_keys = [_match_key(block) for block in after]
        if before_keys == after_keys:
            opcodes = [("equal", 0, len(before), 0, len(after))]
        elif len(before) * len(after) <= _MAX_ALIGNMENT_PRODUCT:
            opcodes = SequenceMatcher(None, before_keys, after_keys, autojunk=False).get_opcodes()
        else:
            matches = _greedy_matches(before_keys, after_keys)
            opcodes = _opcodes_from_matches(len(before), len(after), matches)
            diagnostics.append(
                ComparisonDiagnostic(
                    "alignment_workload_fallback",
                    "large story used bounded monotone alignment; "
                    "repeated block matches may be approximate",
                    "both",
                )
            )
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                for i, j in zip(range(i1, i2), range(j1, j2), strict=True):
                    add_pair(before[i], after[j], after[j].order)
            elif tag == "delete":
                position = _next_order(after, j1, before[i1].order if i1 < i2 else 0)
                for block in before[i1:i2]:
                    add_pair(block, None, position)
            elif tag == "insert":
                for block in after[j1:j2]:
                    add_pair(None, block, block.order)
            else:
                left_slice = before[i1:i2]
                right_slice = after[j1:j2]
                shared = min(len(left_slice), len(right_slice))
                for index in range(shared):
                    if left_slice[index].story_id == right_slice[index].story_id:
                        add_pair(left_slice[index], right_slice[index], right_slice[index].order)
                    else:
                        add_pair(left_slice[index], None, right_slice[index].order)
                        add_pair(None, right_slice[index], right_slice[index].order)
                remainder_order = _next_order(
                    after, j2, left_slice[shared].order if shared < len(left_slice) else 0
                )
                for block in left_slice[shared:]:
                    add_pair(block, None, remainder_order)
                for block in right_slice[shared:]:
                    add_pair(None, block, block.order)

    ordered_pairs.sort(key=lambda item: (item[0], item[1]))
    pairs = [item[2] for item in ordered_pairs]
    left_pair_ids = {pair.left_block_id: pair.pair_id for pair in pairs if pair.left_block_id}
    right_pair_ids = {pair.right_block_id: pair.pair_id for pair in pairs if pair.right_block_id}
    left_blocks = tuple(replace(block, pair_id=left_pair_ids[block.id]) for block in left.blocks)
    right_blocks = tuple(replace(block, pair_id=right_pair_ids[block.id]) for block in right.blocks)
    left = _assign_comment_pairs(left, left_blocks)
    right = _assign_comment_pairs(right, right_blocks)
    return AlignmentResult(left, right, tuple(pairs), tuple(diagnostics))


def _shifted_table_ordinal_diagnostics(
    left: DocumentView, right: DocumentView
) -> list[ComparisonDiagnostic]:
    left_tables = _table_signatures(left.blocks)
    right_tables = _table_signatures(right.blocks)
    left_unmatched: dict[tuple[tuple[int | None, int | None, str], ...], list[int]] = {}
    right_unmatched: dict[tuple[tuple[int | None, int | None, str], ...], list[int]] = {}
    for ordinal, signature in left_tables.items():
        if signature and right_tables.get(ordinal) != signature:
            left_unmatched.setdefault(signature, []).append(ordinal)
    for ordinal, signature in right_tables.items():
        if signature and left_tables.get(ordinal) != signature:
            right_unmatched.setdefault(signature, []).append(ordinal)

    diagnostics = []
    for signature, left_ordinals in left_unmatched.items():
        right_ordinals = right_unmatched.get(signature, ())
        for left_ordinal, right_ordinal in zip(left_ordinals, right_ordinals, strict=False):
            diagnostics.append(
                ComparisonDiagnostic(
                    "table_ordinal_alignment_coarse",
                    "matching table content has a different ordinal; "
                    "table block alignment may be coarse",
                    "both",
                    locator=f"table:{left_ordinal}->table:{right_ordinal}",
                )
            )
    return diagnostics


def _table_signatures(
    blocks: tuple[DocumentBlock, ...],
) -> dict[int, tuple[tuple[int | None, int | None, str], ...]]:
    grouped: dict[int, list[tuple[int | None, int | None, str]]] = {}
    for block in blocks:
        table_index = block.coordinate.table_index
        if table_index is None:
            continue
        grouped.setdefault(table_index, []).append(
            (block.coordinate.row_index, block.coordinate.cell_index, block.text)
        )
    return {index: tuple(signature) for index, signature in grouped.items()}


def diff_text(
    pairs: tuple[BlockPair, ...], left: DocumentView, right: DocumentView
) -> tuple[tuple[TextChange, ...], tuple[ComparisonDiagnostic, ...]]:
    left_by_id = {block.id: block for block in left.blocks}
    right_by_id = {block.id: block for block in right.blocks}
    changes: list[TextChange] = []
    diagnostics: list[ComparisonDiagnostic] = []
    for pair in pairs:
        before = left_by_id.get(pair.left_block_id or "")
        after = right_by_id.get(pair.right_block_id or "")
        if before is None:
            if after is not None and after.text:
                changes.append(_single_side_change(pair, None, after, inserted=True))
            continue
        if after is None:
            if before.text:
                changes.append(_single_side_change(pair, before, None, inserted=False))
            continue
        if before.text == after.text:
            continue
        if len(_tokens(before.text)) * len(_tokens(after.text)) > _MAX_TOKEN_PRODUCT:
            changes.append(_bounded_text_change(pair, before, after))
            diagnostics.append(
                ComparisonDiagnostic(
                    "text_diff_workload_fallback",
                    "large paragraph used a bounded common-prefix/suffix diff span",
                    "both",
                    locator=pair.pair_id,
                )
            )
            continue
        changes.extend(_span_changes(pair, before, after))
    return tuple(changes), tuple(diagnostics)


def _group_order(
    key: str,
    left_groups: dict[str, list[DocumentBlock]],
    right_groups: dict[str, list[DocumentBlock]],
) -> tuple[float, str]:
    orders = [block.order for block in left_groups.get(key, ())]
    orders.extend(block.order for block in right_groups.get(key, ()))
    return (min(orders, default=inf), key)


def _next_order(blocks: list[DocumentBlock], next_index: int, fallback: int) -> float:
    if next_index < len(blocks):
        return float(blocks[next_index].order)
    return float(fallback) + 0.5


def _group_blocks(blocks: tuple[DocumentBlock, ...]) -> dict[str, list[DocumentBlock]]:
    groups: dict[str, list[DocumentBlock]] = {}
    for block in blocks:
        group = (
            "main"
            if block.story_id == "body" or block.story_id.startswith("table:")
            else block.story_id
        )
        groups.setdefault(group, []).append(block)
    return groups


def _match_key(block: DocumentBlock) -> tuple[str, str]:
    return block.story_id, block.text


def _greedy_matches(
    before: list[tuple[str, str]], after: list[tuple[str, str]]
) -> list[tuple[int, int]]:
    positions: dict[tuple[str, str], list[int]] = {}
    for index, key in enumerate(after):
        positions.setdefault(key, []).append(index)
    cursor_by_key: dict[tuple[str, str], int] = {}
    cursor = 0
    matches = []
    for before_index, key in enumerate(before):
        values = positions.get(key, ())
        position = cursor_by_key.get(key, 0)
        while position < len(values) and values[position] < cursor:
            position += 1
        if position == len(values):
            cursor_by_key[key] = position
            continue
        after_index = values[position]
        cursor_by_key[key] = position + 1
        matches.append((before_index, after_index))
        cursor = after_index + 1
    return matches


def _opcodes_from_matches(
    before_size: int, after_size: int, matches: list[tuple[int, int]]
) -> list[tuple[str, int, int, int, int]]:
    opcodes: list[tuple[str, int, int, int, int]] = []
    before_cursor = after_cursor = 0
    for before_index, after_index in [*matches, (before_size, after_size)]:
        if before_cursor < before_index or after_cursor < after_index:
            if before_cursor == before_index:
                tag = "insert"
            elif after_cursor == after_index:
                tag = "delete"
            else:
                tag = "replace"
            opcodes.append((tag, before_cursor, before_index, after_cursor, after_index))
        if before_index < before_size and after_index < after_size:
            opcodes.append(("equal", before_index, before_index + 1, after_index, after_index + 1))
        before_cursor = before_index + 1
        after_cursor = after_index + 1
    return opcodes


def _pair(
    before: DocumentBlock | None, after: DocumentBlock | None, comparison_id: str
) -> BlockPair:
    left_id = before.id if before is not None else None
    right_id = after.id if after is not None else None
    pair_id = _stable_id("pair", comparison_id, left_id or "", right_id or "")
    return BlockPair(
        pair_id,
        before.story_id if before is not None else after.story_id if after is not None else "",
        left_id,
        right_id,
    )


def _assign_comment_pairs(view: DocumentView, blocks: tuple[DocumentBlock, ...]) -> DocumentView:
    pair_by_block = {block.id: block.pair_id for block in blocks}
    comments = []
    for comment in view.comments:
        if comment.anchor is None or comment.anchor.block_id is None:
            comments.append(comment)
            continue
        comments.append(
            replace(
                comment,
                anchor=replace(
                    comment.anchor,
                    pair_id=pair_by_block.get(comment.anchor.block_id),
                ),
            )
        )
    return replace(view, blocks=blocks, comments=tuple(comments))


def _single_side_change(
    pair: BlockPair,
    before: DocumentBlock | None,
    after: DocumentBlock | None,
    *,
    inserted: bool,
) -> TextChange:
    text = (
        after.text if inserted and after is not None else before.text if before is not None else ""
    )
    if inserted:
        left = None
        right = (
            TextAnchor(after.id, pair.pair_id, 0, len(text), text) if after is not None else None
        )
    else:
        left = (
            TextAnchor(before.id, pair.pair_id, 0, len(text), text) if before is not None else None
        )
        right = None
    kind = TextChangeKind.INSERTED if inserted else TextChangeKind.DELETED
    return TextChange(
        _change_id(
            pair.pair_id,
            kind,
            text,
            "",
            left_span=(0, len(text)) if not inserted else None,
            right_span=(0, len(text)) if inserted else None,
        ),
        pair.pair_id,
        kind,
        left,
        right,
    )


def _span_changes(pair: BlockPair, before: DocumentBlock, after: DocumentBlock) -> list[TextChange]:
    before_tokens = _tokens(before.text)
    after_tokens = _tokens(after.text)
    matcher = SequenceMatcher(
        None,
        [token[0] for token in before_tokens],
        [token[0] for token in after_tokens],
        autojunk=False,
    )
    result = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        left_start, left_end = _token_range(before_tokens, i1, i2, len(before.text))
        right_start, right_end = _token_range(after_tokens, j1, j2, len(after.text))
        left_text = before.text[left_start:left_end]
        right_text = after.text[right_start:right_end]
        kind = (
            TextChangeKind.INSERTED
            if not left_text
            else TextChangeKind.DELETED
            if not right_text
            else TextChangeKind.REPLACED
        )
        left_anchor = (
            TextAnchor(before.id, pair.pair_id, left_start, left_end, left_text)
            if left_text
            else None
        )
        right_anchor = (
            TextAnchor(after.id, pair.pair_id, right_start, right_end, right_text)
            if right_text
            else None
        )
        result.append(
            TextChange(
                _change_id(
                    pair.pair_id,
                    kind,
                    left_text,
                    right_text,
                    left_span=(left_start, left_end),
                    right_span=(right_start, right_end),
                ),
                pair.pair_id,
                kind,
                left_anchor,
                right_anchor,
            )
        )
    return result


def _bounded_text_change(
    pair: BlockPair, before: DocumentBlock, after: DocumentBlock
) -> TextChange:
    before_tokens = _tokens(before.text)
    after_tokens = _tokens(after.text)
    prefix = 0
    while (
        prefix < min(len(before_tokens), len(after_tokens))
        and before_tokens[prefix][0] == after_tokens[prefix][0]
    ):
        prefix += 1
    suffix = 0
    while (
        suffix < min(len(before_tokens) - prefix, len(after_tokens) - prefix)
        and before_tokens[len(before_tokens) - suffix - 1][0]
        == after_tokens[len(after_tokens) - suffix - 1][0]
    ):
        suffix += 1
    left_start, left_end = _token_range(
        before_tokens, prefix, len(before_tokens) - suffix, len(before.text)
    )
    right_start, right_end = _token_range(
        after_tokens, prefix, len(after_tokens) - suffix, len(after.text)
    )
    left_text = before.text[left_start:left_end]
    right_text = after.text[right_start:right_end]
    kind = (
        TextChangeKind.REPLACED
        if left_text and right_text
        else TextChangeKind.INSERTED
        if right_text
        else TextChangeKind.DELETED
    )
    return TextChange(
        _change_id(
            pair.pair_id,
            kind,
            left_text,
            right_text,
            left_span=(left_start, left_end),
            right_span=(right_start, right_end),
        ),
        pair.pair_id,
        kind,
        TextAnchor(before.id, pair.pair_id, left_start, left_end, left_text) if left_text else None,
        TextAnchor(after.id, pair.pair_id, right_start, right_end, right_text)
        if right_text
        else None,
    )


def _tokens(text: str) -> list[tuple[str, int, int]]:
    return [(match.group(), match.start(), match.end()) for match in _TOKEN_PATTERN.finditer(text)]


def _token_range(
    tokens: list[tuple[str, int, int]], start: int, end: int, text_length: int
) -> tuple[int, int]:
    if start == end:
        position = tokens[start][1] if start < len(tokens) else text_length
        return position, position
    return tokens[start][1], tokens[end - 1][2]


def _change_id(
    pair_id: str,
    kind: TextChangeKind,
    left: str,
    right: str,
    *,
    left_span: tuple[int, int] | None,
    right_span: tuple[int, int] | None,
) -> str:
    left_coordinates = ("", "") if left_span is None else (str(left_span[0]), str(left_span[1]))
    right_coordinates = ("", "") if right_span is None else (str(right_span[0]), str(right_span[1]))
    return _stable_id(
        "change",
        pair_id,
        kind.value,
        left,
        right,
        *left_coordinates,
        *right_coordinates,
    )


def _stable_id(prefix: str, *parts: str) -> str:
    return f"{prefix}-{sha256(chr(31).join(parts).encode('utf-8')).hexdigest()[:20]}"
