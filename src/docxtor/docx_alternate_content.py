from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from .docx_ns import MC_NS, W_P, W_SDT, W_SDT_CONTENT
from .docx_xml import _is_text_box_container

SUPPORTED_REQUIRES = frozenset(
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://schemas.microsoft.com/office/word/2006/wordml",
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape3d",
    }
)


@dataclass(frozen=True)
class AlternateContentCoverage:
    regions: int = 0
    choice_selected: int = 0
    fallback_selected: int = 0
    unknown_requirements: tuple[str, ...] = ()
    unsupported_requirements: tuple[str, ...] = ()

    @property
    def complete(self) -> bool:
        return not self.unknown_requirements and not self.unsupported_requirements


class DocxAlternateContentOperations:
    @property
    def alternate_content_coverage(self) -> AlternateContentCoverage:
        """MC read projection census; unknown Requires namespaces are explicit."""
        return self._alternate_content_coverage

    def _replace_full_segment(self, index: int, text: str) -> None:
        from .docx_units import _paragraph_visible_text, _replace_plain_range

        ref = self._refs[index]
        paragraph = ref.paragraph
        full = _paragraph_visible_text(paragraph)
        if self._replace_alternate_content_text(paragraph._p, text):
            self._refresh_after_edit(index)
            return
        if full:
            _replace_plain_range(paragraph._p, 0, len(full), text)
        elif paragraph.runs:
            paragraph.runs[0].text = text
        else:
            paragraph.add_run(text)
        self._refresh_after_edit(index)

    def _replace_alternate_content_text(self, paragraph: Any, text: str) -> bool:
        from .docx_units import _paragraph_visible_text, _replace_plain_range
        parent = paragraph.getparent()
        if parent is None or parent.tag not in {
            f"{{{MC_NS}}}Choice",
            f"{{{MC_NS}}}Fallback",
        }:
            return False
        region = parent.getparent()
        if region is None or region.tag != f"{{{MC_NS}}}AlternateContent":
            return False
        branch = select_alternate_branch(region)
        selected_paragraphs = [node for node in branch.selected.iter() if node.tag == W_P]
        if paragraph not in selected_paragraphs:
            raise ValueError("target paragraph is outside selected mc:AlternateContent branch")
        ordinal = selected_paragraphs.index(paragraph)
        branches = [
            child
            for child in region
            if child.tag in {f"{{{MC_NS}}}Choice", f"{{{MC_NS}}}Fallback"}
        ]
        targets = []
        for candidate_branch in branches:
            paragraphs = [node for node in candidate_branch.iter() if node.tag == W_P]
            if len(paragraphs) != len(selected_paragraphs):
                raise ValueError(
                    "mc:AlternateContent branches have incompatible paragraph structures"
                )
            targets.append(paragraphs[ordinal])
        for target in targets:
            current = _paragraph_visible_text(Paragraph(target, self._doc._body))
            if current:
                _replace_plain_range(target, 0, len(current), text)
            else:
                run = OxmlElement("w:r")
                node = OxmlElement("w:t")
                node.text = text
                run.append(node)
                target.append(run)
        return True


@dataclass(frozen=True)
class AlternateBranch:
    selected: Any
    selected_is_choice: bool
    unknown_requirements: tuple[str, ...]
    unsupported_requirements: tuple[str, ...]


def select_alternate_branch(element: Any) -> AlternateBranch:
    """Choose the first Choice whose Requires namespaces are supported, else Fallback."""
    fallback = None
    unknown: set[str] = set()
    unsupported: set[str] = set()
    for child in element:
        if child.tag == f"{{{MC_NS}}}Choice":
            prefixes = (child.get("Requires") or "").split()
            namespaces = child.nsmap
            if prefixes and all(prefix in namespaces for prefix in prefixes):
                required = {namespaces[p] for p in prefixes}
                unsupported.update(required - SUPPORTED_REQUIRES)
                if required <= SUPPORTED_REQUIRES:
                    return AlternateBranch(child, True, (), ())
                continue
            else:
                unknown.add("mc:Choice has missing or unbound mc:Requires prefix")
        elif child.tag == f"{{{MC_NS}}}Fallback":
            if fallback is not None:
                raise ValueError("mc:AlternateContent must not contain multiple Fallback elements")
            fallback = child
    if fallback is None:
        raise ValueError("mc:AlternateContent has no supported Choice and no Fallback")
    return AlternateBranch(
        fallback,
        False,
        tuple(sorted(unknown)),
        tuple(sorted(unsupported)),
    )


def iter_alternate_aware_paragraphs(container: Any, *, skip_text_boxes: bool = False):
    """Yield paragraphs in document order, selecting one MC branch at each region."""

    def walk(node: Any):
        for child in node:
            if skip_text_boxes and _is_text_box_container(child.tag):
                continue
            if child.tag == W_P:
                yield child
            elif child.tag == f"{{{MC_NS}}}AlternateContent":
                yield from walk(select_alternate_branch(child).selected)
            elif child.tag == W_SDT:
                for content in child.iterchildren(W_SDT_CONTENT):
                    yield from walk(content)

    yield from walk(container)


def iter_alternate_aware_blocks(container: Any):
    """Yield authored blocks, projecting one branch for each MC region."""
    for child in container.iterchildren():
        if child.tag == f"{{{MC_NS}}}AlternateContent":
            yield from select_alternate_branch(child).selected.iterchildren()
        else:
            yield child


def alternate_coverage_for(root: Any) -> AlternateContentCoverage:
    """Census all AlternateContent regions, including those outside indexed stories."""
    regions = list(root.iter(f"{{{MC_NS}}}AlternateContent"))
    choices = fallbacks = 0
    unknown: set[str] = set()
    unsupported: set[str] = set()
    for region in regions:
        branch = select_alternate_branch(region)
        if branch.selected_is_choice:
            choices += 1
        else:
            fallbacks += 1
        unknown.update(branch.unknown_requirements)
        unsupported.update(branch.unsupported_requirements)
    return AlternateContentCoverage(
        regions=len(regions),
        choice_selected=choices,
        fallback_selected=fallbacks,
        unknown_requirements=tuple(sorted(unknown)),
        unsupported_requirements=tuple(sorted(unsupported)),
    )
