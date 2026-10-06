from __future__ import annotations

from collections.abc import Iterable, Sequence
from copy import deepcopy
from dataclasses import replace
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypeVar
from zipfile import ZipFile

from lxml import etree

from .docx_comment_mutations import (
    CommentAuthor,
    CommentMutationResult,
    CommentRange,
)
from .docx_comment_mutations import (
    add_comment as add_comment_bytes,
)
from .docx_comment_mutations import (
    remove_comments as remove_comments_bytes,
)
from .docx_comment_mutations import (
    update_comment as update_comment_bytes,
)
from .docx_package import PackageEntry, PackageError, parse_package_xml, read_package_entries
from .docx_publish import PublishError, PublishReceipt, publish_docx
from .docx_review_inventory import inventory_review_markup
from .docx_review_models import OperationReceipt
from .docx_review_transaction import ReviewCommand, apply_review_batch
from .docx_revision_mutations import (
    RevisionAuthor,
    RevisionMutationResult,
    RevisionPosition,
    RevisionRange,
)
from .docx_revision_mutations import (
    delete_revision as delete_revision_bytes,
)
from .docx_revision_mutations import (
    insert_revision as insert_revision_bytes,
)
from .docx_revisions import inventory_revisions_bytes

if TYPE_CHECKING:
    from .docx_review_geometry import PhysicalCommentRange, PhysicalReviewGeometry
    from .docx_review_projection import DocxReviewProjection

_MutationResult = TypeVar("_MutationResult", CommentMutationResult, RevisionMutationResult)
_LEGACY_SERIALIZER_OPERATIONS = frozenset(
    {"add_comment", "update_comment", "remove_comments", "insert_revision", "delete_revision"}
)


class DocxHandleOperations:
    """Comment and native tracked-change mutations on one open DOCX handle."""

    filename: str
    _source_bytes: bytes | None
    _source_path: Path | None

    def project_review(self) -> DocxReviewProjection:
        """Project review data and geometry from one immutable current snapshot."""
        from .docx_review_projection import project_docx_for_review

        return project_docx_for_review(self._current_review_bytes())

    def review_geometry(self) -> PhysicalReviewGeometry:
        """Return physical ranges bound to the current document's exact bytes."""
        from .docx_review_geometry import project_docx_review_geometry

        return project_docx_review_geometry(self._current_review_bytes())

    def add_physical_comment(
        self,
        target: PhysicalCommentRange,
        text: str,
        author: CommentAuthor,
    ) -> CommentMutationResult:
        """Add a comment only while its captured physical geometry is current."""
        from .docx_review_geometry_comments import add_physical_comment_bytes

        result = add_physical_comment_bytes(self._current_review_bytes(), target, text, author)
        self._adopt_bytes(result.data)
        return result

    def apply_review_batch(self, commands: Sequence[ReviewCommand]) -> None:
        source = self._current_review_bytes()
        receipt = apply_review_batch(source, [self._preserving_command(item) for item in commands])
        self._adopt_bytes(receipt.data)

    def _preserving_command(self, command: ReviewCommand) -> ReviewCommand:
        def mutate(data: bytes) -> tuple[bytes, OperationReceipt]:
            rendered, receipt = command.mutate(data)
            preserved = self._preserve_review_output(
                data,
                rendered,
                removed_parts=frozenset(receipt.affected_parts),
                updated_parts=(
                    frozenset(receipt.affected_parts)
                    if receipt.operation not in _LEGACY_SERIALIZER_OPERATIONS
                    else frozenset()
                ),
            )
            return preserved, _receipt_for_payload(data, preserved, receipt)

        return replace(command, mutate=mutate)

    def add_comment(
        self,
        target: CommentRange,
        text: str,
        author: CommentAuthor,
    ) -> CommentMutationResult:
        source = self._current_review_bytes()
        result = self._preserve_mutation_result(
            source, add_comment_bytes(source, target, text, author)
        )
        self._adopt_bytes(result.data)
        return result

    def update_comment(
        self,
        comment_id: str,
        text: str,
        *,
        expected_text: str | None = None,
        author: CommentAuthor | None = None,
    ) -> CommentMutationResult:
        source = self._current_review_bytes()
        result = self._preserve_mutation_result(
            source,
            update_comment_bytes(
                source,
                comment_id,
                text,
                expected_text=expected_text,
                author=author,
            ),
        )
        self._adopt_bytes(result.data)
        return result

    def delete_comment(self, comment_id: str) -> CommentMutationResult:
        return self.remove_comments({comment_id})

    def remove_comments(self, comment_ids: set[str] | None = None) -> CommentMutationResult:
        source = self._current_review_bytes()
        result = self._preserve_mutation_result(source, remove_comments_bytes(source, comment_ids))
        self._adopt_bytes(result.data)
        return result

    def insert_revision(
        self,
        position: RevisionPosition,
        text: str,
        reviewer: RevisionAuthor,
    ) -> RevisionMutationResult:
        source = self._current_review_bytes()
        result = self._preserve_mutation_result(
            source, insert_revision_bytes(source, position, text, reviewer)
        )
        self._adopt_bytes(result.data)
        return result

    def delete_revision(
        self,
        target: RevisionRange,
        reviewer: RevisionAuthor,
    ) -> RevisionMutationResult:
        source = self._current_review_bytes()
        result = self._preserve_mutation_result(
            source, delete_revision_bytes(source, target, reviewer)
        )
        self._adopt_bytes(result.data)
        return result

    def replace_revision(
        self,
        target: RevisionRange,
        replacement: str,
        reviewer: RevisionAuthor,
    ) -> tuple[RevisionMutationResult, RevisionMutationResult]:
        source = self._current_review_bytes()
        deleted = self._preserve_mutation_result(
            source, delete_revision_bytes(source, target, reviewer)
        )
        inserted = self._preserve_mutation_result(
            deleted.data,
            insert_revision_bytes(
                deleted.data,
                RevisionPosition(target.locator, target.start_offset),
                replacement,
                reviewer,
            ),
        )
        self._adopt_bytes(inserted.data)
        return deleted, inserted

    def publish(
        self,
        path: str | Path | None = None,
        *,
        validators: Iterable[Any] = (),
    ) -> PublishReceipt:
        """Validate and atomically replace the opened path, or ``path`` if given."""
        destination = Path(path) if path is not None else self._source_path
        if destination is None:
            raise PublishError(
                "publish() requires a destination path when the document was not opened from a file"
            )
        return publish_docx(
            self._current_review_bytes(),
            destination,
            source=self._source_bytes,
            validators=validators,
        )

    def save_docx(self, path: str | Path) -> None:
        self.publish(path)

    def _current_review_bytes(self) -> bytes:
        """Overlay in-memory edits without losing unreachable source entries.

        Comparing with the same source round-trip isolates edits from serializer
        normalization. An unchanged handle retains its original archive digest;
        changed snapshots use stable ZIP metadata so projection and mutation see
        identical bytes even across separate calls.
        """
        rendered = self.to_bytes()
        source = self._source_bytes
        if source is None:
            return _package_bytes(read_package_entries(rendered))
        return self._preserve_review_output(source, rendered)

    def _preserve_review_output(
        self,
        source: bytes,
        rendered: bytes,
        *,
        removed_parts: frozenset[str] = frozenset(),
        updated_parts: frozenset[str] = frozenset(),
    ) -> bytes:
        """Preserve source entries around serializer output, honoring reachable removals."""
        if rendered == source:
            return source
        original = read_package_entries(source)
        rendered_entries = read_package_entries(rendered)
        baseline = read_package_entries(type(self).open_bytes(source).to_bytes())
        current_by_name = {entry.name: entry for entry in rendered_entries}
        baseline_by_name = {entry.name: entry for entry in baseline}
        if (
            not (removed_parts - current_by_name.keys())
            and current_by_name.keys() == baseline_by_name.keys()
            and all(
                entry.data == baseline_by_name[name].data for name, entry in current_by_name.items()
            )
        ):
            return source

        entries = []
        original_names = {entry.name for entry in original}
        removed_names = (baseline_by_name.keys() | removed_parts) - current_by_name.keys()
        retained_names = original_names - removed_names
        retained_names.update(current_by_name)
        for entry in original:
            if entry.name in removed_names:
                continue
            before = baseline_by_name.get(entry.name)
            after = current_by_name.get(entry.name)
            if before is None:
                # python-docx cannot reach this part; its absence is not a removal.
                if after is not None and after.data != entry.data:
                    if entry.name not in updated_parts:
                        raise PackageError(
                            f"mutation output overwrites an unreachable source entry: {entry.name}"
                        )
                    entries.append(replace(entry, data=after.data))
                else:
                    entries.append(entry)
            elif after is not None:
                data = entry.data
                if entry.name == "[Content_Types].xml" and (
                    after.data != before.data or removed_names
                ):
                    data = _merge_content_types(
                        entry.data, before.data, after.data, retained_names, removed_names
                    )
                elif after.data != before.data:
                    data = after.data
                entries.append(replace(entry, data=data))
        entries.extend(entry for entry in rendered_entries if entry.name not in original_names)
        return _package_bytes(tuple(entries))

    def _preserve_mutation_result(
        self, source: bytes, result: _MutationResult
    ) -> _MutationResult:
        data = self._preserve_review_output(source, result.data)
        receipt = _receipt_for_payload(source, data, result.receipt)
        if isinstance(result, CommentMutationResult):
            return replace(
                result, data=data, receipt=receipt, comments=inventory_review_markup(data).comments
            )
        return replace(
            result,
            data=data,
            receipt=receipt,
            before=inventory_revisions_bytes(source),
            after=inventory_revisions_bytes(data),
        )

    def _adopt_bytes(self, data: bytes) -> None:
        source_path = self._source_path
        replacement = type(self).open_bytes(data, filename=self.filename)
        self.__dict__.update(replacement.__dict__)
        self._source_path = source_path


def _package_bytes(entries: tuple[PackageEntry, ...]) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w") as archive:
        for entry in entries:
            archive.writestr(entry.zip_info(), entry.data)
    payload = buffer.getvalue()
    read_package_entries(payload)
    return payload


def _receipt_for_payload(source: bytes, data: bytes, receipt: OperationReceipt) -> OperationReceipt:
    before_parts = {entry.name: entry.data for entry in read_package_entries(source)}
    after_parts = {entry.name: entry.data for entry in read_package_entries(data)}
    return replace(
        receipt,
        before_sha256=sha256(source).hexdigest(),
        after_sha256=sha256(data).hexdigest(),
        affected_parts=tuple(
            sorted(
                name
                for name in before_parts.keys() | after_parts.keys()
                if before_parts.get(name) != after_parts.get(name)
            )
        ),
    )


def _merge_content_types(
    original: bytes,
    before: bytes,
    after: bytes,
    retained_names: set[str],
    removed_names: set[str],
) -> bytes:
    """Apply serializer-visible declarations while retaining orphan declarations."""
    root = parse_package_xml(original, part_name="[Content_Types].xml")
    before_root = parse_package_xml(before, part_name="[Content_Types].xml")
    after_root = parse_package_xml(after, part_name="[Content_Types].xml")

    def declarations(tree: Any) -> dict[tuple[str, str | None], Any]:
        return {
            (
                node.tag,
                node.get("PartName") if node.tag.endswith("Override") else node.get("Extension"),
            ): node
            for node in tree
            if isinstance(node.tag, str)
            and node.tag.rsplit("}", 1)[-1] in {"Default", "Override"}
        }

    original_by_key = declarations(root)
    before_by_key = declarations(before_root)
    after_by_key = declarations(after_root)
    retained_extensions = {
        Path(name).suffix.removeprefix(".").casefold() for name in retained_names
    }
    for key, node in before_by_key.items():
        current = after_by_key.get(key)
        if current is not None and dict(node.attrib) == dict(current.attrib):
            continue
        if (
            current is None
            and node.tag.endswith("Default")
            and (node.get("Extension") or "").casefold() in retained_extensions
        ):
            # A source orphan can still depend on a default the renderer dropped.
            continue
        original_node = original_by_key.get(key)
        if original_node is not None:
            root.remove(original_node)
        if current is not None:
            root.append(deepcopy(current))
    for key, node in after_by_key.items():
        if key not in before_by_key:
            original_node = original_by_key.get(key)
            if original_node is not None:
                root.remove(original_node)
            root.append(deepcopy(node))
    for node in list(root):
        if (
            isinstance(node.tag, str)
            and node.tag.endswith("Override")
            and (node.get("PartName") or "").lstrip("/") in removed_names
        ):
            root.remove(node)
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
