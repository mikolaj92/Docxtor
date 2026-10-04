from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any

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
from .docx_publish import PublishError, PublishReceipt, publish_docx
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
from .docx_revision_mutations import (
    replace_revision as replace_revision_bytes,
)


class DocxHandleOperations:
    """Comment and native tracked-change mutations on one open DOCX handle."""

    filename: str
    _source_bytes: bytes | None
    _source_path: Path | None

    def apply_review_batch(self, commands: Sequence[ReviewCommand]) -> None:
        receipt = apply_review_batch(self.to_bytes(), commands)
        self._adopt_bytes(receipt.data)

    def add_comment(
        self,
        target: CommentRange,
        text: str,
        author: CommentAuthor,
    ) -> CommentMutationResult:
        result = add_comment_bytes(self.to_bytes(), target, text, author)
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
        result = update_comment_bytes(
            self.to_bytes(),
            comment_id,
            text,
            expected_text=expected_text,
            author=author,
        )
        self._adopt_bytes(result.data)
        return result

    def delete_comment(self, comment_id: str) -> CommentMutationResult:
        return self.remove_comments({comment_id})

    def remove_comments(self, comment_ids: set[str] | None = None) -> CommentMutationResult:
        result = remove_comments_bytes(self.to_bytes(), comment_ids)
        self._adopt_bytes(result.data)
        return result

    def insert_revision(
        self,
        position: RevisionPosition,
        text: str,
        reviewer: RevisionAuthor,
    ) -> RevisionMutationResult:
        result = insert_revision_bytes(self.to_bytes(), position, text, reviewer)
        self._adopt_bytes(result.data)
        return result

    def delete_revision(
        self,
        target: RevisionRange,
        reviewer: RevisionAuthor,
    ) -> RevisionMutationResult:
        result = delete_revision_bytes(self.to_bytes(), target, reviewer)
        self._adopt_bytes(result.data)
        return result

    def replace_revision(
        self,
        target: RevisionRange,
        replacement: str,
        reviewer: RevisionAuthor,
    ) -> tuple[RevisionMutationResult, RevisionMutationResult]:
        deleted, inserted = replace_revision_bytes(self.to_bytes(), target, replacement, reviewer)
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
            self.to_bytes(),
            destination,
            source=self._source_bytes,
            validators=validators,
        )

    def save_docx(self, path: str | Path) -> None:
        self.publish(path)

    def _adopt_bytes(self, data: bytes) -> None:
        source_path = self._source_path
        replacement = type(self).open_bytes(data, filename=self.filename)
        self.__dict__.update(replacement.__dict__)
        self._source_path = source_path
