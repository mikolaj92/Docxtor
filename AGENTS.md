# AGENTS

Docxtor is the **sole mechanical DOCX layer**. It owns physical OOXML addressing,
inventory, mutation, facts, and package publication.

It does **not** decide review meaning, PII, or law. ReviewKit (and Dike via it)
layers review semantics. Temida consumers stay thin adapters.

## Non-negotiable

1. **Sole mechanical DOCX layer.** All paragraph, run, offset, and package work
   lives here. Consumers must not re-implement run splitting, offset math, or
   paragraph mutation.
2. **No review or law semantics.** Report physical values, locators, and
   capabilities only. Do not classify PII, legal relevance, or review decisions.
3. **One public API.** Import from `docxtor`. Do not grow a second public
   surface (CLI, extra packages, or private helper re-exports).

## Local proof

The mill runs `[tool.lokay] test` → `uv run ruff check . && uv run pytest -q`.
That is the same gate as [`.github/workflows/ci.yml`](.github/workflows/ci.yml)
(ruff + pytest).
