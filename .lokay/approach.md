# Approach plan

<!-- lokay-approach source=deterministic repo=mikolaj92/Docxtor issue=165 -->

Repository: `mikolaj92/Docxtor`  
Issue: #165 — Cleanup: AGENTS.md + [tool.lokay] test

## Goal

Add AGENTS.md (mechanical-DOCX-only law) and `[tool.lokay] test = ["uv", "run", "pytest", "-q"]` (or ruff+pytest) so Lokay does not skip Docxtor.

## Files likely touched

- `AGENTS.md`
- `pyproject.toml`

## Test plan

- AGENTS.md states: sole mechanical DOCX layer; no review/law semantics; one public API
- `[tool.lokay] test` runs the same gate as CI (`.github/workflows/ci.yml`: ruff + pytest)
- No product code change required
- `uv run ruff check . && uv run pytest -q`

## Non-goals

- New features; expanding mypy file list (optional follow-up).

## Notes

- Trust intentional issue; this plan is evidence for later review, not a human gate.
- Coding agent may refine details but should stay on the stated goal and non-goals.
- Collector boundary: if implementation introduces unbounded collection, ship only a bounded collector patch that starts durably in the background after merge. The coding agent and lokay must not populate data or wait for collection to finish.
