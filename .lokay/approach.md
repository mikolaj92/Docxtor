# Approach plan

<!-- lokay-approach source=deterministic repo=mikolaj92/Docxtor issue=24 -->

Repository: `mikolaj92/Docxtor`  
Issue: #24 — Tag v0.4.1 nie zgadza sie z version 0.4.0 w paczce

## Goal

Align published package metadata with a version that can be tagged and pinned
without repeating the v0.4.1 / dist 0.4.0 mismatch.

Tag `v0.4.1` (commit `d125429`, DocToText → Docxtor rename) still has
`pyproject.toml version = 0.4.0`. That historical tag cannot be rewritten from
this PR. Main was still `0.4.0`, so a pin of `v0.4.1` installs dist `0.4.0`.

## Files likely touched

- `pyproject.toml` — bump `project.version` to `0.4.2`
- `src/docxtor/__init__.py` — export matching `__version__`
- `tests/test_version.py` — keep metadata and public version in lockstep

## Test plan

- `uv run pytest tests/test_version.py`
- `uv run pytest` if the environment is already synced

## Non-goals

- Do not move, delete, or retag `v0.4.1`
- Do not create a GitHub release from this coding session
- Do not change document-processing behavior

## Notes

- Next release should be tagged `v0.4.2` so the pin and dist versions match.
- Trust intentional issue; this plan is evidence for later review, not a human gate.
