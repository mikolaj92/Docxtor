# Approach plan

<!-- lokay-approach source=deterministic repo=mikolaj92/Docxtor issue=30 -->

Repository: `mikolaj92/Docxtor`  
Issue: #30 — Tag v0.4.1 nie zgadza sie z version 0.4.0 w paczce

## Goal

Record the historical pin mismatch honestly so consumers stop pinning `v0.4.1`.
Do not invent a new package version. Do not rewrite the old tag.

Tag `v0.4.1` (commit `d125429`, DocToText → Docxtor rename) still has
`pyproject.toml version = 0.4.0`. That tag cannot be rewritten from this PR.
Main and tag `v0.4.3` already agree (`version = 0.4.3`). Bumping past `0.4.3`
would be untruthful: there is no new release in this change.

## Files likely touched

- `README.md` — document that `v0.4.1` is a bad pin; use `v0.4.3` or later
- `tests/test_version.py` — keep the README warning from disappearing

## Test plan

- `uv run pytest tests/test_version.py`
- `uv run pytest` if the environment is already synced

## Non-goals

- Do not move, delete, or retag `v0.4.1`
- Do not bump `project.version` or `__version__`
- Do not create a GitHub release from this coding session
- Do not change document-processing behavior

## Notes

- Historical `v0.4.1` will keep installing dist `0.4.0`. That is the truth this
  PR documents rather than pretending to rewrite.
- Trust intentional issue; this plan is evidence for later review, not a human gate.
