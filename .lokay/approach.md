# Approach plan

<!-- lokay-approach source=deterministic repo=mikolaj92/Docxtor issue=164 -->

Repository: `mikolaj92/Docxtor`  
Issue: #164 — Cleanup: angielskie komunikaty DocumentError (engines/text/pdf)

## Goal

Replace Polish user strings in the library core with stable English errors (one locale in the package).

## Files likely touched

- `tests/test_documents.py`

## Test plan

- No Polish `Nie …` strings under src/docxtor/
- Tests that assert message text updated
- Behavior (exception type / chaining) unchanged
- `uv run pytest -q tests/test_documents.py`
- `rg -n 'Nie ' src/docxtor` → empty

## Non-goals

- i18n framework; Temida UI copy.

## Notes

- Trust intentional issue; this plan is evidence for later review, not a human gate.
- Coding agent may refine details but should stay on the stated goal and non-goals.
- Collector boundary: if implementation introduces unbounded collection, ship only a bounded collector patch that starts durably in the background after merge. The coding agent and lokay must not populate data or wait for collection to finish.
