# Approach plan

<!-- lokay-approach source=deterministic repo=mikolaj92/Docxtor issue=32 -->

Repository: `mikolaj92/Docxtor`  
Issue: #32 — Make nested revision and hyperlink text write-addressable without unwrapping OOXML

## Goal

- mikolaj92/Temida#5076

## Files likely touched

- (infer from repo inspection)

## Test plan

- Text inside `w:ins`, `w:del` / `w:delText`, and inline `w:hyperlink` is returned with a stable container/span identity and a typed mechanical role.
- Partial and whole replacements address the exact extracted characters and preserve surrounding wrapper elements, `w:rPr`, revision author/date/ID, hyperlink anchor/relationship, and document order.
- Internal table-of-contents hyperlinks remain functional after replacing their display text.
- The contract works across supported body, table, header, footer, and text-box story parts.
- Unsupported structural revision forms fail closed without writing a partial output.
- Public API documentation and model-only fixtures cover extraction, replacement, round-trip structure, and untouched-document behavior.
- No review, anonymization, Dike, Posejdon, or Polish-law semantics enter Docxtor.

## Non-goals

- (none stated)

## Notes

- Trust intentional issue; this plan is evidence for later review, not a human gate.
- Coding agent may refine details but should stay on the stated goal and non-goals.
- Collector boundary: if implementation introduces unbounded collection, ship only a bounded collector patch that starts durably in the background after merge. The coding agent and mill must not populate data or wait for collection to finish.
- No explicit file paths in issue; infer from repo inspection.
