# Body appendix pagination

`BodyAppendix` can opt into paragraph pagination hints when its content is added
with `append_body_appendix`:

```python
appendix = BodyAppendix(
    heading="Appendix",
    paragraphs=("Introduction", "First item", "Second item"),
    keep_paragraphs_together=True,
    keep_heading_with_next=True,
    keep_first_paragraph_with_next=True,
)
```

All three options default to `False`, preserving the existing layout.

- `keep_paragraphs_together=True` asks the word processor to keep each appendix
  paragraph's lines on one page when they fit. Each paragraph is controlled
  independently; items are not linked into one group. A paragraph taller than a
  page can still split across pages, and a fitting paragraph can move to the
  next page when it does not fit in the remaining space.
- `keep_heading_with_next=True` keeps the heading with the first appendix
  paragraph.
- `keep_first_paragraph_with_next=True` keeps the first appendix paragraph with
  the paragraph after it. It does not link later paragraphs together.

These are OOXML pagination hints; the word processor determines the final page
layout. Keeping paragraphs together can add a page or leave a short continuation
page for content that is taller than one page. A word processor may also relax a
keep-with-next link when the next paragraph is taller than the available page
area; in that case the linked paragraph can remain with its predecessor while
the oversized paragraph starts on the following page.
