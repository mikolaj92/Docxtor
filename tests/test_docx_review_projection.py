from docxtor import create_docx_from_paragraphs, project_docx_for_review


def test_projects_typed_paragraphs_without_policy() -> None:
    projection = project_docx_for_review(create_docx_from_paragraphs(("One.", "Two.")))
    assert [p.text for p in projection.paragraphs] == ["One.", "Two."]
    assert projection.coverage.value == "complete"
