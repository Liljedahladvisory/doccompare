import pytest
import html as html_module
from pathlib import Path
from doccompare.models import (
    ComparisonResult, DiffElement, DiffSegment, DiffType, ElementType
)
from doccompare.rendering.html_builder import HtmlBuilder


def make_simple_result():
    seg1 = DiffSegment(diff_type=DiffType.UNCHANGED, text="Hello ")
    seg2 = DiffSegment(diff_type=DiffType.ADDED, text="world")
    seg3 = DiffSegment(diff_type=DiffType.DELETED, text="earth")
    elem = DiffElement(
        element_type=ElementType.PARAGRAPH,
        segments=[seg1, seg2, seg3],
        diff_type=DiffType.MODIFIED,
    )
    return ComparisonResult(
        diff_elements=[elem],
        summary={"added_words": 1, "deleted_words": 1, "moved_words": 0, "unchanged_words": 1},
    )


def test_html_output_contains_correct_css_classes():
    builder = HtmlBuilder()
    result = make_simple_result()
    html_out = builder.build(result, Path("original.docx"), Path("modified.docx"))
    assert 'class="added"' in html_out
    assert 'class="deleted"' in html_out


def test_special_characters_are_escaped():
    seg = DiffSegment(diff_type=DiffType.ADDED, text="<script>alert('xss')</script>")
    elem = DiffElement(element_type=ElementType.PARAGRAPH, segments=[seg], diff_type=DiffType.ADDED)
    result = ComparisonResult(diff_elements=[elem], summary={})
    builder = HtmlBuilder()
    html_out = builder.build(result, Path("a.docx"), Path("b.docx"))
    assert "<script>" not in html_out
    assert "&lt;script&gt;" in html_out


def test_pdf_renders_without_error(tmp_path):
    from doccompare.rendering.pdf_renderer import render_pdf
    css_path = Path(__file__).parent.parent / "src" / "doccompare" / "rendering" / "styles.css"
    if not css_path.exists():
        pytest.skip("CSS file not found")
    html_content = "<html><body><p>Test</p></body></html>"
    output = tmp_path / "test.pdf"
    render_pdf(html_content, css_path, output)
    assert output.exists()
    assert output.stat().st_size > 0
