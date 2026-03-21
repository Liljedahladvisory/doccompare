import pytest
from doccompare.models import ParsedDocument, DocumentElement, TextRun, ElementType, DiffType
from doccompare.comparison.differ import Differ


def make_doc(*texts):
    elements = []
    for i, text in enumerate(texts):
        elem = DocumentElement(
            element_type=ElementType.PARAGRAPH,
            runs=[TextRun(text=text)],
            element_id=f"p_{i}",
        )
        elements.append(elem)
    return ParsedDocument(elements=elements)


def test_identical_documents():
    doc = make_doc("Hello world.", "Second paragraph.")
    differ = Differ()
    result = differ.compare(doc, doc)
    for elem in result.diff_elements:
        for seg in elem.segments:
            assert seg.diff_type == DiffType.UNCHANGED


def test_detects_added_paragraph():
    orig = make_doc("First.", "Third.")
    mod = make_doc("First.", "Second.", "Third.")
    differ = Differ()
    result = differ.compare(orig, mod)
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.ADDED in all_types


def test_detects_deleted_paragraph():
    orig = make_doc("First.", "Second.", "Third.")
    mod = make_doc("First.", "Third.")
    differ = Differ()
    result = differ.compare(orig, mod)
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.DELETED in all_types


def test_word_level_diff_simple_change():
    orig = make_doc("The quick brown fox.")
    mod = make_doc("The quick red fox.")
    differ = Differ()
    result = differ.compare(orig, mod)
    assert len(result.diff_elements) >= 1
    segs = result.diff_elements[0].segments
    types = [s.diff_type for s in segs]
    assert DiffType.DELETED in types or DiffType.ADDED in types


def test_empty_documents():
    orig = ParsedDocument(elements=[])
    mod = ParsedDocument(elements=[])
    differ = Differ()
    result = differ.compare(orig, mod)
    assert result.diff_elements == []


def test_completely_different_documents():
    orig = make_doc("Apple banana cherry.")
    mod = make_doc("Zebra yacht xylophone.")
    differ = Differ()
    result = differ.compare(orig, mod)
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.ADDED in all_types or DiffType.DELETED in all_types
