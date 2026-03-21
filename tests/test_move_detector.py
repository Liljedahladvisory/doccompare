import pytest
from doccompare.models import (
    ParsedDocument, DocumentElement, TextRun, ElementType,
    DiffType, DiffSegment, DiffElement, ComparisonResult
)
from doccompare.comparison.differ import Differ
from doccompare.comparison.move_detector import MoveDetector


def make_comparison_result(deleted_text: str, added_text: str) -> ComparisonResult:
    del_seg = DiffSegment(diff_type=DiffType.DELETED, text=deleted_text)
    add_seg = DiffSegment(diff_type=DiffType.ADDED, text=added_text)
    del_elem = DiffElement(element_type=ElementType.PARAGRAPH, segments=[del_seg], diff_type=DiffType.DELETED)
    add_elem = DiffElement(element_type=ElementType.PARAGRAPH, segments=[add_seg], diff_type=DiffType.ADDED)
    return ComparisonResult(diff_elements=[del_elem, add_elem], summary={})


def test_detects_moved_paragraph():
    long_text = "This is a long paragraph that has been moved from one location to another in the document."
    result = make_comparison_result(long_text, long_text)
    detector = MoveDetector(threshold=85.0)
    result = detector.detect(result)
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.MOVED_FROM in all_types
    assert DiffType.MOVED_TO in all_types


def test_respects_threshold():
    result = make_comparison_result(
        "This is completely different from anything else in the document really yes.",
        "Absolutely nothing in common with the other text at all whatsoever indeed.",
    )
    detector = MoveDetector(threshold=85.0)
    result = detector.detect(result)
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.MOVED_FROM not in all_types
    assert DiffType.MOVED_TO not in all_types


def test_no_false_positives_on_short_text():
    result = make_comparison_result("Hi.", "Hi.")
    detector = MoveDetector(threshold=85.0)
    result = detector.detect(result)
    # Short text (<= 20 chars) should not be classified as move
    all_types = [seg.diff_type for elem in result.diff_elements for seg in elem.segments]
    assert DiffType.MOVED_FROM not in all_types
