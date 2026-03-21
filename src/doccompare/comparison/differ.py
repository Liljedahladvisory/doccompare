from doccompare.models import (
    ParsedDocument, DocumentElement, DiffElement, DiffSegment,
    DiffType, ElementType, ComparisonResult
)
import diff_match_patch as dmp_module


def _lcs_match(original_elements: list, modified_elements: list) -> list:
    """Match elements using LCS to find corresponding pairs."""
    n, m = len(original_elements), len(modified_elements)
    dp = [[0] * (m + 1) for _ in range(n + 1)]

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            o = original_elements[i - 1]
            mod = modified_elements[j - 1]
            if (o.element_type == mod.element_type and
                    o.level == mod.level and
                    _similarity(o.plain_text, mod.plain_text) > 0.3):
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    # Backtrack
    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        o = original_elements[i - 1]
        mod = modified_elements[j - 1]
        if (o.element_type == mod.element_type and
                o.level == mod.level and
                _similarity(o.plain_text, mod.plain_text) > 0.3 and
                dp[i][j] == dp[i - 1][j - 1] + 1):
            matches.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] > dp[i][j - 1]:
            i -= 1
        else:
            j -= 1

    matches.reverse()
    return matches


def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    from rapidfuzz import fuzz
    return fuzz.ratio(a, b) / 100.0


def _diff_chars(original: str, modified: str) -> list:
    """Character-level diff using diff-match-patch. No semantic cleanup to preserve exact chars."""
    dmp = dmp_module.diff_match_patch()
    diffs = dmp.diff_main(original, modified)
    # Do NOT call diff_cleanupSemantic — it expands char-level diffs to word-level
    return diffs


def _element_to_diff(elem: DocumentElement, diff_type: DiffType) -> DiffElement:
    """Convert a whole element to a DiffElement with a single diff type."""
    # Build one segment per run to preserve inline formatting
    segments = []
    for run in elem.runs:
        if run.text:
            segments.append(DiffSegment(
                diff_type=diff_type,
                text=run.text,
                original_formatting=run.formatting,
            ))
    if not segments:
        segments = [DiffSegment(diff_type=diff_type, text=elem.plain_text)]
    return DiffElement(
        element_type=elem.element_type,
        level=elem.level,
        segments=segments,
        diff_type=diff_type,
    )


def _build_run_intervals(elem: DocumentElement) -> list:
    """Returns [(start, end, formatting)] for each run, by character position."""
    intervals = []
    pos = 0
    for run in elem.runs:
        if run.text:
            intervals.append((pos, pos + len(run.text), run.formatting))
            pos += len(run.text)
    return intervals


def _fmt_at(abs_pos: int, intervals: list) -> set:
    """Return formatting at absolute character position."""
    for start, end, fmt in intervals:
        if start <= abs_pos < end:
            return fmt
    return intervals[-1][2] if intervals else set()


def _split_with_fmt(text: str, text_start: int, intervals: list, diff_type: DiffType) -> list:
    """Split text into DiffSegments at run boundaries, preserving exact formatting."""
    segments = []
    offset = 0
    while offset < len(text):
        abs_pos = text_start + offset
        fmt = _fmt_at(abs_pos, intervals)
        # Find where this run ends
        run_end_abs = text_start + len(text)  # fallback: rest of text
        for start, end, _ in intervals:
            if start <= abs_pos < end:
                run_end_abs = end
                break
        chunk_end = min(run_end_abs - text_start, len(text))
        chunk = text[offset:chunk_end]
        if not chunk:
            break
        segments.append(DiffSegment(diff_type=diff_type, text=chunk, original_formatting=fmt))
        offset = chunk_end
    if not segments:
        segments = [DiffSegment(diff_type=diff_type, text=text)]
    return segments


def _diff_matched_elements(orig: DocumentElement, mod: DocumentElement) -> DiffElement:
    """Character-level diff preserving exact per-run formatting from each document."""
    orig_text = orig.plain_text
    mod_text = mod.plain_text
    orig_intervals = _build_run_intervals(orig)
    mod_intervals = _build_run_intervals(mod)

    if orig_text == mod_text:
        # Identical text: emit one segment per run of the modified document
        segments = []
        for start, end, fmt in mod_intervals:
            chunk = mod_text[start:end]
            if chunk:
                segments.append(DiffSegment(diff_type=DiffType.UNCHANGED, text=chunk, original_formatting=fmt))
        if not segments:
            segments = [DiffSegment(diff_type=DiffType.UNCHANGED, text=orig_text)]
        return DiffElement(
            element_type=orig.element_type,
            level=orig.level,
            segments=segments,
            diff_type=DiffType.UNCHANGED,
        )

    raw_diffs = _diff_chars(orig_text, mod_text)
    segments = []
    has_changes = False
    orig_pos = 0
    mod_pos = 0

    for op, text in raw_diffs:
        if op == 0:   # unchanged — use mod formatting
            segments.extend(_split_with_fmt(text, mod_pos, mod_intervals, DiffType.UNCHANGED))
            orig_pos += len(text)
            mod_pos += len(text)
        elif op == 1:  # added — use mod formatting
            segments.extend(_split_with_fmt(text, mod_pos, mod_intervals, DiffType.ADDED))
            mod_pos += len(text)
            has_changes = True
        elif op == -1:  # deleted — use orig formatting
            segments.extend(_split_with_fmt(text, orig_pos, orig_intervals, DiffType.DELETED))
            orig_pos += len(text)
            has_changes = True

    diff_type = DiffType.MODIFIED if has_changes else DiffType.UNCHANGED
    return DiffElement(
        element_type=orig.element_type,
        level=orig.level,
        segments=segments,
        diff_type=diff_type,
    )


class Differ:
    def compare(self, original: ParsedDocument, modified: ParsedDocument) -> ComparisonResult:
        orig_elems = original.elements
        mod_elems = modified.elements

        matches = _lcs_match(orig_elems, mod_elems)
        matched_orig = {i for i, _ in matches}
        matched_mod = {j for _, j in matches}

        # Build ordered result
        diff_elements = []
        match_idx = 0
        orig_ptr = 0
        mod_ptr = 0

        while match_idx < len(matches):
            orig_i, mod_j = matches[match_idx]

            # Emit deleted elements before this match
            while orig_ptr < orig_i:
                diff_elements.append(_element_to_diff(orig_elems[orig_ptr], DiffType.DELETED))
                orig_ptr += 1

            # Emit added elements before this match
            while mod_ptr < mod_j:
                diff_elements.append(_element_to_diff(mod_elems[mod_ptr], DiffType.ADDED))
                mod_ptr += 1

            # Emit the matched pair
            diff_elements.append(_diff_matched_elements(orig_elems[orig_i], mod_elems[mod_j]))
            orig_ptr = orig_i + 1
            mod_ptr = mod_j + 1
            match_idx += 1

        # Remaining elements
        while orig_ptr < len(orig_elems):
            diff_elements.append(_element_to_diff(orig_elems[orig_ptr], DiffType.DELETED))
            orig_ptr += 1
        while mod_ptr < len(mod_elems):
            diff_elements.append(_element_to_diff(mod_elems[mod_ptr], DiffType.ADDED))
            mod_ptr += 1

        summary = self._compute_summary(diff_elements)
        return ComparisonResult(diff_elements=diff_elements, summary=summary)

    def _compute_summary(self, diff_elements: list) -> dict:
        added = deleted = moved = modified = unchanged = 0
        for elem in diff_elements:
            for seg in elem.segments:
                if seg.diff_type == DiffType.ADDED:
                    added += len(seg.text.split())
                elif seg.diff_type == DiffType.DELETED:
                    deleted += len(seg.text.split())
                elif seg.diff_type in (DiffType.MOVED_FROM, DiffType.MOVED_TO):
                    moved += len(seg.text.split())
                elif seg.diff_type == DiffType.MODIFIED:
                    modified += len(seg.text.split())
                elif seg.diff_type == DiffType.UNCHANGED:
                    unchanged += len(seg.text.split())
        return {
            "added_words": added,
            "deleted_words": deleted,
            "moved_words": moved,
            "modified_words": modified,
            "unchanged_words": unchanged,
        }
