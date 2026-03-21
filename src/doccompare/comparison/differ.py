import re
from doccompare.models import (
    ParsedDocument, DocumentElement, DiffElement, DiffSegment,
    DiffType, ElementType, ComparisonResult
)
import diff_match_patch as dmp_module


_TEXT_TYPES = {ElementType.PARAGRAPH, ElementType.LIST_ITEM}


def _can_match(o: DocumentElement, mod: DocumentElement) -> bool:
    """Decide if two elements are candidates for matching.

    Rules:
    - Headings: must share the same type and level; similarity > 0.4
    - Text elements (paragraph / list item): type may differ freely;
      similarity > 0.6 (high bar prevents matching different clauses
      that share legal boilerplate)
    - Table rows: same type; similarity > 0.4
    """
    sim = _similarity(o.plain_text, mod.plain_text)
    if o.element_type == ElementType.HEADING or mod.element_type == ElementType.HEADING:
        return (o.element_type == mod.element_type
                and o.level == mod.level
                and sim > 0.4)
    if o.element_type in _TEXT_TYPES and mod.element_type in _TEXT_TYPES:
        return sim > 0.6
    return o.element_type == mod.element_type and sim > 0.4


def _lcs_match(original_elements: list, modified_elements: list) -> list:
    """Match elements using LCS to find corresponding pairs."""
    n, m = len(original_elements), len(modified_elements)
    dp = [[0] * (m + 1) for _ in range(n + 1)]

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if _can_match(original_elements[i - 1], modified_elements[j - 1]):
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    # Backtrack
    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        if (_can_match(original_elements[i - 1], modified_elements[j - 1])
                and dp[i][j] == dp[i - 1][j - 1] + 1):
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


# Minimum similarity (0–100) between two words to apply char-level sub-diff.
# Below this threshold, the words are shown as full delete + insert.
_CHAR_DIFF_THRESHOLD = 50


def _tokenize(text: str) -> list:
    """Split text into alternating word and non-word (space/punct) tokens."""
    return re.findall(r'\w+|[^\w]+', text)


def _lcs_token_ops(a: list, b: list) -> list:
    """LCS diff on token lists. Returns [(op, token)] where op is -1/0/1."""
    n, m = len(a), len(b)
    # Cap to avoid O(n²) slowdown on very long paragraphs
    if n * m > 40000:
        return [(-1, t) for t in a] + [(1, t) for t in b]
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if a[i - 1] == b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])
    ops = []
    i, j = n, m
    while i > 0 or j > 0:
        if i > 0 and j > 0 and a[i - 1] == b[j - 1]:
            ops.append((0, a[i - 1]))
            i -= 1; j -= 1
        elif j > 0 and (i == 0 or dp[i][j - 1] >= dp[i - 1][j]):
            ops.append((1, b[j - 1]))
            j -= 1
        else:
            ops.append((-1, a[i - 1]))
            i -= 1
    ops.reverse()
    return ops


def _diff_hybrid(original: str, modified: str) -> list:
    """
    Hybrid diff: word-level structure with char-level precision inside changed word pairs.

    - Unchanged tokens: EQUAL
    - Changed whitespace/punctuation: DELETE old + INSERT new
    - Changed word pair with similarity >= threshold: char-level sub-diff
    - Changed word pair below threshold: full word DELETE + INSERT
    """
    from rapidfuzz import fuzz

    orig_tokens = _tokenize(original)
    mod_tokens = _tokenize(modified)
    token_ops = _lcs_token_ops(orig_tokens, mod_tokens)

    result = []
    i = 0
    while i < len(token_ops):
        op, tok = token_ops[i]
        # Look for a DELETE immediately followed by an INSERT
        if op == -1 and i + 1 < len(token_ops) and token_ops[i + 1][0] == 1:
            del_tok = tok
            ins_tok = token_ops[i + 1][1]
            # Char-level sub-diff only for actual word tokens that are similar enough
            if del_tok.strip() and ins_tok.strip() and fuzz.ratio(del_tok, ins_tok) >= _CHAR_DIFF_THRESHOLD:
                dmp = dmp_module.diff_match_patch()
                result.extend(dmp.diff_main(del_tok, ins_tok))
            else:
                result.append((-1, del_tok))
                result.append((1, ins_tok))
            i += 2
        else:
            result.append((op, tok))
            i += 1
    return result


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
                font_size=run.font_size,
            ))
    if not segments:
        segments = [DiffSegment(diff_type=diff_type, text=elem.plain_text)]
    return DiffElement(
        element_type=elem.element_type,
        level=elem.level,
        segments=segments,
        diff_type=diff_type,
        list_style=elem.list_style,
        list_numid=elem.list_numid,
        list_lvl_text=elem.list_lvl_text,
    )


def _build_run_intervals(elem: DocumentElement) -> list:
    """Returns [(start, end, formatting, font_size)] for each run, by character position."""
    intervals = []
    pos = 0
    for run in elem.runs:
        if run.text:
            intervals.append((pos, pos + len(run.text), run.formatting, run.font_size))
            pos += len(run.text)
    return intervals


def _run_at(abs_pos: int, intervals: list):
    """Return (formatting, font_size) at absolute character position."""
    for start, end, fmt, fs in intervals:
        if start <= abs_pos < end:
            return fmt, fs
    if intervals:
        return intervals[-1][2], intervals[-1][3]
    return set(), None


def _split_with_fmt(text: str, text_start: int, intervals: list, diff_type: DiffType) -> list:
    """Split text into DiffSegments at run boundaries, preserving exact formatting."""
    segments = []
    offset = 0
    while offset < len(text):
        abs_pos = text_start + offset
        fmt, font_size = _run_at(abs_pos, intervals)
        # Find where this run ends
        run_end_abs = text_start + len(text)  # fallback: rest of text
        for start, end, _, _fs in intervals:
            if start <= abs_pos < end:
                run_end_abs = end
                break
        chunk_end = min(run_end_abs - text_start, len(text))
        chunk = text[offset:chunk_end]
        if not chunk:
            break
        segments.append(DiffSegment(diff_type=diff_type, text=chunk, original_formatting=fmt, font_size=font_size))
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
        for start, end, fmt, font_size in mod_intervals:
            chunk = mod_text[start:end]
            if chunk:
                segments.append(DiffSegment(diff_type=DiffType.UNCHANGED, text=chunk, original_formatting=fmt, font_size=font_size))
        if not segments:
            segments = [DiffSegment(diff_type=DiffType.UNCHANGED, text=orig_text)]
        return DiffElement(
            element_type=orig.element_type,
            level=orig.level,
            segments=segments,
            diff_type=DiffType.UNCHANGED,
            list_style=mod.list_style,
            list_numid=mod.list_numid,
        list_lvl_text=mod.list_lvl_text,
        )

    raw_diffs = _diff_hybrid(orig_text, mod_text)
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
        list_style=mod.list_style,
        list_numid=mod.list_numid,
        list_lvl_text=mod.list_lvl_text,
    )


def _rematch_adjacent(diff_elements: list, orig_elems: list, mod_elems: list) -> list:
    """Post-processing: find adjacent DELETE+ADD blocks and re-match similar pairs as inline diffs.

    The LCS algorithm sometimes misses locally optimal matches because it optimizes globally.
    This pass finds contiguous runs of DELETED elements followed by ADDED elements, computes
    pairwise similarity, and converts high-similarity pairs into inline diffs.
    """
    result = []
    i = 0
    while i < len(diff_elements):
        # Collect a contiguous block of DELETED elements
        deleted_block = []
        while i < len(diff_elements) and diff_elements[i].diff_type == DiffType.DELETED:
            deleted_block.append(diff_elements[i])
            i += 1

        # Collect a contiguous block of ADDED elements right after
        added_block = []
        while i < len(diff_elements) and diff_elements[i].diff_type == DiffType.ADDED:
            added_block.append(diff_elements[i])
            i += 1

        if not deleted_block and not added_block:
            # Not a delete/add block — just pass through
            result.append(diff_elements[i])
            i += 1
            continue

        if not deleted_block or not added_block:
            # Only deletes or only adds — no re-matching possible
            result.extend(deleted_block)
            result.extend(added_block)
            continue

        # Try to match deleted elements with added elements by similarity
        # Find the original DocumentElements for proper re-diffing
        _rematch_blocks(deleted_block, added_block, result, orig_elems, mod_elems)

    return result


def _seg_plain_text(elem: DiffElement) -> str:
    """Extract plain text from a DiffElement's segments."""
    return "".join(seg.text for seg in elem.segments)


def _find_orig_element(diff_elem: DiffElement, orig_elems: list) -> "DocumentElement | None":
    """Find the original DocumentElement that corresponds to a deleted DiffElement."""
    target_text = _seg_plain_text(diff_elem)
    for elem in orig_elems:
        if elem.plain_text == target_text and elem.element_type.value == diff_elem.element_type.value:
            return elem
    # Fallback: closest match
    best, best_sim = None, 0.0
    for elem in orig_elems:
        if elem.element_type.value == diff_elem.element_type.value:
            sim = _similarity(elem.plain_text, target_text)
            if sim > best_sim:
                best_sim = sim
                best = elem
    return best


def _find_mod_element(diff_elem: DiffElement, mod_elems: list) -> "DocumentElement | None":
    """Find the modified DocumentElement that corresponds to an added DiffElement."""
    target_text = _seg_plain_text(diff_elem)
    for elem in mod_elems:
        if elem.plain_text == target_text and elem.element_type.value == diff_elem.element_type.value:
            return elem
    best, best_sim = None, 0.0
    for elem in mod_elems:
        if elem.element_type.value == diff_elem.element_type.value:
            sim = _similarity(elem.plain_text, target_text)
            if sim > best_sim:
                best_sim = sim
                best = elem
    return best


def _rematch_blocks(deleted_block: list, added_block: list, result: list,
                    orig_elems: list, mod_elems: list):
    """Match deleted and added elements by similarity and produce inline diffs for good matches."""
    n_del = len(deleted_block)
    n_add = len(added_block)

    # Compute similarity matrix
    sim_matrix = []
    for d_idx in range(n_del):
        row = []
        d_text = _seg_plain_text(deleted_block[d_idx])
        for a_idx in range(n_add):
            a_text = _seg_plain_text(added_block[a_idx])
            row.append(_similarity(d_text, a_text))
        sim_matrix.append(row)

    # Greedy matching: repeatedly pick the highest-similarity pair above threshold
    THRESHOLD = 0.5
    matched_del = set()
    matched_add = set()
    pairs = []  # (del_idx, add_idx)

    while True:
        best_sim = THRESHOLD
        best_pair = None
        for d_idx in range(n_del):
            if d_idx in matched_del:
                continue
            for a_idx in range(n_add):
                if a_idx in matched_add:
                    continue
                if sim_matrix[d_idx][a_idx] > best_sim:
                    best_sim = sim_matrix[d_idx][a_idx]
                    best_pair = (d_idx, a_idx)
        if best_pair is None:
            break
        pairs.append(best_pair)
        matched_del.add(best_pair[0])
        matched_add.add(best_pair[1])

    # Emit results in order of the added block (modified document order)
    # First: unmatched deletes
    for d_idx in range(n_del):
        if d_idx not in matched_del:
            result.append(deleted_block[d_idx])

    # Then: added elements in order, replacing matched ones with inline diffs
    for a_idx in range(n_add):
        if a_idx not in matched_add:
            result.append(added_block[a_idx])
        else:
            # Find the corresponding delete
            d_idx = next(d for d, a in pairs if a == a_idx)
            # Get original DocumentElements for proper formatting-aware diff
            orig_doc_elem = _find_orig_element(deleted_block[d_idx], orig_elems)
            mod_doc_elem = _find_mod_element(added_block[a_idx], mod_elems)
            if orig_doc_elem and mod_doc_elem:
                result.append(_diff_matched_elements(orig_doc_elem, mod_doc_elem))
            else:
                # Fallback: emit as-is
                result.append(deleted_block[d_idx])
                result.append(added_block[a_idx])


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

        # Post-processing: re-match adjacent DELETE+ADD blocks as inline diffs
        diff_elements = _rematch_adjacent(diff_elements, orig_elems, mod_elems)

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
