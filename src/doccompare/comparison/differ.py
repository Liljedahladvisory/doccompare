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
    """Weighted LCS: match elements maximising total similarity, not just count.

    Standard LCS treats all matchable pairs as equal (weight 1), so it may pair
    a 95%-similar paragraph with a 65%-similar one if that yields more total
    matches.  Weighted LCS uses the actual similarity score as edge weight,
    so a near-perfect match is never sacrificed for quantity.
    """
    n, m = len(original_elements), len(modified_elements)

    # Pre-compute similarity scores (avoids redundant calls during backtrack)
    sim_cache: dict = {}
    for i in range(n):
        for j in range(m):
            if _can_match(original_elements[i], modified_elements[j]):
                sim_cache[(i, j)] = _similarity(
                    original_elements[i].plain_text,
                    modified_elements[j].plain_text,
                )

    # DP with float weights
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            sim = sim_cache.get((i - 1, j - 1))
            if sim is not None:
                dp[i][j] = max(dp[i - 1][j - 1] + sim, dp[i - 1][j], dp[i][j - 1])
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    # Backtrack
    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        sim = sim_cache.get((i - 1, j - 1))
        if sim is not None and abs(dp[i][j] - (dp[i - 1][j - 1] + sim)) < 1e-9:
            matches.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] > dp[i][j - 1] + 1e-9:
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

    For long texts that exceed the token-LCS cap, falls back to
    diff_match_patch directly (efficient Myers diff on the full strings).
    """
    from rapidfuzz import fuzz

    orig_tokens = _tokenize(original)
    mod_tokens = _tokenize(modified)

    # For long texts, skip token-level LCS entirely and use diff_match_patch
    # which implements an efficient O(ND) Myers algorithm
    if len(orig_tokens) * len(mod_tokens) > 40000:
        dmp = dmp_module.diff_match_patch()
        diffs = dmp.diff_main(original, modified)
        dmp.diff_cleanupSemantic(diffs)
        return diffs

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


def _copy_para_fmt(src: DocumentElement, dest: DiffElement):
    """Copy paragraph-level formatting from a DocumentElement to a DiffElement."""
    dest.alignment = src.alignment
    dest.left_indent_pt = src.left_indent_pt
    dest.right_indent_pt = src.right_indent_pt
    dest.first_line_indent_pt = src.first_line_indent_pt
    dest.space_before_pt = src.space_before_pt
    dest.space_after_pt = src.space_after_pt
    dest.line_spacing = src.line_spacing


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
                font_name=run.font_name,
            ))
    if not segments:
        segments = [DiffSegment(diff_type=diff_type, text=elem.plain_text)]
    de = DiffElement(
        element_type=elem.element_type,
        level=elem.level,
        segments=segments,
        diff_type=diff_type,
        list_style=elem.list_style,
        list_numid=elem.list_numid,
        list_lvl_text=elem.list_lvl_text,
    )
    _copy_para_fmt(elem, de)
    return de


def _build_run_intervals(elem: DocumentElement) -> list:
    """Returns [(start, end, formatting, font_size, font_name)] for each run, by character position."""
    intervals = []
    pos = 0
    for run in elem.runs:
        if run.text:
            intervals.append((pos, pos + len(run.text), run.formatting, run.font_size, run.font_name))
            pos += len(run.text)
    return intervals


def _run_at(abs_pos: int, intervals: list):
    """Return (formatting, font_size, font_name) at absolute character position."""
    for start, end, fmt, fs, fn in intervals:
        if start <= abs_pos < end:
            return fmt, fs, fn
    if intervals:
        return intervals[-1][2], intervals[-1][3], intervals[-1][4]
    return set(), None, None


def _split_with_fmt(text: str, text_start: int, intervals: list, diff_type: DiffType) -> list:
    """Split text into DiffSegments at run boundaries, preserving exact formatting."""
    segments = []
    offset = 0
    while offset < len(text):
        abs_pos = text_start + offset
        fmt, font_size, font_name = _run_at(abs_pos, intervals)
        # Find where this run ends
        run_end_abs = text_start + len(text)  # fallback: rest of text
        for start, end, *_ in intervals:
            if start <= abs_pos < end:
                run_end_abs = end
                break
        chunk_end = min(run_end_abs - text_start, len(text))
        chunk = text[offset:chunk_end]
        if not chunk:
            break
        segments.append(DiffSegment(diff_type=diff_type, text=chunk, original_formatting=fmt, font_size=font_size, font_name=font_name))
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
        for start, end, fmt, font_size, font_name in mod_intervals:
            chunk = mod_text[start:end]
            if chunk:
                segments.append(DiffSegment(diff_type=DiffType.UNCHANGED, text=chunk, original_formatting=fmt, font_size=font_size, font_name=font_name))
        if not segments:
            segments = [DiffSegment(diff_type=DiffType.UNCHANGED, text=orig_text)]
        de = DiffElement(
            element_type=orig.element_type,
            level=orig.level,
            segments=segments,
            diff_type=DiffType.UNCHANGED,
            list_style=mod.list_style,
            list_numid=mod.list_numid,
            list_lvl_text=mod.list_lvl_text,
        )
        _copy_para_fmt(mod, de)
        return de

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
    de = DiffElement(
        element_type=orig.element_type,
        level=orig.level,
        segments=segments,
        diff_type=diff_type,
        list_style=mod.list_style,
        list_numid=mod.list_numid,
        list_lvl_text=mod.list_lvl_text,
    )
    _copy_para_fmt(mod, de)
    return de


def _seg_plain_text(elem: DiffElement) -> str:
    """Extract plain text from a DiffElement's segments."""
    return "".join(seg.text for seg in elem.segments)


def _find_source_element(diff_elem: DiffElement, source_elems: list) -> "DocumentElement | None":
    """Find the DocumentElement in source_elems that produced this DiffElement."""
    target_text = _seg_plain_text(diff_elem)
    # Exact match first
    for elem in source_elems:
        if elem.plain_text == target_text and elem.element_type.value == diff_elem.element_type.value:
            return elem
    # Fallback: closest match
    best, best_sim = None, 0.0
    for elem in source_elems:
        if elem.element_type.value == diff_elem.element_type.value:
            sim = _similarity(elem.plain_text, target_text)
            if sim > best_sim:
                best_sim = sim
                best = elem
    return best


def _global_rematch(diff_elements: list, orig_elems: list, mod_elems: list) -> list:
    """Global post-processing: find ALL unmatched DELETE/ADD elements and re-pair similar ones.

    Unlike adjacent-only matching, this scans the entire list so it catches pairs
    separated by matched (UNCHANGED/MODIFIED) elements in between.

    Strategy:
    1. Collect indices of every DELETED and every ADDED element.
    2. Compute pairwise similarity.
    3. Greedily pair highest-similarity matches above threshold.
    4. For each pair: remove the DELETED element, replace the ADDED element
       with an inline diff (preserving modified-document position).
    """
    THRESHOLD = 0.5

    # Step 1: collect indices
    del_indices = [i for i, e in enumerate(diff_elements) if e.diff_type == DiffType.DELETED]
    add_indices = [i for i, e in enumerate(diff_elements) if e.diff_type == DiffType.ADDED]

    if not del_indices or not add_indices:
        return diff_elements

    # Step 2: compute pairwise similarity and build candidate list
    candidates = []  # (similarity, del_list_pos, add_list_pos)
    del_texts = [_seg_plain_text(diff_elements[i]) for i in del_indices]
    add_texts = [_seg_plain_text(diff_elements[i]) for i in add_indices]

    for di, d_text in enumerate(del_texts):
        for ai, a_text in enumerate(add_texts):
            sim = _similarity(d_text, a_text)
            if sim > THRESHOLD:
                candidates.append((sim, di, ai))

    # Step 3: greedy match — highest similarity first
    candidates.sort(key=lambda x: x[0], reverse=True)
    matched_del = set()  # positions in del_indices
    matched_add = set()  # positions in add_indices
    pairs = []  # (del_elem_index, add_elem_index)

    for sim, di, ai in candidates:
        if di in matched_del or ai in matched_add:
            continue
        pairs.append((del_indices[di], add_indices[ai]))
        matched_del.add(di)
        matched_add.add(ai)

    if not pairs:
        return diff_elements

    # Step 4: build result — remove matched DELETEs, replace matched ADDs with inline diffs
    del_to_remove = {d for d, _ in pairs}
    add_to_replace = {}  # add_index -> (orig_doc_elem, mod_doc_elem)
    for d_idx, a_idx in pairs:
        orig_doc = _find_source_element(diff_elements[d_idx], orig_elems)
        mod_doc = _find_source_element(diff_elements[a_idx], mod_elems)
        if orig_doc and mod_doc:
            add_to_replace[a_idx] = (orig_doc, mod_doc)
        else:
            # Can't find source elements — undo this pair
            del_to_remove.discard(d_idx)

    result = []
    for i, elem in enumerate(diff_elements):
        if i in del_to_remove:
            continue  # skip — will be shown as inline diff at the ADD position
        if i in add_to_replace:
            orig_doc, mod_doc = add_to_replace[i]
            result.append(_diff_matched_elements(orig_doc, mod_doc))
        else:
            result.append(elem)

    return result


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

        # Post-processing: re-match unmatched DELETE/ADD pairs as inline diffs
        diff_elements = _global_rematch(diff_elements, orig_elems, mod_elems)

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
