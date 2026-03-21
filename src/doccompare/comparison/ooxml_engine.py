"""OOXML-native document comparison engine producing Track Changes output.

Compares two .docx files by manipulating the underlying Office Open XML tree
directly, producing a .docx with standard Microsoft Word revision markup
(<w:ins>, <w:del>).

Architecture:
  1. Use the newer document as the technical baseline
  2. Normalize runs (merge adjacent runs with identical rPr)
  3. Match block-level elements between old and new via weighted LCS
  4. Diff matched paragraphs at character level (diff_match_patch)
  5. Inject w:ins / w:del markup into the baseline XML
  6. Write the modified document as a valid .docx
"""

import copy
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree
from rapidfuzz import fuzz
import diff_match_patch as dmp_module

# ── OOXML namespaces ────────────────────────────────────────────────────
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _qn(tag: str) -> str:
    """'{W_NS}p' from 'w:p', etc."""
    prefix, local = tag.split(":")
    ns = {"w": W_NS, "xml": XML_NS}[prefix]
    return f"{{{ns}}}{local}"


# Pre-computed qualified names
W_P = _qn("w:p")
W_R = _qn("w:r")
W_T = _qn("w:t")
W_RPR = _qn("w:rPr")
W_PPR = _qn("w:pPr")
W_TBL = _qn("w:tbl")
W_TR = _qn("w:tr")
W_TC = _qn("w:tc")
W_INS = _qn("w:ins")
W_DEL = _qn("w:del")
W_DEL_TEXT = _qn("w:delText")
W_BODY = _qn("w:body")
W_SECT_PR = _qn("w:sectPr")
W_SDT = _qn("w:sdt")
W_SDT_CONTENT = _qn("w:sdtContent")
W_HYPERLINK = _qn("w:hyperlink")
W_BOOKMARK_START = _qn("w:bookmarkStart")
W_BOOKMARK_END = _qn("w:bookmarkEnd")


# ── Revision-ID generator ──────────────────────────────────────────────
class _RevId:
    def __init__(self, start=1):
        self._n = start

    def next(self) -> str:
        val = self._n
        self._n += 1
        return str(val)


def _rev_attrs(rid: _RevId, author: str, date: str) -> dict:
    return {
        _qn("w:id"): rid.next(),
        _qn("w:author"): author,
        _qn("w:date"): date,
    }


# ── Public API ──────────────────────────────────────────────────────────
def compare(
    old_path,
    new_path,
    output_path,
    author: str = "DocCompare",
) -> Path:
    """Compare two .docx files and produce a .docx with Track Changes.

    Uses the newer document as the baseline and injects deletions from
    the older document.
    """
    old_path = Path(old_path)
    new_path = Path(new_path)
    output_path = Path(output_path) if output_path else None
    date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rid = _RevId()

    # 1. Read document.xml from both
    old_tree = _read_doc_xml(old_path)
    new_tree = _read_doc_xml(new_path)
    old_body = old_tree.find(W_BODY)
    new_body = new_tree.find(W_BODY)

    # 2. Normalize runs (merge adjacent with identical rPr)
    _normalize_runs(old_body)
    _normalize_runs(new_body)

    # 3. Extract block-level elements
    old_blocks = _get_blocks(old_body)
    new_blocks = _get_blocks(new_body)

    # 4. Match blocks via weighted LCS
    matches = _match_blocks(old_blocks, new_blocks)
    matched_old = {i for i, _ in matches}
    matched_new = {j for _, j in matches}

    # 5a. Inline diff for matched paragraph pairs
    for oi, nj in matches:
        ob, nb = old_blocks[oi], new_blocks[nj]
        if ob.tag == W_P and nb.tag == W_P:
            _diff_para(ob, nb, author, date, rid)
        elif ob.tag == W_TBL and nb.tag == W_TBL:
            _diff_table(ob, nb, author, date, rid)

    # 5b. Mark unmatched new blocks as inserted
    for j in range(len(new_blocks)):
        if j not in matched_new:
            _mark_block_inserted(new_blocks[j], author, date, rid)

    # 5c. Insert deleted old blocks into new_body
    _insert_deletions(
        old_blocks, new_blocks, matches, matched_old, new_body, author, date, rid,
    )

    # 6. Compute summary from the modified XML
    summary = _compute_summary(new_tree)

    # 7. Optionally write output .docx
    if output_path:
        _write_docx(new_path, output_path, new_tree)

    return new_tree, summary


# ── I/O helpers ─────────────────────────────────────────────────────────
def _read_doc_xml(path: Path) -> etree._Element:
    with zipfile.ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def _write_docx(template: Path, output: Path, doc_tree: etree._Element):
    """Copy template .docx, replacing document.xml with modified tree."""
    xml_bytes = etree.tostring(
        doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True,
    )
    with zipfile.ZipFile(template, "r") as zin:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, xml_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))


# ── Summary computation ─────────────────────────────────────────────────
def _compute_summary(doc_tree: etree._Element) -> dict:
    """Count inserted/deleted/unchanged words from the Track Changes markup."""
    ns = {"w": W_NS}
    added_words = 0
    deleted_words = 0
    unchanged_words = 0

    body = doc_tree.find(W_BODY)
    for p in body.iter(W_P):
        for child in p:
            if child.tag == W_INS:
                for t in child.iter(W_T):
                    if t.text:
                        added_words += len(t.text.split())
            elif child.tag == W_DEL:
                for dt in child.iter(W_DEL_TEXT):
                    if dt.text:
                        deleted_words += len(dt.text.split())
            elif child.tag == W_R:
                t = child.find(W_T)
                if t is not None and t.text:
                    unchanged_words += len(t.text.split())

    return {
        "added_words": added_words,
        "deleted_words": deleted_words,
        "unchanged_words": unchanged_words,
    }


# ── Run normalization ───────────────────────────────────────────────────
def _rpr_sig(rpr) -> str:
    """Canonical signature for an rPr element (for equality comparison)."""
    if rpr is None:
        return ""
    return etree.tostring(rpr).decode()


def _normalize_runs(body: etree._Element):
    """Merge adjacent runs with identical rPr within each paragraph."""
    for para in body.iter(W_P):
        direct_runs = [c for c in para if c.tag == W_R]
        if len(direct_runs) < 2:
            continue
        i = 0
        while i < len(direct_runs) - 1:
            r1, r2 = direct_runs[i], direct_runs[i + 1]
            # Only merge truly adjacent siblings
            children = list(para)
            try:
                idx1, idx2 = children.index(r1), children.index(r2)
            except ValueError:
                i += 1
                continue
            if idx2 != idx1 + 1:
                i += 1
                continue
            if _rpr_sig(r1.find(W_RPR)) == _rpr_sig(r2.find(W_RPR)):
                t1, t2 = r1.find(W_T), r2.find(W_T)
                if t1 is not None and t2 is not None:
                    t1.text = (t1.text or "") + (t2.text or "")
                    # Preserve xml:space
                    txt = t1.text or ""
                    if txt and (txt[0] == " " or txt[-1] == " "):
                        t1.set(f"{{{XML_NS}}}space", "preserve")
                    para.remove(r2)
                    direct_runs.pop(i + 1)
                    continue
            i += 1


# ── Block extraction & text helpers ─────────────────────────────────────
def _get_blocks(body: etree._Element) -> list:
    """Get direct block-level children of body."""
    blocks = []
    for child in body:
        if child.tag in (W_P, W_TBL):
            blocks.append(child)
        elif child.tag == W_SDT:
            content = child.find(W_SDT_CONTENT)
            if content is not None:
                for sub in content:
                    if sub.tag in (W_P, W_TBL):
                        blocks.append(sub)
    return blocks


def _para_text(p: etree._Element) -> str:
    """Get plain text of a paragraph (all w:t descendants)."""
    return "".join(t.text or "" for t in p.iter(W_T))


def _block_text(b: etree._Element) -> str:
    return "".join(t.text or "" for t in b.iter(W_T))


# ── Matching ────────────────────────────────────────────────────────────
def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return fuzz.ratio(a, b) / 100.0


def _match_blocks(old_blocks: list, new_blocks: list) -> list:
    """Weighted LCS matching of block elements."""
    n, m = len(old_blocks), len(new_blocks)
    old_texts = [_block_text(b) for b in old_blocks]
    new_texts = [_block_text(b) for b in new_blocks]

    sim_cache: dict = {}
    for i in range(n):
        for j in range(m):
            if old_blocks[i].tag != new_blocks[j].tag:
                continue
            s = _similarity(old_texts[i], new_texts[j])
            if s > 0.4:
                sim_cache[(i, j)] = s

    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            s = sim_cache.get((i - 1, j - 1))
            if s is not None:
                dp[i][j] = max(dp[i - 1][j - 1] + s, dp[i - 1][j], dp[i][j - 1])
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        s = sim_cache.get((i - 1, j - 1))
        if s is not None and abs(dp[i][j] - (dp[i - 1][j - 1] + s)) < 1e-9:
            matches.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] > dp[i][j - 1] + 1e-9:
            i -= 1
        else:
            j -= 1
    matches.reverse()
    return matches


# ── Run intervals & splitting ───────────────────────────────────────────
def _collect_runs(para: etree._Element) -> list:
    """Collect all w:r elements in a paragraph in document order,
    including those nested inside w:hyperlink, w:ins, w:del, etc."""
    runs = []
    for elem in para.iter(W_R):
        runs.append(elem)
    return runs


def _run_intervals(para: etree._Element) -> list:
    """Return [(start, end, rPr_element)] for each run's text."""
    intervals = []
    pos = 0
    for r in _collect_runs(para):
        t = r.find(W_T)
        text = (t.text or "") if t is not None else ""
        rpr = r.find(W_RPR)
        if text:
            intervals.append((pos, pos + len(text), rpr))
            pos += len(text)
    return intervals


def _rpr_at(pos: int, intervals: list):
    """Return rPr element at character position, or None."""
    for s, e, rpr in intervals:
        if s <= pos < e:
            return rpr
    return intervals[-1][2] if intervals else None


def _split_by_runs(text: str, start: int, intervals: list):
    """Yield (chunk_text, rPr) for text split at run boundaries."""
    if not intervals:
        yield text, None
        return
    offset = 0
    while offset < len(text):
        apos = start + offset
        rpr = _rpr_at(apos, intervals)
        run_end = start + len(text)  # fallback
        for s, e, _ in intervals:
            if s <= apos < e:
                run_end = e
                break
        end = min(run_end - start, len(text))
        chunk = text[offset:end]
        if not chunk:
            break
        yield chunk, rpr
        offset = end


# ── XML element construction ───────────────────────────────────────────
def _make_run(text: str, rpr) -> etree._Element:
    """Create a <w:r> with optional rPr and <w:t>."""
    r = etree.Element(W_R)
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, W_T)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    return r


def _make_del_run(text: str, rpr) -> etree._Element:
    """Create a <w:r> with <w:delText> for deleted content."""
    r = etree.Element(W_R)
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    dt = etree.SubElement(r, W_DEL_TEXT)
    dt.text = text
    if text and (text[0] == " " or text[-1] == " "):
        dt.set(f"{{{XML_NS}}}space", "preserve")
    return r


# ── Paragraph-level diff ───────────────────────────────────────────────
def _diff_para(
    old_p: etree._Element,
    new_p: etree._Element,
    author: str,
    date: str,
    rid: _RevId,
):
    """Apply inline track-changes markup to new_p based on diff with old_p."""
    old_text = _para_text(old_p)
    new_text = _para_text(new_p)

    if old_text == new_text:
        return
    if not old_text and not new_text:
        return

    old_iv = _run_intervals(old_p)
    new_iv = _run_intervals(new_p)

    dmp = dmp_module.diff_match_patch()
    diffs = dmp.diff_main(old_text, new_text)
    dmp.diff_cleanupSemantic(diffs)

    children = []
    old_pos = 0
    new_pos = 0

    for op, text in diffs:
        if not text:
            continue
        if op == 0:  # EQUAL — use new document's formatting
            for chunk, rpr in _split_by_runs(text, new_pos, new_iv):
                children.append(_make_run(chunk, rpr))
            old_pos += len(text)
            new_pos += len(text)
        elif op == -1:  # DELETE — use old document's formatting
            d = etree.Element(W_DEL, _rev_attrs(rid, author, date))
            for chunk, rpr in _split_by_runs(text, old_pos, old_iv):
                d.append(_make_del_run(chunk, rpr))
            children.append(d)
            old_pos += len(text)
        elif op == 1:  # INSERT — use new document's formatting
            ins = etree.Element(W_INS, _rev_attrs(rid, author, date))
            for chunk, rpr in _split_by_runs(text, new_pos, new_iv):
                ins.append(_make_run(chunk, rpr))
            children.append(ins)
            new_pos += len(text)

    # Replace paragraph content (keep pPr, drop old runs/hyperlinks/etc.)
    for child in list(new_p):
        if child.tag != W_PPR:
            new_p.remove(child)
    for child in children:
        new_p.append(child)


# ── Table-level diff ───────────────────────────────────────────────────
def _diff_table(
    old_t: etree._Element,
    new_t: etree._Element,
    author: str,
    date: str,
    rid: _RevId,
):
    """Diff two tables by matching rows positionally and diffing cells."""
    old_rows = [r for r in old_t if r.tag == W_TR]
    new_rows = [r for r in new_t if r.tag == W_TR]

    for i in range(min(len(old_rows), len(new_rows))):
        old_cells = [c for c in old_rows[i] if c.tag == W_TC]
        new_cells = [c for c in new_rows[i] if c.tag == W_TC]
        for j in range(min(len(old_cells), len(new_cells))):
            old_paras = [p for p in old_cells[j] if p.tag == W_P]
            new_paras = [p for p in new_cells[j] if p.tag == W_P]
            for k in range(min(len(old_paras), len(new_paras))):
                _diff_para(old_paras[k], new_paras[k], author, date, rid)


# ── Whole-block insertion/deletion marking ──────────────────────────────
def _mark_para_inserted(p: etree._Element, author: str, date: str, rid: _RevId):
    """Wrap all runs in <w:ins> and mark paragraph mark as inserted."""
    if not _para_text(p).strip():
        return
    runs = [c for c in p if c.tag == W_R]
    if not runs:
        return

    ins = etree.Element(W_INS, _rev_attrs(rid, author, date))
    first_idx = list(p).index(runs[0])
    for r in runs:
        p.remove(r)
        ins.append(r)
    p.insert(first_idx, ins)

    # Mark paragraph mark as inserted
    _mark_ppr(p, W_INS, author, date, rid)


def _mark_para_deleted(p: etree._Element, author: str, date: str, rid: _RevId):
    """Wrap all runs in <w:del>, convert w:t to w:delText."""
    runs = [c for c in p if c.tag == W_R]
    if not runs:
        return

    d = etree.Element(W_DEL, _rev_attrs(rid, author, date))
    first_idx = list(p).index(runs[0])
    for r in runs:
        p.remove(r)
        # Convert w:t → w:delText
        for t in r.findall(W_T):
            dt = etree.SubElement(r, W_DEL_TEXT)
            dt.text = t.text
            if t.text and (t.text[0] == " " or t.text[-1] == " "):
                dt.set(f"{{{XML_NS}}}space", "preserve")
            r.remove(t)
        d.append(r)
    p.insert(first_idx, d)

    # Mark paragraph mark as deleted
    _mark_ppr(p, W_DEL, author, date, rid)


def _mark_ppr(p, mark_tag, author, date, rid):
    """Add ins/del mark to pPr/rPr of a paragraph."""
    ppr = p.find(W_PPR)
    if ppr is None:
        ppr = etree.SubElement(p, W_PPR)
        p.insert(0, ppr)
    rpr = ppr.find(W_RPR)
    if rpr is None:
        rpr = etree.SubElement(ppr, W_RPR)
    mark = etree.SubElement(rpr, mark_tag)
    for k, v in _rev_attrs(rid, author, date).items():
        mark.set(k, v)


def _mark_block_inserted(block: etree._Element, author: str, date: str, rid: _RevId):
    if block.tag == W_P:
        _mark_para_inserted(block, author, date, rid)
    elif block.tag == W_TBL:
        for p in block.iter(W_P):
            _mark_para_inserted(p, author, date, rid)


# ── Inserting deleted blocks ────────────────────────────────────────────
def _insert_deletions(
    old_blocks, new_blocks, matches, matched_old, new_body, author, date, rid,
):
    """Deep-copy unmatched old blocks, mark as deleted, inject into new_body."""
    # Group by insertion point (= index of next matched new block, or None)
    groups: dict = {}
    for i in range(len(old_blocks)):
        if i in matched_old:
            continue
        if not _block_text(old_blocks[i]).strip():
            continue

        insert_before = None
        for mi, mj in matches:
            if mi > i:
                insert_before = mj
                break

        groups.setdefault(insert_before, [])
        db = copy.deepcopy(old_blocks[i])
        if db.tag == W_P:
            _mark_para_deleted(db, author, date, rid)
        elif db.tag == W_TBL:
            for p in db.iter(W_P):
                _mark_para_deleted(p, author, date, rid)
        groups[insert_before].append(db)

    # Insert into new_body
    for ref_j, del_blocks in groups.items():
        if ref_j is not None:
            ref = new_blocks[ref_j]
            parent = ref.getparent()
            idx = list(parent).index(ref)
            for k, db in enumerate(del_blocks):
                parent.insert(idx + k, db)
        else:
            # Append at end (before sectPr if present)
            sect = new_body.find(W_SECT_PR)
            if sect is not None:
                idx = list(new_body).index(sect)
            else:
                idx = len(list(new_body))
            for k, db in enumerate(del_blocks):
                new_body.insert(idx + k, db)
