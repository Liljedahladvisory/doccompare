"""PDF pipeline: render OOXML Track-Changes .docx to PDF.

Primary path (Word available):
  1. Write track-changes .docx to temp file
  2. Use Microsoft Word headless (AppleScript) to export as PDF
  3. Generate summary/legend page via WeasyPrint
  4. Merge with pypdf
  5. Clean up temp files

Fallback path (no Word):
  Walk the XML tree, produce HTML with inline styles, render via WeasyPrint.

Style and numbering resolution reads word/styles.xml and word/numbering.xml
from the source DOCX to faithfully reproduce inherited formatting.
"""

import html as html_mod
import os
import subprocess
import tempfile
import time
import zipfile

from loguru import logger
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from lxml import etree

# ── OOXML namespaces & tags ─────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _qn(t):
    p, l = t.split(":")
    ns = {"w": W, "xml": XML_NS}[p]
    return f"{{{ns}}}{l}"


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
W_HYPERLINK = _qn("w:hyperlink")

_W = f"{{{W}}}"


# ── Helper: read bool-ish OOXML element ─────────────────────────────────
def _is_on(elem):
    """Return True if an OOXML toggle element is 'on'."""
    if elem is None:
        return False
    val = elem.get(f"{_W}val", "true")
    return val not in ("0", "false")


# ── Style Resolver ──────────────────────────────────────────────────────
class StyleResolver:
    """Resolve OOXML paragraph/run styles by walking the basedOn chain."""

    def __init__(self, docx_path):
        self._styles = {}  # styleId -> etree Element
        if docx_path is None:
            return
        try:
            with zipfile.ZipFile(docx_path) as zf:
                if "word/styles.xml" in zf.namelist():
                    tree = etree.parse(zf.open("word/styles.xml"))
                    for style_el in tree.iter(f"{_W}style"):
                        sid = style_el.get(f"{_W}styleId")
                        if sid:
                            self._styles[sid] = style_el
        except Exception:
            pass

    # ── public API ──────────────────────────────────────────────────────
    def resolve_paragraph_style(self, style_id):
        """Return (ppr_dict, rpr_dict) for a paragraph style, walking basedOn."""
        ppr = {}
        rpr = {}
        if not style_id or style_id not in self._styles:
            return ppr, rpr
        chain = self._basedOn_chain(style_id)
        # Apply from root ancestor to leaf (most-specific last)
        for sid in reversed(chain):
            el = self._styles.get(sid)
            if el is None:
                continue
            ppr_el = el.find(f"{_W}pPr")
            if ppr_el is not None:
                _merge_ppr_dict(ppr, _parse_ppr(ppr_el))
            rpr_el = el.find(f"{_W}rPr")
            if rpr_el is not None:
                _merge_rpr_dict(rpr, _parse_rpr(rpr_el))
        return ppr, rpr

    def resolve_run_style(self, style_id):
        """Return rpr_dict for a character style, walking basedOn."""
        rpr = {}
        if not style_id or style_id not in self._styles:
            return rpr
        chain = self._basedOn_chain(style_id)
        for sid in reversed(chain):
            el = self._styles.get(sid)
            if el is None:
                continue
            rpr_el = el.find(f"{_W}rPr")
            if rpr_el is not None:
                _merge_rpr_dict(rpr, _parse_rpr(rpr_el))
        return rpr

    def get_style_numpr(self, style_id):
        """Walk the style basedOn chain to find numPr (numId, ilvl).

        numId and ilvl may come from different levels in the chain.
        E.g. Heading2NotinTOC has ilvl=1, basedOn Heading1NotinTOC which has numId=24.
        We collect them independently: first non-None wins for each.
        """
        if not style_id:
            return None, None
        chain = self._basedOn_chain(style_id)
        found_nid = None
        found_ilvl = None
        # Walk from leaf to root — first found wins for each property
        for sid in chain:
            el = self._styles.get(sid)
            if el is None:
                continue
            ppr_el = el.find(f"{_W}pPr")
            if ppr_el is None:
                continue
            numpr = ppr_el.find(f"{_W}numPr")
            if numpr is not None:
                nid_el = numpr.find(f"{_W}numId")
                ilvl_el = numpr.find(f"{_W}ilvl")
                if found_nid is None and nid_el is not None:
                    found_nid = nid_el.get(f"{_W}val")
                if found_ilvl is None and ilvl_el is not None:
                    found_ilvl = ilvl_el.get(f"{_W}val")
            if found_nid is not None and found_ilvl is not None:
                break
        # Suppress if numId=0
        if found_nid == "0":
            return None, None
        return found_nid, found_ilvl

    # ── internals ───────────────────────────────────────────────────────
    def _basedOn_chain(self, style_id, _seen=None):
        """Return list [style_id, parent, grandparent, ...] (leaf-first)."""
        if _seen is None:
            _seen = set()
        if style_id in _seen or style_id not in self._styles:
            return []
        _seen.add(style_id)
        chain = [style_id]
        el = self._styles[style_id]
        based = el.find(f"{_W}basedOn")
        if based is not None:
            parent_id = based.get(f"{_W}val")
            if parent_id:
                chain.extend(self._basedOn_chain(parent_id, _seen))
        return chain


# ── Numbering Resolver ──────────────────────────────────────────────────
class NumberingResolver:
    """Resolve OOXML numbering definitions and generate labels."""

    def __init__(self, docx_path):
        self._abstract_nums = {}   # abstractNumId -> etree Element
        self._num_to_abstract = {}  # numId -> abstractNumId (str)
        self._counters = {}         # (numId, ilvl) -> current count
        self._last_ilvl = {}        # numId -> last ilvl used (for reset)
        if docx_path is None:
            return
        try:
            with zipfile.ZipFile(docx_path) as zf:
                if "word/numbering.xml" in zf.namelist():
                    tree = etree.parse(zf.open("word/numbering.xml"))
                    root = tree.getroot()
                    for an in root.iter(f"{_W}abstractNum"):
                        aid = an.get(f"{_W}abstractNumId")
                        if aid:
                            self._abstract_nums[aid] = an
                    for num in root.iter(f"{_W}num"):
                        nid = num.get(f"{_W}numId")
                        anr = num.find(f"{_W}abstractNumId")
                        if nid and anr is not None:
                            self._num_to_abstract[nid] = anr.get(f"{_W}val")
        except Exception:
            pass

    def generate_label(self, num_id, ilvl):
        """Return the numbering label string (e.g. '1.2') or '' if none."""
        if num_id is None or num_id == "0":
            return ""
        ilvl = int(ilvl or 0)
        lvl_el = self._find_level(num_id, ilvl)
        if lvl_el is None:
            return ""

        num_fmt = lvl_el.findtext(f"{_W}numFmt", default="")
        if not num_fmt:
            fmt_el = lvl_el.find(f"{_W}numFmt")
            if fmt_el is not None:
                num_fmt = fmt_el.get(f"{_W}val", "")

        lvl_text_el = lvl_el.find(f"{_W}lvlText")
        lvl_text = lvl_text_el.get(f"{_W}val", "") if lvl_text_el is not None else ""
        start_el = lvl_el.find(f"{_W}start")
        start = int(start_el.get(f"{_W}val", "1")) if start_el is not None else 1

        # Reset deeper levels when a shallower level is hit
        last = self._last_ilvl.get(num_id, -1)
        if ilvl <= last:
            # Reset all deeper levels
            for deeper in range(ilvl + 1, 10):
                self._counters.pop((num_id, deeper), None)
        self._last_ilvl[num_id] = ilvl

        # Increment counter
        key = (num_id, ilvl)
        if key not in self._counters:
            self._counters[key] = start
        else:
            self._counters[key] += 1
        current = self._counters[key]

        # Build label by substituting %1, %2, etc.
        label = lvl_text
        for lvl_idx in range(ilvl + 1):
            counter_val = self._counters.get((num_id, lvl_idx), 1)
            fmt = self._get_fmt_for_level(num_id, lvl_idx)
            formatted = self._format_number(counter_val, fmt)
            label = label.replace(f"%{lvl_idx + 1}", formatted)

        return label

    def _find_level(self, num_id, ilvl):
        """Find the w:lvl element, following numStyleLink chains."""
        abstract_id = self._num_to_abstract.get(str(num_id))
        if abstract_id is None:
            return None
        return self._find_level_in_abstract(abstract_id, ilvl, set())

    def _find_level_in_abstract(self, abstract_id, ilvl, seen):
        if abstract_id in seen or abstract_id not in self._abstract_nums:
            return None
        seen.add(abstract_id)
        an = self._abstract_nums[abstract_id]

        # Check for numStyleLink — follow to the linked abstract
        nsl = an.find(f"{_W}numStyleLink")
        if nsl is not None:
            link_style = nsl.get(f"{_W}val", "")
            target_aid = self._find_abstract_by_style_link(link_style)
            if target_aid is not None:
                return self._find_level_in_abstract(target_aid, ilvl, seen)

        # Find the level directly
        for lvl in an.iter(f"{_W}lvl"):
            if lvl.get(f"{_W}ilvl") == str(ilvl):
                return lvl
        return None

    def _find_abstract_by_style_link(self, style_name):
        """Find abstractNumId that has a styleLink matching the given name."""
        for aid, an in self._abstract_nums.items():
            sl = an.find(f"{_W}styleLink")
            if sl is not None and sl.get(f"{_W}val", "") == style_name:
                return aid
        return None

    def _get_fmt_for_level(self, num_id, lvl_idx):
        """Get numFmt for a specific level."""
        lvl_el = self._find_level(num_id, lvl_idx)
        if lvl_el is None:
            return "decimal"
        fmt_el = lvl_el.find(f"{_W}numFmt")
        if fmt_el is not None:
            return fmt_el.get(f"{_W}val", "decimal")
        return "decimal"

    @staticmethod
    def _format_number(n, fmt):
        if fmt == "decimal":
            return str(n)
        elif fmt == "lowerLetter":
            return chr(ord('a') + (n - 1) % 26)
        elif fmt == "upperLetter":
            return chr(ord('A') + (n - 1) % 26)
        elif fmt == "lowerRoman":
            return _to_roman(n).lower()
        elif fmt == "upperRoman":
            return _to_roman(n)
        elif fmt == "none":
            return ""
        return str(n)


def _to_roman(n):
    result = ""
    for val, rom in [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                     (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                     (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]:
        while n >= val:
            result += rom
            n -= val
    return result


# ── Parsing helpers: XML element → dict ─────────────────────────────────
def _parse_ppr(ppr_el):
    """Parse a w:pPr element into a dict."""
    d = {}
    if ppr_el is None:
        return d

    jc = ppr_el.find(f"{_W}jc")
    if jc is not None:
        d["alignment"] = jc.get(f"{_W}val", "")

    ind = ppr_el.find(f"{_W}ind")
    if ind is not None:
        for attr, key in [("left", "left_indent"), ("start", "left_indent"),
                          ("right", "right_indent"), ("end", "right_indent"),
                          ("hanging", "hanging"), ("firstLine", "first_line")]:
            val = ind.get(f"{_W}{attr}")
            if val is not None:
                try:
                    d[key] = int(val)
                except ValueError:
                    pass

    spacing = ppr_el.find(f"{_W}spacing")
    if spacing is not None:
        for attr, key in [("before", "space_before"), ("after", "space_after"),
                          ("line", "line_val")]:
            val = spacing.get(f"{_W}{attr}")
            if val is not None:
                try:
                    d[key] = int(val)
                except ValueError:
                    pass
        lr = spacing.get(f"{_W}lineRule")
        if lr:
            d["line_rule"] = lr

    return d


def _parse_rpr(rpr_el):
    """Parse a w:rPr element into a dict."""
    d = {}
    if rpr_el is None:
        return d

    b = rpr_el.find(f"{_W}b")
    if b is not None:
        d["bold"] = _is_on(b)

    i = rpr_el.find(f"{_W}i")
    if i is not None:
        d["italic"] = _is_on(i)

    u = rpr_el.find(f"{_W}u")
    if u is not None:
        val = u.get(f"{_W}val", "single")
        d["underline"] = val != "none"

    strike = rpr_el.find(f"{_W}strike")
    if strike is not None:
        d["strike"] = _is_on(strike)

    sz = rpr_el.find(f"{_W}sz")
    if sz is not None:
        try:
            d["font_size_half_pts"] = int(sz.get(f"{_W}val", "0"))
        except (ValueError, TypeError):
            pass

    rfonts = rpr_el.find(f"{_W}rFonts")
    if rfonts is not None:
        name = (rfonts.get(f"{_W}ascii")
                or rfonts.get(f"{_W}hAnsi")
                or rfonts.get(f"{_W}cs"))
        if name:
            d["font_name"] = name

    color = rpr_el.find(f"{_W}color")
    if color is not None:
        val = color.get(f"{_W}val", "")
        if val and val != "auto":
            d["color"] = val

    return d


def _merge_ppr_dict(base, overlay):
    """Merge overlay ppr_dict into base (overlay wins)."""
    base.update({k: v for k, v in overlay.items() if v is not None})


def _merge_rpr_dict(base, overlay):
    """Merge overlay rpr_dict into base (overlay wins)."""
    base.update({k: v for k, v in overlay.items() if v is not None})


# ── Dict → CSS conversion ──────────────────────────────────────────────
def _ppr_dict_to_css(d):
    """Convert a ppr_dict to inline CSS string."""
    parts = []
    align_map = {"left": "left", "center": "center", "right": "right",
                 "both": "justify", "justify": "justify"}
    alignment = d.get("alignment")
    if alignment:
        css = align_map.get(alignment)
        if css:
            parts.append(f"text-align:{css}")

    left = d.get("left_indent")
    if left:
        parts.append(f"margin-left:{left / 20:.1f}pt")

    right = d.get("right_indent")
    if right:
        parts.append(f"margin-right:{right / 20:.1f}pt")

    hanging = d.get("hanging")
    first_line = d.get("first_line")
    if hanging:
        parts.append(f"text-indent:-{hanging / 20:.1f}pt")
    elif first_line:
        parts.append(f"text-indent:{first_line / 20:.1f}pt")

    sb = d.get("space_before")
    if sb is not None:
        parts.append(f"margin-top:{sb / 20:.1f}pt")

    sa = d.get("space_after")
    if sa is not None:
        parts.append(f"margin-bottom:{sa / 20:.1f}pt")

    line_val = d.get("line_val")
    if line_val:
        lr = d.get("line_rule", "")
        if lr in ("exact", "atLeast"):
            parts.append(f"line-height:{line_val / 20:.1f}pt")
        else:
            parts.append(f"line-height:{line_val / 240:.2f}")

    return ";".join(parts)


def _rpr_dict_to_css(d):
    """Convert an rpr_dict to inline CSS string."""
    parts = []

    if d.get("bold"):
        parts.append("font-weight:bold")
    if d.get("italic"):
        parts.append("font-style:italic")
    if d.get("underline"):
        parts.append("text-decoration:underline")
    if d.get("strike"):
        # If already underline, combine
        if d.get("underline"):
            parts[-1] = "text-decoration:underline line-through"
        else:
            parts.append("text-decoration:line-through")

    sz = d.get("font_size_half_pts")
    if sz and sz > 0:
        parts.append(f"font-size:{sz / 2:.1f}pt")

    fn = d.get("font_name")
    if fn:
        parts.append(f"font-family:'{fn}', serif")

    color = d.get("color")
    if color and len(color) == 6:
        parts.append(f"color:#{color}")

    return ";".join(parts)


# ── Element rendering (style-aware) ────────────────────────────────────
def _render_run(run, css_class=None, default_rpr=None, style_resolver=None):
    """Render a w:r to HTML, merging style + direct formatting."""
    rpr_el = run.find(W_RPR)
    effective = dict(default_rpr) if default_rpr else {}

    # Resolve rStyle if present
    if rpr_el is not None and style_resolver is not None:
        rstyle = rpr_el.find(f"{_W}rStyle")
        if rstyle is not None:
            sid = rstyle.get(f"{_W}val")
            if sid:
                style_rpr = style_resolver.resolve_run_style(sid)
                _merge_rpr_dict(effective, style_rpr)

    # Merge direct formatting on top
    if rpr_el is not None:
        direct = _parse_rpr(rpr_el)
        _merge_rpr_dict(effective, direct)

    style = _rpr_dict_to_css(effective)

    t = run.find(W_T)
    text = html_mod.escape(t.text or "") if t is not None else ""
    if not text:
        return ""

    classes = f' class="{css_class}"' if css_class else ""
    style_attr = f' style="{style}"' if style else ""
    return f"<span{classes}{style_attr}>{text}</span>"


def _render_del_run(run, default_rpr=None, style_resolver=None):
    """Render a deleted w:r (with w:delText) to HTML."""
    rpr_el = run.find(W_RPR)
    effective = dict(default_rpr) if default_rpr else {}

    if rpr_el is not None and style_resolver is not None:
        rstyle = rpr_el.find(f"{_W}rStyle")
        if rstyle is not None:
            sid = rstyle.get(f"{_W}val")
            if sid:
                style_rpr = style_resolver.resolve_run_style(sid)
                _merge_rpr_dict(effective, style_rpr)

    if rpr_el is not None:
        direct = _parse_rpr(rpr_el)
        _merge_rpr_dict(effective, direct)

    style = _rpr_dict_to_css(effective)

    dt = run.find(f"{_W}delText")
    text = html_mod.escape(dt.text or "") if dt is not None else ""
    if not text:
        return ""

    style_attr = f' style="{style}"' if style else ""
    return f'<span class="deleted"{style_attr}>{text}</span>'


def _render_paragraph(para, style_resolver=None, numbering_resolver=None) -> str:
    """Render a w:p to HTML <p>, with full style + numbering resolution."""
    ppr_el = para.find(W_PPR)

    # 1. Get paragraph style ID
    p_style_id = None
    if ppr_el is not None:
        pstyle = ppr_el.find(f"{_W}pStyle")
        if pstyle is not None:
            p_style_id = pstyle.get(f"{_W}val")

    # 2. Resolve style chain → effective ppr_dict and rpr_dict (defaults for runs)
    effective_ppr = {}
    default_rpr = {}
    if style_resolver and p_style_id:
        effective_ppr, default_rpr = style_resolver.resolve_paragraph_style(p_style_id)

    # 3. Merge DIRECT pPr on top
    if ppr_el is not None:
        direct_ppr = _parse_ppr(ppr_el)
        _merge_ppr_dict(effective_ppr, direct_ppr)

    style = _ppr_dict_to_css(effective_ppr)

    # 4. Numbering
    num_label = ""
    if numbering_resolver is not None:
        num_id = None
        ilvl = None
        # Check direct numPr first
        if ppr_el is not None:
            numpr = ppr_el.find(f"{_W}numPr")
            if numpr is not None:
                nid_el = numpr.find(f"{_W}numId")
                ilvl_el = numpr.find(f"{_W}ilvl")
                if nid_el is not None:
                    num_id = nid_el.get(f"{_W}val")
                if ilvl_el is not None:
                    ilvl = ilvl_el.get(f"{_W}val")
        # Fallback: check style chain for numPr
        if num_id is None and style_resolver and p_style_id:
            num_id, ilvl = style_resolver.get_style_numpr(p_style_id)
        if num_id and num_id != "0":
            num_label = numbering_resolver.generate_label(num_id, ilvl)

    # 5. Check if entire paragraph is inserted or deleted
    para_class = ""
    if ppr_el is not None:
        rpr = ppr_el.find(W_RPR)
        if rpr is not None:
            if rpr.find(W_INS) is not None:
                para_class = " element-added"
            elif rpr.find(W_DEL) is not None:
                para_class = " element-deleted"

    # 6. Render child runs
    content = []
    if num_label:
        content.append(f'<span class="numbering">{html_mod.escape(num_label)}\u00a0</span>')

    for child in para:
        if child.tag == W_R:
            content.append(_render_run(child, default_rpr=default_rpr,
                                       style_resolver=style_resolver))
        elif child.tag == W_INS:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_run(r, css_class="added",
                                               default_rpr=default_rpr,
                                               style_resolver=style_resolver))
        elif child.tag == W_DEL:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_del_run(r, default_rpr=default_rpr,
                                                   style_resolver=style_resolver))
        elif child.tag == W_HYPERLINK:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_run(r, default_rpr=default_rpr,
                                               style_resolver=style_resolver))

    inner = "".join(content)
    if not inner.strip():
        return ""
    class_attr = f' class="{para_class.strip()}"' if para_class else ""
    style_attr = f' style="{style}"' if style else ""
    return f"<p{class_attr}{style_attr}>{inner}</p>"


def _render_table(tbl, style_resolver=None, numbering_resolver=None) -> str:
    """Render a w:tbl to HTML <table>."""
    rows = []
    for tr in tbl:
        if tr.tag != W_TR:
            continue
        cells = []
        for tc in tr:
            if tc.tag != W_TC:
                continue
            cell_html = []
            for p in tc:
                if p.tag == W_P:
                    rendered = _render_paragraph(
                        p, style_resolver=style_resolver,
                        numbering_resolver=numbering_resolver,
                    )
                    if rendered:
                        cell_html.append(rendered)
            cells.append(f"<td>{''.join(cell_html)}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    return f"<table>{''.join(rows)}</table>"


def render_tracked_changes_html(
    doc_tree,
    summary: dict,
    original_name: str,
    modified_name: str,
    docx_path=None,
) -> str:
    """Walk the OOXML tree and produce full HTML with tracked changes."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    body = doc_tree.find(W_BODY)

    # Build resolvers from the source DOCX (if provided)
    style_resolver = StyleResolver(docx_path)
    numbering_resolver = NumberingResolver(docx_path)

    # Render all block elements
    body_parts = []
    for child in body:
        if child.tag == W_P:
            rendered = _render_paragraph(
                child, style_resolver=style_resolver,
                numbering_resolver=numbering_resolver,
            )
            if rendered:
                body_parts.append(rendered)
        elif child.tag == W_TBL:
            body_parts.append(_render_table(
                child, style_resolver=style_resolver,
                numbering_resolver=numbering_resolver,
            ))

    content_html = "\n".join(body_parts)

    # Summary + legend (on a new page at the end)
    added = summary.get("added_words", 0)
    deleted = summary.get("deleted_words", 0)
    unchanged = summary.get("unchanged_words", 0)

    summary_html = f"""
    <div class="summary-page">
        <h2 class="summary-title">DocCompare &mdash; Sammanfattning</h2>
        <div class="summary-meta">
            <strong>Original:</strong> {html_mod.escape(original_name)} &nbsp;|&nbsp;
            <strong>Modifierat:</strong> {html_mod.escape(modified_name)} &nbsp;|&nbsp;
            <strong>Datum:</strong> {now}
        </div>
        <div class="summary-stats">
            <div class="stat stat-added">+{added} ord tillagda</div>
            <div class="stat stat-deleted">&minus;{deleted} ord borttagna</div>
            <div class="stat">{unchanged} ord of&ouml;r&auml;ndrade</div>
        </div>
        <div class="summary-legend">
            <h3>Legend</h3>
            <p><span class="added">Tillagd text</span> &mdash;
               text som finns i det modifierade dokumentet men inte i originalet.</p>
            <p><span class="deleted">Borttagen text</span> &mdash;
               text som finns i originalet men inte i det modifierade dokumentet.</p>
            <p>Of&ouml;r&auml;ndrad text &mdash;
               text som &auml;r identisk i b&aring;da dokumenten.</p>
        </div>
        <div class="summary-footer">
            Genererad av DocCompare &mdash; Liljedahl Advisory AB
        </div>
    </div>
    """

    css_path = Path(__file__).parent / "styles.css"
    css = css_path.read_text(encoding="utf-8")

    return f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8">
<title>DocCompare &mdash; {html_mod.escape(original_name)} vs {html_mod.escape(modified_name)}</title>
<style>
{css}

/* Summary page — forced onto new page */
.summary-page {{
    page-break-before: always;
    font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
}}
.summary-title {{
    font-size: 18pt;
    font-weight: 700;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-meta {{
    font-size: 9pt;
    color: #555;
    border-top: 1px solid #bdc3c7;
    padding-top: 8pt;
    margin-bottom: 16pt;
}}
.summary-stats {{
    margin-bottom: 20pt;
}}
.summary-stats .stat {{
    display: block;
    margin-bottom: 4pt;
    font-size: 11pt;
}}
.summary-stats .stat-added {{ color: #0047ab; }}
.summary-stats .stat-deleted {{ color: #c0392b; }}
.summary-legend h3 {{
    font-size: 13pt;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-legend p {{
    font-size: 10pt;
    margin-bottom: 6pt;
}}
.summary-footer {{
    margin-top: 24pt;
    padding-top: 8pt;
    border-top: 1px solid #bdc3c7;
    font-size: 8pt;
    color: #888;
}}
</style>
</head>
<body>
{content_html}
{summary_html}
</body>
</html>"""


def _word_temp_dir() -> Path:
    """Return a temp directory inside Word's own sandbox container.

    ~/Library/Group Containers/UBF8T346G9.Office/ is Word's sandbox —
    it has GUARANTEED read/write access with zero permission dialogs.
    """
    d = Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office" / "DocCompare"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _pdf_to_docx_headless(pdf_path: Path, docx_path: Path) -> bool:
    """Convert .pdf to .docx using Microsoft Word via AppleScript (headless).

    Word opens the PDF (which triggers its built-in PDF-to-DOCX converter),
    then saves as .docx. No dialogs, no user interaction.
    Returns True on success, False if Word is unavailable.
    """
    pdf_abs = str(pdf_path.resolve())
    docx_abs = str(docx_path.resolve())

    script = f'''
    tell application "Microsoft Word"
        open POSIX file "{pdf_abs}" as alias
        delay 2
        set theDoc to active document
        set outputPath to POSIX file "{docx_abs}" as text
        save as theDoc file name outputPath file format format document
        close theDoc saving no
    end tell
    '''

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=90,
        )
        if result.returncode != 0:
            logger.warning(f"Word PDF→DOCX failed (rc={result.returncode}): {result.stderr.strip()}")
            return False
        if not docx_path.exists():
            logger.warning("Word PDF→DOCX succeeded but DOCX file not found")
            return False
        logger.info(f"Word converted PDF to DOCX: {docx_path.name}")
        return True
    except subprocess.TimeoutExpired:
        logger.warning("Word PDF→DOCX timed out after 90s")
        return False
    except FileNotFoundError:
        logger.warning("osascript not found — not on macOS?")
        return False


def _word_to_pdf_headless(docx_path: Path, pdf_path: Path) -> bool:
    """Convert .docx to PDF using Microsoft Word via AppleScript (headless).

    No System Events, no activate — Word stays in the background.
    Temp files go to ~/Library/Caches/DocCompare/ to avoid sandbox dialogs.
    Returns True on success, False if Word is unavailable.
    """
    docx_abs = str(docx_path.resolve())
    pdf_abs = str(pdf_path.resolve())

    # AppleScript: use 'open ... without activate' pattern.
    # No System Events (avoids permission dialog).
    # No 'activate' (Word stays in background).
    script = f'''
    tell application "Microsoft Word"
        -- 'open' without 'activate' keeps Word in the background
        open POSIX file "{docx_abs}" as alias
        delay 1
        set theDoc to active document
        set outputPath to POSIX file "{pdf_abs}" as text
        save as theDoc file name outputPath file format format PDF
        close theDoc saving no
    end tell
    '''

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            logger.warning(f"Word AppleScript failed (rc={result.returncode}): {result.stderr.strip()}")
            return False
        if not pdf_path.exists():
            logger.warning("Word AppleScript succeeded but PDF file not found")
            return False
        return True
    except subprocess.TimeoutExpired:
        logger.warning("Word AppleScript timed out after 60s")
        return False
    except FileNotFoundError:
        logger.warning("osascript not found — not on macOS?")
        return False


def _render_summary_pdf(summary: dict, original_name: str, modified_name: str) -> bytes:
    """Render summary/legend page as a standalone PDF via WeasyPrint."""
    from doccompare.rendering.pdf_renderer import render_pdf

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    added = summary.get("added_words", 0)
    deleted = summary.get("deleted_words", 0)
    unchanged = summary.get("unchanged_words", 0)

    css_path = Path(__file__).parent / "styles.css"
    css = css_path.read_text(encoding="utf-8")

    html_content = f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8">
<style>
{css}
.summary-page {{
    font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
}}
.summary-title {{
    font-size: 18pt;
    font-weight: 700;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-meta {{
    font-size: 9pt;
    color: #555;
    border-top: 1px solid #bdc3c7;
    padding-top: 8pt;
    margin-bottom: 16pt;
}}
.summary-stats {{
    margin-bottom: 20pt;
}}
.summary-stats .stat {{
    display: block;
    margin-bottom: 4pt;
    font-size: 11pt;
}}
.summary-stats .stat-added {{ color: #0047ab; }}
.summary-stats .stat-deleted {{ color: #c0392b; }}
.summary-legend h3 {{
    font-size: 13pt;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-legend p {{
    font-size: 10pt;
    margin-bottom: 6pt;
}}
.summary-footer {{
    margin-top: 24pt;
    padding-top: 8pt;
    border-top: 1px solid #bdc3c7;
    font-size: 8pt;
    color: #888;
}}
</style>
</head>
<body>
    <div class="summary-page">
        <h2 class="summary-title">DocCompare &mdash; Summary</h2>
        <div class="summary-meta">
            <strong>Original:</strong> {html_mod.escape(original_name)} &nbsp;|&nbsp;
            <strong>Modified:</strong> {html_mod.escape(modified_name)} &nbsp;|&nbsp;
            <strong>Date:</strong> {now}
        </div>
        <div class="summary-stats">
            <div class="stat stat-added">+{added} words added</div>
            <div class="stat stat-deleted">&minus;{deleted} words deleted</div>
            <div class="stat">{unchanged} words unchanged</div>
        </div>
        <div class="summary-legend">
            <h3>Legend</h3>
            <p><span style="color:#0047ab;text-decoration:underline">Added text</span> &mdash;
               text present in the modified document but not in the original.</p>
            <p><span style="color:#c0392b;text-decoration:line-through">Deleted text</span> &mdash;
               text present in the original but not in the modified document.</p>
            <p>Unchanged text &mdash;
               text identical in both documents.</p>
        </div>
        <div class="summary-footer">
            Generated by DocCompare &mdash; a Liljedahl Legal Tech Tool from Liljedahl Advisory AB
        </div>
    </div>
</body>
</html>"""

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        render_pdf(html_content, css_path, tmp_path)
        return tmp_path.read_bytes()
    finally:
        tmp_path.unlink(missing_ok=True)


def _merge_pdfs(main_pdf: Path, summary_pdf_bytes: bytes, output_pdf: Path):
    """Merge the main document PDF with the summary page PDF."""
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()

    # Add all pages from the main document
    reader = PdfReader(str(main_pdf))
    for page in reader.pages:
        writer.add_page(page)

    # Add summary page(s)
    import io
    summary_reader = PdfReader(io.BytesIO(summary_pdf_bytes))
    for page in summary_reader.pages:
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)


def produce_pdf(
    doc_tree,
    output_pdf: Path,
    summary: dict,
    original_name: str,
    modified_name: str,
    docx_path=None,
):
    """Produce final PDF: Word headless export + summary/legend merge.

    Falls back to WeasyPrint HTML rendering if Word is unavailable.
    """
    from doccompare.comparison.ooxml_engine import _write_docx

    output_pdf = Path(output_pdf)

    # Temp files go inside Word's own sandbox container so that
    # Word can read/write them without triggering macOS permission dialogs.
    word_dir = _word_temp_dir()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    tmp_docx_path = word_dir / f".doccompare_tc_{ts}.docx"
    tmp_pdf_path = word_dir / f".doccompare_tc_{ts}.pdf"

    try:
        # Step 1: Write tracked-changes .docx
        if docx_path:
            _write_docx(Path(docx_path), tmp_docx_path, doc_tree)
        else:
            logger.warning("No docx_path provided — falling back to WeasyPrint")
            return _produce_pdf_weasyprint(
                doc_tree, output_pdf, summary, original_name, modified_name, docx_path,
            )

        # Step 2: Convert to PDF via Word (headless)
        logger.info(f"Word headless: {tmp_docx_path} → {tmp_pdf_path}")
        word_ok = _word_to_pdf_headless(tmp_docx_path, tmp_pdf_path)
        if not word_ok:
            logger.warning("Word headless conversion failed — falling back to WeasyPrint")
            return _produce_pdf_weasyprint(
                doc_tree, output_pdf, summary, original_name, modified_name, docx_path,
            )
        logger.info("Word headless conversion succeeded")

        # Step 3: Render summary/legend page
        summary_bytes = _render_summary_pdf(summary, original_name, modified_name)

        # Step 4: Merge
        _merge_pdfs(tmp_pdf_path, summary_bytes, output_pdf)

    finally:
        tmp_docx_path.unlink(missing_ok=True)
        tmp_pdf_path.unlink(missing_ok=True)


def _produce_pdf_weasyprint(
    doc_tree,
    output_pdf: Path,
    summary: dict,
    original_name: str,
    modified_name: str,
    docx_path=None,
):
    """Fallback: render tracked-changes XML tree to PDF via WeasyPrint."""
    from doccompare.rendering.pdf_renderer import render_pdf

    html_content = render_tracked_changes_html(
        doc_tree, summary, original_name, modified_name,
        docx_path=docx_path,
    )
    css_path = Path(__file__).parent / "styles.css"
    render_pdf(html_content, css_path, output_pdf)
