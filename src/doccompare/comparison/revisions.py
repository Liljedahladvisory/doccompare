"""Read Word's actual revisions and verify its accepted/rejected projections.

This module never rebuilds paragraphs. The Word-produced package remains the
layout source; text projections are only independent validation evidence.
"""
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from hashlib import sha256
import posixpath
import re
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
STORY = re.compile(r'word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$')
CONTENT_REVISIONS = {'ins', 'del', 'moveFrom', 'moveTo'}
PROPERTY_REVISIONS = {'rPrChange', 'pPrChange', 'tblPrChange', 'tblGridChange',
                      'trPrChange', 'tcPrChange', 'sectPrChange', 'numberingChange',
                      'cellIns', 'cellDel', 'cellMerge'}


class ComparisonQualityError(RuntimeError):
    """The result could not be demonstrated to represent both inputs."""


def local(node):
    return etree.QName(node).localname


def read_stories(path):
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if STORY.fullmatch(name):
                yield name, etree.fromstring(archive.read(name))


def preflight(path):
    path = Path(path)
    if path.suffix.lower() != '.docx':
        raise ComparisonQualityError('Välj två Word-dokument i formatet .docx.')
    try:
        with zipfile.ZipFile(path) as archive:
            if archive.testzip() or 'word/document.xml' not in archive.namelist():
                raise ValueError('invalid package')
        for name, tree in read_stories(path):
            for node in tree.iter():
                if not isinstance(node.tag, str):
                    continue
                tag = local(node)
                if tag in CONTENT_REVISIONS | PROPERTY_REVISIONS:
                    raise ComparisonQualityError(
                        f'{path.name} innehåller redan spårade ändringar. '
                        'Acceptera eller avvisa dem i en separat kopia före jämförelsen.')
                if tag in {'altChunk', 'object'}:
                    raise ComparisonQualityError(
                        f'{path.name} innehåller inbäddat innehåll som denna version '
                        'inte kan jämföra säkert.')
    except (zipfile.BadZipFile, etree.XMLSyntaxError, ValueError, KeyError) as exc:
        raise ComparisonQualityError(f'{path.name} är inte en läsbar DOCX-fil.') from exc


@dataclass(frozen=True)
class Revision:
    id: str
    kind: str
    part: str
    paragraph: int
    text: str
    words: int
    structural: bool = False
    context: str = ""
    scope: str = ""


def visible_text(node):
    """Read text without introducing spaces between formatting runs."""
    pieces = []
    for child in node.iter():
        if child.tag in {f'{{{W}}}t', f'{{{W}}}delText'}:
            pieces.append(child.text or '')
        elif child.tag == f'{{{W}}}tab':
            pieces.append('\t')
        elif child.tag == f'{{{W}}}br':
            pieces.append('\n')
    return ''.join(pieces)


def revision_ledger(path):
    revisions = []
    for part, tree in read_stories(path):
        paragraphs = {p: i + 1 for i, p in enumerate(tree.iter(f'{{{W}}}p'))}
        for node in tree.iter():
            if not isinstance(node.tag, str):
                continue
            tag = local(node)
            if tag not in CONTENT_REVISIONS | PROPERTY_REVISIONS:
                continue
            # Old properties can contain historical markup, not new revisions.
            if any(local(a) in PROPERTY_REVISIONS for a in node.iterancestors()):
                continue
            parent = node.getparent()
            structural = local(parent) in {'rPr', 'trPr', 'tcPr', 'pPr'} or tag in PROPERTY_REVISIONS
            paragraph = next((paragraphs[a] for a in node.iterancestors() if a in paragraphs), 0)
            text = '' if structural else visible_text(node)
            context_node = next((a for a in node.iterancestors() if local(a) in {'p', 'tr', 'tc'}), parent)
            revisions.append(Revision(
                id=f'{part}:{node.get(f"{{{W}}}id", str(len(revisions)))}',
                kind=tag, part=part, paragraph=paragraph, text=text,
                words=len(text.split()), structural=structural,
                context=(' | '.join(visible_text(p) for p in context_node.iter(f'{{{W}}}p')) or visible_text(context_node))[:180] if structural else '', scope=local(parent),
            ))
    return revisions


def summarize(revisions):
    words = Counter()
    for revision in revisions:
        words[revision.kind] += revision.words
    return {
        'added_words': words['ins'], 'deleted_words': words['del'],
        'moved_words': words['moveTo'],
        'revision_count': len(revisions),
        'format_revision_count': sum(r.kind in PROPERTY_REVISIONS for r in revisions),
        'revisions': [asdict(r) for r in revisions],
        'structure_changes': [asdict(r) for r in revisions if r.part != 'word/document.xml'],
    }


def _paragraph_tokens(paragraph, relationships=None):
    """Ignore calculated field results but retain field instructions and anchors.

    Page numbers/TOC/REF results can change on repagination. A changed field
    instruction must still fail projection validation. Run boundaries are not
    semantic, so adjacent text is joined before normalization.
    """
    relationships = relationships or {}
    pieces = []
    depth = 0
    for node in paragraph.iter():
        tag = local(node)
        if tag == 'fldChar':
            t = node.get(f'{{{W}}}fldCharType')
            if t == 'begin':
                depth += 1
                pieces.append('[FIELD:')
            elif t == 'end':
                depth = max(0, depth - 1)
                pieces.append(']')
        elif tag == 'instrText':
            pieces.append(node.text or '')
        elif tag == 't' and not depth:
            # fldSimple result is also recalculated by Word.
            if not any(local(a) == 'fldSimple' for a in node.iterancestors()):
                pieces.append(node.text or '')
        elif tag == 'hyperlink':
            rid = node.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            pieces.append('[LINK:' + relationships.get(rid, '') + '#' + node.get(f'{{{W}}}anchor', '') + ']')
        elif tag in {'blip', 'imagedata'}:
            for attr in ('embed', 'link', 'id'):
                rid = node.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}' + attr)
                if rid:
                    pieces.append('[IMAGE:' + relationships.get(rid, 'MISSING') + ']')
        elif tag == 'fldSimple':
            pieces.append('[FIELD:' + re.sub(r'\s+', ' ', node.get(f'{{{W}}}instr', '').strip()) + ']')
        elif tag == 'tab' and not depth:
            pieces.append('\t')
        elif tag == 'br' and not depth:
            pieces.append('[BREAK:' + node.get(f'{{{W}}}type', 'textWrapping') + ']')
        elif tag in {'footnoteReference', 'endnoteReference'}:
            # Word can renumber note IDs; note contents are checked separately.
            pieces.append('[' + tag + ']')
    value = ''.join(pieces).replace('\r', '').strip()
    return re.sub(r'\[FIELD:(.*?)\]', lambda m: '[FIELD:' + re.sub(r'\s+', ' ', m[1].strip()) + ']', value, flags=re.S)


def _relationships(path, part):
    relfile = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    result = {}
    with zipfile.ZipFile(path) as archive:
        if relfile not in archive.namelist():
            return result
        for rel in etree.fromstring(archive.read(relfile)):
            target = rel.get('Target', '')
            if rel.get('TargetMode') == 'External':
                value = target
            else:
                target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip('/')
                value = sha256(archive.read(target)).hexdigest() if target in archive.namelist() else 'MISSING:' + target
            result[rel.get('Id')] = value
    return result


def content_projection(path):
    """Text + paragraph/cell boundaries per story, insensitive to ZIP part IDs."""
    return _project_stories(path, read_stories(path))


def _project_stories(path, parts):
    parts = dict(parts)
    stories = []
    by_part = {}
    for part, tree in parts.items():
        family = re.sub(r'\d+', '', Path(part).stem)
        tokens = []
        relationships = _relationships(path, part)
        for node in tree.iter():
            tag = local(node)
            if tag == 'p':
                value = _paragraph_tokens(node, relationships)
                # Empty paragraphs remain part of layout, but Word may add a
                # terminal paragraph to a table cell. Text fidelity ignores these.
                if value:
                    tokens.append(('p', value))
            elif tag in {'tr', 'tc'}:
                tokens.append(('row' if tag == 'tr' else 'cell', ''))
        by_part[part] = tuple(tokens)
        if tokens and family not in {'header', 'footer'}:
            stories.append((family, tuple(tokens)))
    # Word may merge identical header/footer ZIP parts on save. Compare the
    # effective content for every section and variant, not the part count.
    # Section identities also detect swapped, missing and misrouted stories.
    document = parts.get('word/document.xml')
    if document is not None:
        with zipfile.ZipFile(path) as archive:
            relfile = 'word/_rels/document.xml.rels'
            rels = etree.fromstring(archive.read(relfile)) if relfile in archive.namelist() else []
        targets = {rel.get('Id'): posixpath.normpath(posixpath.join('word', rel.get('Target', ''))).lstrip('/')
                   for rel in rels if rel.get('TargetMode') != 'External'}
        effective = {}
        sections = [s for s in document.iter(f'{{{W}}}sectPr')
                    if not any(local(a) in PROPERTY_REVISIONS for a in s.iterancestors())]
        for index, section in enumerate(sections):
            for ref in section:
                tag = local(ref)
                if tag not in {'headerReference', 'footerReference'}:
                    continue
                rid = ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = targets.get(rid)
                family = tag.removesuffix('Reference')
                if target not in by_part or not Path(target).name.startswith(family):
                    raise ComparisonQualityError('En sidhuvuds- eller sidfotsreferens saknar innehåll.')
                effective[(family, ref.get(f'{{{W}}}type', 'default'))] = by_part[target]
            for (family, variant), tokens in effective.items():
                if tokens:
                    stories.append((family, (('section', f'{index}:{variant}'),) + tokens))
    return Counter(stories)


def _rejected_column_residue(original, rejected, tracked):
    """Normalize only proven empty inserted columns left by Word after rejection.

    Never changes a DOCX or the PDF. All retained cell contents/positions, rows,
    stories, fields and links must still match the source projection exactly.
    Merged/nested tables and unsupported inserted content remain fail-closed.
    """
    source_parts = dict(read_stories(original))
    result_parts = dict(read_stories(rejected))
    tracked_parts = dict(read_stories(tracked))
    part = 'word/document.xml'
    if any(part not in parts for parts in (source_parts, result_parts, tracked_parts)):
        return None, []
    tables = [list(parts[part].iter(f'{{{W}}}tbl'))
              for parts in (source_parts, result_parts, tracked_parts)]
    if len({len(items) for items in tables}) != 1:
        return None, []
    notes = []
    for index, (source, result, redline) in enumerate(zip(*tables), 1):
        source_rows, result_rows, tracked_rows = [t.findall('w:tr', NS) for t in (source, result, redline)]
        if not source_rows or len({len(rows) for rows in (source_rows, result_rows, tracked_rows)}) != 1:
            continue
        old_count = len(source_rows[0].findall('w:tc', NS))
        new_count = len(result_rows[0].findall('w:tc', NS))
        if new_count <= old_count or redline.find('w:tblGrid/w:tblGridChange', NS) is None:
            continue
        if any(t.xpath('.//w:tbl | .//w:gridSpan[not(ancestor::w:tcPrChange)] | '
                       './/w:hMerge[not(ancestor::w:tcPrChange)] | '
                       './/w:vMerge[not(ancestor::w:tcPrChange)]', namespaces=NS)
               for t in (source, result, redline)):
            continue
        from itertools import combinations
        from math import comb
        row_sets = [list(zip(*(row.findall('w:tc', NS) for row in rows)))
                    for rows in (source_rows, result_rows, tracked_rows)]
        if any(len(row.findall('w:tc', NS)) != count
               for rows, count in ((source_rows, old_count), (result_rows, new_count), (tracked_rows, new_count))
               for row in rows):
            continue
        candidates = []
        for column, (result_cells, evidence_cells) in enumerate(zip(row_sets[1], row_sets[2])):
            for cell, evidence in zip(result_cells, evidence_cells):
                if any(_paragraph_tokens(p) for p in cell.iter(f'{{{W}}}p')):
                    break
                # Only plain inserted text can justify an empty residual cell.
                if evidence.xpath('.//w:del | .//w:moveFrom | .//w:fldChar | .//w:fldSimple | '
                                  './/w:hyperlink | .//w:drawing | .//w:pict | .//w:object | '
                                  './/w:footnoteReference | .//w:endnoteReference', namespaces=NS):
                    break
                text_nodes = evidence.xpath('.//w:t | .//w:tab | .//w:br', namespaces=NS)
                if not text_nodes or any(not n.xpath('ancestor::w:ins', namespaces=NS) for n in text_nodes):
                    break
            else:
                candidates.append(column)
        extra = new_count - old_count
        if len(candidates) < extra or comb(len(candidates), extra) > 256:
            continue
        source_rels, result_rels = _relationships(original, part), _relationships(rejected, part)
        def cell_value(cell, relationships):
            return tuple(value for p in cell.iter(f'{{{W}}}p')
                         if (value := _paragraph_tokens(p, relationships)))
        matches = []
        for columns in combinations(candidates, extra):
            if all([cell_value(c, source_rels) for c in old_row.findall('w:tc', NS)] ==
                   [cell_value(c, result_rels) for ci, c in enumerate(result_row.findall('w:tc', NS))
                    if ci not in columns]
                   for old_row, result_row in zip(source_rows, result_rows)):
                matches.append(columns)
        # An ambiguous alignment is not sufficient evidence.
        if len(matches) != 1:
            continue
        for column in matches[0]:
            for cell in row_sets[1][column]:
                cell.getparent().remove(cell)
        notes.append({'table': index, 'old_columns': old_count, 'new_columns': new_count})
    return _project_stories(rejected, result_parts.items()), notes


def verify_projections(original, modified, accepted, rejected, *, tracked=None):
    notes = []
    for source, projection, label in [(modified, accepted, 'nya'), (original, rejected, 'gamla')]:
        expected = content_projection(source)
        if expected != content_projection(projection):
            if label == 'gamla' and tracked is not None:
                normalized, notes = _rejected_column_residue(source, projection, tracked)
                if notes and expected == normalized:
                    continue
            raise ComparisonQualityError(
                f'Word-resultatet kunde inte stämmas av mot den {label} versionen. '
                'Ingen PDF har ersatt din resultatfil. Dokumentparet behöver granskas.')
    return notes
