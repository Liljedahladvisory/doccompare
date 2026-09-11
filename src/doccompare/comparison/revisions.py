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
    stories = []
    for part, tree in read_stories(path):
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
        if tokens:
            stories.append((family, tuple(tokens)))
    return Counter(stories)


def verify_projections(original, modified, accepted, rejected):
    for source, projection, label in [(modified, accepted, 'nya'), (original, rejected, 'gamla')]:
        if content_projection(source) != content_projection(projection):
            raise ComparisonQualityError(
                f'Word-resultatet kunde inte stämmas av mot den {label} versionen. '
                'Ingen PDF har ersatt din resultatfil. Dokumentparet behöver granskas.')
