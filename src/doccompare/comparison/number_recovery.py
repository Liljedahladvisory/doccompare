"""Recover a narrowly proven Word omission of a replaced literal heading number."""
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import re
import zipfile

from lxml import etree

from .revisions import NS, W, content_projection, local, visible_text


def _plain(paragraph, *, deleted=False):
    """Only ordinary body paragraphs; no fields, lists, links or other revisions."""
    if paragraph.getparent().tag != f'{{{W}}}body':
        return False
    if paragraph.find('w:pPr/w:numPr', NS) is not None:
        return False
    for child in paragraph:
        if local(child) == 'pPr':
            if child.xpath('.//w:pPrChange | .//w:sectPr | .//w:rPr/w:ins | .//w:rPr/w:del', namespaces=NS):
                return False
            continue
        runs = list(child) if deleted and local(child) == 'del' else [child]
        for run in runs:
            if local(run) != 'r' or any(local(n) not in {'rPr', 't', 'delText'} for n in run):
                return False
            if run.xpath('.//w:rPrChange', namespaces=NS):
                return False
    return True


def recover_missing_numbers(original, modified, accepted, rejected, tracked, destination, author):
    """Write a candidate, never a validated report. Caller MUST re-run Word checks.

    Requires exact old projection and equal new story/row/cell/paragraph sequence,
    except missing leading digits before a period/parenthesis. A unique plain
    body paragraph must contain Word's deletion of the old digits and the exact
    retained suffix. All other ZIP members and paragraph content are preserved.
    Ambiguous or unrelated damage returns zero without creating a candidate.
    """
    if Path(tracked).resolve() == Path(destination).resolve():
        raise ValueError('Återhämtningskopian måste vara en separat fil.')
    if content_projection(original) != content_projection(rejected):
        return 0
    expected, actual = content_projection(modified), content_projection(accepted)
    old_bodies = [tokens for (family, tokens), count in actual.items() if family == 'document' and count == 1]
    new_bodies = [tokens for (family, tokens), count in expected.items() if family == 'document' and count == 1]
    if len(old_bodies) != 1 or len(new_bodies) != 1:
        return 0
    before, after = old_bodies[0], new_bodies[0]
    if len(before) != len(after):
        return 0
    actual.pop(('document', before))
    expected.pop(('document', after))
    if actual != expected:
        return 0
    gaps = []
    for a, b in zip(before, after):
        if a == b:
            continue
        if a[0] != 'p' or b[0] != 'p':
            return 0
        match = re.fullmatch(r'([0-9]{1,4})([.)]\s.+)', b[1])
        if not match or a[1] != match[2]:
            return 0
        gaps.append((a[1], b[1], match[1]))
    if not gaps or len(gaps) > 20:
        return 0
    def document(path):
        with zipfile.ZipFile(path) as archive:
            return etree.fromstring(archive.read('word/document.xml'))
    source, prior, tree = document(modified), document(original), document(tracked)
    changes = []
    for suffix, text, digits in gaps:
        sources = [p for p in source.iter(f'{{{W}}}p') if visible_text(p) == text]
        candidates = [p for p in tree.iter(f'{{{W}}}p')
                      if ''.join(visible_text(r) for r in p.findall('w:r', NS)) == suffix]
        if len(sources) != 1 or len(candidates) != 1:
            return 0
        src, paragraph = sources[0], candidates[0]
        if not _plain(src) or not _plain(paragraph, deleted=True):
            return 0
        children = [n for n in paragraph if local(n) != 'pPr']
        if not children or local(children[0]) != 'del' or any(local(n) != 'r' for n in children[1:]):
            return 0
        removed = visible_text(children[0])
        if not re.fullmatch(r'[0-9]{1,4}', removed) or removed == digits:
            return 0
        originals = [p for p in prior.iter(f'{{{W}}}p') if visible_text(p) == removed + suffix]
        if len(originals) != 1 or not _plain(originals[0]):
            return 0
        changes.append((paragraph, children[0], digits))
    ids = [int(n.get(f'{{{W}}}id')) for n in tree.iter()
           if (n.get(f'{{{W}}}id') or '').isdigit()]
    next_id = max(ids, default=0) + 1
    for offset, (paragraph, deletion, digits) in enumerate(changes):
        insertion = etree.Element(f'{{{W}}}ins', {
            f'{{{W}}}id': str(next_id + offset), f'{{{W}}}author': author,
            f'{{{W}}}date': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        })
        run = etree.SubElement(insertion, f'{{{W}}}r')
        properties = paragraph.find('w:r/w:rPr', NS)
        if properties is not None:
            run.append(deepcopy(properties))
        etree.SubElement(run, f'{{{W}}}t').text = digits
        paragraph.insert(paragraph.index(deletion) + 1, insertion)
    with zipfile.ZipFile(tracked) as old, zipfile.ZipFile(destination, 'w') as new:
        for info in old.infolist():
            data = (etree.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
                    if info.filename == 'word/document.xml' else old.read(info))
            new.writestr(info, data)
    return len(changes)
