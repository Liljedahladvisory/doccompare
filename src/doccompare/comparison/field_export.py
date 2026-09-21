"""A display-only snapshot for Word's PDF bug with revised complex fields.

The validated tracked package remains untouched. Only field control nodes are
removed in a private render copy; cached text, its revisions, runs, paragraphs,
tables, drawings and section properties stay exactly where Word put them.
"""
from pathlib import Path
import zipfile

from lxml import etree

from .revisions import ComparisonQualityError, CONTENT_REVISIONS, STORY, W, local


def _snapshot_fields(tree):
    stack, fields = [], []
    for node in tree.iter():
        tag = local(node)
        if tag == 'fldChar':
            kind = node.get(f'{{{W}}}fldCharType')
            if kind == 'begin':
                stack.append({'nodes': [node], 'code': '', 'result': False,
                              'nested_code': False, 'in_code': any(not f['result'] for f in stack)})
                if len(stack) > 1 and not stack[-2]['result']:
                    stack[-2]['nested_code'] = True
            elif not stack or kind not in {'separate', 'end'}:
                raise ComparisonQualityError('Word-exporten innehåller ett ofullständigt fält.')
            else:
                field = stack[-1]
                field['nodes'].append(node)
                if kind == 'separate':
                    if field['result']:
                        raise ComparisonQualityError('Word-exporten innehåller ett tvetydigt fält.')
                    field['result'] = True
                else:
                    fields.append(stack.pop())
        elif tag in {'instrText', 'delInstrText'} and stack:
            stack[-1]['nodes'].append(node)
            stack[-1]['code'] += node.text or ''
    if stack:
        raise ComparisonQualityError('Word-exporten innehåller ett oavslutat fält.')

    frozen = 0
    for field in fields:
        code = field['code'].strip().split()
        revised = any(local(a) in CONTENT_REVISIONS
                      for n in field['nodes'] for a in n.iterancestors())
        # Pagination must describe the comparison PDF, not an input document.
        if not revised or (code and code[0].upper() in {'PAGE', 'NUMPAGES', 'SECTIONPAGES'}):
            continue
        if not code or not field['result'] or field['nested_code'] or field['in_code']:
            # Leave fields without a safely separable cached result active.
            # If Word still cannot export them, the second attempt fails closed.
            continue
        for node in field['nodes']:
            # A field control may contain legacy form metadata, but never
            # remove any visible payload, even from an unusual input package.
            if any(local(child) in {'t', 'delText', 'drawing', 'pict', 'object'}
                   for child in node.iterdescendants()):
                raise ComparisonQualityError('Ett fält innehåller oväntat visningsinnehåll.')
            node.getparent().remove(node)
        frozen += 1
    return frozen


def prepare_export_copy(source, destination):
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve():
        raise ValueError('Exportkopian måste vara en separat fil.')
    parts, frozen = [], 0
    with zipfile.ZipFile(source) as archive:
        for info in archive.infolist():
            data = archive.read(info)
            if STORY.fullmatch(info.filename):
                tree = etree.fromstring(data)
                count = _snapshot_fields(tree)
                if count:
                    data = etree.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
                    frozen += count
            parts.append((info, data))
    if frozen:
        with zipfile.ZipFile(destination, 'w') as archive:
            for info, data in parts:
                archive.writestr(info, data)
    return frozen
