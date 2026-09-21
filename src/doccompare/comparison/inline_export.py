"""Remove formatting/cell revision callouts from a private Word render copy."""
from pathlib import Path
import zipfile

from lxml import etree

from .revisions import (ComparisonQualityError, PROPERTY_REVISIONS, STORY,
                        content_projection, local)


def prepare_inline_copy(source, destination):
    """Keep current properties and all inline content revisions, omit callouts.

    Word for Mac prints cell-change balloons even with in-line revisions enabled.
    Removing only the revision metadata avoids them without accepting/rejecting
    document content. Cell metadata is removable only when its visible payload
    already has the corresponding inline revision. The validated tracked source
    remains untouched and is still used for validation and summary statistics.
    """
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve():
        raise ValueError('Exportkopian måste vara en separat fil.')
    parts, count = [], 0
    visible = {'t', 'delText', 'tab', 'br', 'cr', 'drawing', 'pict', 'object',
               'sym', 'footnoteReference', 'endnoteReference'}
    payload = {'r', 't', 'delText', 'instrText', 'delInstrText', 'drawing',
               'pict', 'object', 'sym', 'footnoteReference', 'endnoteReference'}
    with zipfile.ZipFile(source) as archive:
        for info in archive.infolist():
            data = archive.read(info)
            if STORY.fullmatch(info.filename):
                tree = etree.fromstring(data)
                markers = [n for n in tree.iter() if local(n) in PROPERTY_REVISIONS
                           and not any(local(a) in PROPERTY_REVISIONS for a in n.iterancestors())]
                for marker in markers:
                    kind = local(marker)
                    if kind in {'cellIns', 'cellDel'}:
                        cell = marker.getparent().getparent()
                        required = {'ins', 'moveTo'} if kind == 'cellIns' else {'del', 'moveFrom'}
                        if local(cell) != 'tc':
                            raise ComparisonQualityError('En celländring saknar en giltig tabellcell.')
                        for node in cell.iter():
                            if local(node) not in visible or any(local(a) in PROPERTY_REVISIONS for a in node.iterancestors()):
                                continue
                            if not any(local(a) in required for a in node.iterancestors()):
                                raise ComparisonQualityError('En celländring saknar motsvarande ändringsmarkering i texten. Ingen PDF har ersatt din resultatfil.')
                    # Historical properties must not hide unexpected visible data.
                    if any(local(n) in payload for n in marker.iterdescendants()):
                        raise ComparisonQualityError('En formatändring innehåller oväntat dokumentinnehåll.')
                    marker.getparent().remove(marker)
                    count += 1
                if markers:
                    data = etree.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            parts.append((info, data))
    with zipfile.ZipFile(destination, 'w') as archive:
        for info, data in parts:
            archive.writestr(info, data)
    if content_projection(source) != content_projection(destination):
        destination.unlink(missing_ok=True)
        raise ComparisonQualityError('Exportkopians innehåll ändrades oväntat när formatmarkeringar togs bort.')
    return count
