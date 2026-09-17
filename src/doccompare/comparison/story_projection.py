"""Preserve inherited stories when Word removes revised section breaks."""
from copy import deepcopy
from pathlib import Path
import zipfile

from lxml import etree

from .revisions import ComparisonQualityError, PROPERTY_REVISIONS, W, content_projection, local


def prepare_projection_copy(source, destination):
    """Make existing section inheritance explicit in the validation copy only.

    Word can discard an inherited header/footer reference when accepting a
    deleted section break. Copying that same reference to each inheriting
    section preserves the displayed document and Word's actual story revisions.
    No header/footer text, relationship target or historical properties change.
    """
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve():
        raise ValueError('Kontrollkopian måste vara en separat fil.')
    count = 0
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(destination, 'w') as new:
        for info in old.infolist():
            data = old.read(info)
            if info.filename == 'word/document.xml':
                tree = etree.fromstring(data)
                effective = {}
                for section in tree.iter(f'{{{W}}}sectPr'):
                    if any(local(a) in PROPERTY_REVISIONS for a in section.iterancestors()):
                        continue
                    explicit = {}
                    for ref in section:
                        if local(ref) in {'headerReference', 'footerReference'}:
                            key = (ref.tag, ref.get(f'{{{W}}}type', 'default'))
                            if key in explicit:
                                raise ComparisonQualityError('Ett avsnitt har motstridiga sidhuvuds- eller sidfotsreferenser.')
                            explicit[key] = ref
                    effective.update(explicit)
                    for key, ref in effective.items():
                        if key not in explicit:
                            section.insert(0, deepcopy(ref))
                            count += 1
                if count:
                    data = etree.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            new.writestr(info, data)
    if content_projection(source) != content_projection(destination):
        raise ComparisonQualityError('Kontrollkopians sidhuvuden eller sidfötter ändrades oväntat.')
    return count
