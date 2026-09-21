from io import BytesIO

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject,
                           IndirectObject, NameObject, NumberObject, TextStringObject)

from doccompare.rendering.quality_report import assemble_pdf, remove_broken_internal_links


def pdf_with_links():
    writer = PdfWriter()
    for _ in range(2):
        page = writer.add_blank_page(300, 400)
        stream = DecodedStreamObject()
        stream.set_data(b'q Q')
        page[NameObject('/Contents')] = writer._add_object(stream)
    def annotation():
        return DictionaryObject({NameObject('/Subtype'): NameObject('/Link'),
                                 NameObject('/Rect'): ArrayObject([NumberObject(v) for v in (0, 0, 30, 30)])})
    valid, broken, external = [annotation() for _ in range(3)]
    valid[NameObject('/Dest')] = ArrayObject([writer.pages[1].indirect_reference, NameObject('/Fit')])
    broken[NameObject('/Dest')] = IndirectObject(9999, 0, writer)
    external[NameObject('/A')] = DictionaryObject({NameObject('/S'): NameObject('/URI'),
                                                 NameObject('/URI'): TextStringObject('https://example.com/reference')})
    writer.pages[0][NameObject('/Annots')] = writer._add_object(ArrayObject([
        writer._add_object(link) for link in (valid, broken, external)]))
    result = BytesIO()
    writer.write(result)
    return result.getvalue()


def test_broken_destination_does_not_remove_valid_links_or_page_content(tmp_path):
    source = tmp_path / 'word.pdf'
    source.write_bytes(pdf_with_links())
    reader = PdfReader(source, strict=True)
    original_content = reader.pages[0].get_contents().get_data()
    assert remove_broken_internal_links(reader) == 1
    assert len(reader.pages[0]['/Annots']) == 2
    assert reader.pages[0].get_contents().get_data() == original_content
    note_writer = PdfWriter()
    page = note_writer.add_blank_page(300, 400)
    stream = DecodedStreamObject()
    stream.set_data(b'q Q')
    page[NameObject('/Contents')] = note_writer._add_object(stream)
    note = BytesIO()
    note_writer.write(note)
    output = tmp_path / 'report.pdf'
    assemble_pdf(source, note.getvalue(), output)
    merged = PdfReader(output, strict=True)
    links = [ref.get_object() for ref in merged.pages[0]['/Annots']]
    assert len(merged.pages) == 3 and len(links) == 2
    assert links[0]['/Dest'][0] == merged.pages[1].indirect_reference
    assert links[1]['/A']['/URI'] == 'https://example.com/reference'
    assert merged.pages[0].get_contents().get_data() == original_content
