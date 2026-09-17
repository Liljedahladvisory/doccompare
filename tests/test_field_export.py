"""Export recovery preserves visible text, markup and layout in a separate copy."""
import zipfile

from docx import Document
from lxml import etree
import pytest

from doccompare.comparison.field_export import prepare_export_copy, _snapshot_fields
from doccompare.comparison.revisions import NS, W, ComparisonQualityError
from doccompare.comparison import word_bridge as bridge


def field(parent, code, value, *, revision='ins', ident='1'):
    container = etree.SubElement(parent, f'{{{W}}}{revision}', {f'{{{W}}}id': ident}) if revision else parent
    for tag, text in [('fldChar', 'begin'), ('instrText', code), ('fldChar', 'separate'),
                      ('delText' if revision == 'del' else 't', value), ('fldChar', 'end')]:
        run = etree.SubElement(container, f'{{{W}}}r')
        props = etree.SubElement(run, f'{{{W}}}rPr')
        etree.SubElement(props, f'{{{W}}}b')
        node = etree.SubElement(run, f'{{{W}}}{tag}')
        if tag == 'fldChar': node.set(f'{{{W}}}fldCharType', text)
        else: node.text = text
    return container


def package(tmp_path):
    path = tmp_path / 'tracked.docx'
    doc = Document()
    p = doc.add_paragraph('Hänvisning: ')
    field(p._p, ' REF old ', '1.1', revision='del')
    field(p._p, ' REF new ', '2.1', ident='2')
    field(p._p, ' PAGE ', '1', ident='3')
    field(p._p, ' NUMPAGES ', '5', ident='4')
    field(p._p, ' REF unchanged ', '3.1', revision=None)
    doc.save(path)
    return path


def test_only_revised_field_controls_are_removed(tmp_path):
    source = package(tmp_path)
    before = source.read_bytes()
    out = tmp_path / 'render.docx'
    assert prepare_export_copy(source, out) == 2
    assert source.read_bytes() == before
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(out) as new:
        for name in old.namelist():
            if name != 'word/document.xml':
                assert old.read(name) == new.read(name)
        expected = etree.fromstring(old.read('word/document.xml'))
        actual = etree.fromstring(new.read('word/document.xml'))
        for rev in expected.xpath('.//w:ins[@w:id="2"] | .//w:del[@w:id="1"]', namespaces=NS):
            for control in rev.xpath('.//w:fldChar | .//w:instrText', namespaces=NS):
                control.getparent().remove(control)
        assert etree.tostring(expected, method='c14n') == etree.tostring(actual, method='c14n')


@pytest.mark.parametrize('damage', ['unclosed', 'no_cache', 'nested_code', 'visible_control'])
def test_unsafe_fields_stop_recovery(damage):
    root = etree.Element(f'{{{W}}}p')
    rev = field(root, ' IF test ', 'visat värde')
    controls = root.findall('.//w:fldChar', NS)
    if damage == 'unclosed':
        controls[-1].getparent().remove(controls[-1])
    elif damage == 'no_cache':
        controls[1].getparent().remove(controls[1])
    elif damage == 'visible_control':
        etree.SubElement(controls[0], f'{{{W}}}t').text = 'innehåll'
    else:
        inner = etree.Element(f'{{{W}}}p')
        field(inner, ' REF x ', '1', revision=None)
        for node in reversed(list(inner)):
            rev.insert(1, node)
    if damage in {'no_cache', 'nested_code'}:
        before = etree.tostring(root)
        assert _snapshot_fields(root) == 0
        assert etree.tostring(root) == before
    else:
        with pytest.raises(ComparisonQualityError):
            _snapshot_fields(root)


def test_source_cannot_be_overwritten(tmp_path):
    source = package(tmp_path)
    with pytest.raises(ValueError):
        prepare_export_copy(source, source)


@pytest.mark.parametrize('code', ['-1743', '-1712', '-50'])
def test_other_word_errors_never_trigger_field_recovery(tmp_path, monkeypatch, code):
    monkeypatch.setattr(bridge, '_run_script', lambda *a: (_ for _ in ()).throw(bridge.WordScriptError(f'({code})')))
    with pytest.raises(bridge.WordScriptError):
        bridge.export_document(tmp_path)
    assert not list(tmp_path.iterdir())


def test_pdf_retry_is_once_and_preserves_tracked_file(tmp_path, monkeypatch):
    source = package(tmp_path)
    tracked = tmp_path / (tmp_path.name + '-tracked.docx')
    source.rename(tracked)
    before = tracked.read_bytes()
    pdf = tmp_path / (tmp_path.name + '-document.pdf')
    calls = []
    def run(script):
        calls.append(script)
        if len(calls) == 1:
            pdf.write_bytes(b'incomplete')
            raise bridge.WordScriptError('(-1708)')
        assert not pdf.exists()
        assert '-render.docx' in script
        pdf.write_bytes(b'PDF validated by caller')
    monkeypatch.setattr(bridge, '_run_script', run)
    assert bridge.export_document(tmp_path) == {'export_mode': 'word-field-snapshot', 'frozen_fields': 2}
    assert len(calls) == 2
    assert tracked.read_bytes() == before
