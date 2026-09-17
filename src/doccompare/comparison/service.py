"""Single, fail-closed comparison path for GUI, CLI and adapter clients."""
from hashlib import sha256
from pathlib import Path
import os
import shutil
import sys
import tempfile

from .revisions import preflight, revision_ledger, summarize, verify_projections
from .word_bridge import run_comparison, export_document, word_lock, word_workspace
from doccompare.rendering.quality_report import assemble_pdf, render_note, validate_pdf, remove_broken_internal_links


def compare_documents(original, modified, output, *, author='DocCompare',
                      original_name='', modified_name='', progress=None):
    original, modified, output = [Path(p).expanduser().resolve() for p in (original, modified, output)]
    if output.suffix.lower() != '.pdf' or output in {original, modified}:
        raise ValueError('Resultatet måste vara en separat PDF-fil.')
    for source in (original, modified):
        preflight(source)
    if sys.platform != 'darwin' or not Path('/Applications/Microsoft Word.app').exists():
        raise RuntimeError('Denna version kräver Microsoft Word för macOS för att bevara dokumentlayouten.')
    if not output.parent.is_dir():
        raise ValueError('Mappen för resultatfilen finns inte.')
    status = progress or (lambda message: None)
    root = word_workspace()
    with word_lock(root), tempfile.TemporaryDirectory(prefix='job-', dir=root) as directory:
        folder = Path(directory)
        def file(name):
            return folder / (folder.name + '-' + name)
        for source, name in ((original, 'original.docx'), (modified, 'modified.docx')):
            shutil.copyfile(source, file(name))
            # Validate the snapshot, not just a file that may have changed since preflight.
            preflight(file(name))
        status('Jämför med Microsoft Word…')
        version = run_comparison(folder, author)
        status('Kontrollerar resultatet mot båda källversionerna…')
        column_notes = verify_projections(file('original.docx'), file('modified.docx'),
                                          file('accepted.docx'), file('rejected.docx'),
                                          tracked=file('tracked.docx'))
        summary = summarize(revision_ledger(file('tracked.docx')))
        for note in column_notes:
            summary['revisions'].append({
                'kind': 'tblGridChange', 'part': 'word/document.xml', 'paragraph': 0,
                'structural': True, 'text': '', 'scope': 'tblGrid',
                'context': f'Tabell {note["table"]}: {note["old_columns"]} → {note["new_columns"]} kolumner.',
            })
        summary['column_validation_notes'] = column_notes
        status('Exporterar den kontrollerade jämförelsen med Microsoft Word…')
        summary.update(export_document(folder))
        document_pdf = validate_pdf(file('document.pdf'))
        summary['unavailable_internal_links'] = remove_broken_internal_links(document_pdf)
        summary.update(word_version=version, document_pages=len(document_pdf.pages),
                       original_sha256=sha256(file('original.docx').read_bytes()).hexdigest(),
                       modified_sha256=sha256(file('modified.docx').read_bytes()).hexdigest(),
                       engine='microsoft-word', validation='text-projections-passed')
        status('Skapar jämförelsebilaga…')
        note = render_note(summary, original_name or original.name, modified_name or modified.name)
        # Stage beside destination: an atomic rename preserves an existing report
        # on any compare, validation, rendering or write failure.
        fd, staged_name = tempfile.mkstemp(prefix='.doccompare-', suffix='.pdf', dir=output.parent)
        os.close(fd)
        staged = Path(staged_name)
        try:
            assemble_pdf(file('document.pdf'), note, staged)
            os.replace(staged, output)
        finally:
            staged.unlink(missing_ok=True)
        return summary
