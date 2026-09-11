"""Append a comparison note without reflowing any Word-rendered page."""
from datetime import datetime
from html import escape
from io import BytesIO

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError
from pypdf.generic import ArrayObject, NameObject, NullObject


def remove_broken_internal_links(reader):
    """Drop only unusable Link annotations emitted by Word for deleted targets.

    Page content, external links and valid internal links are retained. Other
    PDF parsing errors still propagate. The count is shown in the summary.
    """
    removed = 0
    for page in reader.pages:
        annotations = page['/Annots'] if '/Annots' in page else []
        kept = []
        for reference in annotations:
            annotation = reference.get_object()
            if annotation.get('/Subtype') == '/Link' and '/Dest' in annotation:
                try:
                    target = annotation['/Dest']
                    if isinstance(target, ArrayObject) and target:
                        target = target[0].get_object()
                    broken = target is None or isinstance(target, NullObject)
                except PdfReadError:
                    broken = True
                if broken:
                    removed += 1
                    continue
            kept.append(reference)
        if len(kept) != len(annotations):
            page[NameObject('/Annots')] = ArrayObject(kept)
    return removed


def validate_pdf(path):
    reader = PdfReader(str(path), strict=True)
    if reader.is_encrypted or not reader.pages:
        raise RuntimeError('Word skapade inte en läsbar PDF med innehåll.')
    for page in reader.pages:
        if float(page.mediabox.width) <= 0 or float(page.mediabox.height) <= 0:
            raise RuntimeError('Word skapade en PDF med ogiltigt sidformat.')
        if page.get_contents() is None:
            raise RuntimeError('Word skapade en PDF med en ofullständig sida.')
    return reader


def render_note(summary, original_name, modified_name):
    from weasyprint import HTML
    added_label = 'ord tillagt' if summary['added_words'] == 1 else 'ord tillagda'
    deleted_label = 'ord borttaget' if summary['deleted_words'] == 1 else 'ord borttagna'
    moved = summary['moved_words']
    moved_label = 'ord flyttat' if moved == 1 else 'ord flyttade'
    moved_stat = f'<p class="moved">{moved} {moved_label}</p>' if moved else ''
    moved_legend = ('<p><span class="moved double">Flyttad text</span>: text som har flyttats inom dokumentet.</p>') if moved else ''
    missing_links = summary.get('unavailable_internal_links', 0)
    link_note = (f'<p class="link-note">{missing_links} interna länkar saknar ett giltigt mål i Words PDF-export. '
                 'Länktexten finns kvar, men dessa länkar går inte att klicka på.</p>') if missing_links else ''
    html = f'''<!doctype html><html lang="sv"><meta charset="utf-8"><style>
    @page {{size:A4; margin:20mm; @bottom-left {{content:"DocCompare";font-family:Arial,sans-serif;font-size:8pt;color:#888}}
    @bottom-right {{content:"Sammanfattning " counter(page) " av " counter(pages);font-family:Arial,sans-serif;font-size:8pt;color:#888}}}}
    body {{font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size:10pt;line-height:1.45;color:#222;margin:0}}
    h1 {{font-size:18pt;line-height:1.25;margin:0 0 8pt;color:#2c3e50}}
    h2 {{font-size:12pt;margin:0 0 8pt;color:#2c3e50;break-after:avoid}}
    p {{margin:0 0 5pt;orphans:2;widows:2}}
    .meta {{font-size:9pt;color:#555;border-top:.6pt solid #bdc3c7;padding-top:8pt;margin-bottom:18pt}}
    .meta p {{margin-bottom:3pt;overflow-wrap:anywhere}}
    .stats {{font-size:11pt;margin-bottom:22pt;break-inside:avoid}}
    .added {{color:#2e97d3}} .deleted {{color:#b5082e}} .moved {{color:#1a7a3f}}
    .underline {{text-decoration:underline}} .strike {{text-decoration:line-through}} .double {{text-decoration:underline double}}
    .legend {{break-inside:avoid}} .link-note {{margin-top:12pt;color:#555}}
    footer {{margin-top:24pt;border-top:.6pt solid #bdc3c7;padding-top:8pt;color:#888;font-size:8pt;break-inside:avoid}}
    </style><body><h1>DocCompare – Sammanfattning</h1>
    <div class="meta"><p><b>Original:</b> {escape(original_name)}</p><p><b>Modifierat:</b> {escape(modified_name)}</p>
    <p><b>Datum:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p></div>
    <div class="stats"><p class="added">+{summary['added_words']} {added_label}</p>
    <p class="deleted">−{summary['deleted_words']} {deleted_label}</p>{moved_stat}</div>
    <section class="legend"><h2>Teckenförklaring</h2>
    <p><span class="added underline">Tillagd text</span>: text som finns i den modifierade versionen men inte i originalet.</p>
    <p><span class="deleted strike">Borttagen text</span>: text som finns i originalet men inte i den modifierade versionen.</p>
    {moved_legend}
    <p>Oförändrad text: text som är identisk i båda versionerna.</p></section>
    {link_note}
    <footer>Genererad av DocCompare · a Liljedahl Legal Tech product</footer>
    </body></html>'''
    # Keep technical provenance available without placing it on the client-facing page.
    writer = PdfWriter()
    writer.append(PdfReader(BytesIO(HTML(string=html).write_pdf())), import_outline=False)
    writer.add_metadata({
        '/DocCompareOriginalSHA256': summary['original_sha256'],
        '/DocCompareModifiedSHA256': summary['modified_sha256'],
        '/DocCompareWordVersion': summary['word_version'],
        '/DocCompareValidation': summary.get('validation', 'not-provided'),
    })
    result = BytesIO()
    writer.write(result)
    return result.getvalue()


def assemble_pdf(main_pdf, note, destination):
    main = validate_pdf(main_pdf)
    remove_broken_internal_links(main)
    writer = PdfWriter()
    writer.append(main, import_outline=False)
    note_reader = PdfReader(BytesIO(note))
    writer.append(note_reader, import_outline=False)
    writer.add_outline_item('Jämförelsedokument', 0)
    writer.add_outline_item('Sammanfattning', len(main.pages))
    writer.add_metadata({key: value for key, value in (note_reader.metadata or {}).items()
                         if key.startswith('/DocCompare') and isinstance(value, str)})
    writer.add_metadata({'/Title': 'DocCompare – dokumentjämförelse', '/Creator': 'DocCompare / Microsoft Word'})
    with open(destination, 'wb') as stream:
        writer.write(stream)
    result = validate_pdf(destination)
    if len(result.pages) != len(main.pages) + len(PdfReader(BytesIO(note)).pages):
        raise RuntimeError('Sidantalet ändrades oväntat vid sammanfogningen.')
    for source, merged in zip(main.pages, result.pages):
        if source.mediabox != merged.mediabox or source.get_contents().get_data() != merged.get_contents().get_data():
            raise RuntimeError('Dokumentsidornas innehåll ändrades vid sammanfogningen.')
