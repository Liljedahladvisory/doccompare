"""Append a comparison note without reflowing any Word-rendered page."""
from datetime import datetime
from html import escape
from io import BytesIO
import re

from pypdf import PdfReader, PdfWriter


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
    names = {'document': 'Dokument', 'header': 'Sidhuvud', 'footer': 'Sidfot',
             'footnotes': 'Fotnoter', 'endnotes': 'Slutnoter'}
    kinds = {'ins': 'Tillagt', 'del': 'Borttaget', 'moveFrom': 'Flyttat från',
             'moveTo': 'Flyttat till', 'rPrChange': 'Teckenformat',
             'pPrChange': 'Styckeformat', 'tblPrChange': 'Tabellformat',
             'tblGridChange': 'Tabellkolumner', 'trPrChange': 'Radformat',
             'tcPrChange': 'Cellformat', 'sectPrChange': 'Avsnittsformat',
             'numberingChange': 'Numrering', 'cellIns': 'Tillagd cell',
             'cellDel': 'Borttagen cell', 'cellMerge': 'Sammanfogade celler'}
    rows = []
    # Include non-body and structural revisions because inline markup alone
    # cannot explain every formatting, table or header/footer change.
    for r in summary['revisions']:
        if r['structural'] and r['kind'] in {'ins', 'del'} and r.get('scope') in {'rPr', 'pPr'}:
            continue
        if not r['structural'] and r['part'] == 'word/document.xml':
            continue
        part = re.sub(r'\d+', '', r['part'].split('/')[-1].replace('.xml', ''))
        location = names.get(part, part)
        if r['paragraph']:
            location += f", stycke {r['paragraph']}"
        description = r['text'] or r.get('context') or 'Struktur eller formatering ändrad'
        if r.get('scope') == 'trPr':
            location += ', tabellrad'
        rows.append(f'<tr><td>{escape(location)}</td><td>{escape(kinds.get(r["kind"], "Formatändring"))}</td><td>{escape(description)}</td></tr>')
    details = ('<h2>Struktur, format och övriga dokumentdelar</h2>'
               '<p>Styckenumren är interna positioner i jämförelsedokumentet, inte klausulnummer. '
               'Words formatändringar kan även omfatta normalisering av dokumentformat.</p>'
               '<table><thead><tr><th>Plats</th><th>Ändring</th><th>Text / sammanhang</th></tr></thead><tbody>'
               + ''.join(rows) + '</tbody></table>') if rows else ''
    html = f'''<!doctype html><html lang="sv"><meta charset="utf-8"><style>
    @page {{size:A4; margin:22mm 20mm 20mm; @bottom-left {{content:"Liljedahl Advisory · DocCompare";font-size:8pt;color:#697580}}
    @bottom-right {{content:"Jämförelsebilaga " counter(page);font-size:8pt;color:#697580}}}}
    body {{font-family:"Arial",sans-serif;font-size:10pt;line-height:1.45;color:#243443}}
    h1 {{font-size:24pt;line-height:1.1;margin:0 0 8mm;color:#162a3b}}
    h2 {{font-size:12pt;margin:6mm 0 3mm;color:#162a3b}}
    .eyebrow {{font-size:9pt;letter-spacing:1.4pt;color:#526977;margin-bottom:5mm}}
    .meta {{border-top:1pt solid #9aabb5;border-bottom:1pt solid #d9e0e5;padding:4mm 0;overflow-wrap:anywhere}}
    .stats {{font-size:14pt;margin:6mm 0}} .add {{color:#2e97d3;text-decoration:underline}} .del {{color:#b5082e;text-decoration:line-through}}
    p {{margin:2mm 0}} .muted {{font-size:8.5pt;color:#536775}} .hash {{font-family:monospace;font-size:7pt;overflow-wrap:anywhere}}
    table {{border-collapse:collapse;width:100%;font-size:8.5pt;table-layout:fixed}}
    th,td {{text-align:left;vertical-align:top;padding:2mm 2mm;border-bottom:.5pt solid #d9e0e5;overflow-wrap:anywhere}}
    th {{background:#edf1f4}} th:nth-child(1) {{width:25%}} th:nth-child(2) {{width:22%}} tr {{break-inside:avoid}} thead {{display:table-header-group}}
    </style><body><div class="eyebrow">DOCCOMPARE / JÄMFÖRELSEBILAGA</div><h1>Jämförelseöversikt</h1>
    <div class="meta"><p><b>Tidigare version:</b> {escape(original_name)}</p><p><b>Ny version:</b> {escape(modified_name)}</p>
    <p><b>Skapad:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p></div>
    <p class="stats">+{summary['added_words']} tillagda ord &nbsp; · &nbsp; −{summary['deleted_words']} borttagna ord</p>
    <p>{summary['revision_count']} revisionsposter, varav {summary['format_revision_count']} format- eller celländringar.
    Flyttad text: {summary['moved_words']} ord enligt Words flyttmarkeringar.</p>
    <h2>Så läser du dokumentet</h2><p><span class="add">Blå, understruken text</span> har lagts till.
    <span class="del">Röd, överstruken text</span> har tagits bort. Grön dubbelmarkering visar flyttad text när Word identifierar en flytt. Violett markerar ändrad formatering.</p>
    <p>Dokumentets sidor har satts av Microsoft Word med ändringarna i löptexten.
    Ändringarnas längd kan påverka rad- och sidbrytningar.</p>
    <h2>Utförda kontroller</h2><p>Text, fältinstruktioner, länkmål, bildreferenser och cellgränser har stämts av i båda riktningarna:
    accepterade ändringar mot den nya versionen och avvisade ändringar mot den tidigare versionen.</p>
    <p class="muted">Kontrollen är ingen fullständig visuell avstämning. Granska särskilt numrering,
    tabellgeometri, bilder och sidbrytningar före extern leverans. Ordräkningen summerar Words revisionsfragment.</p>
    {details}
    <h2>Spårbarhet</h2><p class="muted">Motor: Microsoft Word {escape(summary['word_version'])} · {summary['document_pages']} dokumentsidor.
    Bilagan skapas separat och fogas sist utan att sätta om dokumentet.</p>
    <p class="hash">Tidigare SHA-256: {summary['original_sha256']}<br>Ny SHA-256: {summary['modified_sha256']}</p>
    </body></html>'''
    return HTML(string=html).write_pdf()


def assemble_pdf(main_pdf, note, destination):
    main = validate_pdf(main_pdf)
    writer = PdfWriter()
    writer.append(main, import_outline=False)
    writer.append(PdfReader(BytesIO(note)), import_outline=False)
    writer.add_outline_item('Jämförelsedokument', 0)
    writer.add_outline_item('Jämförelsebilaga', len(main.pages))
    writer.add_metadata({'/Title': 'DocCompare – dokumentjämförelse', '/Creator': 'DocCompare / Microsoft Word'})
    with open(destination, 'wb') as stream:
        writer.write(stream)
    result = validate_pdf(destination)
    if len(result.pages) != len(main.pages) + len(PdfReader(BytesIO(note)).pages):
        raise RuntimeError('Sidantalet ändrades oväntat vid sammanfogningen.')
    for source, merged in zip(main.pages, result.pages):
        if source.mediabox != merged.mediabox or source.get_contents().get_data() != merged.get_contents().get_data():
            raise RuntimeError('Dokumentsidornas innehåll ändrades vid sammanfogningen.')
