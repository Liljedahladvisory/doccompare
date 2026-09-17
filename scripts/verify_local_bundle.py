"""Check imports through the bundle's actual py2app boot hooks, not plain Python.

Run using the bundle's Contents/MacOS/python. Optional synthetic input paths
exercise the same comparison service that the GUI worker imports.
"""
import argparse
import ast
import importlib
from pathlib import Path
import os
import sys

parser = argparse.ArgumentParser()
parser.add_argument('app', type=Path)
parser.add_argument('--original', type=Path)
parser.add_argument('--modified', type=Path)
parser.add_argument('--output', type=Path)
args = parser.parse_args()
app = args.app.resolve()
inputs = [p.resolve() if p else None for p in (args.original, args.modified, args.output)]
if any(inputs) and not all(inputs):
    parser.error('Supply original, modified and output together.')
resources = app / 'Contents/Resources'
lib_dir = resources / 'lib' / f'python{sys.version_info.major}.{sys.version_info.minor}'
sys.path[:0] = [str(lib_dir), str(lib_dir / 'lib-dynload'),
               str(resources / 'lib' / f'python{sys.version_info.major}{sys.version_info.minor}.zip')]
os.environ.update(RESOURCEPATH=str(resources), ARGVZERO=str(app / 'Contents/MacOS/DocCompare'))
sys.frozen = 'macosx_app'
boot = resources / '__boot__.py'
tree = ast.parse(boot.read_text(), filename=str(boot))
# Execute the shipping bootstrap, only replacing its final GUI entry invocation.
last = tree.body[-1]
if not (isinstance(last, ast.Expr) and isinstance(last.value, ast.Call)
        and isinstance(last.value.func, ast.Name) and last.value.func.id == '_run'):
    raise RuntimeError('Unrecognized py2app bootstrap; verifier needs review.')
tree.body.pop()
exec(compile(tree, str(boot), 'exec'), {'__file__': str(boot), '__name__': '__main__'})

modules = ['comparison.service', 'comparison.revisions', 'comparison.word_bridge',
           'comparison.field_export', 'comparison.story_projection',
           'rendering.quality_report', 'parsers.docx_parser', 'gui']
for name in modules:
    module = importlib.import_module('doccompare.' + name)
    loaded = Path(module.__file__).resolve()
    expected = lib_dir / 'doccompare' / Path(*name.split('.')).with_suffix('.py')
    if not loaded.is_relative_to(resources) or loaded.read_bytes() != expected.read_bytes():
        raise RuntimeError(f'Stale or external module loaded: {name}: {loaded}')
    print(f'BOOT IMPORT OK {name}: {loaded}', flush=True)

if all(inputs):
    from doccompare.comparison.service import compare_documents
    summary = compare_documents(*inputs, progress=lambda text: print(text, flush=True))
    from pypdf import PdfReader
    reader = PdfReader(inputs[2])
    assert len(reader.pages) > summary['document_pages']
    assert summary['validation'] == 'text-projections-passed'
    print(f'COMPARISON OK: {len(reader.pages)} PDF pages; '
          f'+{summary["added_words"]}/-{summary["deleted_words"]} words', flush=True)
