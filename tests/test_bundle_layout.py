"""Regression: py2app hooks must never import an obsolete subpackage."""
import importlib.util
from pathlib import Path
import subprocess
import sys

import pytest

spec = importlib.util.spec_from_file_location(
    'bundle_layout', Path(__file__).parents[1] / 'scripts/bundle_layout.py')
layout = importlib.util.module_from_spec(spec)
spec.loader.exec_module(layout)


def test_relocated_import_uses_current_source(tmp_path):
    source = tmp_path / 'source'
    package = source / 'comparison'
    package.mkdir(parents=True)
    (package / '__init__.py').write_text('')
    (package / 'service.py').write_text('VERSION = "current"\n')
    lib = tmp_path / 'lib'
    parent = lib / 'doccompare'
    parent.mkdir(parents=True)
    (parent / '__init__.py').write_text('')
    stale = lib / 'doccompare.comparison'
    stale.mkdir()
    (stale / '__init__.py').write_text('')
    (stale / 'obsolete.py').write_text('VERSION = "old"\n')
    boot = tmp_path / 'boot.py'
    boot.write_text('_path_hooks = ["doccompare.comparison", "unrelated"]\n')
    probe = '''
import importlib.util, importlib.machinery, sys
from pathlib import Path
lib = Path(sys.argv[1])
sys.path.insert(0, str(lib))
class Finder:
    def find_spec(self, fullname, path=None, target=None):
        if fullname == 'doccompare.comparison':
            loader = importlib.machinery.SourceFileLoader(fullname, str(lib / fullname / '__init__.py'))
            return importlib.util.spec_from_loader(fullname, loader)
sys.meta_path.insert(0, Finder())
from doccompare.comparison.service import VERSION
assert VERSION == 'current'
'''
    before = subprocess.run([sys.executable, '-c', probe, str(lib)], capture_output=True, text=True)
    assert before.returncode != 0
    assert "No module named 'doccompare.comparison.service'" in before.stderr
    assert layout.refresh_relocated_packages(source, lib, boot) == ['doccompare.comparison']
    subprocess.run([sys.executable, '-c', probe, str(lib)], check=True)
    assert not (stale / 'obsolete.py').exists()


def test_missing_relocated_source_stops_build(tmp_path):
    boot = tmp_path / 'boot.py'
    boot.write_text('_path_hooks = ["doccompare.retired"]\n')
    with pytest.raises(ValueError, match='no current source'):
        layout.refresh_relocated_packages(tmp_path / 'src', tmp_path / 'lib', boot)
