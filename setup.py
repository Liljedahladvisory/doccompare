"""
py2app build configuration for DocCompare.

Build the .app bundle:
    pip install py2app
    python setup.py py2app

The resulting app will be in dist/DocCompare.app
"""
from setuptools import setup

APP = ['src/doccompare/app.py']
DATA_FILES = [
    ('doccompare/rendering', ['src/doccompare/rendering/styles.css']),
    ('doccompare/assets', ['src/doccompare/assets/logo.png']),
]
OPTIONS = {
    'argv_emulation': False,
    'packages': [
        'doccompare',
        'doccompare.parsers',
        'doccompare.comparison',
        'doccompare.rendering',
        # Core dependencies
        'docx',              # python-docx
        'diff_match_patch',
        'rapidfuzz',
        'weasyprint',
        'click',
        'loguru',
        'rich',
        'PIL',               # Pillow
        'pypdf',
        'lxml',
        # WeasyPrint transitive dependencies
        'cssselect2',
        'tinycss2',
        'cffi',
        'pydyf',
    ],
    'includes': [
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'json',
        'pathlib',
    ],
    'excludes': [
        'test',
        'tests',
        'pytest',
        'setuptools',
        'pip',
    ],
    'iconfile': 'DocCompare.icns',
    'plist': {
        'CFBundleName': 'DocCompare',
        'CFBundleDisplayName': 'DocCompare',
        'CFBundleIdentifier': 'se.liljedahladvisory.doccompare',
        'CFBundleVersion': '0.1.0',
        'CFBundleShortVersionString': '0.1.0',
        'NSHighResolutionCapable': True,
        'NSHumanReadableCopyright': '\u00a9 2025 Liljedahl Advisory AB',
        'LSMinimumSystemVersion': '12.0',
    },
    'frameworks': [],
    'resources': [],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
)
