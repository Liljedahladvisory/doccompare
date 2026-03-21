from setuptools import setup

APP = ['src/doccompare/app.py']
DATA_FILES = [('', ['src/doccompare/rendering/styles.css'])]
OPTIONS = {
    'argv_emulation': False,
    'packages': ['doccompare'],
    'iconfile': None,
    'plist': {
        'CFBundleName': 'DocCompare',
        'CFBundleDisplayName': 'DocCompare',
        'CFBundleVersion': '0.1.0',
        'NSHighResolutionCapable': True,
    },
}
setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
