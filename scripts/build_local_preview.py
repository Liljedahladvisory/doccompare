"""Build an isolated local pilot using an explicitly supplied existing runtime.

This overlays source onto an existing DocCompare.app. It is not a reproducible
release build or a notarized distribution. Never target /Applications.
"""
import argparse
from pathlib import Path
import plistlib
import shutil
import subprocess
import tempfile

parser = argparse.ArgumentParser()
parser.add_argument('--runtime', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True, help='New .zip for the signed local preview')
args = parser.parse_args()
source = Path(__file__).resolve().parents[1] / 'src/doccompare'
runtime, archive = args.runtime.resolve(), args.output.resolve()
if archive.exists() or archive.suffix != '.zip':
    parser.error('Output must be a new .zip file.')
if not (runtime / 'Contents/Resources/lib/python3.12/doccompare').is_dir():
    parser.error('Expected an existing Python 3.12 DocCompare runtime.')
# iCloud/FileProvider adds FinderInfo back to .app directories under Documents.
# Sign on the local temporary volume and deliver an archive preserving that build.
staging = tempfile.TemporaryDirectory(prefix='doccompare-preview-')
output = Path(staging.name) / 'DocCompare Preview.app'
shutil.copytree(runtime, output, symlinks=True)
package = output / 'Contents/Resources/lib/python3.12/doccompare'
shutil.rmtree(package)
shutil.copytree(source, package, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
# py2app also ships the executable entry script at the resource root.
shutil.copyfile(source / 'app.py', output / 'Contents/Resources/app.py')
boot = output / 'Contents/Resources/__boot__.py'
boot.write_text("import os\nos.environ['DOCCOMPARE_PREVIEW'] = '1'\n" + boot.read_text())
plist = output / 'Contents/Info.plist'
info = plistlib.loads(plist.read_bytes())
info.update(CFBundleName='DocCompare Preview', CFBundleDisplayName='DocCompare Preview',
            CFBundleIdentifier='se.liljedahladvisory.doccompare.preview',
            CFBundleVersion='0.3.0', CFBundleShortVersionString='0.3.0',
            NSAppleEventsUsageDescription='DocCompare använder Word för att jämföra arbetskopior och bevara dokumentlayouten.')
plist.write_bytes(plistlib.dumps(info))
# Remove only Finder metadata on paths inside this newly generated bundle.
# Do not mutate metadata through runtime symlinks pointing outside the bundle.
attributes = subprocess.run(['xattr', '-lr', str(output)], capture_output=True, text=True, check=True).stdout
for line in attributes.splitlines():
    for attribute in ('com.apple.FinderInfo', 'com.apple.ResourceFork'):
        marker = ': ' + attribute + ':'
        if marker in line:
            path = Path(line.split(marker, 1)[0])
            if path.resolve().is_relative_to(output):
                subprocess.run(['xattr', '-d', attribute, str(path)], check=True)
subprocess.run(['codesign', '--force', '--deep', '--sign', '-', str(output)], check=True)
subprocess.run(['codesign', '--verify', '--deep', '--strict', str(output)], check=True)
subprocess.run(['ditto', '-c', '-k', '--norsrc', '--keepParent', str(output), str(archive)], check=True)
staging.cleanup()
print(archive)
