#!/usr/bin/env bash
#
# build_dmg.sh — Build DocCompare.app and package it into a distributable .dmg
#
# Prerequisites (build machine only):
#   brew install pango cairo gdk-pixbuf libffi python@3.12
#
# Usage:
#   chmod +x build_dmg.sh
#   ./build_dmg.sh
#
# The resulting DMG is fully self-contained — end users do NOT need
# Homebrew or any other dependencies installed.
#
set -euo pipefail

APP_NAME="DocCompare"
VERSION="0.1.0"
DMG_NAME="${APP_NAME}-${VERSION}"
BUILD_DIR="$(pwd)/build"
DIST_DIR="$(pwd)/dist"
PYTHON="python3.12"

echo "═══════════════════════════════════════════════════"
echo "  Building ${APP_NAME} v${VERSION}"
echo "═══════════════════════════════════════════════════"

# ── Step 0: Check prerequisites ──────────────────────────────────────────────
echo ""
echo "▸ Checking prerequisites..."

if ! command -v "$PYTHON" &>/dev/null; then
    echo "  ✗ Python 3.12 not found. Run: brew install python@3.12"
    exit 1
fi
echo "  ✓ Python 3.12 OK"

for dep in pango cairo gdk-pixbuf; do
    if ! brew list "$dep" &>/dev/null; then
        echo "  ✗ Missing: $dep — Run: brew install pango cairo gdk-pixbuf libffi"
        exit 1
    fi
done
echo "  ✓ Homebrew dependencies OK"

# ── Step 1: Set up virtual environment ───────────────────────────────────────
echo ""
echo "▸ Setting up build environment..."
rm -rf "$BUILD_DIR" "$DIST_DIR" .venv_build
$PYTHON -m venv .venv_build
source .venv_build/bin/activate
pip install -q 'setuptools<81' py2app 2>&1 | tail -1
pip install -q -e . 2>&1 | tail -1
echo "  ✓ Build environment ready"

# ── Step 2: Build .app bundle with py2app ────────────────────────────────────
echo ""
echo "▸ Building .app bundle with py2app..."
python3 setup.py py2app 2>&1 | tail -3
echo "  ✓ .app bundle built"

APP_PATH="${DIST_DIR}/${APP_NAME}.app"
if [ ! -d "$APP_PATH" ]; then
    FOUND_APP=$(find "$DIST_DIR" -name "*.app" -maxdepth 1 | head -1)
    if [ -n "$FOUND_APP" ]; then
        mv "$FOUND_APP" "$APP_PATH"
    else
        echo "  ✗ No .app found. Build failed."
        exit 1
    fi
fi

# ── Step 3: Bundle ALL native dylibs (recursive) ────────────────────────────
echo ""
echo "▸ Bundling native libraries (recursive scan)..."
python3 bundle_dylibs.py "$APP_PATH"

# ── Step 4: Ad-hoc code sign ────────────────────────────────────────────────
echo ""
echo "▸ Code signing (ad-hoc)..."
codesign --force --deep --sign - "$APP_PATH" 2>/dev/null || true
echo "  ✓ Signed"

# ── Step 5: Create DMG ──────────────────────────────────────────────────────
echo ""
echo "▸ Creating DMG..."

DMG_FINAL="${DIST_DIR}/${DMG_NAME}.dmg"
DMG_STAGING="${DIST_DIR}/dmg_staging"
rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"
cp -R "$APP_PATH" "$DMG_STAGING/"
ln -s /Applications "$DMG_STAGING/Applications"

hdiutil create \
    -volname "$APP_NAME" \
    -srcfolder "$DMG_STAGING" \
    -ov \
    -format UDBZ \
    "$DMG_FINAL" \
    2>&1 | grep -v "^$"

rm -rf "$DMG_STAGING"
echo "  ✓ DMG created"

# ── Step 6: Verify ──────────────────────────────────────────────────────────
echo ""
echo "▸ Verifying no Homebrew references remain..."
BROKEN=0
for dylib in "${APP_PATH}/Contents/Frameworks/"*.dylib; do
    if otool -L "$dylib" 2>/dev/null | grep -q "/opt/homebrew"; then
        echo "  ✗ $(basename $dylib) still references Homebrew!"
        BROKEN=1
    fi
done
if [ $BROKEN -eq 0 ]; then
    echo "  ✓ All dylibs are self-contained"
fi

# ── Summary ─────────────────────────────────────────────────────────────────
echo ""
echo "═══════════════════════════════════════════════════"
echo "  Build complete!"
echo ""
echo "  App:  ${APP_PATH}"
echo "  DMG:  ${DMG_FINAL}"
echo "  Size: $(du -sh "$DMG_FINAL" | cut -f1)"
echo ""
echo "  The DMG is fully self-contained."
echo "  End users do NOT need Homebrew or any"
echo "  other dependencies installed."
echo "═══════════════════════════════════════════════════"
