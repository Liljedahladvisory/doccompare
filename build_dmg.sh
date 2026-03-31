#!/usr/bin/env bash
#
# build_dmg.sh — Build DocCompare.app and package it into a distributable .dmg
#
# Prerequisites:
#   brew install pango cairo gdk-pixbuf libffi
#   pip install py2app
#   pip install -e .   (install doccompare in dev mode first)
#
# Usage:
#   chmod +x build_dmg.sh
#   ./build_dmg.sh
#
set -euo pipefail

APP_NAME="DocCompare"
VERSION="0.1.0"
DMG_NAME="${APP_NAME}-${VERSION}"
BUILD_DIR="$(pwd)/build"
DIST_DIR="$(pwd)/dist"

echo "═══════════════════════════════════════════════════"
echo "  Building ${APP_NAME} v${VERSION}"
echo "═══════════════════════════════════════════════════"

# ── Step 0: Check prerequisites ──────────────────────────────────────────────
echo ""
echo "▸ Checking prerequisites..."

# Check Homebrew dependencies for WeasyPrint
for dep in pango cairo gdk-pixbuf; do
    if ! brew list "$dep" &>/dev/null; then
        echo "  ✗ Missing Homebrew dependency: $dep"
        echo "  Run: brew install pango cairo gdk-pixbuf libffi"
        exit 1
    fi
done
echo "  ✓ Homebrew dependencies OK"

# Check py2app
if ! python3 -c "import py2app" 2>/dev/null; then
    echo "  ✗ py2app not installed. Run: pip install py2app"
    exit 1
fi
echo "  ✓ py2app OK"

# Check doccompare is importable
if ! python3 -c "import doccompare" 2>/dev/null; then
    echo "  ✗ doccompare not installed. Run: pip install -e ."
    exit 1
fi
echo "  ✓ doccompare OK"

# ── Step 1: Clean previous builds ───────────────────────────────────────────
echo ""
echo "▸ Cleaning previous builds..."
rm -rf "$BUILD_DIR" "$DIST_DIR"
echo "  ✓ Clean"

# ── Step 2: Build .app bundle with py2app ────────────────────────────────────
echo ""
echo "▸ Building .app bundle with py2app..."
python3 setup.py py2app 2>&1 | tail -5
echo "  ✓ .app bundle built"

APP_PATH="${DIST_DIR}/${APP_NAME}.app"
if [ ! -d "$APP_PATH" ]; then
    echo "  ✗ Expected ${APP_PATH} not found!"
    echo "  Checking dist/ contents:"
    ls -la "$DIST_DIR"/
    # Try to find the actual .app
    FOUND_APP=$(find "$DIST_DIR" -name "*.app" -maxdepth 1 | head -1)
    if [ -n "$FOUND_APP" ]; then
        echo "  Found: $FOUND_APP — renaming to ${APP_PATH}"
        mv "$FOUND_APP" "$APP_PATH"
    else
        echo "  ✗ No .app found in dist/. Build failed."
        exit 1
    fi
fi

# ── Step 3: Copy Homebrew dylibs for WeasyPrint (pango, cairo, etc.) ─────────
echo ""
echo "▸ Bundling native libraries..."

FRAMEWORKS_DIR="${APP_PATH}/Contents/Frameworks"
mkdir -p "$FRAMEWORKS_DIR"

# Find Homebrew lib path
BREW_PREFIX="$(brew --prefix)"
BREW_LIB="${BREW_PREFIX}/lib"

# Copy key dylibs that WeasyPrint needs
DYLIBS=(
    libpango-1.0.0.dylib
    libpangocairo-1.0.0.dylib
    libpangoft2-1.0.0.dylib
    libcairo.2.dylib
    libcairo-gobject.2.dylib
    libgdk_pixbuf-2.0.0.dylib
    libgobject-2.0.0.dylib
    libglib-2.0.0.dylib
    libgio-2.0.0.dylib
    libffi.8.dylib
    libharfbuzz.0.dylib
    libfontconfig.1.dylib
    libfreetype.6.dylib
    libpixman-1.0.dylib
    libpng16.16.dylib
    libintl.8.dylib
)

for lib in "${DYLIBS[@]}"; do
    src="${BREW_LIB}/${lib}"
    if [ -f "$src" ]; then
        cp "$src" "$FRAMEWORKS_DIR/"
        echo "  ✓ ${lib}"
    else
        # Try to find it in subdirectories
        found=$(find "$BREW_PREFIX" -name "$lib" -type f 2>/dev/null | head -1)
        if [ -n "$found" ]; then
            cp "$found" "$FRAMEWORKS_DIR/"
            echo "  ✓ ${lib} (from ${found})"
        else
            echo "  ⚠ ${lib} not found — may not be needed"
        fi
    fi
done

# Fix dylib rpaths to look inside the bundle
echo ""
echo "▸ Fixing dylib load paths..."
for dylib in "$FRAMEWORKS_DIR"/*.dylib; do
    [ -f "$dylib" ] || continue
    # Change the install name to @executable_path/../Frameworks/
    basename_lib=$(basename "$dylib")
    install_name_tool -id "@executable_path/../Frameworks/${basename_lib}" "$dylib" 2>/dev/null || true

    # Fix references to other Homebrew libs
    otool -L "$dylib" 2>/dev/null | grep "${BREW_PREFIX}" | awk '{print $1}' | while read -r ref; do
        ref_base=$(basename "$ref")
        if [ -f "${FRAMEWORKS_DIR}/${ref_base}" ]; then
            install_name_tool -change "$ref" "@executable_path/../Frameworks/${ref_base}" "$dylib" 2>/dev/null || true
        fi
    done
done
echo "  ✓ Load paths fixed"

# Also fix the main executable
MAIN_EXEC="${APP_PATH}/Contents/MacOS/${APP_NAME}"
if [ ! -f "$MAIN_EXEC" ]; then
    # py2app might name it differently
    MAIN_EXEC=$(find "${APP_PATH}/Contents/MacOS/" -type f | head -1)
fi
if [ -f "$MAIN_EXEC" ]; then
    otool -L "$MAIN_EXEC" 2>/dev/null | grep "${BREW_PREFIX}" | awk '{print $1}' | while read -r ref; do
        ref_base=$(basename "$ref")
        if [ -f "${FRAMEWORKS_DIR}/${ref_base}" ]; then
            install_name_tool -change "$ref" "@executable_path/../Frameworks/${ref_base}" "$MAIN_EXEC" 2>/dev/null || true
        fi
    done
fi

# ── Step 4: Ad-hoc code sign ────────────────────────────────────────────────
echo ""
echo "▸ Code signing (ad-hoc)..."
codesign --force --deep --sign - "$APP_PATH" 2>/dev/null || true
echo "  ✓ Signed"

# ── Step 5: Create DMG ──────────────────────────────────────────────────────
echo ""
echo "▸ Creating DMG..."

DMG_TEMP="${DIST_DIR}/${DMG_NAME}-temp.dmg"
DMG_FINAL="${DIST_DIR}/${DMG_NAME}.dmg"

# Create a temporary directory for DMG contents
DMG_STAGING="${DIST_DIR}/dmg_staging"
rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"

# Copy the app
cp -R "$APP_PATH" "$DMG_STAGING/"

# Create a symlink to /Applications for drag-and-drop install
ln -s /Applications "$DMG_STAGING/Applications"

# Create the DMG
hdiutil create \
    -volname "$APP_NAME" \
    -srcfolder "$DMG_STAGING" \
    -ov \
    -format UDBZ \
    "$DMG_FINAL" \
    2>&1 | grep -v "^$"

# Clean up staging
rm -rf "$DMG_STAGING" "$DMG_TEMP"

echo "  ✓ DMG created"

# ── Step 6: Summary ─────────────────────────────────────────────────────────
echo ""
echo "═══════════════════════════════════════════════════"
echo "  Build complete!"
echo ""
echo "  App:  ${APP_PATH}"
echo "  DMG:  ${DMG_FINAL}"
echo "  Size: $(du -sh "$DMG_FINAL" | cut -f1)"
echo ""
echo "  To install: Open the DMG and drag DocCompare"
echo "  to your Applications folder."
echo "═══════════════════════════════════════════════════"
