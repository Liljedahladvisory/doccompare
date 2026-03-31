#!/usr/bin/env python3
"""
Recursively find and bundle ALL Homebrew dylibs needed by WeasyPrint
into the .app bundle, then fix all load paths so the app is self-contained.
"""
import subprocess
import os
import sys
import shutil

BREW = "/opt/homebrew"
APP_PATH = sys.argv[1] if len(sys.argv) > 1 else "dist/DocCompare.app"
FRAMEWORKS = os.path.join(APP_PATH, "Contents", "Frameworks")

# ── Step 1: Recursively find all required dylibs ────────────────────────────

SEEDS = [
    "libpango-1.0.dylib", "libpangocairo-1.0.dylib", "libpangoft2-1.0.dylib",
    "libcairo.2.dylib", "libcairo-gobject.2.dylib", "libgdk_pixbuf-2.0.dylib",
    "libgobject-2.0.dylib", "libglib-2.0.dylib", "libgio-2.0.dylib",
    "libfontconfig.1.dylib", "libharfbuzz.0.dylib", "libfreetype.6.dylib",
    "libpixman-1.0.dylib", "libpng16.16.dylib",
]

seen = {}  # basename -> realpath


def scan_lib(path):
    rp = os.path.realpath(path)
    if not os.path.isfile(rp):
        return
    bn = os.path.basename(rp)
    if bn in seen:
        return
    seen[bn] = rp
    try:
        out = subprocess.check_output(["otool", "-L", rp], text=True)
    except Exception:
        return
    for line in out.strip().split("\n")[1:]:
        dep = line.strip().split(" ")[0]
        if dep.startswith(BREW):
            dep_real = os.path.realpath(dep)
            if os.path.isfile(dep_real):
                scan_lib(dep_real)


print(f"Scanning dylib dependencies from {len(SEEDS)} seed libraries...")
for s in SEEDS:
    scan_lib(os.path.join(BREW, "lib", s))

print(f"Found {len(seen)} dylibs to bundle.\n")

# ── Step 2: Copy all dylibs to Frameworks/ ──────────────────────────────────

os.makedirs(FRAMEWORKS, exist_ok=True)

for bn, src in sorted(seen.items()):
    dst = os.path.join(FRAMEWORKS, bn)
    shutil.copy2(src, dst)
    os.chmod(dst, 0o755)
    print(f"  Copied: {bn}")

# ── Step 3: Build a mapping from old install names to new ones ──────────────
# We need to map both the versioned names AND the /opt/homebrew/opt/ paths

remap = {}  # old_path -> @executable_path/../Frameworks/basename
for bn, src in seen.items():
    new = f"@executable_path/../Frameworks/{bn}"
    # Map the realpath
    remap[src] = new
    # Map the /opt/homebrew/lib/ symlink path
    remap[os.path.join(BREW, "lib", bn)] = new
    # Also map /opt/homebrew/opt/*/lib/* paths
    try:
        out = subprocess.check_output(["otool", "-D", src], text=True)
        install_name = out.strip().split("\n")[-1].strip()
        if install_name and install_name != src:
            remap[install_name] = new
    except Exception:
        pass

# Also add common symlink patterns
for bn, src in seen.items():
    # e.g., libcairo.2.dylib might be referenced as libcairo.dylib
    parts = bn.split(".")
    # Add all /opt/homebrew/opt/ patterns we can find
    try:
        out = subprocess.check_output(["otool", "-L", src], text=True)
        for line in out.strip().split("\n")[1:]:
            dep = line.strip().split(" ")[0]
            if dep.startswith(BREW):
                dep_real = os.path.realpath(dep)
                dep_bn = os.path.basename(dep_real)
                if dep_bn in seen:
                    remap[dep] = f"@executable_path/../Frameworks/{dep_bn}"
    except Exception:
        pass

print(f"\nFixing load paths ({len(remap)} remappings)...")

# ── Step 4: Fix all dylibs in Frameworks/ ───────────────────────────────────

for bn in sorted(seen):
    dylib = os.path.join(FRAMEWORKS, bn)
    new_id = f"@executable_path/../Frameworks/{bn}"

    # Fix install name
    subprocess.run(["install_name_tool", "-id", new_id, dylib],
                   capture_output=True)

    # Fix all references
    try:
        out = subprocess.check_output(["otool", "-L", dylib], text=True)
    except Exception:
        continue

    for line in out.strip().split("\n")[1:]:
        old_ref = line.strip().split(" ")[0]
        if old_ref in remap:
            subprocess.run(
                ["install_name_tool", "-change", old_ref, remap[old_ref], dylib],
                capture_output=True)

print("  Frameworks/ dylibs fixed.")

# ── Step 5: Fix any .so files in the bundle that reference Homebrew ─────────

print("Scanning .so files in bundle...")
fixed_so = 0
for root, dirs, files in os.walk(APP_PATH):
    for f in files:
        if f.endswith(".so") or f.endswith(".dylib"):
            fpath = os.path.join(root, f)
            if fpath.startswith(FRAMEWORKS):
                continue  # already fixed
            try:
                out = subprocess.check_output(["otool", "-L", fpath], text=True)
            except Exception:
                continue
            for line in out.strip().split("\n")[1:]:
                old_ref = line.strip().split(" ")[0]
                if old_ref in remap:
                    subprocess.run(
                        ["install_name_tool", "-change", old_ref, remap[old_ref], fpath],
                        capture_output=True)
                    fixed_so += 1

print(f"  Fixed {fixed_so} references in .so files.")

# ── Step 6: Patch WeasyPrint to look in Frameworks/ ─────────────────────────
# WeasyPrint uses ctypes.util.find_library() which won't find our bundled libs.
# We write a small wrapper that intercepts the library loading.

print("Patching WeasyPrint library loading...")

# Find weasyprint in the bundle
wp_dirs = []
for root, dirs, files in os.walk(APP_PATH):
    if "weasyprint" in dirs:
        wp_dirs.append(os.path.join(root, "weasyprint"))

for wp_dir in wp_dirs:
    # WeasyPrint loads libs in text.py or __init__.py via ffi
    # We need to create a _lib_paths.py that sets DYLD paths
    # Better approach: create a sitecustomize that patches ctypes.util.find_library
    pass

# Create a startup hook that sets DYLD_LIBRARY_PATH
boot_py = os.path.join(APP_PATH, "Contents", "Resources", "__boot__.py")
if os.path.exists(boot_py):
    with open(boot_py, "r") as f:
        content = f.read()

    if "DYLD_LIBRARY_PATH" not in content:
        # Prepend our Frameworks path to the environment
        patch = '''
import os as _os
_fw = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), "Frameworks")
_os.environ["DYLD_LIBRARY_PATH"] = _fw + ":" + _os.environ.get("DYLD_LIBRARY_PATH", "")
# Also patch ctypes.util.find_library to check Frameworks/ first
import ctypes.util as _cu
_orig_find = _cu.find_library
def _patched_find(name):
    import glob
    for pattern in [f"lib{name}*.dylib", f"{name}*.dylib"]:
        matches = glob.glob(_os.path.join(_fw, pattern))
        if matches:
            return matches[0]
    return _orig_find(name)
_cu.find_library = _patched_find

'''
        with open(boot_py, "w") as f:
            f.write(patch + content)
        print("  Patched __boot__.py with library path setup.")
    else:
        print("  __boot__.py already patched.")
else:
    print(f"  WARNING: {boot_py} not found!")

print(f"\nDone! {len(seen)} dylibs bundled into {FRAMEWORKS}")
