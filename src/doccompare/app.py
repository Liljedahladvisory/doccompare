"""macOS .app entry point for DocCompare."""
import sys
import os


def _set_macos_app_name():
    """Override the macOS menu bar to show 'DocCompare' regardless of bundle name."""
    try:
        from Foundation import NSProcessInfo, NSBundle
        NSProcessInfo.processInfo().setProcessName_("DocCompare")
        bundle = NSBundle.mainBundle()
        for d in filter(None, [bundle.localizedInfoDictionary(),
                                bundle.infoDictionary()]):
            d["CFBundleName"]        = "DocCompare"
            d["CFBundleDisplayName"] = "DocCompare"
    except Exception:
        pass


if __name__ == "__main__":
    _set_macos_app_name()

    # Ensure the package is importable when running from .app bundle
    app_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.dirname(app_dir)
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)

    from doccompare.gui import main
    main()
