"""DocCompare GUI — modern macOS desktop app."""
import tkinter as tk
from tkinter import ttk, filedialog
import threading
from pathlib import Path
from datetime import datetime, date, timedelta
import sys
import os
import json
import random
import hashlib
import hmac as _hmac
import base64
import uuid
import urllib.request
import urllib.error
import ssl

# ── Config ──────────────────────────────────────────────────────────────────
_APP_DATA = os.path.expanduser("~/.doccompare_llt")
CONFIG_PATH = os.path.expanduser("~/.doccompare_llt.json")
LICENSE_PATH = os.path.join(_APP_DATA, "license.json")
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzevrhw7cJdjHZLn_OCfSR7RVSaU2hgV8RKU3hzeSkGouKYa-0ioo85eYoWkiAkHVpB/exec"
VERIFICATION_TIMEOUT = 600  # 10 min

# HMAC secret for license signing (same pattern as Meeting Recorder LLT)
_LICENSE_HMAC_SECRET = bytes.fromhex(
    "d4e7a1c9f03b48d6b5fc81927364e0a2"
    "29c85d3f6a7e41089bdf02c5174a63e8"
)


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ── License system ─────────────────────────────────────────────────────────

def _get_machine_id() -> str:
    """Get a stable machine identifier."""
    try:
        return str(uuid.getnode())
    except Exception:
        return "unknown"


def _generate_license_key(name: str, email: str, company: str = "",
                          days: int = 365) -> str:
    """Generate an HMAC-signed license key valid for `days` days."""
    created = date.today().isoformat()
    expires = (date.today() + timedelta(days=days)).isoformat()
    payload = {
        "company": company or name,
        "created": created,
        "email": email,
        "expires": expires,
        "trial": True,
    }
    payload_json = json.dumps(payload, separators=(",", ":"), sort_keys=True)
    payload_bytes = payload_json.encode("utf-8")
    sig = _hmac.new(_LICENSE_HMAC_SECRET, payload_bytes, hashlib.sha256).digest()
    combined = payload_bytes + b"|" + sig
    key_b64 = base64.urlsafe_b64encode(combined).decode("ascii")
    chunks = [key_b64[i:i+4] for i in range(0, len(key_b64), 4)]
    return "LLT." + ".".join(chunks)


def _verify_license(key_str: str) -> dict | None:
    """Verify a license key (HMAC-SHA256). Returns payload dict or None."""
    try:
        raw = key_str.strip().replace("\n", "").replace("\r", "").replace(" ", "")
        if raw.upper().startswith("LLT."):
            raw = raw[4:]
        raw = raw.replace(".", "")
        pad = 4 - len(raw) % 4
        if pad != 4:
            raw += "=" * pad
        combined = base64.urlsafe_b64decode(raw)
        if len(combined) < 34:
            return None
        sig = combined[-32:]
        if combined[-33:-32] != b"|":
            return None
        payload_bytes = combined[:-33]
        expected = _hmac.new(_LICENSE_HMAC_SECRET, payload_bytes,
                             hashlib.sha256).digest()
        if not _hmac.compare_digest(sig, expected):
            return None
        return json.loads(payload_bytes.decode("utf-8"))
    except Exception:
        return None


def _save_license(key_str: str, payload: dict):
    """Save license data to disk."""
    os.makedirs(_APP_DATA, exist_ok=True)
    license_data = {
        "key": key_str,
        "machine_id": _get_machine_id(),
        "activated": date.today().isoformat(),
        "company": payload.get("company", ""),
        "email": payload.get("email", ""),
        "expires": payload.get("expires", ""),
    }
    with open(LICENSE_PATH, "w", encoding="utf-8") as f:
        json.dump(license_data, f, indent=2, ensure_ascii=False)


def _check_license_file() -> tuple[bool, str, dict | None]:
    """Check stored license. Returns (valid, message, payload)."""
    if not os.path.isfile(LICENSE_PATH):
        return False, "No license found.", None
    try:
        data = json.loads(open(LICENSE_PATH, encoding="utf-8").read())
        key_str = data.get("key", "")
        stored_machine = data.get("machine_id", "")
    except Exception:
        return False, "Could not read license file.", None

    payload = _verify_license(key_str)
    if payload is None:
        return False, "Invalid license key.", None

    current_machine = _get_machine_id()
    if stored_machine and stored_machine != current_machine:
        return False, "License is registered on a different computer.", None

    expires = payload.get("expires", "2000-01-01")
    if date.fromisoformat(expires) < date.today():
        return False, f"Your license expired on {expires}. Contact Liljedahl Legal Tech for renewal.", None

    return True, "ok", payload


# ── Translations ────────────────────────────────────────────────────────────

STRINGS = {
    "sv": {
        "subtitle":           "Jämför .docx-dokument och generera en diff-rapport",
        "original_doc":       "Originaldokument",
        "modified_doc":        "Ändrat dokument",
        "save_report":         "Spara rapport som",
        "select_file":         "Välj fil",
        "choose_location":     "Välj plats",
        "no_file":             "Ingen fil vald",
        "desktop_auto":        "Skrivbordet — automatiskt filnamn",
        "optional":            "valfritt",
        "compare_btn":         "Jämför dokument",
        "comparing":           "Jämför dokument\u2026",
        "rendering":           "Renderar PDF\u2026",
        "done":                "Klart! Rapport sparad: {filename}",
        "added":               "tillagda",
        "deleted":             "borttagna",
        "unchanged":           "oförändrade",
        "error_prefix":        "Fel",
        "format_mismatch":     "Båda filerna måste vara samma format ({ext1} \u2260 {ext2})",
        "format_unsupported":  "Format stöds ej: {ext}",
        "not_verified":        "Programmet är inte aktiverat. Starta om och verifiera din e-post.",
        "verify_first":        "Verifiera din e-post först.",
        "settings_title":      "Inställningar",
        "settings_msg":        "Ändra namn eller företagsnamn:",
        "save":                "Spara",
        "cancel":              "Avbryt",
        "pick_original_title": "Välj originaldokument",
        "pick_modified_title": "Välj ändrat dokument",
        "pick_output_title":   "Spara rapport",
        "word_docs":           "Word-dokument",
        "language":            "Språk",
    },
    "en": {
        "subtitle":           "Compare .docx documents and generate a diff report",
        "original_doc":       "Original document",
        "modified_doc":        "Modified document",
        "save_report":         "Save report as",
        "select_file":         "Select file",
        "choose_location":     "Choose location",
        "no_file":             "No file selected",
        "desktop_auto":        "Desktop — automatic filename",
        "optional":            "optional",
        "compare_btn":         "Compare Documents",
        "comparing":           "Comparing documents\u2026",
        "rendering":           "Rendering PDF\u2026",
        "done":                "Done! Report saved: {filename}",
        "added":               "added",
        "deleted":             "deleted",
        "unchanged":           "unchanged",
        "error_prefix":        "Error",
        "format_mismatch":     "Both files must be the same format ({ext1} \u2260 {ext2})",
        "format_unsupported":  "Unsupported format: {ext}",
        "not_verified":        "App not activated. Restart and verify your email.",
        "verify_first":        "Verify your email first.",
        "settings_title":      "Settings",
        "settings_msg":        "Change name or company name:",
        "save":                "Save",
        "cancel":              "Cancel",
        "pick_original_title": "Select original document",
        "pick_modified_title": "Select modified document",
        "pick_output_title":   "Save report",
        "word_docs":           "Word Documents",
        "language":            "Language",
    },
}


def _t(key: str, lang: str = "sv", **kw) -> str:
    """Look up a translated string."""
    s = STRINGS.get(lang, STRINGS["sv"]).get(key, key)
    if kw:
        s = s.format(**kw)
    return s


# ── Email verification ──────────────────────────────────────────────────────

def _generate_code() -> str:
    return str(random.randint(1000, 9999))


def _send_verification_code(email: str, name: str, code: str) -> bool:
    payload = {
        "action": "send_verification_email",
        "email": email,
        "name": name,
        "code": code,
    }
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}

    ctx = ssl.create_default_context()
    try:
        req = urllib.request.Request(WEBHOOK_URL, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result.get("status") == "ok"
    except Exception:
        pass

    try:
        ctx_noverify = ssl.create_default_context()
        ctx_noverify.check_hostname = False
        ctx_noverify.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(WEBHOOK_URL, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=15, context=ctx_noverify) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result.get("status") == "ok"
    except Exception:
        return False


def _send_registration(data: dict) -> bool:
    """Send full registration data to GAS webhook (stored in spreadsheet)."""
    payload = json.dumps(data).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    for verify in (True, False):
        try:
            ctx = ssl.create_default_context()
            if not verify:
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
            req = urllib.request.Request(WEBHOOK_URL, data=payload, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
                result = json.loads(resp.read().decode("utf-8"))
                if result.get("status") == "ok":
                    return True
        except Exception:
            pass
    return False


# ── Colour palette ──────────────────────────────────────────────────────────
BG      = "#0E0D0C"
BG2     = "#161412"
BG3     = "#1E1B18"
BG4     = "#272320"
BORDER  = "#332E28"
BORDER2 = "#4A4238"
FG      = "#F2EEE8"
FG2     = "#C8B89A"
FG3     = "#E0D4C0"
FG_DIM  = "#9A8A72"
ACCENT  = "#E07820"
ACCENT2 = "#C05E0A"
RED     = "#D95050"
GREEN   = "#4AB870"

# ── Fonts ───────────────────────────────────────────────────────────────────
FONT_LOGO1   = ("Helvetica Neue", 14, "bold")
FONT_LOGO2   = ("Helvetica Neue", 14)
FONT_POWERED = ("Helvetica Neue", 9, "italic")
FONT_SECTION = ("Helvetica Neue", 8, "bold")
FONT_H       = ("Helvetica Neue", 11, "bold")
FONT_B       = ("Helvetica Neue", 11)
FONT_S       = ("Helvetica Neue", 10)
FONT_XS      = ("Helvetica Neue", 9)
FONT_M       = ("Menlo", 10)
FONT_MS      = ("Menlo", 9)
FONT_TITLE   = ("Helvetica Neue", 26, "bold")
FONT_SUB     = ("Helvetica Neue", 12)
FONT_BTN     = ("Helvetica Neue", 14, "bold")
FONT_GEAR    = ("Helvetica Neue", 18)

ICON_CHECK = "\u2713"


def _style_widgets():
    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "Orange.Horizontal.TProgressbar",
        troughcolor=BG3, background=ACCENT, bordercolor=BG,
        lightcolor=ACCENT, darkcolor=ACCENT,
    )


class RoundedButton(tk.Canvas):
    """Canvas-based button with smooth rounded corners."""

    def __init__(self, parent, text, command=None, style="solid",
                 bg=None, fg="#FFFFFF", radius=12,
                 font_spec=None, padx=28, pady=12,
                 state="normal", fixed_width=None):
        import tkinter.font as tkfont
        self._style    = style
        self._bg       = bg or ACCENT
        self._fg       = fg
        self._radius   = radius
        self._padx     = padx
        self._pady     = pady
        self._command  = command
        self._enabled  = (state == "normal")
        self._text     = text
        self._hovering = False
        self._fspec    = font_spec or ("Helvetica Neue", 13)
        weight = "bold" if len(self._fspec) > 2 and "bold" in self._fspec[2] else "normal"
        mf = tkfont.Font(family=self._fspec[0], size=self._fspec[1], weight=weight)
        th = mf.metrics("linespace")
        tw = mf.measure(text)
        self._btn_w = fixed_width or (tw + 2 * padx)
        self._btn_h = th + 2 * pady
        super().__init__(parent, width=self._btn_w, height=self._btn_h,
                         bg=parent.cget("bg"), highlightthickness=0, bd=0)
        self._draw()
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>",    self._on_enter)
        self.bind("<Leave>",    self._on_leave)

    def _resolve_fill(self):
        if not self._enabled:
            return (BG3, FG_DIM) if self._style == "solid" else (BG, BORDER)
        if self._hovering:
            if self._style == "solid":
                c = self._bg.lstrip("#")
                r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
                return f"#{int(r*.82):02x}{int(g*.82):02x}{int(b*.82):02x}", self._fg
            return BG2, self._fg
        return (self._bg, self._fg) if self._style == "solid" else (BG, self._fg)

    def _draw(self):
        self.delete("all")
        w, h, r = self._btn_w, self._btn_h, self._radius
        fill, fg = self._resolve_fill()
        outline = fg if self._style == "ghost" else fill
        m = 1
        x0, y0, x1, y1 = m, m, w - m, h - m
        pts = [x0+r, y0, x1-r, y0, x1, y0, x1, y0+r,
               x1, y1-r, x1, y1, x1-r, y1, x0+r, y1,
               x0, y1, x0, y1-r, x0, y0+r, x0, y0]
        self.create_polygon(pts, smooth=True, fill=fill, outline=outline, width=1)
        self.create_text(w // 2, h // 2, text=self._text,
                         font=self._fspec, fill=fg, anchor="center")

    def config(self, **kw):
        changed = False
        for k in ("text", "state", "bg", "fg"):
            if k in kw:
                v = kw.pop(k)
                if k == "text":    self._text = v
                elif k == "state": self._enabled = (v == "normal")
                elif k == "bg":    self._bg = v
                elif k == "fg":    self._fg = v
                changed = True
        if "cursor" in kw:
            super().config(cursor=kw.pop("cursor"))
        if kw:
            super().config(**kw)
        if changed:
            self._draw()

    def cget(self, key):
        if key == "state":  return "normal" if self._enabled else "disabled"
        if key == "text":   return self._text
        if key == "bg":     return self._bg
        return super().cget(key)

    def _on_click(self, _=None):
        if self._enabled and self._command: self._command()
    def _on_enter(self, _=None):
        self._hovering = True; self._draw()
    def _on_leave(self, _=None):
        self._hovering = False; self._draw()


# ─────────────────────────────────────────────────────────────────────────────

class DocCompareApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DocCompare")
        self.root.resizable(True, True)
        self.root.configure(bg=BG)
        self.root.geometry("600x700")
        self.root.minsize(540, 620)

        self._config = load_config()
        self.user_name = self._config.get("user_name", "")
        self._verified = bool(self._config.get("verified"))
        self.lang = self._config.get("language", "sv")

        self.original_path: Path | None = None
        self.modified_path: Path | None = None
        self.output_path: Path | None = None

        _style_widgets()
        self._build_ui()
        self._center_window()

        if not self._verified:
            self.root.after(300, self._show_registration_dialog)
        elif self._verified:
            # Check license validity
            valid, msg, payload = _check_license_file()
            if not valid and os.path.isfile(LICENSE_PATH):
                # License exists but expired/invalid
                self.root.after(300, lambda: self._show_expired_window(msg))
            elif not self._config.get("language"):
                # Verified but no language chosen yet — show language picker
                self.root.after(300, self._show_language_dialog)

    def _s(self, key, **kw):
        return _t(key, self.lang, **kw)

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        self._outer = tk.Frame(root, bg=BG)
        self._outer.pack(padx=32, pady=(24, 24), fill="both", expand=True)
        self._populate_main_ui()

    def _populate_main_ui(self):
        outer = self._outer
        # Clear existing children
        for w in outer.winfo_children():
            w.destroy()

        # ── Header ───────────────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x", pady=(0, 16))

        left = tk.Frame(header, bg=BG)
        left.pack(side="left", fill="y")

        name_row = tk.Frame(left, bg=BG)
        name_row.pack(anchor="w")
        self._logo_name_lbl = tk.Label(
            name_row, text=self._display_name(),
            font=FONT_LOGO1, bg=BG, fg=FG)
        self._logo_name_lbl.pack(side="left")
        tk.Label(name_row, text="  DocCompare",
                 font=FONT_LOGO2, bg=BG, fg=FG2).pack(side="left")

        tk.Label(left, text="Powered by Liljedahl Legal Tech",
                 font=FONT_POWERED, bg=BG, fg=FG_DIM).pack(anchor="w", pady=(3, 0))

        right = tk.Frame(header, bg=BG)
        right.pack(side="right", fill="y")
        gear = tk.Label(right, text="\u2699", font=FONT_GEAR,
                        bg=BG, fg=FG_DIM, cursor="hand2")
        gear.pack(side="right")
        gear.bind("<Enter>",    lambda _: gear.config(fg=FG))
        gear.bind("<Leave>",    lambda _: gear.config(fg=FG_DIM))
        gear.bind("<Button-1>", lambda _: self._show_settings_dialog())

        # ── Title ────────────────────────────────────────────────────────
        title_frame = tk.Frame(outer, bg=BG)
        title_frame.pack(fill="x", pady=(0, 6))
        tk.Label(title_frame, text="DocCompare",
                 font=FONT_TITLE, bg=BG, fg=FG).pack(side="left")
        badge_frame = tk.Frame(title_frame, bg=BG3, padx=8, pady=2,
                               highlightbackground=BORDER, highlightthickness=1)
        badge_frame.pack(side="left", padx=(12, 0), pady=(8, 0))
        tk.Label(badge_frame, text="v0.1", font=FONT_MS, bg=BG3,
                 fg=ACCENT).pack()

        tk.Label(outer, text=self._s("subtitle"),
                 font=FONT_SUB, bg=BG, fg=FG_DIM).pack(anchor="w", pady=(0, 16))

        tk.Frame(outer, bg=ACCENT, height=2).pack(fill="x", pady=(0, 16))

        # ── File cards ───────────────────────────────────────────────────
        self.orig_card, self.orig_label, self.orig_icon = self._file_card(
            outer, self._s("original_doc"), self._s("select_file"),
            self._pick_original,
        )
        self.mod_card, self.mod_label, self.mod_icon = self._file_card(
            outer, self._s("modified_doc"), self._s("select_file"),
            self._pick_modified,
        )
        self.out_card, self.out_label, self.out_icon = self._file_card(
            outer, self._s("save_report"), self._s("choose_location"),
            self._pick_output,
            default_text=self._s("desktop_auto"), optional=True,
        )

        # ── Compare button ───────────────────────────────────────────────
        btn_frame = tk.Frame(outer, bg=BG)
        btn_frame.pack(fill="x", pady=(8, 16))
        self.compare_btn = RoundedButton(
            btn_frame, text=self._s("compare_btn"),
            command=self._run_comparison,
            bg=ACCENT, fg="#FFFFFF", font_spec=FONT_BTN,
            padx=40, pady=14, radius=12,
            state="disabled", fixed_width=536,
        )
        self.compare_btn.pack(fill="x")

        # ── Progress + status ────────────────────────────────────────────
        self.progress = ttk.Progressbar(
            outer, mode="indeterminate", length=400,
            style="Orange.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 4))

        self.status_label = tk.Label(
            outer, text="", font=FONT_B, bg=BG, fg=FG_DIM,
            wraplength=500, justify="left", anchor="w")
        self.status_label.pack(anchor="w", fill="x", pady=(6, 0))

        # ── Footer ───────────────────────────────────────────────────────
        footer = tk.Frame(outer, bg=BG)
        footer.pack(side="bottom", fill="x", pady=(16, 0))
        tk.Label(footer,
                 text="Liljedahl Legal Tech  \u2022  Liljedahl Advisory AB",
                 font=FONT_XS, bg=BG, fg=BORDER2).pack()

    def _file_card(self, parent, title, btn_text, command,
                   default_text=None, optional=False):
        if default_text is None:
            default_text = self._s("no_file")
        card = tk.Frame(parent, bg=BG2,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(card, bg=BG2)
        inner.pack(fill="x", padx=14, pady=12)
        left = tk.Frame(inner, bg=BG2)
        left.pack(side="left", fill="x", expand=True)
        top_row = tk.Frame(left, bg=BG2)
        top_row.pack(anchor="w")
        tk.Label(top_row, text=title, font=FONT_H, bg=BG2, fg=FG).pack(side="left")
        if optional:
            tk.Label(top_row, text=self._s("optional"),
                     font=FONT_XS, bg=BG2, fg=FG_DIM).pack(side="left", padx=(8, 0))
        file_lbl = tk.Label(left, text=default_text,
                            font=FONT_S, bg=BG2, fg=FG_DIM, anchor="w")
        file_lbl.pack(anchor="w", pady=(3, 0))
        btn = RoundedButton(inner, text=btn_text, command=command,
                            style="ghost", bg=BG2, fg=FG2,
                            font_spec=FONT_S, padx=16, pady=8, radius=8)
        btn.pack(side="right", padx=(12, 0))
        icon_lbl = tk.Label(inner, text="",
                            font=("Helvetica Neue", 13), bg=BG2, fg=GREEN)
        icon_lbl.pack(side="right", padx=(0, 4))
        return card, file_lbl, icon_lbl

    def _display_name(self) -> str:
        return self.user_name if self.user_name else "DocCompare"

    def _center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2 - 60}")

    # ── Registration dialog (English, with address/org/VAT) ──────────────

    def _show_registration_dialog(self):
        self._pending_code = None

        dlg = tk.Toplevel(self.root)
        dlg.title("Register DocCompare")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 500, 680
        x = self.root.winfo_x() + (self.root.winfo_width()  - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        # Scrollable content area
        pad = tk.Frame(dlg, bg=BG, padx=36, pady=24)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text="Welcome to DocCompare",
                 font=("Helvetica Neue", 15, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(pad,
                 text="Enter your details below. A verification code\n"
                      "will be sent to your email address.",
                 font=FONT_S, bg=BG, fg=FG2,
                 wraplength=420, justify="left").pack(anchor="w", pady=(6, 16))

        # ── Step 1: all fields ───────────────────────────────────────────
        self._reg_step1 = tk.Frame(pad, bg=BG)
        self._reg_step1.pack(fill="x")

        def _field(parent, label, hint=None):
            tk.Label(parent, text=label, font=FONT_SECTION,
                     bg=BG, fg=FG_DIM).pack(anchor="w")
            if hint:
                tk.Label(parent, text=hint, font=FONT_XS,
                         bg=BG, fg=BORDER2).pack(anchor="w")
            var = tk.StringVar()
            entry = tk.Entry(parent, textvariable=var, font=FONT_M,
                             bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                             highlightthickness=1, highlightbackground=BORDER2,
                             highlightcolor=ACCENT)
            entry.pack(fill="x", ipady=8, pady=(3, 10))
            return var, entry

        name_var, name_entry = _field(self._reg_step1, "NAME / COMPANY")
        name_entry.focus_set()
        email_var, email_entry = _field(self._reg_step1, "EMAIL ADDRESS")
        address_var, address_entry = _field(self._reg_step1, "ADDRESS",
                                            hint="Street, city, postal code")
        org_var, org_entry = _field(self._reg_step1, "ORGANISATION NUMBER",
                                    hint="If company (e.g. 556xxx-xxxx)")
        vat_var, vat_entry = _field(self._reg_step1, "VAT NUMBER",
                                    hint="If company outside Sweden (e.g. DE123456789)")

        # ── Step 2: verification code (hidden initially) ─────────────────
        self._reg_step2 = tk.Frame(pad, bg=BG)

        tk.Label(self._reg_step2, text="VERIFICATION CODE", font=FONT_SECTION,
                 bg=BG, fg=FG_DIM).pack(anchor="w")
        self._code_info_lbl = tk.Label(
            self._reg_step2,
            text="A 4-digit code has been sent to your email.",
            font=FONT_XS, bg=BG, fg=FG2)
        self._code_info_lbl.pack(anchor="w", pady=(2, 6))
        code_var = tk.StringVar()
        code_entry = tk.Entry(self._reg_step2, textvariable=code_var,
                              font=("Menlo", 18),
                              bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                              highlightthickness=1, highlightbackground=BORDER2,
                              highlightcolor=ACCENT, justify="center")
        code_entry.pack(fill="x", ipady=10, pady=(0, 6))

        # Error label
        error_lbl = tk.Label(pad, text="", font=FONT_XS, bg=BG, fg=RED, anchor="w")
        error_lbl.pack(anchor="w", fill="x", pady=(4, 0))

        # Button row
        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x", pady=(10, 0), side="bottom")

        # ── Step 1 action: send code ─────────────────────────────────────
        def send_code():
            name = name_var.get().strip()
            email = email_var.get().strip()

            if not name:
                name_entry.config(highlightbackground=RED)
                error_lbl.config(text="Please enter your name.", fg=RED)
                return
            name_entry.config(highlightbackground=BORDER2)

            if not email or "@" not in email:
                email_entry.config(highlightbackground=RED)
                error_lbl.config(text="Please enter a valid email address.", fg=RED)
                return
            email_entry.config(highlightbackground=BORDER2)

            error_lbl.config(text="Sending verification code\u2026", fg=FG_DIM)
            dlg.update()

            code = _generate_code()
            self._pending_code = code

            def _do_send():
                sent = _send_verification_code(email, name, code)
                dlg.after(0, lambda: _on_sent(sent))

            def _on_sent(sent):
                if sent:
                    self._reg_step1.pack_forget()
                    self._reg_step2.pack(fill="x")
                    self._code_info_lbl.config(
                        text=f"A 4-digit code has been sent to {email}.")
                    error_lbl.config(text="", fg=RED)
                    send_btn.pack_forget()
                    verify_btn.pack(side="right")
                    resend_btn.pack(side="right", padx=(0, 8))
                    code_entry.focus_set()
                else:
                    error_lbl.config(
                        text="Could not send email. Check the address and try again.",
                        fg=RED)

            threading.Thread(target=_do_send, daemon=True).start()

        send_btn = tk.Label(btn_row, text="Send Code", font=FONT_B,
                            bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                            cursor="hand2")
        send_btn.pack(side="right")
        send_btn.bind("<Button-1>", lambda _: send_code())
        send_btn.bind("<Enter>", lambda _: send_btn.config(bg=ACCENT2))
        send_btn.bind("<Leave>", lambda _: send_btn.config(bg=ACCENT))

        email_entry.bind("<Return>", lambda _: send_code())
        name_entry.bind("<Return>", lambda _: email_entry.focus_set())

        # ── Step 2 action: verify ────────────────────────────────────────
        def verify_code():
            entered = code_var.get().strip()
            if not entered:
                code_entry.config(highlightbackground=RED)
                error_lbl.config(text="Please enter the verification code.", fg=RED)
                return

            if entered == self._pending_code:
                name = name_var.get().strip()
                email = email_var.get().strip()
                company = name
                address = address_var.get().strip()
                org_nr = org_var.get().strip()
                vat_nr = vat_var.get().strip()

                # Generate 12-month HMAC-signed license
                license_key = _generate_license_key(name, email, company, days=365)
                license_payload = _verify_license(license_key)
                _save_license(license_key, license_payload)

                license_expires = (date.today() + timedelta(days=365)).isoformat()

                self.user_name = name
                self._verified = True
                self._config["user_name"] = name
                self._config["email"] = email
                self._config["address"] = address
                self._config["org_nr"] = org_nr
                self._config["vat_nr"] = vat_nr
                self._config["verified"] = True
                self._config["registered"] = datetime.now().isoformat()
                self._config["license_expires"] = license_expires
                save_config(self._config)

                # Send registration data to spreadsheet
                def _reg():
                    _send_registration({
                        "name": name,
                        "company": company,
                        "email": email,
                        "address": address,
                        "org_nr": org_nr,
                        "vat_nr": vat_nr,
                        "registered": self._config["registered"],
                        "license_expires": license_expires,
                        "machine_id": _get_machine_id(),
                    })
                threading.Thread(target=_reg, daemon=True).start()

                self._logo_name_lbl.config(text=self._display_name())
                dlg.destroy()

                # Show language picker after registration
                self.root.after(300, self._show_language_dialog)
            else:
                code_entry.config(highlightbackground=RED)
                error_lbl.config(
                    text="Wrong code. Check your email and try again.", fg=RED)

        verify_btn = tk.Label(btn_row, text="Verify", font=FONT_B,
                              bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                              cursor="hand2")
        verify_btn.bind("<Button-1>", lambda _: verify_code())
        verify_btn.bind("<Enter>", lambda _: verify_btn.config(bg=ACCENT2))
        verify_btn.bind("<Leave>", lambda _: verify_btn.config(bg=ACCENT))

        def resend():
            code = _generate_code()
            self._pending_code = code
            email = email_var.get().strip()
            name = name_var.get().strip()
            error_lbl.config(text="Sending new code\u2026", fg=FG_DIM)
            dlg.update()
            def _do():
                sent = _send_verification_code(email, name, code)
                dlg.after(0, lambda: error_lbl.config(
                    text="New code sent!" if sent else "Could not send. Try again.",
                    fg=GREEN if sent else RED))
            threading.Thread(target=_do, daemon=True).start()

        resend_btn = tk.Label(btn_row, text="Resend", font=FONT_B,
                              bg=BG3, fg=FG2, padx=24, pady=8, cursor="hand2")
        resend_btn.bind("<Button-1>", lambda _: resend())
        resend_btn.bind("<Enter>", lambda _: resend_btn.config(bg=BG4, fg=FG))
        resend_btn.bind("<Leave>", lambda _: resend_btn.config(bg=BG3, fg=FG2))

        code_entry.bind("<Return>", lambda _: verify_code())
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)
        dlg.wait_window()

    # ── Language selection dialog ─────────────────────────────────────────

    def _show_language_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Language / Språk")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 400, 220
        x = self.root.winfo_x() + (self.root.winfo_width()  - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        pad = tk.Frame(dlg, bg=BG, padx=36, pady=28)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text="Choose your language",
                 font=("Helvetica Neue", 15, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(pad, text="This can be changed later in Settings.",
                 font=FONT_S, bg=BG, fg=FG2).pack(anchor="w", pady=(6, 24))

        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x")

        def pick(lang):
            self.lang = lang
            self._config["language"] = lang
            save_config(self._config)
            dlg.destroy()
            # Rebuild the UI in the chosen language
            self._populate_main_ui()

        for lang_code, label in [("sv", "Svenska"), ("en", "English")]:
            btn = tk.Label(btn_row, text=label, font=("Helvetica Neue", 13, "bold"),
                           bg=ACCENT if lang_code == "sv" else BG3,
                           fg="#FFFFFF" if lang_code == "sv" else FG2,
                           padx=32, pady=12, cursor="hand2")
            btn.pack(side="left", padx=(0, 12))
            _code = lang_code
            btn.bind("<Button-1>", lambda _, c=_code: pick(c))
            if lang_code == "sv":
                btn.bind("<Enter>", lambda _, b=btn: b.config(bg=ACCENT2))
                btn.bind("<Leave>", lambda _, b=btn: b.config(bg=ACCENT))
            else:
                btn.bind("<Enter>", lambda _, b=btn: b.config(bg=BG4, fg=FG))
                btn.bind("<Leave>", lambda _, b=btn: b.config(bg=BG3, fg=FG2))

        dlg.protocol("WM_DELETE_WINDOW", lambda: pick("sv"))
        dlg.wait_window()

    # ── License expired dialog ────────────────────────────────────────────

    def _show_expired_window(self, msg: str):
        dlg = tk.Toplevel(self.root)
        dlg.title("License Expired")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 460, 260
        x = self.root.winfo_x() + (self.root.winfo_width() - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        pad = tk.Frame(dlg, bg=BG, padx=36, pady=28)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text="License Expired",
                 font=("Helvetica Neue", 15, "bold"),
                 bg=BG, fg=RED).pack(anchor="w")
        tk.Label(pad, text=msg,
                 font=FONT_S, bg=BG, fg=FG2,
                 wraplength=380, justify="left").pack(anchor="w", pady=(10, 6))
        tk.Label(pad,
                 text="Please contact svante@liljedahladvisory.se\nto renew your license.",
                 font=FONT_S, bg=BG, fg=FG_DIM,
                 wraplength=380, justify="left").pack(anchor="w", pady=(6, 20))

        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x")

        def quit_app():
            dlg.destroy()
            self.root.destroy()

        quit_btn = tk.Label(btn_row, text="Quit", font=FONT_B,
                            bg=RED, fg="#FFFFFF", padx=24, pady=8,
                            cursor="hand2")
        quit_btn.pack(side="right")
        quit_btn.bind("<Button-1>", lambda _: quit_app())

        dlg.protocol("WM_DELETE_WINDOW", quit_app)
        dlg.wait_window()

    # ── Settings dialog ──────────────────────────────────────────────────

    def _show_settings_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title(self._s("settings_title"))
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 440, 310
        x = self.root.winfo_x() + (self.root.winfo_width()  - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        pad = tk.Frame(dlg, bg=BG, padx=36, pady=32)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text=self._s("settings_title"),
                 font=("Helvetica Neue", 13, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(pad, text=self._s("settings_msg"),
                 font=FONT_S, bg=BG, fg=FG2,
                 wraplength=368, justify="left").pack(anchor="w", pady=(8, 14))

        entry_var = tk.StringVar(value=self.user_name)
        entry = tk.Entry(pad, textvariable=entry_var, font=FONT_M,
                         bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=BORDER2,
                         highlightcolor=ACCENT)
        entry.pack(fill="x", ipady=9)
        entry.focus_set()
        entry.select_range(0, "end")

        # Language selector
        lang_frame = tk.Frame(pad, bg=BG)
        lang_frame.pack(fill="x", pady=(14, 0))
        tk.Label(lang_frame, text=self._s("language"), font=FONT_S,
                 bg=BG, fg=FG2).pack(side="left")

        lang_var = tk.StringVar(value=self.lang)
        for val, lbl in [("sv", "Svenska"), ("en", "English")]:
            tk.Radiobutton(lang_frame, text=lbl, variable=lang_var, value=val,
                           font=FONT_S, bg=BG3, fg=FG, selectcolor=ACCENT,
                           activebackground=BG4, activeforeground=FG,
                           indicatoron=False, relief="solid", bd=1,
                           padx=10, pady=3, cursor="hand2"
                           ).pack(side="left", padx=(8, 0))

        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x", pady=(18, 0))

        def save():
            name = entry_var.get().strip()
            if not name:
                entry.config(highlightbackground=RED)
                return
            new_lang = lang_var.get()
            lang_changed = (new_lang != self.lang)
            self.user_name = name
            self.lang = new_lang
            self._config["user_name"] = name
            self._config["language"] = new_lang
            save_config(self._config)
            self._logo_name_lbl.config(text=self._display_name())
            dlg.destroy()
            if lang_changed:
                self._populate_main_ui()

        save_lbl = tk.Label(btn_row, text=self._s("save"), font=FONT_B,
                            bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                            cursor="hand2")
        save_lbl.pack(side="right")
        save_lbl.bind("<Button-1>", lambda _: save())
        save_lbl.bind("<Enter>",  lambda _: save_lbl.config(bg=ACCENT2))
        save_lbl.bind("<Leave>",  lambda _: save_lbl.config(bg=ACCENT))

        cancel_lbl = tk.Label(btn_row, text=self._s("cancel"), font=FONT_B,
                              bg=BG3, fg=FG2, padx=24, pady=8, cursor="hand2")
        cancel_lbl.pack(side="right", padx=(0, 8))
        cancel_lbl.bind("<Button-1>", lambda _: dlg.destroy())
        cancel_lbl.bind("<Enter>", lambda _: cancel_lbl.config(bg=BG4, fg=FG))
        cancel_lbl.bind("<Leave>", lambda _: cancel_lbl.config(bg=BG3, fg=FG2))

        entry.bind("<Return>", lambda _: save())
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.wait_window()

    # ── File pickers ─────────────────────────────────────────────────────

    def _pick_original(self):
        path = filedialog.askopenfilename(
            title=self._s("pick_original_title"),
            filetypes=[(self._s("word_docs"), "*.docx")],
        )
        if path:
            self.original_path = Path(path)
            self.orig_label.config(text=self.original_path.name, fg=FG3)
            self.orig_icon.config(text=ICON_CHECK)
            self.orig_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_modified(self):
        path = filedialog.askopenfilename(
            title=self._s("pick_modified_title"),
            filetypes=[(self._s("word_docs"), "*.docx")],
        )
        if path:
            self.modified_path = Path(path)
            self.mod_label.config(text=self.modified_path.name, fg=FG3)
            self.mod_icon.config(text=ICON_CHECK)
            self.mod_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title=self._s("pick_output_title"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=Path.home() / "Desktop",
        )
        if path:
            self.output_path = Path(path)
            self.out_label.config(text=self.output_path.name, fg=FG3)
            self.out_icon.config(text=ICON_CHECK)
            self.out_card.config(highlightbackground=ACCENT)

    # ── Button state ─────────────────────────────────────────────────────

    def _update_button_state(self):
        if not self._verified:
            self.status_label.config(text=self._s("not_verified"), fg=RED)
            self.compare_btn.config(state="disabled")
            return

        # Check license
        valid, msg, _ = _check_license_file()
        if not valid:
            self.status_label.config(text=msg, fg=RED)
            self.compare_btn.config(state="disabled")
            return

        if self.original_path and self.modified_path:
            ext1 = self.original_path.suffix.lower()
            ext2 = self.modified_path.suffix.lower()
            if ext1 != ext2:
                self.status_label.config(
                    text=self._s("format_mismatch", ext1=ext1, ext2=ext2), fg=RED)
                self.compare_btn.config(state="disabled")
                return
            if ext1 not in (".docx",):
                self.status_label.config(
                    text=self._s("format_unsupported", ext=ext1), fg=RED)
                self.compare_btn.config(state="disabled")
                return
            self.status_label.config(text="", fg=FG_DIM)
            self.compare_btn.config(state="normal")
        else:
            self.compare_btn.config(state="disabled")

    def _default_output(self) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return Path.home() / "Desktop" / f"comparison_{ts}.pdf"

    # ── Comparison logic ─────────────────────────────────────────────────

    def _run_comparison(self):
        if not self._verified:
            self.status_label.config(text=self._s("verify_first"), fg=RED)
            return

        valid, msg, _ = _check_license_file()
        if not valid:
            self._show_expired_window(msg)
            return

        self.compare_btn.config(state="disabled")
        self.progress.start(12)
        output = self.output_path or self._default_output()

        def set_status(msg):
            self.root.after(0, lambda: self.status_label.config(text=msg, fg=FG_DIM))

        def worker():
            try:
                from doccompare.comparison.ooxml_engine import compare as ooxml_compare
                from doccompare.rendering.pdf_pipeline import produce_pdf

                set_status(self._s("comparing"))
                doc_tree, summary = ooxml_compare(
                    self.original_path, self.modified_path, None)

                set_status(self._s("rendering"))
                produce_pdf(
                    doc_tree, output, summary,
                    original_name=self.original_path.name,
                    modified_name=self.modified_path.name,
                    docx_path=self.modified_path)

                s = summary
                msg = (
                    f"{self._s('done', filename=output.name)}\n"
                    f"+{s.get('added_words', 0)} {self._s('added')}  "
                    f"\u2212{s.get('deleted_words', 0)} {self._s('deleted')}  "
                    f"{s.get('unchanged_words', 0)} {self._s('unchanged')}"
                )
                self.root.after(0, lambda: self._on_success(msg, output))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, msg: str, output: Path):
        self.progress.stop()
        self.status_label.config(text=msg, fg=GREEN)
        self.compare_btn.config(state="normal")
        import subprocess
        subprocess.run(["open", str(output)])

    def _on_error(self, error: str):
        self.progress.stop()
        self.status_label.config(
            text=f"{self._s('error_prefix')}: {error}", fg=RED)
        self.compare_btn.config(state="normal")


def main():
    root = tk.Tk()
    root.minsize(540, 620)
    DocCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
