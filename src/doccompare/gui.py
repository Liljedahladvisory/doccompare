"""DocCompare GUI — modern macOS desktop app."""
import tkinter as tk
from tkinter import ttk, filedialog
import threading
from pathlib import Path
from datetime import datetime
import sys
import os
import json
import random
import urllib.request
import urllib.error
import ssl

# ── Config ──────────────────────────────────────────────────────────────────
CONFIG_PATH = os.path.expanduser("~/.doccompare_llt.json")
WEBHOOK_URL = "https://script.google.com/macros/s/DITT_ID/exec"  # TODO: replace with real GAS URL
VERIFICATION_TIMEOUT = 600  # 10 min


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


# ── Email verification ──────────────────────────────────────────────────────

def _generate_code() -> str:
    """Generate a 4-digit verification code."""
    return str(random.randint(1000, 9999))


def _send_verification_code(email: str, name: str, code: str) -> bool:
    """Send verification code via Google Apps Script webhook. Returns True if OK."""
    payload = {
        "action": "send_verification_email",
        "email": email,
        "name": name,
        "code": code,
    }
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}

    # Try with normal SSL first
    ctx = ssl.create_default_context()
    try:
        req = urllib.request.Request(WEBHOOK_URL, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result.get("status") == "ok"
    except Exception:
        pass

    # Fallback without SSL verification (bundled Python may lack certs)
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


# ── Colour palette (Meeting Recorder LLT warm dark theme) ──────────────────
BG      = "#0E0D0C"
BG2     = "#161412"
BG3     = "#1E1B18"
BG4     = "#272320"
BORDER  = "#332E28"
BORDER2 = "#4A4238"
FG      = "#F2EEE8"       # primary text — warm white
FG2     = "#C8B89A"       # secondary text
FG3     = "#E0D4C0"       # emphasized text
FG_DIM  = "#9A8A72"       # tertiary — hints
ACCENT  = "#E07820"       # orange brand accent
ACCENT2 = "#C05E0A"       # darker orange for hover/active
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

# ── Icon characters ─────────────────────────────────────────────────────────
ICON_CHECK = "\u2713"     # ✓


def _style_widgets():
    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "Orange.Horizontal.TProgressbar",
        troughcolor=BG3,
        background=ACCENT,
        bordercolor=BG,
        lightcolor=ACCENT,
        darkcolor=ACCENT,
    )


class RoundedButton(tk.Canvas):
    """Canvas-based button with smooth rounded corners and full colour control."""

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
                darker = f"#{int(r*.82):02x}{int(g*.82):02x}{int(b*.82):02x}"
                return darker, self._fg
            else:
                return BG2, self._fg
        if self._style == "solid":
            return self._bg, self._fg
        return BG, self._fg

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
        self.create_polygon(pts, smooth=True,
                            fill=fill, outline=outline, width=1)
        self.create_text(w // 2, h // 2, text=self._text,
                         font=self._fspec, fill=fg, anchor="center")

    def config(self, **kw):
        changed = False
        if "text" in kw:
            self._text = kw.pop("text"); changed = True
        if "state" in kw:
            self._enabled = (kw.pop("state") == "normal"); changed = True
        if "bg" in kw:
            self._bg = kw.pop("bg"); changed = True
        if "fg" in kw:
            self._fg = kw.pop("fg"); changed = True
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
        if self._enabled and self._command:
            self._command()

    def _on_enter(self, _=None):
        self._hovering = True; self._draw()

    def _on_leave(self, _=None):
        self._hovering = False; self._draw()


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

        self.original_path: Path | None = None
        self.modified_path: Path | None = None
        self.output_path: Path | None = None

        _style_widgets()
        self._build_ui()
        self._center_window()

        # Show registration on first run (not yet verified)
        if not self._verified:
            self.root.after(300, self._show_registration_dialog)

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        outer = tk.Frame(root, bg=BG)
        outer.pack(padx=32, pady=(24, 24), fill="both", expand=True)

        # ── Header area ──────────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x", pady=(0, 16))

        left = tk.Frame(header, bg=BG)
        left.pack(side="left", fill="y")

        # Name row — shows user name (bold) + "  DocCompare"
        name_row = tk.Frame(left, bg=BG)
        name_row.pack(anchor="w")

        self._logo_name_lbl = tk.Label(
            name_row, text=self._display_name(),
            font=FONT_LOGO1, bg=BG, fg=FG,
        )
        self._logo_name_lbl.pack(side="left")

        tk.Label(name_row, text="  DocCompare",
                 font=FONT_LOGO2, bg=BG, fg=FG2).pack(side="left")

        # "Powered by Liljedahl Legal Tech" — italic, underneath
        tk.Label(left, text="Powered by Liljedahl Legal Tech",
                 font=FONT_POWERED, bg=BG, fg=FG_DIM).pack(anchor="w", pady=(3, 0))

        # Gear icon (settings) — right side of header
        right = tk.Frame(header, bg=BG)
        right.pack(side="right", fill="y")

        gear = tk.Label(right, text="\u2699", font=FONT_GEAR,
                        bg=BG, fg=FG_DIM, cursor="hand2")
        gear.pack(side="right")
        gear.bind("<Enter>",    lambda _: gear.config(fg=FG))
        gear.bind("<Leave>",    lambda _: gear.config(fg=FG_DIM))
        gear.bind("<Button-1>", lambda _: self._show_settings_dialog())

        # ── Title block ──────────────────────────────────────────────────
        title_frame = tk.Frame(outer, bg=BG)
        title_frame.pack(fill="x", pady=(0, 6))

        tk.Label(
            title_frame, text="DocCompare",
            font=FONT_TITLE, bg=BG, fg=FG,
        ).pack(side="left")

        # Version badge
        badge_frame = tk.Frame(title_frame, bg=BG3, padx=8, pady=2,
                               highlightbackground=BORDER, highlightthickness=1)
        badge_frame.pack(side="left", padx=(12, 0), pady=(8, 0))
        tk.Label(badge_frame, text="v0.1", font=FONT_MS, bg=BG3,
                 fg=ACCENT).pack()

        tk.Label(
            outer, text="Jämför .docx-dokument och generera en diff-rapport",
            font=FONT_SUB, bg=BG, fg=FG_DIM,
        ).pack(anchor="w", pady=(0, 16))

        # ── Thin accent line ─────────────────────────────────────────────
        tk.Frame(outer, bg=ACCENT, height=2).pack(fill="x", pady=(0, 16))

        # ── File picker cards ────────────────────────────────────────────
        self.orig_card, self.orig_label, self.orig_icon = self._file_card(
            outer, "Originaldokument", "Välj fil", self._pick_original,
        )
        self.mod_card, self.mod_label, self.mod_icon = self._file_card(
            outer, "Ändrat dokument", "Välj fil", self._pick_modified,
        )
        self.out_card, self.out_label, self.out_icon = self._file_card(
            outer, "Spara rapport som", "Välj plats", self._pick_output,
            default_text="Skrivbordet — automatiskt filnamn",
            optional=True,
        )

        # ── Compare button ───────────────────────────────────────────────
        btn_frame = tk.Frame(outer, bg=BG)
        btn_frame.pack(fill="x", pady=(8, 16))

        self.compare_btn = RoundedButton(
            btn_frame, text="Jämför dokument",
            command=self._run_comparison,
            bg=ACCENT, fg="#FFFFFF",
            font_spec=FONT_BTN,
            padx=40, pady=14, radius=12,
            state="disabled",
            fixed_width=536,
        )
        self.compare_btn.pack(fill="x")

        # ── Progress bar ─────────────────────────────────────────────────
        self.progress = ttk.Progressbar(
            outer, mode="indeterminate", length=400,
            style="Orange.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(0, 4))

        # ── Status label ─────────────────────────────────────────────────
        self.status_label = tk.Label(
            outer, text="",
            font=FONT_B, bg=BG, fg=FG_DIM,
            wraplength=500, justify="left", anchor="w",
        )
        self.status_label.pack(anchor="w", fill="x", pady=(6, 0))

        # ── Footer ───────────────────────────────────────────────────────
        footer = tk.Frame(outer, bg=BG)
        footer.pack(side="bottom", fill="x", pady=(16, 0))
        tk.Label(
            footer,
            text="Liljedahl Legal Tech  \u2022  Liljedahl Advisory AB",
            font=FONT_XS, bg=BG, fg=BORDER2,
        ).pack()

    def _file_card(
        self, parent, title: str, btn_text: str, command,
        default_text="Ingen fil vald", optional=False,
    ):
        """Create a card-style file picker row."""
        card = tk.Frame(parent, bg=BG2,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", pady=(0, 10))

        inner = tk.Frame(card, bg=BG2)
        inner.pack(fill="x", padx=14, pady=12)

        left = tk.Frame(inner, bg=BG2)
        left.pack(side="left", fill="x", expand=True)

        top_row = tk.Frame(left, bg=BG2)
        top_row.pack(anchor="w")

        tk.Label(
            top_row, text=title,
            font=FONT_H, bg=BG2, fg=FG,
        ).pack(side="left")

        if optional:
            tk.Label(
                top_row, text="valfritt",
                font=FONT_XS, bg=BG2, fg=FG_DIM,
            ).pack(side="left", padx=(8, 0))

        file_lbl = tk.Label(
            left, text=default_text,
            font=FONT_S, bg=BG2, fg=FG_DIM,
            anchor="w",
        )
        file_lbl.pack(anchor="w", pady=(3, 0))

        btn = RoundedButton(
            inner, text=btn_text, command=command,
            style="ghost", bg=BG2, fg=FG2,
            font_spec=FONT_S, padx=16, pady=8, radius=8,
        )
        btn.pack(side="right", padx=(12, 0))

        icon_lbl = tk.Label(
            inner, text="",
            font=("Helvetica Neue", 13), bg=BG2, fg=GREEN,
        )
        icon_lbl.pack(side="right", padx=(0, 4))

        return card, file_lbl, icon_lbl

    def _display_name(self) -> str:
        if self.user_name:
            return self.user_name
        return "DocCompare"

    # ── Window centering ─────────────────────────────────────────────────

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2 - 60}")

    # ── Registration dialog (email verification) ─────────────────────────

    def _show_registration_dialog(self):
        """Two-step registration: 1) name + email → send code, 2) enter code."""
        self._pending_code = None  # the code we sent

        dlg = tk.Toplevel(self.root)
        dlg.title("Registrera DocCompare")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 480, 420
        x = self.root.winfo_x() + (self.root.winfo_width()  - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        pad = tk.Frame(dlg, bg=BG, padx=36, pady=28)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text="Välkommen till DocCompare",
                 font=("Helvetica Neue", 15, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(pad,
                 text="Ange ditt namn och e-postadress.\n"
                      "En verifieringskod skickas till din e-post.",
                 font=FONT_S, bg=BG, fg=FG2,
                 wraplength=400, justify="left").pack(anchor="w", pady=(8, 20))

        # ── Step 1 fields: name + email ──────────────────────────────────
        self._reg_step1 = tk.Frame(pad, bg=BG)
        self._reg_step1.pack(fill="x")

        tk.Label(self._reg_step1, text="NAMN / FÖRETAG", font=FONT_SECTION,
                 bg=BG, fg=FG_DIM).pack(anchor="w")
        name_var = tk.StringVar()
        name_entry = tk.Entry(self._reg_step1, textvariable=name_var, font=FONT_M,
                              bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                              highlightthickness=1, highlightbackground=BORDER2,
                              highlightcolor=ACCENT)
        name_entry.pack(fill="x", ipady=9, pady=(4, 14))
        name_entry.focus_set()

        tk.Label(self._reg_step1, text="E-POSTADRESS", font=FONT_SECTION,
                 bg=BG, fg=FG_DIM).pack(anchor="w")
        email_var = tk.StringVar()
        email_entry = tk.Entry(self._reg_step1, textvariable=email_var, font=FONT_M,
                               bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                               highlightthickness=1, highlightbackground=BORDER2,
                               highlightcolor=ACCENT)
        email_entry.pack(fill="x", ipady=9, pady=(4, 6))

        # ── Step 2 fields: verification code (hidden initially) ──────────
        self._reg_step2 = tk.Frame(pad, bg=BG)
        # Not packed yet — shown after code is sent

        tk.Label(self._reg_step2, text="VERIFIERINGSKOD", font=FONT_SECTION,
                 bg=BG, fg=FG_DIM).pack(anchor="w")
        self._code_info_lbl = tk.Label(
            self._reg_step2,
            text="En 4-siffrig kod har skickats till din e-post.",
            font=FONT_XS, bg=BG, fg=FG2,
        )
        self._code_info_lbl.pack(anchor="w", pady=(2, 6))
        code_var = tk.StringVar()
        code_entry = tk.Entry(self._reg_step2, textvariable=code_var, font=("Menlo", 18),
                              bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                              highlightthickness=1, highlightbackground=BORDER2,
                              highlightcolor=ACCENT, justify="center")
        code_entry.pack(fill="x", ipady=10, pady=(0, 6))

        # Error / status label
        error_lbl = tk.Label(pad, text="", font=FONT_XS, bg=BG, fg=RED, anchor="w")
        error_lbl.pack(anchor="w", fill="x", pady=(4, 0))

        # Buttons
        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x", pady=(12, 0), side="bottom")

        # ── Step 1 action: send code ─────────────────────────────────────
        def send_code():
            name = name_var.get().strip()
            email = email_var.get().strip()

            if not name:
                name_entry.config(highlightbackground=RED)
                error_lbl.config(text="Ange ditt namn.")
                return
            name_entry.config(highlightbackground=BORDER2)

            if not email or "@" not in email:
                email_entry.config(highlightbackground=RED)
                error_lbl.config(text="Ange en giltig e-postadress.")
                return
            email_entry.config(highlightbackground=BORDER2)

            # Generate and send code
            error_lbl.config(text="Skickar verifieringskod\u2026", fg=FG_DIM)
            dlg.update()

            code = _generate_code()
            self._pending_code = code

            def _do_send():
                sent = _send_verification_code(email, name, code)
                dlg.after(0, lambda: _on_sent(sent))

            def _on_sent(sent):
                if sent:
                    # Switch to step 2
                    self._reg_step1.pack_forget()
                    self._reg_step2.pack(fill="x")
                    self._code_info_lbl.config(
                        text=f"En 4-siffrig kod har skickats till {email}.")
                    error_lbl.config(text="", fg=RED)
                    send_btn.pack_forget()
                    verify_btn.pack(side="right")
                    resend_btn.pack(side="right", padx=(0, 8))
                    code_entry.focus_set()
                else:
                    error_lbl.config(
                        text="Kunde inte skicka e-post. Kontrollera adressen och försök igen.",
                        fg=RED)

            threading.Thread(target=_do_send, daemon=True).start()

        send_btn = tk.Label(btn_row, text="Skicka kod", font=FONT_B,
                            bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                            cursor="hand2")
        send_btn.pack(side="right")
        send_btn.bind("<Button-1>", lambda _: send_code())
        send_btn.bind("<Enter>", lambda _: send_btn.config(bg=ACCENT2))
        send_btn.bind("<Leave>", lambda _: send_btn.config(bg=ACCENT))

        email_entry.bind("<Return>", lambda _: send_code())
        name_entry.bind("<Return>", lambda _: email_entry.focus_set())

        # ── Step 2 action: verify code ───────────────────────────────────
        def verify_code():
            entered = code_var.get().strip()
            if not entered:
                code_entry.config(highlightbackground=RED)
                error_lbl.config(text="Ange verifieringskoden.", fg=RED)
                return

            if entered == self._pending_code:
                # Success!
                name = name_var.get().strip()
                email = email_var.get().strip()
                self.user_name = name
                self._verified = True
                self._config["user_name"] = name
                self._config["email"] = email
                self._config["verified"] = True
                save_config(self._config)
                self._logo_name_lbl.config(text=self._display_name())
                dlg.destroy()
            else:
                code_entry.config(highlightbackground=RED)
                error_lbl.config(text="Fel kod. Kontrollera din e-post och försök igen.",
                                 fg=RED)

        verify_btn = tk.Label(btn_row, text="Verifiera", font=FONT_B,
                              bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                              cursor="hand2")
        verify_btn.bind("<Button-1>", lambda _: verify_code())
        verify_btn.bind("<Enter>", lambda _: verify_btn.config(bg=ACCENT2))
        verify_btn.bind("<Leave>", lambda _: verify_btn.config(bg=ACCENT))
        # Not packed yet — shown in step 2

        # Resend button
        def resend():
            code = _generate_code()
            self._pending_code = code
            email = email_var.get().strip()
            name = name_var.get().strip()
            error_lbl.config(text="Skickar ny kod\u2026", fg=FG_DIM)
            dlg.update()

            def _do():
                sent = _send_verification_code(email, name, code)
                dlg.after(0, lambda: error_lbl.config(
                    text="Ny kod skickad!" if sent else "Kunde inte skicka. Försök igen.",
                    fg=GREEN if sent else RED))
            threading.Thread(target=_do, daemon=True).start()

        resend_btn = tk.Label(btn_row, text="Skicka igen", font=FONT_B,
                              bg=BG3, fg=FG2, padx=24, pady=8,
                              cursor="hand2")
        resend_btn.bind("<Button-1>", lambda _: resend())
        resend_btn.bind("<Enter>", lambda _: resend_btn.config(bg=BG4, fg=FG))
        resend_btn.bind("<Leave>", lambda _: resend_btn.config(bg=BG3, fg=FG2))
        # Not packed yet — shown in step 2

        code_entry.bind("<Return>", lambda _: verify_code())

        # Cannot close without verifying
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)
        dlg.wait_window()

    # ── Settings dialog (change name) ────────────────────────────────────

    def _show_settings_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Inställningar")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.grab_set()

        self.root.update_idletasks()
        dw, dh = 440, 240
        x = self.root.winfo_x() + (self.root.winfo_width()  - dw) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dh) // 2
        dlg.geometry(f"{dw}x{dh}+{x}+{y}")

        pad = tk.Frame(dlg, bg=BG, padx=36, pady=32)
        pad.pack(fill="both", expand=True)

        tk.Label(pad, text="Inställningar", font=("Helvetica Neue", 13, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(pad, text="Ändra namn eller företagsnamn:",
                 font=FONT_S, bg=BG, fg=FG2,
                 wraplength=368, justify="left").pack(anchor="w", pady=(8, 18))

        entry_var = tk.StringVar(value=self.user_name)
        entry = tk.Entry(pad, textvariable=entry_var, font=FONT_M,
                         bg=BG3, fg=FG, insertbackground=FG, relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=BORDER2,
                         highlightcolor=ACCENT)
        entry.pack(fill="x", ipady=9)
        entry.focus_set()
        entry.select_range(0, "end")

        btn_row = tk.Frame(pad, bg=BG)
        btn_row.pack(fill="x", pady=(18, 0))

        def save():
            name = entry_var.get().strip()
            if not name:
                entry.config(highlightbackground=RED)
                return
            self.user_name = name
            self._config["user_name"] = name
            save_config(self._config)
            self._logo_name_lbl.config(text=self._display_name())
            dlg.destroy()

        save_lbl = tk.Label(btn_row, text="Spara", font=FONT_B,
                            bg=ACCENT, fg="#FFFFFF", padx=24, pady=8,
                            cursor="hand2")
        save_lbl.pack(side="right")
        save_lbl.bind("<Button-1>", lambda _: save())
        save_lbl.bind("<Enter>",  lambda _: save_lbl.config(bg=ACCENT2))
        save_lbl.bind("<Leave>",  lambda _: save_lbl.config(bg=ACCENT))

        cancel_lbl = tk.Label(btn_row, text="Avbryt", font=FONT_B,
                              bg=BG3, fg=FG2, padx=24, pady=8,
                              cursor="hand2")
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
            title="Välj originaldokument",
            filetypes=[("Word-dokument", "*.docx")],
        )
        if path:
            self.original_path = Path(path)
            self.orig_label.config(text=self.original_path.name, fg=FG3)
            self.orig_icon.config(text=ICON_CHECK)
            self.orig_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_modified(self):
        path = filedialog.askopenfilename(
            title="Välj ändrat dokument",
            filetypes=[("Word-dokument", "*.docx")],
        )
        if path:
            self.modified_path = Path(path)
            self.mod_label.config(text=self.modified_path.name, fg=FG3)
            self.mod_icon.config(text=ICON_CHECK)
            self.mod_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Spara rapport",
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
            self.status_label.config(
                text="Programmet är inte aktiverat. Starta om och verifiera din e-post.",
                fg=RED,
            )
            self.compare_btn.config(state="disabled")
            return

        if self.original_path and self.modified_path:
            ext1 = self.original_path.suffix.lower()
            ext2 = self.modified_path.suffix.lower()
            if ext1 != ext2:
                self.status_label.config(
                    text=f"Båda filerna måste vara samma format ({ext1} \u2260 {ext2})",
                    fg=RED,
                )
                self.compare_btn.config(state="disabled")
                return
            if ext1 not in (".docx",):
                self.status_label.config(
                    text=f"Format stöds ej: {ext1}",
                    fg=RED,
                )
                self.compare_btn.config(state="disabled")
                return
            self.status_label.config(text="", fg=FG_DIM)
            self.compare_btn.config(state="normal")
        else:
            self.compare_btn.config(state="disabled")

    def _default_output(self) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return Path.home() / "Desktop" / f"jämförelse_{ts}.pdf"

    # ── Comparison logic ─────────────────────────────────────────────────

    def _run_comparison(self):
        if not self._verified:
            self.status_label.config(text="Verifiera din e-post först.", fg=RED)
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

                set_status("Jämför dokument\u2026")
                doc_tree, summary = ooxml_compare(
                    self.original_path, self.modified_path, None,
                )

                set_status("Renderar PDF\u2026")
                produce_pdf(
                    doc_tree, output, summary,
                    original_name=self.original_path.name,
                    modified_name=self.modified_path.name,
                    docx_path=self.modified_path,
                )

                s = summary
                msg = (
                    f"Klart! Rapport sparad: {output.name}\n"
                    f"+{s.get('added_words', 0)} tillagda  "
                    f"\u2212{s.get('deleted_words', 0)} borttagna  "
                    f"{s.get('unchanged_words', 0)} oförändrade"
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
        self.status_label.config(text=f"Fel: {error}", fg=RED)
        self.compare_btn.config(state="normal")


def main():
    root = tk.Tk()
    root.minsize(540, 620)
    DocCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
