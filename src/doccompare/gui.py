"""DocCompare GUI — modern macOS desktop app."""
import tkinter as tk
from tkinter import ttk, filedialog
import threading
from pathlib import Path
from datetime import datetime
import sys

# ── Color palette ────────────────────────────────────────────────────────
BG = "#111114"
SURFACE = "#1a1a1f"
CARD = "#222228"
FIELD = "#2a2a32"
FIELD_HOVER = "#32323c"
BORDER = "#333340"
FG = "#f0f0f5"
FG_DIM = "#a0a0b0"
SUBTITLE = "#78788a"
ACCENT = "#e8692a"          # Warm orange
ACCENT_HOVER = "#f07a3a"
ACCENT_MUTED = "#3a2518"    # Very dark orange for subtle backgrounds
SUCCESS = "#34c759"
ERROR = "#ff453a"
BTN_FG = "#ffffff"
DISABLED_BG = "#2a2a32"
DISABLED_FG = "#555566"
FONT = "SF Pro Display" if sys.platform == "darwin" else "Segoe UI"
FONT_MONO = "SF Mono" if sys.platform == "darwin" else "Consolas"

# ── Icon characters (SF Symbols fallback) ────────────────────────────────
ICON_DOC = "\U0001F4C4"      # 📄
ICON_SAVE = "\U0001F4BE"     # 💾
ICON_CHECK = "\u2713"        # ✓
ICON_ARROW = "\u276F"        # ❯


def _style_widgets():
    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "Orange.Horizontal.TProgressbar",
        troughcolor=FIELD,
        background=ACCENT,
        bordercolor=BG,
        lightcolor=ACCENT,
        darkcolor=ACCENT,
    )


class RoundedFrame(tk.Canvas):
    """A canvas that draws a rounded-rectangle background to simulate cards."""

    def __init__(self, parent, bg_color=CARD, corner=12, border_color=BORDER, **kw):
        super().__init__(parent, highlightthickness=0, bg=parent["bg"], **kw)
        self._bg_color = bg_color
        self._border_color = border_color
        self._corner = corner
        self._inner = tk.Frame(self, bg=bg_color)
        self._inner_id = self.create_window(0, 0, anchor="nw", window=self._inner)
        self.bind("<Configure>", self._redraw)

    def _redraw(self, event=None):
        self.delete("bg")
        w, h, r = self.winfo_width(), self.winfo_height(), self._corner
        # Border rectangle
        self._rounded_rect(1, 1, w - 1, h - 1, r, self._border_color, "bg")
        # Fill rectangle
        self._rounded_rect(2, 2, w - 2, h - 2, r - 1, self._bg_color, "bg")
        self.tag_lower("bg")
        self.itemconfigure(self._inner_id, width=w - 8, height=h - 8)
        self.coords(self._inner_id, 4, 4)

    def _rounded_rect(self, x1, y1, x2, y2, r, color, tag):
        self.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, fill=color, outline=color, tags=tag)
        self.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, fill=color, outline=color, tags=tag)
        self.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, fill=color, outline=color, tags=tag)
        self.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, fill=color, outline=color, tags=tag)
        self.create_rectangle(x1 + r, y1, x2 - r, y2, fill=color, outline=color, tags=tag)
        self.create_rectangle(x1, y1 + r, x2, y2 - r, fill=color, outline=color, tags=tag)

    @property
    def inner(self):
        return self._inner


class DocCompareApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DocCompare")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.original_path: Path | None = None
        self.modified_path: Path | None = None
        self.output_path: Path | None = None

        _style_widgets()
        self._build_ui()
        self._center_window()

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root
        outer = tk.Frame(root, bg=BG)
        outer.pack(padx=32, pady=(24, 24), fill="both")

        # ── Header area ──────────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x", pady=(0, 20))

        # Logo — crop transparent padding from the PNG so it aligns flush
        logo_path = Path(__file__).parent / "assets" / "logo.png"
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path).convert("RGBA")
            bbox = img.getbbox()  # crop to visible pixels
            if bbox:
                img = img.crop(bbox)
            target_h = 32
            ratio = target_h / img.height
            img = img.resize((int(img.width * ratio), target_h), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            tk.Label(header, image=self._logo_img, bg=BG).pack(anchor="w")
        except Exception:
            logo_frame = tk.Frame(header, bg=BG)
            logo_frame.pack(anchor="w")
            tk.Label(logo_frame, text="Liljedahl", font=(FONT, 15), bg=BG, fg="#707080").pack(side="left")
            tk.Label(logo_frame, text=" Advisory", font=(FONT, 15, "bold"), bg=BG, fg="#50505e").pack(side="left")

        tk.Label(
            header, text="Liljedahl Legal Tech Tools",
            font=(FONT_MONO, 9, "italic"), bg=BG, fg=SUBTITLE, anchor="w",
        ).pack(anchor="w", pady=(3, 0))

        # ── Title block ──────────────────────────────────────────────────
        title_frame = tk.Frame(outer, bg=BG)
        title_frame.pack(fill="x", pady=(0, 6))

        tk.Label(
            title_frame, text="DocCompare",
            font=(FONT, 26, "bold"), bg=BG, fg=FG,
        ).pack(side="left")

        # Version badge
        badge = tk.Label(
            title_frame, text="v0.1",
            font=(FONT_MONO, 9), bg=ACCENT_MUTED, fg=ACCENT,
            padx=8, pady=2,
        )
        badge.pack(side="left", padx=(12, 0), pady=(8, 0))

        tk.Label(
            outer, text="Compare two documents and generate a PDF diff report",
            font=(FONT, 12), bg=BG, fg=SUBTITLE,
        ).pack(anchor="w", pady=(0, 20))

        # ── Thin accent line ─────────────────────────────────────────────
        tk.Frame(outer, bg=ACCENT, height=2).pack(fill="x", pady=(0, 20))

        # ── File picker cards ────────────────────────────────────────────
        self.orig_card, self.orig_label, self.orig_icon = self._file_card(
            outer, "Original document", "Select file", self._pick_original,
        )
        self.mod_card, self.mod_label, self.mod_icon = self._file_card(
            outer, "Modified document", "Select file", self._pick_modified,
        )
        self.out_card, self.out_label, self.out_icon = self._file_card(
            outer, "Save report as", "Choose location", self._pick_output,
            default_text="Desktop — automatic filename",
            optional=True,
        )

        # ── Compare button ───────────────────────────────────────────────
        self.compare_btn = tk.Button(
            outer, text="Compare Documents",
            command=self._run_comparison,
            bg=ACCENT, fg=BTN_FG,
            activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
            font=(FONT, 14, "bold"),
            relief="flat", cursor="hand2", state="disabled",
            disabledforeground=DISABLED_FG,
            borderwidth=0, highlightthickness=0,
        )
        self.compare_btn.pack(fill="x", ipady=13, pady=(8, 16))
        self._style_button_disabled()

        # ── Progress bar ─────────────────────────────────────────────────
        self.progress = ttk.Progressbar(
            outer, mode="indeterminate", length=400,
            style="Orange.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(0, 4))

        # ── Status label ─────────────────────────────────────────────────
        self.status_label = tk.Label(
            outer, text="",
            font=(FONT, 11), bg=BG, fg=SUBTITLE,
            wraplength=460, justify="left", anchor="w",
        )
        self.status_label.pack(anchor="w", fill="x", pady=(6, 0))

        # ── Footer ───────────────────────────────────────────────────────
        tk.Label(
            outer,
            text="Liljedahl Legal Tech  \u2022  Liljedahl Advisory AB",
            font=(FONT, 9), bg=BG, fg="#44444e",
        ).pack(side="bottom", pady=(16, 0))

    def _file_card(
        self, parent, title: str, btn_text: str, command,
        default_text="No file selected", optional=False,
    ):
        """Create a card-style file picker row."""
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", pady=(0, 10), ipady=0)

        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="x", padx=14, pady=12)

        # Left side: icon + labels
        left = tk.Frame(inner, bg=CARD)
        left.pack(side="left", fill="x", expand=True)

        top_row = tk.Frame(left, bg=CARD)
        top_row.pack(anchor="w")

        title_lbl = tk.Label(
            top_row, text=title,
            font=(FONT, 11, "bold"), bg=CARD, fg=FG,
        )
        title_lbl.pack(side="left")

        if optional:
            tk.Label(
                top_row, text="optional",
                font=(FONT, 9), bg=CARD, fg=SUBTITLE,
            ).pack(side="left", padx=(8, 0))

        file_lbl = tk.Label(
            left, text=default_text,
            font=(FONT, 10), bg=CARD, fg=SUBTITLE,
            anchor="w",
        )
        file_lbl.pack(anchor="w", pady=(3, 0))

        # Right side: button
        btn = tk.Button(
            inner, text=btn_text, command=command,
            bg=FIELD, fg=FG_DIM,
            activebackground=FIELD_HOVER, activeforeground=FG,
            font=(FONT, 10),
            relief="flat", cursor="hand2",
            padx=16, borderwidth=0, highlightthickness=0,
        )
        btn.pack(side="right", ipady=6, padx=(12, 0))

        # Status icon (right of label, hidden initially)
        icon_lbl = tk.Label(
            inner, text="",
            font=(FONT, 13), bg=CARD, fg=SUCCESS,
        )
        icon_lbl.pack(side="right", padx=(0, 4))

        return card, file_lbl, icon_lbl

    # ── Window centering ─────────────────────────────────────────────────

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2 - 60}")

    # ── File pickers ─────────────────────────────────────────────────────

    def _pick_original(self):
        path = filedialog.askopenfilename(
            title="Select original document",
            filetypes=[("Word", "*.docx")],
        )
        if path:
            self.original_path = Path(path)
            self.orig_label.config(text=self.original_path.name, fg=FG)
            self.orig_icon.config(text=ICON_CHECK)
            self.orig_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_modified(self):
        path = filedialog.askopenfilename(
            title="Select modified document",
            filetypes=[("Word", "*.docx")],
        )
        if path:
            self.modified_path = Path(path)
            self.mod_label.config(text=self.modified_path.name, fg=FG)
            self.mod_icon.config(text=ICON_CHECK)
            self.mod_card.config(highlightbackground=ACCENT)
            self._update_button_state()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save report",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=Path.home() / "Desktop",
        )
        if path:
            self.output_path = Path(path)
            self.out_label.config(text=self.output_path.name, fg=FG)
            self.out_icon.config(text=ICON_CHECK)
            self.out_card.config(highlightbackground=ACCENT)

    # ── Button state ─────────────────────────────────────────────────────

    def _update_button_state(self):
        if self.original_path and self.modified_path:
            self.compare_btn.config(
                state="normal", bg=ACCENT,
                disabledforeground=DISABLED_FG,
            )
        else:
            self._style_button_disabled()

    def _style_button_disabled(self):
        self.compare_btn.config(state="disabled", bg=DISABLED_BG)

    def _default_output(self) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return Path.home() / "Desktop" / f"comparison_{ts}.pdf"

    # ── Comparison logic ─────────────────────────────────────────────────

    def _run_comparison(self):
        self.compare_btn.config(state="disabled", bg=DISABLED_BG)
        self.progress.start(12)
        output = self.output_path or self._default_output()

        def set_status(msg):
            self.root.after(0, lambda: self.status_label.config(text=msg, fg=FG_DIM))

        def worker():
            try:
                from doccompare.comparison.ooxml_engine import compare as ooxml_compare
                from doccompare.rendering.pdf_pipeline import produce_pdf

                set_status("Comparing documents\u2026")
                doc_tree, summary = ooxml_compare(
                    self.original_path, self.modified_path, None,
                )

                set_status("Rendering PDF\u2026")
                produce_pdf(
                    doc_tree, output, summary,
                    original_name=self.original_path.name,
                    modified_name=self.modified_path.name,
                    docx_path=self.modified_path,
                )

                s = summary
                msg = (
                    f"Done! Report saved: {output.name}\n"
                    f"+{s.get('added_words', 0)} added  "
                    f"\u2212{s.get('deleted_words', 0)} deleted  "
                    f"{s.get('unchanged_words', 0)} unchanged"
                )
                self.root.after(0, lambda: self._on_success(msg, output))

            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, msg: str, output: Path):
        self.progress.stop()
        self.status_label.config(text=msg, fg=SUCCESS)
        self.compare_btn.config(state="normal", bg=ACCENT)
        import subprocess
        subprocess.run(["open", str(output)])

    def _on_error(self, error: str):
        self.progress.stop()
        self.status_label.config(text=f"Error: {error}", fg=ERROR)
        self.compare_btn.config(state="normal", bg=ACCENT)


def main():
    root = tk.Tk()
    root.minsize(520, 520)
    DocCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
