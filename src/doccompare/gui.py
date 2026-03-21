"""DocCompare GUI — macOS desktop app."""
import tkinter as tk
from tkinter import ttk, filedialog
import threading
from pathlib import Path
from datetime import datetime
import sys


BG = "#1c1c1e"
CARD = "#2c2c2e"
FIELD = "#3a3a3c"
FG = "#f5f5f7"
SUBTITLE = "#aeaeb2"
ACCENT = "#0a84ff"
SUCCESS = "#30d158"
ERROR = "#ff453a"
BTN_FG = "#1c1c1e"
FONT = "SF Pro Display" if sys.platform == "darwin" else "Segoe UI"


def _style_progressbar():
    style = ttk.Style()
    style.theme_use("default")
    style.configure(
        "Dark.Horizontal.TProgressbar",
        troughcolor=FIELD,
        background=ACCENT,
        bordercolor=BG,
        lightcolor=ACCENT,
        darkcolor=ACCENT,
    )


class DocCompareApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DocCompare")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.original_path: Path | None = None
        self.modified_path: Path | None = None
        self.output_path: Path | None = None

        _style_progressbar()
        self._build_ui()
        self._center_window()

    def _build_ui(self):
        root = self.root
        outer = tk.Frame(root, bg=BG)
        outer.pack(padx=36, pady=(28, 28), fill="both")

        # ── Logo ──────────────────────────────────────────────────────────
        logo_path = Path(__file__).parent / "assets" / "logo.png"
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path)
            # Scale to height 40px, preserving aspect ratio
            target_h = 40
            ratio = target_h / img.height
            img = img.resize((int(img.width * ratio), target_h), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            tk.Label(outer, image=self._logo_img, bg=BG).pack(anchor="w", pady=(0, 20))
        except Exception:
            # Fallback to text if image can't be loaded
            logo_frame = tk.Frame(outer, bg=BG)
            logo_frame.pack(anchor="w", pady=(0, 20))
            tk.Label(logo_frame, text="Liljedahl", font=(FONT, 18), bg=BG, fg="#8e8e93").pack(side="left")
            tk.Label(logo_frame, text=" Advisory", font=(FONT, 18, "bold"), bg=BG, fg="#636366").pack(side="left")

        # ── Divider ───────────────────────────────────────────────────────
        tk.Frame(outer, bg="#3a3a3c", height=1).pack(fill="x", pady=(0, 20))

        # ── App title ─────────────────────────────────────────────────────
        tk.Label(outer, text="DocCompare", font=(FONT, 22, "bold"), bg=BG, fg=FG).pack(anchor="w")
        tk.Label(
            outer, text="Jämför två dokument och generera en PDF-rapport",
            font=(FONT, 12), bg=BG, fg=SUBTITLE,
        ).pack(anchor="w", pady=(2, 22))

        # ── File pickers ──────────────────────────────────────────────────
        self.orig_label = self._file_row(outer, "Original dokument", self._pick_original)
        self.mod_label = self._file_row(outer, "Modifierat dokument", self._pick_modified)
        self.out_label = self._file_row(
            outer, "Spara rapport som (valfritt)", self._pick_output,
            default_text="Skrivbordet — automatiskt namn", btn_text="Välj plats",
            btn_color="#48484a",
        )

        # ── Compare button ────────────────────────────────────────────────
        self.compare_btn = tk.Button(
            outer, text="Jämför dokument", command=self._run_comparison,
            bg=ACCENT, fg=BTN_FG, font=(FONT, 13, "bold"),
            relief="flat", cursor="hand2", state="disabled",
            activeforeground=BTN_FG,
        )
        self.compare_btn.pack(fill="x", ipady=11, pady=(4, 14))

        # ── Progress & status ─────────────────────────────────────────────
        self.progress = ttk.Progressbar(
            outer, mode="indeterminate", length=400,
            style="Dark.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x")

        self.status_label = tk.Label(
            outer, text="", font=(FONT, 11), bg=BG, fg=SUBTITLE, wraplength=420, justify="left",
        )
        self.status_label.pack(anchor="w", pady=(8, 0))

    def _file_row(
        self, parent, label: str, command,
        default_text="Ingen fil vald", btn_text="Välj fil", btn_color=None,
    ) -> tk.Label:
        tk.Label(parent, text=label, font=(FONT, 11, "bold"), bg=BG, fg=FG, anchor="w").pack(
            fill="x", pady=(0, 4)
        )
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 14))

        lbl = tk.Label(
            row, text=default_text, font=(FONT, 11), bg=FIELD, fg=SUBTITLE,
            anchor="w", padx=10, relief="flat",
        )
        lbl.pack(side="left", ipady=7, fill="x", expand=True)

        tk.Button(
            row, text=btn_text, command=command,
            bg=btn_color or ACCENT, fg=BTN_FG, font=(FONT, 11),
            relief="flat", padx=14, cursor="hand2",
            activeforeground=BTN_FG,
        ).pack(side="left", padx=(8, 0), ipady=6)

        return lbl

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2 - 60}")

    def _pick_original(self):
        path = filedialog.askopenfilename(
            title="Välj originaldokument",
            filetypes=[("Dokument", "*.docx *.pdf"), ("Word", "*.docx"), ("PDF", "*.pdf")],
        )
        if path:
            self.original_path = Path(path)
            self.orig_label.config(text=self.original_path.name, fg=FG)
            self._update_button_state()

    def _pick_modified(self):
        path = filedialog.askopenfilename(
            title="Välj modifierat dokument",
            filetypes=[("Dokument", "*.docx *.pdf"), ("Word", "*.docx"), ("PDF", "*.pdf")],
        )
        if path:
            self.modified_path = Path(path)
            self.mod_label.config(text=self.modified_path.name, fg=FG)
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
            self.out_label.config(text=self.output_path.name, fg=FG)

    def _update_button_state(self):
        state = "normal" if (self.original_path and self.modified_path) else "disabled"
        self.compare_btn.config(state=state)

    def _default_output(self) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return Path.home() / "Desktop" / f"comparison_{ts}.pdf"

    def _run_comparison(self):
        self.compare_btn.config(state="disabled")
        self.progress.start(12)
        output = self.output_path or self._default_output()

        def set_status(msg):
            self.root.after(0, lambda: self.status_label.config(text=msg, fg=SUBTITLE))

        def worker():
            try:
                from doccompare.parsers import get_parser
                from doccompare.comparison.differ import Differ
                from doccompare.comparison.move_detector import MoveDetector
                from doccompare.rendering.html_builder import HtmlBuilder
                from doccompare.rendering.pdf_renderer import render_pdf

                css_path = Path(__file__).parent / "rendering" / "styles.css"

                set_status("Läser originaldokument…")
                orig_doc = get_parser(self.original_path).parse(self.original_path)

                set_status("Läser modifierat dokument…")
                mod_doc = get_parser(self.modified_path).parse(self.modified_path)

                set_status("Analyserar skillnader…")
                result = Differ().compare(orig_doc, mod_doc)
                result = MoveDetector().detect(result)

                set_status("Genererar PDF-rapport…")
                html_content = HtmlBuilder().build(result, self.original_path, self.modified_path)
                render_pdf(html_content, css_path, output)

                s = result.summary
                msg = (
                    f"Klar! Rapport sparad: {output.name}\n"
                    f"+{s.get('added_words', 0)} tillagda  "
                    f"−{s.get('deleted_words', 0)} borttagna  "
                    f"~{s.get('moved_words', 0)} flyttade ord"
                )
                self.root.after(0, lambda: self._on_success(msg, output))

            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, msg: str, output: Path):
        self.progress.stop()
        self.status_label.config(text=msg, fg=SUCCESS)
        self.compare_btn.config(state="normal")
        import subprocess
        subprocess.run(["open", str(output)])

    def _on_error(self, error: str):
        self.progress.stop()
        self.status_label.config(text=f"Fel: {error}", fg=ERROR)
        self.compare_btn.config(state="normal")


def main():
    root = tk.Tk()
    root.minsize(500, 420)
    DocCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
