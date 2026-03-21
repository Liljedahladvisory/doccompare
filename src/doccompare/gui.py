"""DocCompare GUI — macOS desktop app."""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from datetime import datetime
import sys
import os


SUPPORTED_EXTENSIONS = (".docx", ".pdf")

BG = "#f5f5f7"
ACCENT = "#0071e3"
BTN_FG = "#ffffff"
LABEL_FG = "#1d1d1f"
SUBTITLE_FG = "#6e6e73"
ERROR_FG = "#cc0000"
SUCCESS_FG = "#2e7d32"
FONT_FAMILY = "SF Pro Display" if sys.platform == "darwin" else "Segoe UI"


class DocCompareApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DocCompare")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.original_path: Path | None = None
        self.modified_path: Path | None = None
        self.output_path: Path | None = None

        self._build_ui()
        self._center_window()

    def _build_ui(self):
        root = self.root

        # Title
        tk.Label(root, text="DocCompare", font=(FONT_FAMILY, 22, "bold"),
                 bg=BG, fg=LABEL_FG).pack(pady=(30, 2))
        tk.Label(root, text="Jämför två dokument och generera en PDF-rapport",
                 font=(FONT_FAMILY, 13), bg=BG, fg=SUBTITLE_FG).pack(pady=(0, 24))

        frame = tk.Frame(root, bg=BG, padx=36)
        frame.pack(fill="x")

        # Original
        tk.Label(frame, text="Original dokument", font=(FONT_FAMILY, 12, "bold"),
                 bg=BG, fg=LABEL_FG, anchor="w").pack(fill="x")
        orig_row = tk.Frame(frame, bg=BG)
        orig_row.pack(fill="x", pady=(4, 12))
        self.orig_label = tk.Label(orig_row, text="Ingen fil vald",
                                   font=(FONT_FAMILY, 11), bg="#e8e8ed", fg=SUBTITLE_FG,
                                   anchor="w", padx=10, relief="flat", width=36)
        self.orig_label.pack(side="left", ipady=6, fill="x", expand=True)
        tk.Button(orig_row, text="Välj fil", command=self._pick_original,
                  bg=ACCENT, fg=BTN_FG, font=(FONT_FAMILY, 11),
                  relief="flat", padx=14, cursor="hand2").pack(side="left", padx=(8, 0), ipady=5)

        # Modified
        tk.Label(frame, text="Modifierat dokument", font=(FONT_FAMILY, 12, "bold"),
                 bg=BG, fg=LABEL_FG, anchor="w").pack(fill="x")
        mod_row = tk.Frame(frame, bg=BG)
        mod_row.pack(fill="x", pady=(4, 12))
        self.mod_label = tk.Label(mod_row, text="Ingen fil vald",
                                  font=(FONT_FAMILY, 11), bg="#e8e8ed", fg=SUBTITLE_FG,
                                  anchor="w", padx=10, relief="flat", width=36)
        self.mod_label.pack(side="left", ipady=6, fill="x", expand=True)
        tk.Button(mod_row, text="Välj fil", command=self._pick_modified,
                  bg=ACCENT, fg=BTN_FG, font=(FONT_FAMILY, 11),
                  relief="flat", padx=14, cursor="hand2").pack(side="left", padx=(8, 0), ipady=5)

        # Output
        tk.Label(frame, text="Spara rapport som (valfritt)",
                 font=(FONT_FAMILY, 12, "bold"), bg=BG, fg=LABEL_FG, anchor="w").pack(fill="x")
        out_row = tk.Frame(frame, bg=BG)
        out_row.pack(fill="x", pady=(4, 20))
        self.out_label = tk.Label(out_row, text="Skrivbordet (automatiskt namn)",
                                  font=(FONT_FAMILY, 11), bg="#e8e8ed", fg=SUBTITLE_FG,
                                  anchor="w", padx=10, relief="flat", width=36)
        self.out_label.pack(side="left", ipady=6, fill="x", expand=True)
        tk.Button(out_row, text="Välj plats", command=self._pick_output,
                  bg="#6e6e73", fg=BTN_FG, font=(FONT_FAMILY, 11),
                  relief="flat", padx=14, cursor="hand2").pack(side="left", padx=(8, 0), ipady=5)

        # Compare button
        self.compare_btn = tk.Button(
            root, text="Jämför dokument", command=self._run_comparison,
            bg=ACCENT, fg=BTN_FG, font=(FONT_FAMILY, 14, "bold"),
            relief="flat", padx=24, cursor="hand2", state="disabled",
        )
        self.compare_btn.pack(pady=(0, 16), ipady=10, padx=36, fill="x")

        # Progress bar
        self.progress = ttk.Progressbar(root, mode="indeterminate", length=400)
        self.progress.pack(padx=36, fill="x")

        # Status label
        self.status_label = tk.Label(root, text="", font=(FONT_FAMILY, 11),
                                     bg=BG, fg=SUBTITLE_FG, wraplength=400)
        self.status_label.pack(pady=(8, 24))

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2 - 60
        self.root.geometry(f"+{x}+{y}")

    def _pick_original(self):
        path = filedialog.askopenfilename(
            title="Välj originaldokument",
            filetypes=[("Dokument", "*.docx *.pdf"), ("Word", "*.docx"), ("PDF", "*.pdf")],
        )
        if path:
            self.original_path = Path(path)
            self.orig_label.config(text=self.original_path.name, fg=LABEL_FG)
            self._update_button_state()

    def _pick_modified(self):
        path = filedialog.askopenfilename(
            title="Välj modifierat dokument",
            filetypes=[("Dokument", "*.docx *.pdf"), ("Word", "*.docx"), ("PDF", "*.pdf")],
        )
        if path:
            self.modified_path = Path(path)
            self.mod_label.config(text=self.modified_path.name, fg=LABEL_FG)
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
            self.out_label.config(text=self.output_path.name, fg=LABEL_FG)

    def _update_button_state(self):
        if self.original_path and self.modified_path:
            self.compare_btn.config(state="normal")
        else:
            self.compare_btn.config(state="disabled")

    def _default_output(self) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return Path.home() / "Desktop" / f"comparison_{ts}.pdf"

    def _run_comparison(self):
        self.compare_btn.config(state="disabled")
        self.progress.start(12)
        self.status_label.config(text="Jämför dokument…", fg=SUBTITLE_FG)

        output = self.output_path or self._default_output()

        def worker():
            try:
                from doccompare.parsers import get_parser
                from doccompare.comparison.differ import Differ
                from doccompare.comparison.move_detector import MoveDetector
                from doccompare.rendering.html_builder import HtmlBuilder
                from doccompare.rendering.pdf_renderer import render_pdf

                css_path = Path(__file__).parent / "rendering" / "styles.css"

                self.root.after(0, lambda: self.status_label.config(text="Läser originaldokument…"))
                orig_doc = get_parser(self.original_path).parse(self.original_path)

                self.root.after(0, lambda: self.status_label.config(text="Läser modifierat dokument…"))
                mod_doc = get_parser(self.modified_path).parse(self.modified_path)

                self.root.after(0, lambda: self.status_label.config(text="Analyserar skillnader…"))
                result = Differ().compare(orig_doc, mod_doc)
                result = MoveDetector().detect(result)

                self.root.after(0, lambda: self.status_label.config(text="Genererar PDF-rapport…"))
                html = HtmlBuilder().build(result, self.original_path, self.modified_path)
                render_pdf(html, css_path, output)

                s = result.summary
                msg = (f"Rapport sparad: {output.name}\n"
                       f"+{s.get('added_words',0)} tillagda  "
                       f"−{s.get('deleted_words',0)} borttagna  "
                       f"~{s.get('moved_words',0)} flyttade ord")
                self.root.after(0, lambda: self._on_success(msg, output))

            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, msg: str, output: Path):
        self.progress.stop()
        self.status_label.config(text=msg, fg=SUCCESS_FG)
        self.compare_btn.config(state="normal")
        # Open the PDF
        import subprocess
        subprocess.run(["open", str(output)])

    def _on_error(self, error: str):
        self.progress.stop()
        self.status_label.config(text=f"Fel: {error}", fg=ERROR_FG)
        self.compare_btn.config(state="normal")


def main():
    root = tk.Tk()
    root.minsize(480, 400)
    app = DocCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
