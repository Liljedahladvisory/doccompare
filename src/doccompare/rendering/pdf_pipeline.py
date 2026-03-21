"""PDF pipeline: convert Track-Changes .docx to PDF via Microsoft Word,
then append a summary/legend page."""

import html
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF


def docx_to_pdf(docx_path: Path, pdf_path: Path, timeout: int = 60):
    """Convert .docx to PDF using Microsoft Word via AppleScript.

    Word renders the document with Track Changes markup visible (balloons
    or inline), producing a faithful PDF of the redlined document.
    """
    docx_abs = str(docx_path.resolve())
    pdf_abs = str(pdf_path.resolve())

    # AppleScript: open in Word, save as PDF, close
    script = f'''
        tell application "Microsoft Word"
            activate
            open POSIX file "{docx_abs}"
            delay 2
            set theDoc to active document
            save as theDoc file name POSIX file "{pdf_abs}" file format format PDF
            close theDoc saving no
        end tell
    '''
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Word PDF export failed: {result.stderr.strip()}")

    # Wait for PDF to appear on disk
    for _ in range(20):
        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            return
        time.sleep(0.5)
    raise RuntimeError("Word PDF export timed out — file not found")


def build_summary_pdf(
    summary: dict,
    original_name: str,
    modified_name: str,
) -> bytes:
    """Build a one-page summary/legend PDF using PyMuPDF.

    Returns raw PDF bytes.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    added = summary.get("added_words", 0)
    deleted = summary.get("deleted_words", 0)
    unchanged = summary.get("unchanged_words", 0)

    doc = fitz.open()
    page = doc.new_page(width=595.28, height=841.89)  # A4

    # ── Header ──────────────────────────────────────────────────────
    y = 72  # 1 inch top margin
    page.insert_text(
        (72, y), "DocCompare — Sammanfattning",
        fontname="helv", fontsize=18, color=(0.17, 0.24, 0.31),
    )
    y += 30
    page.draw_line((72, y), (523, y), color=(0.74, 0.76, 0.78), width=0.5)
    y += 20

    # ── Meta ────────────────────────────────────────────────────────
    meta_font = 9
    page.insert_text((72, y), f"Original: {original_name}", fontname="helv", fontsize=meta_font, color=(0.33, 0.33, 0.33))
    y += 14
    page.insert_text((72, y), f"Modifierat: {modified_name}", fontname="helv", fontsize=meta_font, color=(0.33, 0.33, 0.33))
    y += 14
    page.insert_text((72, y), f"Datum: {now}", fontname="helv", fontsize=meta_font, color=(0.33, 0.33, 0.33))
    y += 30

    # ── Statistics ──────────────────────────────────────────────────
    page.insert_text((72, y), "Statistik", fontname="hebo", fontsize=13, color=(0.17, 0.24, 0.31))
    y += 24

    stats = [
        (f"+{added} ord tillagda", (0.0, 0.28, 0.67)),    # blue
        (f"−{deleted} ord borttagna", (0.75, 0.22, 0.17)),  # red
        (f"{unchanged} ord oförändrade", (0.33, 0.33, 0.33)),
    ]
    for text, color in stats:
        page.insert_text((90, y), text, fontname="helv", fontsize=11, color=color)
        y += 20

    y += 20

    # ── Legend ───────────────────────────────────────────────────────
    page.insert_text((72, y), "Legend", fontname="hebo", fontsize=13, color=(0.17, 0.24, 0.31))
    y += 24

    legend_items = [
        ("Tillagd text", "Text som finns i det modifierade dokumentet men inte i originalet.",
         (0.0, 0.28, 0.67)),
        ("Borttagen text", "Text som finns i originalet men inte i det modifierade dokumentet.",
         (0.75, 0.22, 0.17)),
        ("Oförändrad text", "Text som är identisk i båda dokumenten.",
         (0.10, 0.10, 0.10)),
    ]
    for label, desc, color in legend_items:
        page.insert_text((90, y), label, fontname="hebo", fontsize=11, color=color)
        y += 16
        page.insert_text((90, y), desc, fontname="helv", fontsize=9, color=(0.44, 0.44, 0.44))
        y += 22

    y += 20
    page.draw_line((72, y), (523, y), color=(0.74, 0.76, 0.78), width=0.5)
    y += 16
    page.insert_text(
        (72, y),
        "Genererad av DocCompare — Liljedahl Advisory AB",
        fontname="helv", fontsize=8, color=(0.55, 0.55, 0.55),
    )

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes


def merge_pdfs(main_pdf: Path, summary_bytes: bytes, output_pdf: Path):
    """Append the summary page(s) to the main PDF."""
    main = fitz.open(str(main_pdf))
    summary = fitz.open("pdf", summary_bytes)
    main.insert_pdf(summary)
    main.save(str(output_pdf))
    main.close()
    summary.close()


def produce_pdf(
    docx_path: Path,
    output_pdf: Path,
    summary: dict,
    original_name: str,
    modified_name: str,
):
    """Full pipeline: .docx with Track Changes → final PDF with summary page."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        word_pdf = tmp / "word_output.pdf"

        # Step 1: Word → PDF
        docx_to_pdf(docx_path, word_pdf)

        # Step 2: Build summary page
        summary_bytes = build_summary_pdf(summary, original_name, modified_name)

        # Step 3: Merge
        merge_pdfs(word_pdf, summary_bytes, output_pdf)
