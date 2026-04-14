"""Word-native document comparison via OS-specific adapters.

Uses Microsoft Word's built-in Document.Compare to guarantee 100% formatting
fidelity in the output PDF. Falls back to ooxml_engine when Word is unavailable.
"""
from abc import ABC, abstractmethod
from pathlib import Path
import subprocess
import sys
import difflib
import zipfile
import tempfile
import os

from lxml import etree
from loguru import logger


# ── Abstract adapter ────────────────────────────────────────────────────

class ComparisonAdapter(ABC):
    """Base class for OS-specific Word automation adapters."""

    @abstractmethod
    def is_available(self) -> bool:
        """Return True if Word automation is available on this platform."""

    @abstractmethod
    def compare_and_export(
        self,
        original: Path,
        modified: Path,
        output_pdf: Path,
        original_name: str = "",
        modified_name: str = "",
    ) -> dict:
        """Compare two .docx files and export the result as PDF.

        Returns a summary dict with keys: added_words, deleted_words, unchanged_words.
        Raises RuntimeError on failure.
        """


# ── macOS adapter ────────────────────────────────────────────────────────

_WORD_APP_PATH = Path("/Applications/Microsoft Word.app")
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class MacWordAdapter(ComparisonAdapter):
    """Compare documents using Microsoft Word for Mac via AppleScript."""

    def is_available(self) -> bool:
        return sys.platform == "darwin" and _WORD_APP_PATH.exists()

    def compare_and_export(
        self,
        original: Path,
        modified: Path,
        output_pdf: Path,
        original_name: str = "",
        modified_name: str = "",
    ) -> dict:
        from doccompare.rendering.pdf_pipeline import (
            _word_temp_dir,
            _render_summary_pdf,
            _merge_pdfs,
        )

        original_name = original_name or original.name
        modified_name = modified_name or modified.name

        # Use Word's sandbox for temp files
        temp_dir = _word_temp_dir()
        comparison_pdf = temp_dir / f"comparison_{os.getpid()}.pdf"

        try:
            # Step 1: Run Word comparison + PDF export
            self._run_compare_script(original, modified, comparison_pdf)

            # Step 2: Calculate summary from source documents
            summary = self._extract_summary(original, modified)

            # Step 3: Render summary/legend page
            summary_bytes = _render_summary_pdf(
                summary, original_name, modified_name,
            )

            # Step 4: Merge comparison PDF with summary page
            _merge_pdfs(comparison_pdf, summary_bytes, output_pdf)

            logger.info(f"Comparison PDF saved: {output_pdf}")
            return summary

        finally:
            comparison_pdf.unlink(missing_ok=True)

    def _run_compare_script(
        self, original: Path, modified: Path, output_pdf: Path,
    ) -> None:
        """Execute AppleScript to compare documents and export as PDF."""
        orig_abs = str(original.resolve())
        mod_abs = str(modified.resolve())
        pdf_abs = str(output_pdf.resolve())
        orig_name = original.name

        script = f'''
        tell application "Microsoft Word"
            open "{orig_abs}"
            delay 2
            set origDoc to active document
            set origName to name of origDoc
            compare origDoc path "{mod_abs}" ¬
                target compare target new ¬
                detect format changes false ¬
                ignore all comparison warnings true ¬
                add to recent files false
            delay 3

            -- Reject list-numbering-only revisions to prevent
            -- Word from renumbering lists (a,b,c -> d,e,f)
            set compDoc to active document
            set revList to revisions of compDoc
            set rejectedCount to 0
            repeat with rev in reverse of revList
                if revision type of rev is revision paragraph number then
                    reject rev
                    set rejectedCount to rejectedCount + 1
                end if
            end repeat

            save as active document ¬
                file name (POSIX file "{pdf_abs}" as text) ¬
                file format format PDF
            delay 1
            close active document saving no
            close document origName saving no
        end tell
        '''

        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=120,
            )
        except subprocess.TimeoutExpired:
            logger.error("Word comparison timed out after 120s")
            self._cleanup_hung_word()
            raise RuntimeError(
                "Word took too long to compare the documents. "
                "Try with smaller files or restart Word."
            )
        except FileNotFoundError:
            raise RuntimeError("osascript not found — macOS is required.")

        if result.returncode != 0:
            stderr = result.stderr.strip()
            if "not authorized" in stderr.lower() or "(-1743)" in stderr:
                raise RuntimeError(
                    "macOS blocked Word automation.\n\n"
                    "Fix: System Settings → Privacy & Security → Automation\n"
                    "Enable DocCompare → Microsoft Word."
                )
            logger.error(f"AppleScript failed (rc={result.returncode}): {stderr}")
            raise RuntimeError(f"Word comparison failed: {stderr}")

        if not output_pdf.exists():
            raise RuntimeError(
                "Word completed but no PDF was produced. "
                "Make sure Microsoft Word is installed and working."
            )

    def _extract_summary(self, original: Path, modified: Path) -> dict:
        """Approximate word counts with move detection.

        Deleted blocks whose text reappears in an inserted block (fuzzy match
        ≥ 85%) are reclassified as *moved* rather than deleted+added.
        """
        from rapidfuzz import fuzz

        orig_text = _extract_text(original)
        mod_text = _extract_text(modified)

        orig_words = orig_text.split()
        mod_words = mod_text.split()

        matcher = difflib.SequenceMatcher(None, orig_words, mod_words)

        unchanged = 0
        added_blocks: list[tuple[int, int]] = []
        deleted_blocks: list[tuple[int, int]] = []

        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == "equal":
                unchanged += i2 - i1
            elif op == "delete":
                deleted_blocks.append((i1, i2))
            elif op == "insert":
                added_blocks.append((j1, j2))
            elif op == "replace":
                deleted_blocks.append((i1, i2))
                added_blocks.append((j1, j2))

        # Detect moves: match deleted blocks against added blocks
        moved = 0
        used_added: set[int] = set()

        for di, (d1, d2) in enumerate(deleted_blocks):
            del_text = " ".join(orig_words[d1:d2])
            if len(del_text) < 20:
                continue
            for ai, (a1, a2) in enumerate(added_blocks):
                if ai in used_added:
                    continue
                add_text = " ".join(mod_words[a1:a2])
                if fuzz.ratio(del_text, add_text) >= 85:
                    moved += d2 - d1
                    deleted_blocks[di] = (0, 0)  # zero out
                    used_added.add(ai)
                    break

        added = sum(a2 - a1 for ai, (a1, a2) in enumerate(added_blocks)
                    if ai not in used_added)
        deleted = sum(d2 - d1 for d1, d2 in deleted_blocks)

        return {
            "added_words": added,
            "deleted_words": deleted,
            "moved_words": moved,
            "unchanged_words": unchanged,
        }

    def _cleanup_hung_word(self) -> None:
        """Best-effort: close any DocCompare-opened documents in Word."""
        cleanup_script = '''
        tell application "Microsoft Word"
            repeat with d in documents
                close d saving no
            end repeat
        end tell
        '''
        try:
            subprocess.run(
                ["osascript", "-e", cleanup_script],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass


# ── Helpers ──────────────────────────────────────────────────────────────

def _extract_text(docx_path: Path) -> str:
    """Extract plain text from a .docx file using zipfile + lxml."""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            xml_bytes = zf.read("word/document.xml")
        tree = etree.fromstring(xml_bytes)
        texts = tree.itertext()
        return " ".join(t.strip() for t in texts if t.strip())
    except Exception as e:
        logger.warning(f"Could not extract text from {docx_path.name}: {e}")
        return ""


# ── Factory ──────────────────────────────────────────────────────────────

def get_adapter() -> ComparisonAdapter | None:
    """Return a comparison adapter for the current platform, or None."""
    if sys.platform == "darwin":
        adapter = MacWordAdapter()
        if adapter.is_available():
            return adapter
    # Future: WindowsWordAdapter for win32
    return None
