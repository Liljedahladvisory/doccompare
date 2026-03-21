"""Formatting change detection between matched elements."""
from doccompare.models import DocumentElement, TextFormatting


def compare_formatting(orig: DocumentElement, mod: DocumentElement) -> list:
    """Return list of formatting changes between two matched elements."""
    changes = []
    orig_runs = orig.runs
    mod_runs = mod.runs

    for i, (o_run, m_run) in enumerate(zip(orig_runs, mod_runs)):
        if o_run.text == m_run.text:
            added_fmt = m_run.formatting - o_run.formatting
            removed_fmt = o_run.formatting - m_run.formatting
            if added_fmt or removed_fmt:
                changes.append({
                    "text": o_run.text,
                    "added": added_fmt,
                    "removed": removed_fmt,
                })
    return changes
