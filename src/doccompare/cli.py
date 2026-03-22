import sys
from pathlib import Path
from datetime import datetime
import click
from loguru import logger
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console()

SUPPORTED_FORMATS = {".docx", ".pdf"}


@click.command()
@click.argument("original", type=click.Path(exists=True, path_type=Path))
@click.argument("modified", type=click.Path(exists=True, path_type=Path))
@click.option(
    "-o", "--output", type=click.Path(path_type=Path),
    default=None,
    help="Path to output PDF (default: comparison_YYYYMMDD_HHMMSS.pdf)",
)
@click.option(
    "--author", type=str, default="DocCompare",
    help="Author name for track-changes metadata (default: DocCompare)",
)
@click.option(
    "--verbose", "-v", is_flag=True, default=False,
    help="Show detailed logging",
)
@click.version_option(package_name="doccompare")
def compare(original: Path, modified: Path, output: Path, author: str, verbose: bool):
    """Compare two documents and generate a PDF diff report.

    ORIGINAL is the older version, MODIFIED is the newer version.
    Supports .docx and .pdf files. Both files must be the same format.

    Examples:

        doccompare original.docx modified.docx

        doccompare v1.pdf v2.pdf -o diff.pdf

        doccompare draft.docx final.docx --author "Jane Doe"
    """
    if not verbose:
        logger.remove()
        logger.add(sys.stderr, level="WARNING")

    orig_ext = original.suffix.lower()
    mod_ext = modified.suffix.lower()

    for path in [original, modified]:
        if path.suffix.lower() not in SUPPORTED_FORMATS:
            click.echo(
                f"Error: Unsupported format '{path.suffix}'. "
                f"Supported: {', '.join(sorted(SUPPORTED_FORMATS))}",
                err=True,
            )
            sys.exit(1)

    if orig_ext != mod_ext:
        click.echo(
            f"Error: Both files must be the same format. "
            f"Got '{orig_ext}' and '{mod_ext}'.",
            err=True,
        )
        sys.exit(1)

    if output is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = Path(f"comparison_{ts}.pdf")

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
        task = progress.add_task("Comparing documents\u2026", total=None)
        try:
            if orig_ext == ".docx":
                summary = _compare_docx(original, modified, output, author, progress, task)
            else:
                summary = _compare_pdf(original, modified, output, progress, task)
        except Exception as e:
            click.echo(f"Error: {e}", err=True)
            sys.exit(1)
        progress.update(task, description="Done!", completed=1, total=1)

    s = summary
    console.print(f"\n[bold]Comparison complete![/bold] Report saved: [cyan]{output}[/cyan]")
    console.print(f"  [green]+{s.get('added_words', 0)} words added[/green]  "
                  f"[red]-{s.get('deleted_words', 0)} words deleted[/red]  "
                  f"[dim]{s.get('unchanged_words', 0)} words unchanged[/dim]")


def _compare_docx(original, modified, output, author, progress, task):
    """DOCX comparison via OOXML engine + Word headless PDF export."""
    from doccompare.comparison.ooxml_engine import compare as ooxml_compare
    from doccompare.rendering.pdf_pipeline import produce_pdf

    doc_tree, summary = ooxml_compare(original, modified, None, author=author)

    progress.update(task, description="Rendering PDF\u2026")
    produce_pdf(
        doc_tree, output, summary,
        original_name=original.name,
        modified_name=modified.name,
        docx_path=modified,
    )
    return summary


def _compare_pdf(original, modified, output, progress, task):
    """PDF comparison via text extraction + diff_match_patch."""
    from doccompare.comparison.pdf_engine import compare_pdfs

    progress.update(task, description="Comparing PDFs\u2026")
    summary = compare_pdfs(
        original, modified, output,
        original_name=original.name,
        modified_name=modified.name,
    )
    return summary
