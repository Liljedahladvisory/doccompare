import sys
from pathlib import Path
from datetime import datetime
import click
from loguru import logger
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console()

SUPPORTED_FORMATS = {".docx"}


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
    """Compare two .docx documents and generate a PDF diff report.

    ORIGINAL is the older version, MODIFIED is the newer version.

    Examples:

        doccompare original.docx modified.docx

        doccompare draft.docx final.docx -o diff.pdf --author "Jane Doe"
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
            summary = _compare_docx(original, modified, output, author, progress, task)
        except Exception as e:
            click.echo(f"Error: {e}", err=True)
            sys.exit(1)
        progress.update(task, description="Done!", completed=1, total=1)

    s = summary
    console.print(f"\n[bold]Comparison complete![/bold] Report saved: [cyan]{output}[/cyan]")
    moved = s.get('moved_words', 0)
    moved_str = f"  [yellow]{moved} words moved[/yellow]" if moved else ""
    console.print(f"  [green]+{s.get('added_words', 0)} words added[/green]  "
                  f"[red]-{s.get('deleted_words', 0)} words deleted[/red]"
                  f"{moved_str}  "
                  f"[dim]{s.get('unchanged_words', 0)} words unchanged[/dim]")


def _compare_docx(original, modified, output, author, progress, task):
    """DOCX comparison — Word-native adapter with ooxml_engine fallback."""
    from doccompare.comparison.adapters import get_adapter

    adapter = get_adapter()
    if adapter:
        progress.update(task, description="Comparing via Word\u2026")
        try:
            return adapter.compare_and_export(
                original, modified, output,
                original_name=original.name,
                modified_name=modified.name,
            )
        except RuntimeError as e:
            logger.warning(f"Word adapter failed, falling back to OOXML: {e}")

    # Fallback: XML-level comparison
    from doccompare.comparison.ooxml_engine import compare as ooxml_compare
    from doccompare.rendering.pdf_pipeline import produce_pdf

    progress.update(task, description="Comparing documents\u2026")
    doc_tree, summary = ooxml_compare(original, modified, None, author=author)

    progress.update(task, description="Rendering PDF\u2026")
    produce_pdf(
        doc_tree, output, summary,
        original_name=original.name,
        modified_name=modified.name,
        docx_path=modified,
    )
    return summary


