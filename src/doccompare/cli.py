import sys
import tempfile
from pathlib import Path
from datetime import datetime
import click
from loguru import logger
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console()


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
    """Compare two .docx files and generate a PDF with tracked changes.

    ORIGINAL is the older version, MODIFIED is the newer version.
    The engine diffs the OOXML trees natively, injects Track Changes markup,
    converts to PDF via Microsoft Word, and appends a summary page.

    Examples:

        doccompare original.docx modified.docx

        doccompare v1.docx v2.docx -o diff.pdf

        doccompare draft.docx final.docx --author "Jane Doe"
    """
    if not verbose:
        logger.remove()
        logger.add(sys.stderr, level="WARNING")

    # Validate formats
    for path in [original, modified]:
        if path.suffix.lower() != ".docx":
            click.echo(f"Error: Unsupported format '{path.suffix}'. Only .docx is supported.", err=True)
            sys.exit(1)

    if output is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = Path(f"comparison_{ts}.pdf")

    from doccompare.comparison.ooxml_engine import compare as ooxml_compare
    from doccompare.rendering.pdf_pipeline import produce_pdf

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
        task = progress.add_task("Jämför dokument (OOXML-diff)…", total=None)
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_docx = Path(tmpdir) / "tracked.docx"

                # Step 1: OOXML comparison → .docx with track changes
                _, summary = ooxml_compare(original, modified, tmp_docx, author=author)

                # Step 2: Convert to PDF via Word + append summary page
                progress.update(task, description="Konverterar till PDF via Word…")
                produce_pdf(
                    tmp_docx, output, summary,
                    original_name=original.name,
                    modified_name=modified.name,
                )

        except Exception as e:
            click.echo(f"Error: {e}", err=True)
            sys.exit(1)
        progress.update(task, description="Done!", completed=1, total=1)

    # Print summary
    s = summary
    console.print(f"\n[bold]Jämförelse klar![/bold] Rapport sparad: [cyan]{output}[/cyan]")
    console.print(f"  [green]+{s.get('added_words', 0)} ord tillagda[/green]  "
                  f"[red]-{s.get('deleted_words', 0)} ord borttagna[/red]  "
                  f"[dim]{s.get('unchanged_words', 0)} ord oförändrade[/dim]")
