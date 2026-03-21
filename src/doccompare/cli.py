import sys
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
    help="Path to output .docx (default: comparison_YYYYMMDD_HHMMSS.docx)",
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
    """Compare two .docx files and generate a Word document with Track Changes.

    ORIGINAL is the older version, MODIFIED is the newer version.
    The output .docx uses MODIFIED as the baseline with revision markup
    (w:ins / w:del) showing all differences.

    Examples:

        doccompare original.docx modified.docx

        doccompare v1.docx v2.docx -o diff.docx

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
        output = Path(f"comparison_{ts}.docx")

    from doccompare.comparison.ooxml_engine import compare as ooxml_compare

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
        task = progress.add_task("Comparing documents…", total=None)
        try:
            ooxml_compare(original, modified, output, author=author)
        except Exception as e:
            click.echo(f"Error: {e}", err=True)
            sys.exit(1)
        progress.update(task, description="Done!", completed=1, total=1)

    console.print(f"\n[bold]Jämförelse klar![/bold] Rapport sparad: [cyan]{output}[/cyan]")
    console.print("  Öppna i Word för att granska ändringar (Track Changes).")
