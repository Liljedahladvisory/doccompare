import sys
from pathlib import Path
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
    "--move-threshold", type=float, default=85.0,
    help="Similarity threshold (0-100) for classifying text as moved (default: 85)",
)
@click.option(
    "--no-moves", is_flag=True, default=False,
    help="Disable move detection (faster)",
)
@click.option(
    "--verbose", "-v", is_flag=True, default=False,
    help="Show detailed logging",
)
@click.version_option(package_name="doccompare")
def compare(original: Path, modified: Path, output: Path, move_threshold: float, no_moves: bool, verbose: bool):
    """Compare two documents and generate a PDF with color-coded differences.

    ORIGINAL and MODIFIED can be .docx or .pdf files in any combination.

    Examples:

        doccompare original.docx modified.docx

        doccompare v1.pdf v2.pdf -o diff.pdf

        doccompare draft.docx final.pdf --move-threshold 90
    """
    if not verbose:
        logger.remove()
        logger.add(sys.stderr, level="WARNING")

    from doccompare.parsers import get_parser
    from doccompare.comparison.differ import Differ
    from doccompare.comparison.move_detector import MoveDetector
    from doccompare.rendering.html_builder import HtmlBuilder
    from doccompare.rendering.pdf_renderer import render_pdf
    from doccompare.utils import default_output_path

    if output is None:
        output = default_output_path(original, modified)

    # Validate formats
    for path in [original, modified]:
        if path.suffix.lower() not in (".docx", ".pdf"):
            click.echo(f"Error: Unsupported format '{path.suffix}'. Use .docx or .pdf.", err=True)
            sys.exit(1)

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
        task = progress.add_task("Parsing original document...", total=None)
        try:
            orig_parser = get_parser(original)
            orig_doc = orig_parser.parse(original)
        except ValueError as e:
            click.echo(f"Error parsing {original.name}: {e}", err=True)
            sys.exit(1)

        progress.update(task, description="Parsing modified document...")
        try:
            mod_parser = get_parser(modified)
            mod_doc = mod_parser.parse(modified)
        except ValueError as e:
            click.echo(f"Error parsing {modified.name}: {e}", err=True)
            sys.exit(1)

        progress.update(task, description="Comparing documents...")
        differ = Differ()
        result = differ.compare(orig_doc, mod_doc)

        if not no_moves:
            progress.update(task, description="Detecting moved text blocks...")
            detector = MoveDetector(threshold=move_threshold)
            result = detector.detect(result)

        progress.update(task, description="Building HTML output...")
        builder = HtmlBuilder()
        html_content = builder.build(result, original, modified)

        progress.update(task, description="Rendering PDF...")
        css_path = Path(__file__).parent / "rendering" / "styles.css"
        render_pdf(html_content, css_path, output)

        progress.update(task, description="Done!", completed=1, total=1)

    # Print summary
    s = result.summary
    console.print(f"\n[bold]Jämförelse klar![/bold] Rapport sparad: [cyan]{output}[/cyan]")
    console.print(f"  [green]+{s.get('added_words', 0)} ord tillagda[/green]  "
                  f"[red]-{s.get('deleted_words', 0)} ord borttagna[/red]  "
                  f"[blue]~{s.get('moved_words', 0)} ord flyttade[/blue]")
