from pathlib import Path
from weasyprint import HTML, CSS


def render_pdf(html_content: str, css_path: Path, output_path: Path) -> None:
    """Render HTML comparison to PDF using WeasyPrint."""
    h = HTML(string=html_content, base_url=str(css_path.parent))
    css = CSS(filename=str(css_path))
    h.write_pdf(str(output_path), stylesheets=[css])
