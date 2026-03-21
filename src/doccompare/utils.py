from pathlib import Path
from datetime import datetime


def default_output_path(original: Path, modified: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path(f"comparison_{timestamp}.pdf")
