# DocCompare

A Python CLI tool that compares two documents (`.docx` and/or `.pdf`) and generates a PDF report showing all differences with color-coded formatting.

## Color coding

- **Blue + underlined** — added text
- **Red + strikethrough** — deleted text
- **Green** — moved text (shown at both origin and destination)

## Features

- Supports `.docx` and `.pdf` input files in any combination
- Word-level diff using diff-match-patch
- Move detection: identifies blocks of text that have been relocated
- Preserves heading levels, list items, and table structure
- Generates a self-contained PDF report via WeasyPrint
- Summary statistics (added/deleted/moved/unchanged word counts)
- Rich progress display in the terminal

## Requirements

- Python 3.10+
- The following system libraries are required by WeasyPrint:
  - `pango`, `cairo`, `gdk-pixbuf` (Linux/macOS via Homebrew or apt)

## Installation

```bash
pip install doccompare
```

Or install from source:

```bash
git clone https://github.com/yourorg/doccompare.git
cd doccompare
pip install -e .
```

### macOS (Homebrew)

Install WeasyPrint system dependencies:

```bash
brew install pango cairo gdk-pixbuf libffi
```

### Linux (Debian/Ubuntu)

```bash
sudo apt install libpango-1.0-0 libpangoft2-1.0-0 libcairo2 libgdk-pixbuf2.0-0
```

## Usage

### Basic usage

```bash
# Compare two .docx files
doccompare original.docx modified.docx

# Compare two PDFs
doccompare v1.pdf v2.pdf

# Cross-format comparison
doccompare draft.docx final.pdf

# Specify output path
doccompare original.docx modified.docx -o diff_report.pdf
```

### Options

```
Usage: doccompare [OPTIONS] ORIGINAL MODIFIED

  Compare two documents and generate a PDF with color-coded differences.

  ORIGINAL and MODIFIED can be .docx or .pdf files in any combination.

Options:
  -o, --output PATH           Path to output PDF (default: comparison_YYYYMMDD_HHMMSS.pdf)
  --move-threshold FLOAT      Similarity threshold (0-100) for classifying text
                              as moved (default: 85)
  --no-moves                  Disable move detection (faster)
  -v, --verbose               Show detailed logging
  --version                   Show the version and exit.
  --help                      Show this message and exit.
```

### Examples

```bash
# Standard comparison with default settings
doccompare contract_v1.docx contract_v2.docx

# Strict move detection (require near-identical text to count as moved)
doccompare old.docx new.docx --move-threshold 95

# Loose move detection
doccompare old.pdf new.pdf --move-threshold 70

# Skip move detection for large documents (faster)
doccompare big_report_v1.pdf big_report_v2.pdf --no-moves -o report_diff.pdf

# Verbose output for debugging
doccompare a.docx b.docx --verbose
```

## Output

The tool generates a PDF file containing:

1. **Header** — file names and comparison timestamp
2. **Summary box** — word counts for added, deleted, moved, and unchanged text
3. **Legend** — color key
4. **Full document diff** — the complete text of both documents merged, with differences highlighted

## Limitations

- Scanned PDFs (image-only, no text layer) are not supported. Run OCR first (e.g. with `ocrmypdf`).
- Encrypted PDFs must be decrypted before comparison.
- Complex table diffs show row-level changes only, not cell-level diffs.

## Architecture

```
src/doccompare/
├── cli.py              Entry point (Click command)
├── models.py           Data classes (ParsedDocument, DiffElement, etc.)
├── parsers/
│   ├── base.py         Abstract parser interface
│   ├── docx_parser.py  python-docx based parser
│   └── pdf_parser.py   pdfplumber + PyMuPDF based parser
├── comparison/
│   ├── differ.py       LCS element matching + word-level diff
│   ├── move_detector.py  Fuzzy move detection via rapidfuzz
│   └── formatter.py    Formatting change detection
└── rendering/
    ├── html_builder.py HTML generation from diff result
    ├── pdf_renderer.py WeasyPrint PDF rendering
    └── styles.css      Report stylesheet
```

## Development

```bash
# Install with dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Run a specific test file
pytest tests/test_differ.py -v
```

## License

MIT License. See LICENSE for details.
