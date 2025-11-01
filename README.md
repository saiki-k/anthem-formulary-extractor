# Anthem Formulary Extractor

A Python tool for extracting pharmaceutical formulary data provided by [Anthem](https://anthem.com) into structured formats. This tool parses formulary data tables from (a) corresponding PDF(s) containing drug information, organizing them into a hierarchical JSON structure with categories and subcategories, and creates the corresponding Excel file(s).

## Project Structure

```
/
├── src/
│   ├── extract_pdf_tables.py    # PDF extraction script
│   └── create_excel_file.py     # Excel file creation script
├── example/
│   └── Essential_5_Tier_ABCBS.pdf  # Example pharmaceutical formulary PDF
├── output/                       # Default output directory (gitignored)
│   └── PDFFILENAME/              # Organized by PDF filename
│       ├── extracted_data.json
│       ├── extraction_warnings.json
│       ├── table_of_contents.json
│       └── PDFFILENAME.xlsx      # Excel output with multiple sheets
├── .venv/                        # Python virtual environment (gitignored)
├── main.py                       # Main pipeline script
├── requirements.txt              # Python dependencies
├── .gitignore
└── README.md
```

## Installation

### Prerequisites

-   Python 3.10
-   pip (Python package manager)

### Setup

1. **Clone or download this repository**

2. **Create and activate a virtual environment**:

    ```bash
    # Create virtual environment
    python -m venv .venv

    # Activate on Windows (Git Bash)
    source .venv/Scripts/activate

    # Activate on macOS/Linux
    source .venv/bin/activate
    ```

3. **Install dependencies**:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

### Full Pipeline (PDF → JSON + Excel)

Extract data from PDF and automatically create Excel file with multiple sheets:

```bash
# Make sure virtual environment is activated first
# source .venv/Scripts/activate  # Windows Git Bash
# source .venv/bin/activate    # macOS/Linux

python main.py example/Essential_5_Tier_ABCBS.pdf -o example/output/
```

This will create output in `example/output/PDFFILENAME/`:

-   `extracted_data.json` - Main structured data
-   `extraction_warnings.json` - Any skipped/problematic rows
-   `table_of_contents.json` - Extracted table of contents
-   `PDFFILENAME.xlsx` - Excel file with multiple sheets (Categories index + one sheet per category)

### Command-line Options

```
usage: main.py [-h] [-o OUTPUT_DIR] [--json-only] [--excel-only] [--json-path JSON_PATH] pdf_path

Extract pharmaceutical formulary data from PDF and create Excel file.

positional arguments:
  pdf_path              Path to the PDF file to extract

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Output directory for extracted data (default: output)
  --json-only           Only extract JSON, skip Excel creation
  --excel-only          Only create Excel from existing JSON (requires --json-path)
  --json-path JSON_PATH Path to existing JSON file (for --excel-only mode)
```

## Output Format

### extracted_data.json

The main output is an array of category objects, each containing subcategories with drug rows:

```json
[
	{
		"categoryName": "ADHD/ANTI-NARCOLEPSY/ANTI-OBESITY/ANOREXIANTS",
		"subCategories": [
			{
				"subCategoryName": "ADHD AGENT - SELECTIVE ALPHA ADRENERGIC AGONISTS",
				"rows": [
					{
						"drug_name": "clonidine ER (KAPVAY)",
						"tier": "2",
						"notes": "QL",
						"page": 7
					}
				]
			}
		]
	}
]
```

### extraction_warnings.json

Contains any rows that couldn't be properly categorized:

```json
[
	{
		"issue": "Drug row without category/subcategory",
		"data": {
			"drug_name": "orphan drug",
			"tier": "1",
			"notes": "",
			"page": 42
		}
	}
]
```

### table_of_contents.json

Maps category names to their starting page numbers:

```json
{
	"ADHD/ANTI-NARCOLEPSY/ANTI-OBESITY/ANOREXIANTS": 7,
	"AMINOGLYCOSIDES": 8,
	"ANALGESICS - ANTI-INFLAMMATORY": 10
}
```

### PDFFILENAME.xlsx (Excel Output)

Multi-sheet Excel workbook with:

-   **Categories sheet**: Index page with clickable hyperlinks to each category, showing subcategory counts and total drugs
-   **Category sheets**: One sheet per category with columns:
    -   Category
    -   Subcategory
    -   Drug
    -   Tier
    -   Notes

## How It Works

### PDF Structure Detection

The tool works with Anthem's pharmaceutical formulary PDFs with this structure:

-   **Table of Contents**: Pages with `*CATEGORY*....page_number` format
-   **Categories**: Section headers like `*ADHD/ANTI-NARCOLEPSY*`
-   **Subcategories**: Subsection headers like `*ADHD AGENT***`
-   **Drug Rows**: 3-column tables with drug_name, tier, and notes

### Extraction Process

1. **TOC Extraction**: Scans pages for table of contents entries
2. **Category Detection**: Identifies category markers (`*TEXT*`)
3. **Subcategory Detection**: Identifies subcategory markers (`*TEXT***`)
4. **Table Extraction**: Extracts tables from each page, sorting left-to-right
5. **Data Cleaning**:
    - Replaces newlines with spaces
    - Fixes split hyphenated words (e.g., "hyper- tension" → "hyper-tension")
    - Normalizes whitespace
6. **Fuzzy Matching**: Matches extracted categories to TOC entries for consistency
7. **Hierarchical Organization**: Builds nested JSON structure
