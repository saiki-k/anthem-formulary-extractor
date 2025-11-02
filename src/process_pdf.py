#!/usr/bin/env python3
"""
PDF processing module for extracting formulary data and creating output files.
"""

import json
from pathlib import Path

from extract_pdf_tables import extract_structured_data
from create_excel_file import create_excel_from_json
from config import DEBUG_MODE


def print_summary(
    json_output_path: Path,
    warnings_path: Path,
    toc_path: Path,
    data: dict,
    total_subcategories: int,
    total_rows: int,
    excel_output_path: Path = None,
):
    """
    Print a summary of the extracted data.
    """
    print("=" * 80)
    print(f"Extraction complete!")
    print(f"  - Categories saved to: {json_output_path}")
    print(f"  - Warnings saved to: {warnings_path}")
    print(f"  - Table of Contents saved to: {toc_path}")
    if excel_output_path:
        print(f"  - Excel file saved to: {excel_output_path}")

    print(f"\nSummary:")
    print(f"  - TOC entries found: {len(data['table_of_contents'])}")
    print(f"  - Categories found: {len(data['categories'])}")
    print(f"  - Total subcategories: {total_subcategories}")
    print(f"  - Total rows extracted: {total_rows}")
    print(f"  - Total rows processed: {data['total_rows_processed']}")
    print(f"  - Warnings (skipped rows): {len(data['warnings'])}")

    # Validation checks
    print(f"\nValidation:")

    # Check 1: TOC entries should match categories found
    toc_count = len(data["table_of_contents"])
    categories_count = len(data["categories"])
    toc_match = toc_count == categories_count
    print(
        f"  - TOC entries ({toc_count}) == Categories found ({categories_count}) {'✓' if toc_match else '✗'}"
    )

    # Check 2: Total rows processed should equal sum of all components
    expected_total = (
        total_rows + categories_count + total_subcategories + len(data["warnings"])
    )
    actual_total = data["total_rows_processed"]
    rows_match = expected_total == actual_total
    print(
        f"  - Total rows processed ({actual_total}) == Sum of components ({expected_total}) {'✓' if rows_match else '✗'}"
    )
    print(
        f"    Expected: {total_rows} drugs + {categories_count} categories + {total_subcategories} subcategories + {len(data['warnings'])} warnings = {expected_total}"
    )

    # Check 3: Warnings should be zero for clean extraction
    no_warnings = len(data["warnings"]) == 0
    print(f"  - No warnings (clean extraction) {'✓' if no_warnings else '✗'}")

    # Check 4: Overall status
    all_good = toc_match and rows_match and no_warnings
    print(
        f"  - Overall extraction quality: {'✓ EXCELLENT' if all_good else '(!) NEEDS REVIEW'}"
    )

    print("=" * 80 + "\n")


def process_pdf(pdf_path: Path, output_dir: str, json_only: bool):
    """
    Process a single PDF: extract data, save JSON, and optionally create Excel.
    """
    if DEBUG_MODE:
        print("=" * 80)
        print(f"STEP 1: PDF EXTRACTION for {pdf_path.name}")
        print("=" * 80)

    # Create output directory: output/PDFFILENAME/
    pdf_filename = pdf_path.stem
    pdf_output_dir = Path(output_dir) / pdf_filename
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    if DEBUG_MODE:
        print(f"Input PDF: {pdf_path}")
        print(f"Output directory: {pdf_output_dir}")
        print("=" * 80)

    # Extract data from PDF
    data = extract_structured_data(str(pdf_path))

    # Save JSON files
    json_output_path = pdf_output_dir / "extracted_data.json"
    warnings_path = pdf_output_dir / "extraction_warnings.json"
    toc_path = pdf_output_dir / "table_of_contents.json"

    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(data["categories"], f, indent=2, ensure_ascii=False)

    with open(warnings_path, "w", encoding="utf-8") as f:
        json.dump(data["warnings"], f, indent=2, ensure_ascii=False)

    with open(toc_path, "w", encoding="utf-8") as f:
        json.dump(data["table_of_contents"], f, indent=2, ensure_ascii=False)

    # Count total subcategories and rows
    total_subcategories = sum(len(cat["subCategories"]) for cat in data["categories"])
    total_rows = sum(
        len(subcat["rows"])
        for cat in data["categories"]
        for subcat in cat["subCategories"]
    )

    # Stop here if --json-only
    if json_only:
        print_summary(
            json_output_path,
            warnings_path,
            toc_path,
            data,
            total_subcategories,
            total_rows,
        )
        if DEBUG_MODE:
            print("\n" + "=" * 80)
            print("JSON-only mode: Skipping Excel creation")
            print("=" * 80)
        return

    # Excel Creation
    if DEBUG_MODE:
        print("\n" + "=" * 80)
        print("STEP 2: EXCEL CREATION")
        print("=" * 80)

    excel_output_path = pdf_output_dir / f"{pdf_filename}.xlsx"

    try:
        create_excel_from_json(json_output_path, excel_output_path)

        print_summary(
            json_output_path,
            warnings_path,
            toc_path,
            data,
            total_subcategories,
            total_rows,
            excel_output_path,
        )

    except ImportError as e:
        print(f"\nWarning: Could not create Excel file")
        print(f"  Error: {e}")
        print(f"  Please install openpyxl: pip install openpyxl")
        print(f"\n  JSON files have been created successfully in: {pdf_output_dir}")
    except Exception as e:
        print(f"\nError creating Excel file: {e}")
        print(f"  JSON files have been created successfully in: {pdf_output_dir}")
