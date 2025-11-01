#!/usr/bin/env python3
"""
Main entry point for PDF extraction and Excel conversion pipeline.
"""
import sys
from pathlib import Path
import argparse

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Import the extraction and Excel creation modules
from extract_pdf_tables import extract_structured_data
from create_excel_file import create_excel_from_json


def main():
    parser = argparse.ArgumentParser(
        description="Extract pharmaceutical formulary data from PDF and create Excel file."
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        help="Path to the PDF file to extract",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="output",
        help="Output directory for extracted data (default: output)",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="Only extract JSON, skip Excel creation",
    )
    parser.add_argument(
        "--excel-only",
        action="store_true",
        help="Only create Excel from existing JSON (requires --json-path)",
    )
    parser.add_argument(
        "--json-path",
        help="Path to existing JSON file (for --excel-only mode)",
    )

    args = parser.parse_args()

    # Validate input
    if args.excel_only:
        if not args.json_path:
            print("Error: --excel-only requires --json-path")
            return 1
        json_path = Path(args.json_path)
        if not json_path.exists():
            print(f"Error: JSON file not found: {json_path}")
            return 1
    else:
        if not args.pdf_path:
            print("Error: pdf_path is required (unless using --excel-only)")
            return 1
        if not Path(args.pdf_path).exists():
            print(f"Error: PDF file not found: {args.pdf_path}")
            return 1

    # Excel-only mode
    if args.excel_only:
        print("=" * 80)
        print("EXCEL CREATION MODE")
        print("=" * 80)

        json_path = Path(args.json_path)
        excel_path = json_path.parent / f"{json_path.parent.name}.xlsx"

        print(f"Input JSON: {json_path}")
        print(f"Output Excel: {excel_path}")

        create_excel_from_json(json_path, excel_path)
        return 0

    # PDF Extraction
    print("=" * 80)
    print("STEP 1: PDF EXTRACTION")
    print("=" * 80)

    # Create output directory: output/PDFFILENAME/
    pdf_filename = Path(args.pdf_path).stem
    output_dir = Path(args.output_dir) / pdf_filename
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Input PDF: {args.pdf_path}")
    print(f"Output directory: {output_dir}")
    print("=" * 80)

    # Extract data from PDF
    data = extract_structured_data(args.pdf_path)

    # Save JSON files
    json_output_path = output_dir / "extracted_data.json"
    warnings_path = output_dir / "extraction_warnings.json"
    toc_path = output_dir / "table_of_contents.json"

    import json

    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(data["categories"], f, indent=2, ensure_ascii=False)

    with open(warnings_path, "w", encoding="utf-8") as f:
        json.dump(data["warnings"], f, indent=2, ensure_ascii=False)

    with open(toc_path, "w", encoding="utf-8") as f:
        json.dump(data["table_of_contents"], f, indent=2, ensure_ascii=False)

    print("\n" + "=" * 80)
    print(f"Extraction complete!")
    print(f"  - Categories saved to: {json_output_path}")
    print(f"  - Warnings saved to: {warnings_path}")
    print(f"  - Table of Contents saved to: {toc_path}")

    print(f"\nSummary:")
    print(f"  - TOC entries found: {len(data['table_of_contents'])}")
    print(f"  - Categories found: {len(data['categories'])}")
    print(f"  - Warnings (skipped rows): {len(data['warnings'])}")

    # Count total subcategories and rows
    total_subcategories = sum(len(cat["subCategories"]) for cat in data["categories"])
    total_rows = sum(
        len(subcat["rows"])
        for cat in data["categories"]
        for subcat in cat["subCategories"]
    )

    print(f"  - Total subcategories: {total_subcategories}")
    print(f"  - Total rows extracted: {total_rows}")

    # Stop here if --json-only
    if args.json_only:
        print("\n" + "=" * 80)
        print("JSON-only mode: Skipping Excel creation")
        print("=" * 80)
        return 0

    # Excel Creation
    print("\n" + "=" * 80)
    print("STEP 2: EXCEL CREATION")
    print("=" * 80)

    excel_output_path = output_dir / f"{pdf_filename}.xlsx"

    try:
        create_excel_from_json(json_output_path, excel_output_path)

        print("\n" + "=" * 80)
        print("PIPELINE COMPLETE!")
        print("=" * 80)
        print(f"\nOutput files:")
        print(f"  - JSON: {json_output_path}")
        print(f"  - Excel: {excel_output_path}")
        print(f"  - Warnings: {warnings_path}")
        print(f"  - TOC: {toc_path}")

    except ImportError as e:
        print(f"\nWarning: Could not create Excel file")
        print(f"  Error: {e}")
        print(f"  Please install openpyxl: pip install openpyxl")
        print(f"\n  JSON files have been created successfully in: {output_dir}")
        return 0
    except Exception as e:
        print(f"\nError creating Excel file: {e}")
        print(f"  JSON files have been created successfully in: {output_dir}")
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
