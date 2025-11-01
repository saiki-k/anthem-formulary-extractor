import pdfplumber
import json
import re
import argparse
import os
from pathlib import Path


def clean_text(text):
    """Replace all newlines with spaces and clean up whitespace."""
    if not text:
        return ""

    # Join all whitespace-separated parts with spaces
    cleaned = " ".join(text.split())

    # Fix hyphenated words that got split: remove space around hyphens if only one side has space
    # Pattern: "word- word" or "word -word" should become "word-word"
    cleaned = re.sub(r"(\S)-\s+(\S)", r"\1-\2", cleaned)  # "word- word" -> "word-word"
    cleaned = re.sub(r"(\S)\s+-(\S)", r"\1-\2", cleaned)  # "word -word" -> "word-word"

    return cleaned


def extract_table_of_contents(pdf_path):
    """Extract categories and their page numbers from the table of contents."""
    categories = {}
    found_toc = False

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")
            found_any_match = False

            for line in lines:
                # Check if this is the TOC start marker
                if "Table of Contents" in line:
                    found_toc = True
                    print(f"Found Table of Contents on page {page_num}")

                # Match lines with format: *CATEGORY NAME*....page_number
                match = re.search(r"\*([A-Z\/\-\s]+)\*\.+(\d+)", line)
                if match:
                    found_toc = True  # If we find TOC pattern, we're in TOC
                    category_name = clean_text(match.group(1))
                    page_no = int(match.group(2))
                    categories[category_name] = page_no
                    print(f"  Found category: {category_name} -> Page {page_no}")
                    found_any_match = True

            # If we've found TOC before but no matches on this page, stop
            if found_toc and not found_any_match and categories:
                print(
                    f"  No more TOC entries found on page {page_num}, stopping TOC extraction"
                )
                break

    return categories


def is_category(text):
    """Check if text is a category (format: *CATEGORY*)"""
    cleaned = clean_text(text)
    return bool(re.match(r"^\*([A-Z\/\-\s]+)\*$", cleaned))


def is_subcategory(text):
    """
    Check if text is a subcategory (format: *SUBCATEGORY***)
    The ending can be ***, ** *, or * ** due to newline artifacts
    """
    cleaned = clean_text(text)
    # Match: starts with *, has content, ends with 2+ asterisks (with optional spaces between)
    return bool(re.match(r"^\*(.+?)\*[\s*]+$", cleaned))


def extract_category_name(text, toc_categories=None):
    """Extract category name from *CATEGORY* format and find closest TOC match"""
    cleaned = clean_text(text)
    match = re.match(r"^\*([A-Z\/\-\s]+)\*$", cleaned)
    if not match:
        return None

    extracted_name = match.group(1).strip()

    # If no TOC provided, return as-is
    if not toc_categories:
        return extracted_name

    # Try exact match first
    if extracted_name in toc_categories:
        return extracted_name

    # Find closest match using fuzzy matching
    from difflib import get_close_matches

    close_matches = get_close_matches(extracted_name, toc_categories, n=1, cutoff=0.6)

    if close_matches:
        return close_matches[0]

    # If no close match, return the extracted name
    return extracted_name


def extract_subcategory_name(text):
    """Extract subcategory name from *SUBCATEGORY*** format"""
    cleaned = clean_text(text)
    # Match and extract, then remove trailing asterisks and spaces
    match = re.match(r"^\*(.+?)\*[\s*]+$", cleaned)
    if match:
        name = match.group(1).strip()
        # Remove any trailing asterisks that might have been captured
        return name.rstrip("* ")
    return None


def process_row(row, page_num):
    """Process a single table row and return cleaned data."""
    if not row or len(row) < 3:
        return None

    drug_name_raw = (row[0] or "").strip()
    tier = clean_text(row[1] or "")
    notes = clean_text(row[2] or "")

    # Skip header rows and empty rows
    if not drug_name_raw or drug_name_raw.lower() in [
        "drug name",
        "drugname",
        "tier",
        "notes",
    ]:
        return None

    # Clean the drug name (replace newlines with spaces)
    drug_name_cleaned = clean_text(drug_name_raw)

    return {
        "drug_name": drug_name_cleaned,
        "tier": tier,
        "notes": notes,
        "page": page_num,
    }


def classify_row(row_data, toc_categories=None):
    """Classify a row as category, subcategory, or regular drug row."""
    drug_name = row_data["drug_name"]

    if is_category(drug_name):
        return "category", extract_category_name(drug_name, toc_categories)
    elif is_subcategory(drug_name):
        return "subcategory", extract_subcategory_name(drug_name)
    else:
        return "drug", row_data


def extract_tables_from_page(page):
    """Extract and sort tables from a page (left to right)."""
    table_objects = page.find_tables()
    sorted_tables = sorted(table_objects, key=lambda t: t.bbox[0])
    return [t.extract() for t in sorted_tables]


def extract_structured_data(pdf_path):
    """Main function to extract all structured data from the PDF."""
    toc = extract_table_of_contents(pdf_path)
    toc_categories = list(toc.keys()) if toc else None

    start_page = min(toc.values()) if toc else 1
    last_toc_page = max(toc.values()) if toc else 1

    print(f"\nStarting extraction from page {start_page}")
    print(f"Last TOC page: {last_toc_page}")

    categories = []
    current_category = None
    current_subcategory = None
    warnings = []

    consecutive_pages_without_tables = 0
    max_pages_without_tables = 1

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(start_page, len(pdf.pages) + 1):
            page = pdf.pages[page_num - 1]
            print(f"\nProcessing page {page_num}...")

            tables = extract_tables_from_page(page)

            if not tables:
                consecutive_pages_without_tables += 1
                print(
                    f"  No tables found ({consecutive_pages_without_tables}/{max_pages_without_tables})"
                )

                if (
                    page_num > last_toc_page
                    and consecutive_pages_without_tables >= max_pages_without_tables
                ):
                    print(
                        f"  Stopping extraction - no tables for {max_pages_without_tables} consecutive pages"
                    )
                    break
                continue

            consecutive_pages_without_tables = 0

            # Process all rows in all tables on this page
            for table in tables:
                for row in table:
                    row_data = process_row(row, page_num)
                    if not row_data:
                        continue

                    row_type, value = classify_row(row_data, toc_categories)

                    if row_type == "category":
                        # Start a new category
                        current_category = {"categoryName": value, "subCategories": []}
                        categories.append(current_category)
                        current_subcategory = None
                        print(f"  > Category: {value}")

                    elif row_type == "subcategory":
                        # Start a new subcategory within current category
                        if current_category is None:
                            warnings.append(
                                {
                                    "issue": "Subcategory without category",
                                    "data": row_data,
                                }
                            )
                            continue

                        current_subcategory = {"subCategoryName": value, "rows": []}
                        current_category["subCategories"].append(current_subcategory)
                        print(f"    > Subcategory: {value}")

                    elif row_type == "drug":
                        # Add drug row to current subcategory
                        if current_category is None or current_subcategory is None:
                            warnings.append(
                                {
                                    "issue": "Drug row without category/subcategory",
                                    "data": row_data,
                                }
                            )
                            continue

                        current_subcategory["rows"].append(
                            {
                                "drug_name": row_data["drug_name"],
                                "tier": row_data["tier"],
                                "notes": row_data["notes"],
                                "page": row_data["page"],
                            }
                        )
                        print(
                            f"      Row: [{row_data['drug_name']}, {row_data['tier']}, {row_data['notes']}]"
                        )

            # Print progress
            total_rows = sum(
                len(subcat["rows"])
                for cat in categories
                for subcat in cat["subCategories"]
            )
            print(f"  Total rows extracted so far: {total_rows}")

    return {"table_of_contents": toc, "categories": categories, "warnings": warnings}


def main():
    parser = argparse.ArgumentParser(
        description="Extract structured data from pharmaceutical formulary PDFs."
    )
    parser.add_argument(
        "pdf_path",
        help="Path to the PDF file to extract",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="output",
        help="Output directory for extracted data (default: output)",
    )

    args = parser.parse_args()

    # Validate input file
    if not os.path.exists(args.pdf_path):
        print(f"Error: PDF file not found: {args.pdf_path}")
        return 1

    # Create output directory: output/PDFFILENAME/
    pdf_filename = Path(args.pdf_path).stem  # Get filename without extension
    output_dir = Path(args.output_dir) / pdf_filename
    output_dir.mkdir(parents=True, exist_ok=True)

    print("Starting PDF extraction...")
    print("=" * 80)
    print(f"Input PDF: {args.pdf_path}")
    print(f"Output directory: {output_dir}")
    print("=" * 80)

    data = extract_structured_data(args.pdf_path)

    # Save categories to JSON file (new format)
    output_file = output_dir / "extracted_data.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data["categories"], f, indent=2, ensure_ascii=False)

    # Save warnings to a separate file
    warnings_file = output_dir / "extraction_warnings.json"
    with open(warnings_file, "w", encoding="utf-8") as f:
        json.dump(data["warnings"], f, indent=2, ensure_ascii=False)

    # Save TOC to separate file
    toc_file = output_dir / "table_of_contents.json"
    with open(toc_file, "w", encoding="utf-8") as f:
        json.dump(data["table_of_contents"], f, indent=2, ensure_ascii=False)

    print("\n" + "=" * 80)
    print(f"Extraction complete!")
    print(f"  - Categories saved to: {output_file}")
    print(f"  - Warnings saved to: {warnings_file}")
    print(f"  - Table of Contents saved to: {toc_file}")

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

    # Show sample structure
    print(f"\nSample structure:")
    for i, category in enumerate(data["categories"][:2]):
        print(f"\n  Category: {category['categoryName']}")
        for subcat in category["subCategories"][:2]:
            print(f"    Subcategory: {subcat['subCategoryName']}")
            if subcat["rows"]:
                first_row = subcat["rows"][0]
                print(
                    f"      First row: {first_row['drug_name']} (page {first_row['page']})"
                )
                print(f"      Total rows in subcategory: {len(subcat['rows'])}")

    return 0


if __name__ == "__main__":
    exit(main())
