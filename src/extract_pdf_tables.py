import pdfplumber
import json
import re
import argparse
import os
from pathlib import Path
from config import DEBUG_MODE


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
                    if DEBUG_MODE:
                        print(f"Found Table of Contents on page {page_num}")

                # Match lines with format: *CATEGORY NAME*....page_number
                match = re.search(r"\*(.+?)\*\.+(\d+)", line)
                if match:
                    found_toc = True  # If we find TOC pattern, we're in TOC
                    category_name = clean_text(match.group(1))
                    page_no = int(match.group(2))
                    categories[category_name] = page_no
                    if DEBUG_MODE:
                        print(f"  Found category: {category_name} -> Page {page_no}")
                    found_any_match = True

            # If we've found TOC before but no matches on this page, stop
            if found_toc and not found_any_match and categories:
                if DEBUG_MODE:
                    print(
                        f"  No more TOC entries found on page {page_num}, stopping TOC extraction"
                    )
                break

    return categories


def is_category(text, toc_categories=None):
    """
    Check if text is a category by:
    1. Must be enclosed in * and *
    2. Must have a close match in the TOC
    """
    cleaned = clean_text(text)

    # Check if it's enclosed in asterisks (simple * and *)
    if not re.match(r"^\*(.+?)\*$", cleaned):
        return False

    # If no TOC provided, we can't determine if it's a category
    if not toc_categories:
        return False

    # Extract the content between asterisks
    match = re.match(r"^\*(.+?)\*$", cleaned)
    if not match:
        return False

    extracted_name = match.group(1).strip()

    # Check for exact match first
    if extracted_name in toc_categories:
        return True

    # Check for close match using fuzzy matching
    from difflib import get_close_matches

    close_matches = get_close_matches(extracted_name, toc_categories, n=1, cutoff=0.8)

    return len(close_matches) > 0


def is_subcategory(text, toc_categories=None):
    """
    Check if text is a subcategory by:
    1. FIRST PRIORITY: If enclosed in * and *** (or ** *, * **), it's ALWAYS a subcategory
    2. SECOND PRIORITY: If simple * and *, check if it has TOC match (category) or not (subcategory)
    """
    cleaned = clean_text(text)

    # FIRST PRIORITY: Check for * and *** patterns (ALWAYS subcategory)
    # Match patterns like: *SUBCATEGORY***, *SUBCATEGORY** *, *SUBCATEGORY* **
    if re.match(r"^\*(.+?)\*[\s*]+$", cleaned):
        # If it has trailing asterisks/spaces after the closing *, it's ALWAYS a subcategory
        return True

    # SECOND PRIORITY: Check simple * and * format
    if re.match(r"^\*(.+?)\*$", cleaned):
        # If it matches category criteria (TOC match), it's not a subcategory
        if is_category(text, toc_categories):
            return False
        # If it doesn't match TOC, it's a subcategory
        return True

    # If it doesn't match any asterisk patterns, it's not a subcategory
    return False


def extract_category_name(text, toc_categories=None):
    """Extract category name from *CATEGORY* format and find closest TOC match"""
    cleaned = clean_text(text)
    match = re.match(r"^\*(.+?)\*$", cleaned)
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

    close_matches = get_close_matches(extracted_name, toc_categories, n=1, cutoff=0.8)

    if close_matches:
        return close_matches[0]

    # If no close match, return the extracted name
    return extracted_name


def extract_subcategory_name(text):
    """Extract subcategory name from *SUBCATEGORY* or *SUBCATEGORY*** format"""
    cleaned = clean_text(text)

    # Match anything that starts with * and has content before the next *
    # Then optionally followed by more asterisks and spaces
    match = re.match(r"^\*(.+?)\*.*$", cleaned)
    if match:
        # Just return the content between the first * and first closing *
        return match.group(1).strip()

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
    tier = row_data["tier"]
    notes = row_data["notes"]

    # Only check for category/subcategory if tier and notes columns are empty
    # Categories and subcategories should not have tier or notes data
    if not tier.strip() and not notes.strip():
        # Check subcategory first (handles both *** priority and simple * * patterns)
        if is_subcategory(drug_name, toc_categories):
            return "subcategory", extract_subcategory_name(drug_name)
        elif is_category(drug_name, toc_categories):
            return "category", extract_category_name(drug_name, toc_categories)

    # If tier/notes are not empty, or if it doesn't match category/subcategory patterns, it's a drug
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

    if DEBUG_MODE:
        print(f"\nStarting extraction from page {start_page}")
        print(f"Last TOC page: {last_toc_page}")

    categories = []
    current_category = None
    current_subcategory = None
    warnings = []
    total_rows_processed = 0

    consecutive_pages_without_tables = 0
    max_pages_without_tables = 1

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(start_page, len(pdf.pages) + 1):
            page = pdf.pages[page_num - 1]
            if DEBUG_MODE:
                print(f"\nProcessing page {page_num}...")

            tables = extract_tables_from_page(page)

            if not tables:
                consecutive_pages_without_tables += 1
                if DEBUG_MODE:
                    print(
                        f"  No tables found ({consecutive_pages_without_tables}/{max_pages_without_tables})"
                    )

                if (
                    page_num > last_toc_page
                    and consecutive_pages_without_tables >= max_pages_without_tables
                ):
                    if DEBUG_MODE:
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

                    total_rows_processed += 1
                    row_type, value = classify_row(row_data, toc_categories)

                    if row_type == "category":
                        # Start a new category
                        current_category = {"categoryName": value, "subCategories": []}
                        categories.append(current_category)
                        current_subcategory = None
                        if DEBUG_MODE:
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
                        if DEBUG_MODE:
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
                        if DEBUG_MODE:
                            print(
                                f"      Row: [{row_data['drug_name']}, {row_data['tier']}, {row_data['notes']}]"
                            )

            # Print progress
            total_rows = sum(
                len(subcat["rows"])
                for cat in categories
                for subcat in cat["subCategories"]
            )
            if DEBUG_MODE:
                print(f"  Total rows extracted so far: {total_rows}")

    return {
        "table_of_contents": toc,
        "categories": categories,
        "warnings": warnings,
        "total_rows_processed": total_rows_processed,
    }


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

    if DEBUG_MODE:
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

    if DEBUG_MODE:
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
        total_subcategories = sum(
            len(cat["subCategories"]) for cat in data["categories"]
        )
        total_rows = sum(
            len(subcat["rows"])
            for cat in data["categories"]
            for subcat in cat["subCategories"]
        )

        print(f"  - Total subcategories: {total_subcategories}")
        print(f"  - Total rows extracted: {total_rows}")
        print(f"  - Total rows processed: {data['total_rows_processed']}")

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
