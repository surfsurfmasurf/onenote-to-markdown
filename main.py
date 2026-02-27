"""
main.py
-------
Entry point for the OneNote -> Markdown migration tool.

Orchestrates the full pipeline:
    1. Authenticate with Microsoft via Device Code Flow
        2. Fetch the complete OneNote notebook / section / page hierarchy
            3. Download the HTML content of every page
                4. Convert each page to Obsidian-compatible Markdown
                    5. Write the Markdown files to disk, mirroring the OneNote structure
                        6. Generate an _index.md (table of contents) per notebook

                        Usage:
                            python main.py

                            See README.md for setup instructions (Azure App Registration, .env config).
                            """

import sys
from tqdm import tqdm

from config import OUTPUT_DIR
from onenote_fetcher import get_access_token, OneNoteFetcher
from converter import HtmlToMarkdownConverter
from exporter import MarkdownExporter


def main() -> None:
      print("=" * 60)
      print("  OneNote -> Obsidian / Notion Migration Tool")
      print("=" * 60)

    # ------------------------------------------------------------------
      # Step 1: Authenticate
      # ------------------------------------------------------------------
      try:
                token = get_access_token()
except RuntimeError as exc:
        print(f"\n[ERROR] Authentication failed: {exc}")
        sys.exit(1)

    # ------------------------------------------------------------------
    # Step 2: Discover notebook structure
    # ------------------------------------------------------------------
    fetcher = OneNoteFetcher(token)
    converter = HtmlToMarkdownConverter()
    exporter = MarkdownExporter(OUTPUT_DIR)

    print("\nAnalysing OneNote structure...")
    structure = fetcher.get_all_structure()

    total_pages = sum(
              len(section["pages"])
              for notebook in structure
              for section in notebook["sections"]
    )
    notebook_count = len(structure)
    print(
              f"\nFound {notebook_count} notebook(s) containing "
              f"{total_pages} page(s) in total.\n"
    )

    if total_pages == 0:
              print("Nothing to export. Exiting.")
              sys.exit(0)

    # ------------------------------------------------------------------
    # Step 3-5: Download, convert, and save each page
    # ------------------------------------------------------------------
    errors: list[str] = []

    with tqdm(total=total_pages, desc="Exporting", unit="page") as progress:
              for notebook in structure:
                            for section in notebook["sections"]:
                                              for page in section["pages"]:
                                                                    title = page["title"]
                                                                    progress.set_description(
                                                                        f"Converting: {title[:35]}{'...' if len(title) > 35 else ''}"
                                                                    )

                                                  try:
                                                                            html = fetcher.get_page_content(page["id"])
                                                                            markdown = converter.convert(
                                                                                html,
                                                                                title=title,
                                                                                created=page.get("createdDateTime"),
                                                                                modified=page.get("lastModifiedDateTime"),
                                                                            )
                                                                            exporter.save_page(
                                                                                notebook_name=notebook["name"],
                                                                                section_name=section["name"],
                                                                                page_title=title,
                                                                                markdown_content=markdown,
                                                                            )
except Exception as exc:  # noqa: BLE001
                        msg = f"{notebook['name']} / {section['name']} / {title}: {exc}"
                        errors.append(msg)
                        tqdm.write(f"  [WARNING] Skipped – {msg}")

                    progress.update(1)

    # ------------------------------------------------------------------
    # Step 6: Generate per-notebook index files
    # ------------------------------------------------------------------
    print("\nGenerating index files...")
    exporter.create_index(structure)

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    exported = total_pages - len(errors)
    print(f"\n{'=' * 60}")
    print(f"  Export complete!")
    print(f"  Exported : {exported} / {total_pages} pages")
    print(f"  Skipped  : {len(errors)} page(s) due to errors")
    print(f"  Output   : {OUTPUT_DIR}")
    print(f"{'=' * 60}")
    print()
    print("Next steps:")
    print("  Obsidian -> Open the output folder as a new Vault")
    print("  Notion   -> Settings > Import > Markdown & CSV, then select the output folder")

    if errors:
              print("\nPages that could not be exported:")
              for err in errors:
                            print(f"  - {err}")


if __name__ == "__main__":
      main()
