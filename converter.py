"""
converter.py
------------
Converts OneNote page HTML content into Obsidian-compatible Markdown.

Each converted page includes a YAML front matter block with metadata
(title, creation date, modification date, source tag) that is
understood by both Obsidian and Notion's Markdown importer.
"""

import re
import html2text
from datetime import datetime, timezone


class HtmlToMarkdownConverter:
      """
          Converts raw OneNote HTML into clean Markdown with YAML front matter.

              Uses the html2text library for the heavy lifting and applies several
                  post-processing steps to remove OneNote-specific artefacts.
                      """

    def __init__(self) -> None:
              h = html2text.HTML2Text()
              h.ignore_links = False       # Preserve hyperlinks
        h.ignore_images = False      # Preserve image references
        h.body_width = 0             # Disable forced line-wrapping
        h.protect_links = True       # Don't modify link URLs
        h.wrap_links = False         # Keep links on one line
        h.unicode_snob = True        # Use Unicode chars instead of ASCII
        self._h = h

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def convert(
              self,
              html_content: str,
              title: str = "",
              created: str | None = None,
              modified: str | None = None,
    ) -> str:
              """
                      Convert a OneNote page HTML string to Obsidian Markdown.

                              Args:
                                          html_content: Raw HTML string returned by the Graph API.
                                                      title:        Page title (used in front matter and as H1).
                                                                  created:      ISO-8601 creation timestamp from Graph API.
                                                                              modified:     ISO-8601 last-modified timestamp from Graph API.

                                                                                      Returns:
                                                                                                  A complete Markdown document string with YAML front matter.
                                                                                                          """
              front_matter = self._build_front_matter(title, created, modified)
              body = self._h.handle(html_content)
              body = self._post_process(body)
              return front_matter + body

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _build_front_matter(
              self,
              title: str,
              created: str | None,
              modified: str | None,
    ) -> str:
              """Build the YAML front matter block for the Markdown file."""
              lines = ["---"]

        if title:
                      # Escape double-quotes inside the title
                      safe_title = title.replace('"', '\\"')
                      lines.append(f'title: "{safe_title}"')

        for key, value in (("created", created), ("modified", modified)):
                      if value:
                                        formatted = self._format_timestamp(value)
                                        lines.append(f"{key}: {formatted}")

                  lines.append("source: OneNote")
        lines.append("---\n")

        return "\n".join(lines)

    @staticmethod
    def _format_timestamp(iso_string: str) -> str:
              """
                      Parse an ISO-8601 timestamp (e.g. "2024-03-15T10:30:00Z") and
                              return it in a human-readable format suitable for YAML front matter.
                                      """
              try:
                            dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
                            # Convert to local time for readability
                            return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
except (ValueError, AttributeError):
            return iso_string  # Fall back to the raw string if parsing fails

    @staticmethod
    def _post_process(text: str) -> str:
              """
                      Apply a series of regex-based clean-ups to remove OneNote artefacts
                              and normalise the Markdown output.
                                      """
        # Remove OneNote anchor fragments such as []( #section-id )
        text = re.sub(r"\[\]\(#[^)]*\)", "", text)

        # Collapse 3+ consecutive blank lines into 2
        text = re.sub(r"\n{3,}", "\n\n", text)

        # Strip leading/trailing whitespace from the whole document
        text = text.strip()

        return text
