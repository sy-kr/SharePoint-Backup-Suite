#!/usr/bin/env python3
"""
html_to_md.py — Convert HTML exported from Microsoft Loop to clean Markdown.

Usage:
    python tools/html_to_md.py --in page.html --out page.md

Deterministic, stable output. Strips scripts/styles, preserves headings,
lists, tables, and links.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from bs4 import BeautifulSoup, Comment
except ImportError:
    print(
        "ERROR: beautifulsoup4 is not installed.\n"
        "Run:  pip install -r tools/requirements.txt",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    from markdownify import markdownify as md, MarkdownConverter
except ImportError:
    print(
        "ERROR: markdownify is not installed.\n"
        "Run:  pip install -r tools/requirements.txt",
        file=sys.stderr,
    )
    sys.exit(1)


# ---------------------------------------------------------------------------
# Custom converter for cleaner output
# ---------------------------------------------------------------------------
class LoopMarkdownConverter(MarkdownConverter):
    """Subclass markdownify's converter for Loop-specific tweaks."""

    def convert_table(self, el, text, parent_tags):
        """Ensure tables are converted with proper GFM formatting."""
        return super().convert_table(el, text, parent_tags)

    def convert_a(self, el, text, parent_tags):
        """Preserve links; strip empty anchors."""
        href = el.get("href", "")
        title = el.get("title", "")
        if not text.strip() and not href:
            return ""
        if not href:
            return text
        if title:
            return f'[{text}]({href} "{title}")'
        return f"[{text}]({href})"

    def convert_img(self, el, text, parent_tags):
        """Preserve images with alt text."""
        alt = el.get("alt", "")
        src = el.get("src", "")
        title = el.get("title", "")
        if not src:
            return ""
        if title:
            return f'![{alt}]({src} "{title}")'
        return f"![{alt}]({src})"


def convert_html(html_content: str) -> str:
    """Convert HTML to Markdown with Loop-specific cleaning."""

    soup = BeautifulSoup(html_content, "html.parser")

    # ── Strip unwanted elements ──────────────────────────────────────────
    for tag in soup.find_all(["script", "style", "noscript", "iframe"]):
        tag.decompose()

    # Remove HTML comments
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()

    # Remove hidden elements
    for tag in soup.find_all(attrs={"style": re.compile(r"display\s*:\s*none", re.I)}):
        tag.decompose()

    # Remove Microsoft-specific meta/link tags that add noise
    for tag in soup.find_all("meta"):
        tag.decompose()
    for tag in soup.find_all("link"):
        tag.decompose()

    # ── Extract body if present ──────────────────────────────────────────
    body = soup.find("body")
    if body:
        html_str = str(body)
    else:
        html_str = str(soup)

    # ── Convert to Markdown ──────────────────────────────────────────────
    markdown = LoopMarkdownConverter(
        heading_style="atx",
        bullets="-",
        strong_em_symbol="*",
        sub_symbol="",
        sup_symbol="",
        newline_style="backslash",
        strip=["span", "script", "style", "noscript", "iframe"],
    ).convert(html_str)

    # ── Post-process ─────────────────────────────────────────────────────
    # Collapse excessive blank lines (more than 2 consecutive) to 2
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)

    # Strip trailing whitespace on each line
    lines = [line.rstrip() for line in markdown.splitlines()]
    markdown = "\n".join(lines)

    # Ensure file ends with single newline
    markdown = markdown.strip() + "\n"

    return markdown


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert HTML (from Microsoft Loop export) to Markdown."
    )
    parser.add_argument(
        "--in",
        dest="input_file",
        required=True,
        help="Path to input HTML file.",
    )
    parser.add_argument(
        "--out",
        dest="output_file",
        required=True,
        help="Path to output Markdown file.",
    )
    args = parser.parse_args()

    input_path = Path(args.input_file)
    output_path = Path(args.output_file)

    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        return 1

    try:
        html_content = input_path.read_text(encoding="utf-8")
    except Exception as exc:
        print(f"ERROR: Could not read {input_path}: {exc}", file=sys.stderr)
        return 1

    if not html_content.strip():
        print(f"WARNING: Input file is empty: {input_path}", file=sys.stderr)
        output_path.write_text("", encoding="utf-8")
        return 0

    try:
        markdown = convert_html(html_content)
    except Exception as exc:
        print(f"ERROR: Conversion failed: {exc}", file=sys.stderr)
        return 1

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(markdown, encoding="utf-8")
    except Exception as exc:
        print(f"ERROR: Could not write {output_path}: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
