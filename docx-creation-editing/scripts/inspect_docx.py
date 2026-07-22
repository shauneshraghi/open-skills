#!/usr/bin/env python3
"""
inspect_docx.py — Extract and pretty-print OOXML parts from a .docx file.

Usage:
    python inspect_docx.py <file.docx> [part ...]

Examples:
    python inspect_docx.py report.docx
    python inspect_docx.py report.docx word/styles.xml word/numbering.xml
    python inspect_docx.py report.docx --list

Default parts printed when none specified:
    word/document.xml

Common parts:
    word/document.xml       Body content
    word/styles.xml         Style definitions
    word/numbering.xml      List/outline numbering
    word/header1.xml        Default header
    word/header2.xml        Even-page header
    word/footer1.xml        Default footer
    word/footer2.xml        Even-page footer
    word/settings.xml       Document settings (compat flags, etc.)
    [Content_Types].xml     Package content type registry
"""

import sys
import zipfile
import argparse
import re
from pathlib import Path
from lxml import etree

SKIP_NS_PREFIXES = (
    "xmlns:wpc", "xmlns:cx", "xmlns:aink", "xmlns:am3d",
    "xmlns:oel", "xmlns:o", "xmlns:v", "xmlns:wne",
)

RSID_ATTRS = re.compile(r'\bw(?:14)?:\w*[Rr]sid\w*="\w+"')
PARA_ID_ATTRS = re.compile(r'\bw14:(?:paraId|textId)="\w+"')


def strip_noise(xml_bytes: bytes) -> str:
    """Pretty-print XML, stripping rsid/paraId tracking attrs for readability."""
    root = etree.fromstring(xml_bytes)
    pretty = etree.tostring(root, pretty_print=True).decode()
    pretty = RSID_ATTRS.sub("", pretty)
    pretty = PARA_ID_ATTRS.sub("", pretty)
    # Collapse lines that became empty after attr removal
    lines = [l for l in pretty.splitlines() if l.strip() or not l.strip() == ""]
    return "\n".join(lines)


def list_parts(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as z:
        return sorted(z.namelist())


def extract_part(docx_path: Path, part: str) -> str | None:
    with zipfile.ZipFile(docx_path) as z:
        if part not in z.namelist():
            return None
        return strip_noise(z.read(part))


def main():
    parser = argparse.ArgumentParser(
        description="Extract and pretty-print OOXML parts from a .docx file."
    )
    parser.add_argument("docx", help="Path to the .docx file")
    parser.add_argument(
        "parts",
        nargs="*",
        default=["word/document.xml"],
        help="OOXML parts to extract (default: word/document.xml)",
    )
    parser.add_argument(
        "--list", action="store_true", help="List all parts in the package and exit"
    )
    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"ERROR: file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    if args.list:
        parts = list_parts(docx_path)
        print(f"Parts in {docx_path.name}:")
        for p in parts:
            print(f"  {p}")
        return

    for part in args.parts:
        xml = extract_part(docx_path, part)
        if xml is None:
            print(f"\n=== {part} — NOT FOUND ===\n", file=sys.stderr)
            continue
        print(f"\n{'='*72}")
        print(f"  {part}")
        print(f"{'='*72}\n")
        print(xml)


if __name__ == "__main__":
    main()
