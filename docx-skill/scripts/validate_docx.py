"""
validate_docx.py — Validate .docx file structure and content.

Design decisions sourced from:
- ECMA-376 Part 2 §10: Open Packaging Conventions — required parts and relationships
- ECMA-376 §17.13.4: comment markup — commentRangeStart/End must have matching w:comment
- python-docx oxml/ns.py: namespace prefix constants
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Union

from lxml import etree

_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Required parts per ECMA-376 Part 2 §13.2 / OPC spec
_REQUIRED_PARTS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
}


def validate(path: Union[str, Path]) -> dict:
    """Validate a .docx file and return a result dict.

    Returns:
        {
            "valid": bool,
            "errors": [str],
            "warnings": [str],
            "stats": {
                "paragraphs": int,
                "images": int,
                "comments": int,
                "tracked_insertions": int,
                "tracked_deletions": int,
            }
        }
    """
    path = Path(path)
    result = {
        "valid": True,
        "errors": [],
        "warnings": [],
        "stats": {
            "paragraphs": 0,
            "images": 0,
            "comments": 0,
            "tracked_insertions": 0,
            "tracked_deletions": 0,
        },
    }

    # ── 1. ZIP structure check (ECMA-376 Part 2 §10) ─────────────────────────
    if not path.exists():
        result["errors"].append(f"File not found: {path}")
        result["valid"] = False
        return result

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        result["errors"].append(f"Not a valid ZIP/OOXML package: {e}")
        result["valid"] = False
        return result

    with zf:
        names = set(zf.namelist())

        # ── 2. Required parts ─────────────────────────────────────────────────
        for part in _REQUIRED_PARTS:
            if part not in names:
                result["errors"].append(f"Missing required part: {part}")
                result["valid"] = False

        if not result["valid"]:
            return result

        # ── 3. Parse document.xml ─────────────────────────────────────────────
        doc_xml = zf.read("word/document.xml")
        try:
            doc_el = etree.fromstring(doc_xml)
        except etree.XMLSyntaxError as e:
            result["errors"].append(f"document.xml parse error: {e}")
            result["valid"] = False
            return result

        # ── 4. Count paragraphs ───────────────────────────────────────────────
        result["stats"]["paragraphs"] = len(doc_el.findall(f".//{{{_W}}}p"))

        # ── 5. Count and validate images ─────────────────────────────────────
        # Images are referenced via a:blip r:embed="rId..." (ECMA-376 §20.1.8.13)
        rels_path = "word/_rels/document.xml.rels"
        rid_to_target: dict[str, str] = {}
        if rels_path in names:
            rels_xml = zf.read(rels_path)
            rels_el  = etree.fromstring(rels_xml)
            for rel in rels_el:
                rid    = rel.get("Id", "")
                target = rel.get("Target", "")
                rid_to_target[rid] = target

        image_rids = set()
        for blip in doc_el.findall(f".//{{{_A}}}blip"):
            rid = blip.get(f"{{{_R}}}embed")
            if rid:
                image_rids.add(rid)

        result["stats"]["images"] = len(image_rids)

        for rid in image_rids:
            target = rid_to_target.get(rid, "")
            # Targets are relative to word/, so prepend
            media_path = f"word/{target}" if not target.startswith("/") else target.lstrip("/")
            if media_path not in names:
                result["errors"].append(
                    f"Image relationship {rid} → {target} target not found in package"
                )
                result["valid"] = False

        # ── 6. Validate comments ──────────────────────────────────────────────
        if "word/comments.xml" in names:
            comments_xml = zf.read("word/comments.xml")
            comments_el  = etree.fromstring(comments_xml)
            comment_ids  = {
                c.get(f"{{{_W}}}id")
                for c in comments_el.findall(f"{{{_W}}}comment")
            }
            result["stats"]["comments"] = len(comment_ids)

            # Every commentRangeStart in document.xml should have a matching w:comment
            for rs in doc_el.findall(f".//{{{_W}}}commentRangeStart"):
                cid = rs.get(f"{{{_W}}}id")
                if cid not in comment_ids:
                    result["warnings"].append(
                        f"commentRangeStart id={cid} has no matching w:comment in comments.xml"
                    )
        else:
            # Check for orphaned commentRangeStart elements
            orphans = doc_el.findall(f".//{{{_W}}}commentRangeStart")
            if orphans:
                result["warnings"].append(
                    f"{len(orphans)} commentRangeStart element(s) found but comments.xml is absent"
                )

        # ── 7. Count track changes ────────────────────────────────────────────
        result["stats"]["tracked_insertions"] = len(
            doc_el.findall(f".//{{{_W}}}ins")
        )
        result["stats"]["tracked_deletions"] = len(
            doc_el.findall(f".//{{{_W}}}del")
        )

    return result


def validate_and_print(path: Union[str, Path]) -> bool:
    """Validate and pretty-print results. Returns True if valid."""
    import json
    r = validate(path)
    print(json.dumps(r, indent=2))
    return r["valid"]


def _cli() -> None:
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Validate a .docx file")
    parser.add_argument("file", nargs="+", help=".docx file(s) to validate")
    args = parser.parse_args()

    all_ok = True
    for f in args.file:
        print(f"\n=== {f} ===")
        ok = validate_and_print(f)
        if not ok:
            all_ok = False

    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    _cli()
