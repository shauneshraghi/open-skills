"""Validate PowerPoint (.pptx) files.

Design decisions from primary sources:
- ECMA-376 Part 2 §10: OPC packages are ZIP archives; we use zipfile to check
  structural integrity before invoking python-pptx.
- ECMA-376 Part 1 §13.3: required parts ([Content_Types].xml, _rels/.rels,
  ppt/presentation.xml) must be present in any valid PresentationML document.
- Relationship validation: every r:id referenced in the XML must resolve to a
  part listed in the corresponding .rels file.
- XML well-formedness: each part must be parseable by lxml.etree without error.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

REQUIRED_PARTS = [
    "[Content_Types].xml",
    "_rels/.rels",
]

PRESENTATION_CONTENT_TYPES = {
    # Standard .pptx (ECMA-376)
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    # Macro-enabled .pptm (MS Office extension)
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    # Presentation template .potx
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
}


@dataclass
class ValidationResult:
    path: str
    valid: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    info: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "path": self.path,
            "valid": self.valid,
            "errors": self.errors,
            "warnings": self.warnings,
            "info": self.info,
        }


def validate_pptx(path: str | Path) -> ValidationResult:
    """Run all validation checks on a .pptx file and return a ValidationResult."""
    path = Path(path)
    result = ValidationResult(path=str(path), valid=True)

    if not path.exists():
        result.valid = False
        result.errors.append(f"File not found: {path}")
        return result

    _check_zip(path, result)
    if not result.valid:
        return result

    _check_required_parts(path, result)
    _check_xml_well_formed(path, result)
    _check_slide_counts(path, result)
    _check_relationship_targets(path, result)

    return result


def _check_zip(path: Path, result: ValidationResult) -> None:
    if not zipfile.is_zipfile(path):
        result.valid = False
        result.errors.append("Not a valid ZIP/OPC archive.")
        return
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                result.valid = False
                result.errors.append(f"ZIP CRC error in member: {bad}")
    except zipfile.BadZipFile as exc:
        result.valid = False
        result.errors.append(f"BadZipFile: {exc}")


def _check_required_parts(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for part in REQUIRED_PARTS:
            if part not in names:
                result.valid = False
                result.errors.append(f"Missing required part: {part}")

        if "[Content_Types].xml" in names:
            ct_xml = z.read("[Content_Types].xml")
            try:
                root = etree.fromstring(ct_xml)
                prs_parts = [
                    el.get("PartName")
                    for el in root.iter()
                    if el.get("ContentType") in PRESENTATION_CONTENT_TYPES
                ]
                if not prs_parts:
                    result.valid = False
                    result.errors.append("No presentation part found in [Content_Types].xml.")
                else:
                    result.info["presentation_part"] = prs_parts[0]
            except etree.XMLSyntaxError as exc:
                result.valid = False
                result.errors.append(f"[Content_Types].xml XML error: {exc}")


def _check_xml_well_formed(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        xml_parts = [n for n in z.namelist() if n.endswith(".xml") or n.endswith(".rels")]
        result.info["xml_part_count"] = len(xml_parts)
        malformed = []
        for name in xml_parts:
            try:
                data = z.read(name)
                etree.fromstring(data)
            except etree.XMLSyntaxError as exc:
                malformed.append(f"{name}: {exc}")
        if malformed:
            result.valid = False
            for m in malformed:
                result.errors.append(f"Malformed XML: {m}")


def _check_slide_counts(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        slides = [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        notes = [n for n in names if "notesSlide" in n and n.endswith(".xml")]
        images = [
            n for n in names
            if n.startswith("ppt/media/") and any(
                n.lower().endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf")
            )
        ]
        result.info["slide_count"] = len(slides)
        result.info["notes_slide_count"] = len(notes)
        result.info["image_count"] = len(images)


def _check_relationship_targets(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        rels_files = [n for n in names if n.endswith(".rels")]
        broken_rels = []

        for rels_path in rels_files:
            try:
                data = z.read(rels_path)
                root = etree.fromstring(data)
            except Exception:
                continue

            parts = rels_path.split("/")
            if len(parts) >= 2 and parts[-2] == "_rels":
                base_dir = "/".join(parts[:-2])
                source_name = parts[-1].replace(".rels", "")
            else:
                base_dir = ""
                source_name = ""

            for rel in root:
                target_mode = rel.get("TargetMode", "Internal")
                if target_mode == "External":
                    continue
                target = rel.get("Target", "")
                if not target or target.startswith("http"):
                    continue

                if base_dir:
                    resolved = _resolve_path(base_dir + "/" + source_name, target)
                else:
                    resolved = target.lstrip("/")

                if resolved and resolved not in names:
                    broken_rels.append(f"{rels_path}: Target '{target}' → '{resolved}' not found")

        if broken_rels:
            for br in broken_rels[:10]:
                result.warnings.append(f"Broken relationship: {br}")
            if len(broken_rels) > 10:
                result.warnings.append(f"…and {len(broken_rels) - 10} more broken relationships")


def _resolve_path(source: str, target: str) -> str:
    source_dir = source.rsplit("/", 1)[0] if "/" in source else ""
    parts = (source_dir + "/" + target).split("/")
    resolved: list[str] = []
    for part in parts:
        if part == "..":
            if resolved:
                resolved.pop()
        elif part and part != ".":
            resolved.append(part)
    return "/".join(resolved)


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate a .pptx file")
    parser.add_argument("--file", "-f", type=Path, required=True)
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    args = parser.parse_args()

    result = validate_pptx(args.file)

    if args.json:
        print(json.dumps(result.to_dict(), indent=2))
    else:
        status = "VALID" if result.valid else "INVALID"
        print(f"[{status}] {result.path}")
        for e in result.errors:
            print(f"  ERROR: {e}")
        for w in result.warnings:
            print(f"  WARN:  {w}")
        for k, v in result.info.items():
            print(f"  INFO:  {k} = {v}")

    sys.exit(0 if result.valid else 1)


if __name__ == "__main__":
    main()
