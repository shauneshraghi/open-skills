"""Validate Excel (.xlsx / .xlsm / .xltx) files.

Design decisions from primary sources:
- ECMA-376 Part 2 §10: OPC packages are ZIP archives; we use zipfile to check
  structural integrity before invoking openpyxl.
- ECMA-376 Part 1 §13.3: required parts ([Content_Types].xml, _rels/.rels)
  must be present in any valid SpreadsheetML document.
- Three SpreadsheetML content types are recognised (ECMA-376 §11.1.1):
    .xlsx  application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml
    .xlsm  application/vnd.ms-excel.sheet.macroEnabled.main+xml
    .xltx  application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml
- XML well-formedness: every .xml and .rels part is parsed by lxml.etree.
- Relationship target resolution uses the same _resolve_path() logic as
  validate_pptx.py; external targets are skipped.
- Sheet and image counts are stored in result.info for callers.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree

REQUIRED_PARTS = [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
]

WORKSHEET_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
DRAWING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawing+xml"
)
SUPPORTED_MEDIA_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".emf",
    ".wmf",
    ".tiff",
    ".tif",
}
MEDIA_CONTENT_TYPES = {
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/bmp",
    "image/x-emf",
    "image/x-wmf",
    "image/tiff",
}
REL_WORKBOOK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_WORKSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
REL_DRAWING = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

SPREADSHEET_CONTENT_TYPES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
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


def validate_xlsx(path: str | Path) -> ValidationResult:
    """Run all validation checks on an .xlsx file and return a ValidationResult."""
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
    _check_counts(path, result)
    _check_relationship_targets(path, result)
    _check_workbook_structure(path, result)
    _check_drawing_and_media_integrity(path, result)

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
                wb_parts = [
                    el.get("PartName")
                    for el in root.iter()
                    if el.get("ContentType") in SPREADSHEET_CONTENT_TYPES
                ]
                if not wb_parts:
                    result.valid = False
                    result.errors.append(
                        "No spreadsheet workbook part found in [Content_Types].xml."
                    )
                else:
                    result.info["workbook_part"] = wb_parts[0]
                    workbook_part = wb_parts[0].lstrip("/")
                    if workbook_part not in names:
                        result.valid = False
                        result.errors.append(
                            f"Workbook part declared in [Content_Types].xml is missing: {workbook_part}"
                        )
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


def _check_counts(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        sheets = [
            n for n in names
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        ]
        images = [
            n for n in names
            if n.startswith("xl/media/") and Path(n).suffix.lower() in SUPPORTED_MEDIA_EXTENSIONS
        ]
        result.info["sheet_count"] = len(sheets)
        result.info["image_count"] = len(images)


def _check_relationship_targets(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        rels_files = [n for n in names if n.endswith(".rels")]
        broken_rels: list[str] = []

        for rels_path in rels_files:
            try:
                data = z.read(rels_path)
                root = etree.fromstring(data)
            except Exception:
                continue

            parts = rels_path.split("/")
            if len(parts) >= 2 and parts[-2] == "_rels":
                base_source = "/".join(parts[:-2]) + "/" + parts[-1].replace(".rels", "")
            else:
                base_source = parts[-1].replace(".rels", "")

            for rel in root:
                if rel.get("TargetMode", "Internal") == "External":
                    continue
                target = rel.get("Target", "")
                if not target or target.startswith("http"):
                    continue
                resolved = _resolve_path(base_source, target)
                if resolved and resolved not in names:
                    broken_rels.append(
                        f"{rels_path}: Target '{target}' → '{resolved}' not found"
                    )

        if broken_rels:
            result.valid = False
            for br in broken_rels[:10]:
                result.errors.append(f"Broken relationship: {br}")
            if len(broken_rels) > 10:
                result.errors.append(
                    f"…and {len(broken_rels) - 10} more broken relationships"
                )


def _check_workbook_structure(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        workbook_part = result.info.get("workbook_part", "/xl/workbook.xml").lstrip("/")
        if workbook_part not in names:
            result.valid = False
            result.errors.append(f"Workbook part missing: {workbook_part}")
            return

        workbook_rels = _rels_part_for_part(workbook_part)
        if workbook_rels not in names:
            result.valid = False
            result.errors.append(f"Missing workbook relationships part: {workbook_rels}")
            return

        workbook_root = etree.fromstring(z.read(workbook_part))
        workbook_rels_root = etree.fromstring(z.read(workbook_rels))
        worksheet_targets = {
            rel.get("Id", ""): _resolve_path(workbook_part, rel.get("Target", ""))
            for rel in workbook_rels_root
            if _rel_type(rel) == REL_WORKSHEET
        }
        workbook_links = [
            rel
            for rel in workbook_rels_root
            if _rel_type(rel) == REL_WORKBOOK
        ]
        if workbook_links:
            result.warnings.append(
                "Workbook relationships part unexpectedly contains officeDocument links; package structure may be unusual."
            )

        sheet_elements = workbook_root.findall(
            ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
        )
        if not sheet_elements:
            result.valid = False
            result.errors.append("Workbook contains no <sheet> elements.")
            return

        for sheet in sheet_elements:
            rel_id = sheet.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            sheet_name = sheet.get("name", "<unnamed>")
            if not rel_id or rel_id not in worksheet_targets:
                result.valid = False
                result.errors.append(
                    f"Workbook sheet {sheet_name!r} is missing a valid worksheet relationship."
                )
                continue
            worksheet_part = worksheet_targets[rel_id]
            if worksheet_part not in names:
                result.valid = False
                result.errors.append(
                    f"Workbook sheet {sheet_name!r} points to missing worksheet part: {worksheet_part}"
                )


def _check_drawing_and_media_integrity(path: Path, result: ValidationResult) -> None:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        ct_map = _content_type_map(z)
        worksheet_parts = [
            name
            for name in names
            if name.startswith("xl/worksheets/") and name.endswith(".xml")
        ]

        for worksheet_part in worksheet_parts:
            worksheet_rels = _rels_part_for_part(worksheet_part)
            if worksheet_rels not in names:
                continue
            worksheet_rels_root = etree.fromstring(z.read(worksheet_rels))
            for rel in worksheet_rels_root:
                if _rel_type(rel) != REL_DRAWING:
                    continue
                drawing_part = _resolve_path(worksheet_part, rel.get("Target", ""))
                if drawing_part not in names:
                    result.valid = False
                    result.errors.append(
                        f"Worksheet drawing part missing: {worksheet_part} -> {drawing_part}"
                    )
                    continue
                drawing_type = ct_map.get(drawing_part)
                if drawing_type is not None and drawing_type != DRAWING_CONTENT_TYPE:
                    result.valid = False
                    result.errors.append(
                        f"Drawing part content type mismatch for {drawing_part}: {drawing_type}"
                    )

                drawing_rels = _rels_part_for_part(drawing_part)
                if drawing_rels not in names:
                    result.valid = False
                    result.errors.append(f"Missing drawing relationships part: {drawing_rels}")
                    continue

                drawing_rels_root = etree.fromstring(z.read(drawing_rels))
                for drawing_rel in drawing_rels_root:
                    if _rel_type(drawing_rel) != REL_IMAGE:
                        continue
                    media_part = _resolve_path(drawing_part, drawing_rel.get("Target", ""))
                    if media_part not in names:
                        result.valid = False
                        result.errors.append(
                            f"Drawing image target missing: {drawing_part} -> {media_part}"
                        )
                        continue

                    media_ext = Path(media_part).suffix.lower()
                    if media_ext not in SUPPORTED_MEDIA_EXTENSIONS:
                        result.warnings.append(
                            f"Media part uses unrecognized extension {media_ext}: {media_part}"
                        )

                    media_type = ct_map.get(media_part)
                    if media_type is not None and media_type not in MEDIA_CONTENT_TYPES:
                        result.valid = False
                        result.errors.append(
                            f"Media part content type mismatch for {media_part}: {media_type}"
                        )


def _resolve_path(source: str, target: str) -> str:
    source_dir = source.rsplit("/", 1)[0] if "/" in source else ""
    combined = source_dir + "/" + target if source_dir else target
    parts = combined.split("/")
    resolved: list[str] = []
    for part in parts:
        if part == "..":
            if resolved:
                resolved.pop()
        elif part and part != ".":
            resolved.append(part)
    return "/".join(resolved)


def _rels_part_for_part(part_name: str) -> str:
    part_path = Path(part_name)
    return str(part_path.parent / "_rels" / f"{part_path.name}.rels").replace("\\", "/")


def _content_type_map(zf: zipfile.ZipFile) -> dict[str, str]:
    root = etree.fromstring(zf.read("[Content_Types].xml"))
    defaults = {
        child.get("Extension", "").lower(): child.get("ContentType", "")
        for child in root
        if child.tag.endswith("Default")
    }
    overrides = {
        child.get("PartName", "").lstrip("/"): child.get("ContentType", "")
        for child in root
        if child.tag.endswith("Override")
    }

    content_types: dict[str, str] = {}
    for name in zf.namelist():
        if name in overrides:
            content_types[name] = overrides[name]
            continue
        ext = Path(name).suffix.lower().lstrip(".")
        if ext and ext in defaults:
            content_types[name] = defaults[ext]
    return content_types


def _rel_type(rel: Any) -> str:
    return rel.get("Type", "")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate an .xlsx file")
    parser.add_argument("--file", "-f", type=Path, required=True)
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    args = parser.parse_args()

    result = validate_xlsx(args.file)

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
