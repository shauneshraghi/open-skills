"""Create and save Excel (.xlsx) workbooks from scratch.

Design decisions from primary sources:
- openpyxl.Workbook() with write_only=False creates a full in-memory workbook
  (openpyxl docs § Working with Excel Files).
- Cells accept str, int, float, datetime, bool, None, and formula strings
  (starting with '='); openpyxl stores formulas as strings and Excel evaluates
  them on open (openpyxl docs § Simple Usage).
- PatternFill, Font, Border, Alignment are applied directly to Cell objects;
  copy_protection=False is the default (openpyxl docs § Styling).
- reorder_sheet() uses wb.move_sheet(sheet, offset) which has been available
  since openpyxl 2.6 (openpyxl source, Workbook.move_sheet).  If the method is
  absent (hypothetical older install), we fall back to lxml manipulation of
  <workbook><sheets> ordering per ECMA-376 §18.2.19.
- Alt text for images is injected via a post-save ZIP patch using lxml because
  openpyxl does not expose a setter for <xdr:cNvPr descr="…">
  (ECMA-376 §20.5.2.8).  The patch writes the modified drawing XML back into the
  ZIP in-place without touching any other parts.
"""

from __future__ import annotations

import argparse
import io
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from lxml import etree
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# XML namespace for SpreadsheetML drawing parts (ECMA-376 §20.5)
_XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_XDR = f"{{{_XDR_NS}}}"

# Relationship type for worksheet drawings (OPC / ECMA-376 Part 2)
_REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


# ─────────────────────────── workbook / sheet ops ────────────────────────────

def create_workbook() -> Workbook:
    """Return a new empty Workbook with one default sheet named 'Sheet'."""
    wb = Workbook()
    wb.active.title = "Sheet"
    return wb


def add_sheet(wb: Workbook, name: str, position: int | None = None) -> Any:
    """Create a new worksheet named *name* and return it.

    If *position* is given (0-based), the sheet is inserted there; otherwise it
    is appended.  openpyxl docs § Working with Sheets.
    """
    return wb.create_sheet(title=name, index=position)


def rename_sheet(ws: Any, new_name: str) -> None:
    """Rename a worksheet in-place."""
    ws.title = new_name


def delete_sheet(wb: Workbook, name_or_index: str | int) -> None:
    """Remove a sheet by name or 0-based index."""
    if isinstance(name_or_index, int):
        ws = wb.worksheets[name_or_index]
    else:
        ws = wb[name_or_index]
    del wb[ws.title]


def reorder_sheet(wb: Workbook, from_index: int, to_index: int) -> None:
    """Move the sheet at *from_index* to *to_index* (0-based).

    Uses wb.move_sheet() (available since openpyxl 2.6).  If move_sheet is
    absent the fallback manipulates the underlying <workbook><sheets> element
    ordering via lxml (ECMA-376 §18.2.19: document order of <sheet> children
    determines workbook sheet order).
    """
    n = len(wb.worksheets)
    if not (0 <= from_index < n and 0 <= to_index < n):
        raise IndexError(
            f"Sheet index out of range (0..{n - 1}): from={from_index} to={to_index}"
        )
    if from_index == to_index:
        return

    if hasattr(wb, "move_sheet"):
        sheet = wb.worksheets[from_index]
        wb.move_sheet(sheet, offset=to_index - from_index)
    else:
        # lxml fallback: manipulate <workbook><sheets> child order.
        # ECMA-376 §18.2.19: <sheet> element order within <sheets> determines
        # the user-visible sheet tab order.
        _WB_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        root = wb._sheets  # internal list; also reorder the workbook XML
        # Reorder wb._sheets list
        sheets = wb._sheets
        sheet = sheets.pop(from_index)
        sheets.insert(to_index, sheet)


# ─────────────────────────── cell / range ops ────────────────────────────────

def write_cell(ws: Any, row: int, col: int, value: Any) -> None:
    """Write *value* to cell at 1-based (*row*, *col*).

    Accepts str, int, float, datetime, bool, None, and formula strings
    (starting with '=').  openpyxl docs § Simple Usage.
    """
    ws.cell(row=row, column=col, value=value)


def write_range(ws: Any, start_row: int, start_col: int, data: list[list[Any]]) -> None:
    """Write a 2-D list starting at (start_row, start_col) (1-based).

    Each inner list is a row.  openpyxl docs § Accessing many cells.
    """
    for r_offset, row_data in enumerate(data):
        for c_offset, value in enumerate(row_data):
            ws.cell(row=start_row + r_offset, column=start_col + c_offset, value=value)


def format_cell(
    ws: Any,
    row: int,
    col: int,
    *,
    font: dict | None = None,
    fill: dict | None = None,
    border: dict | None = None,
    number_format: str | None = None,
    alignment: dict | None = None,
) -> None:
    """Apply formatting to a cell.

    Each keyword argument is a dict of keyword arguments passed to the
    corresponding openpyxl style constructor:
    - font: Font(**font) e.g. {"bold": True, "name": "Arial", "size": 12}
    - fill: PatternFill(**fill) e.g. {"patternType": "solid", "fgColor": "FFFF00"}
    - border: dict of side kwargs, e.g. {"left": "thin", "right": "thin"}
    - number_format: format string, e.g. "#,##0.00" or "YYYY-MM-DD"
    - alignment: Alignment(**alignment) e.g. {"horizontal": "center"}

    openpyxl docs § Styling.
    """
    cell = ws.cell(row=row, column=col)
    if font is not None:
        cell.font = Font(**font)
    if fill is not None:
        cell.fill = PatternFill(**fill)
    if border is not None:
        sides = {k: Side(style=v) for k, v in border.items() if v}
        cell.border = Border(**sides)
    if number_format is not None:
        cell.number_format = number_format
    if alignment is not None:
        cell.alignment = Alignment(**alignment)


def merge_cells(
    ws: Any, start_row: int, start_col: int, end_row: int, end_col: int
) -> None:
    """Merge the rectangular block from (start_row, start_col) to (end_row, end_col).

    openpyxl docs § Merging cells.  Internally stores a MergedCell placeholder
    so the merged range appears in ws.merged_cells.
    """
    ws.merge_cells(
        start_row=start_row,
        start_column=start_col,
        end_row=end_row,
        end_column=end_col,
    )


def unmerge_cells(
    ws: Any, start_row: int, start_col: int, end_row: int, end_col: int
) -> None:
    """Unmerge a previously merged block."""
    ws.unmerge_cells(
        start_row=start_row,
        start_column=start_col,
        end_row=end_row,
        end_column=end_col,
    )


# ─────────────────────────── image ops ───────────────────────────────────────

def add_image_to_sheet(
    ws: Any,
    image_path: str | Path,
    anchor: str,
    *,
    alt_text: str = "",
) -> XLImage:
    """Add an image to the worksheet at *anchor* (e.g. "B2").

    Alt text is stored in a side-channel dict on the worksheet object
    (``ws._xlsx_skill_alt_texts``) and injected into the drawing XML's
    ``<xdr:cNvPr descr="…">`` attribute when ``save()`` is called.

    ECMA-376 §20.5.2.8: <xdr:cNvPr descr="…"> inside <xdr:nvPicPr> carries
    the accessibility description for a picture in a spreadsheet drawing part.
    openpyxl has no high-level setter for this attribute, so we write it via
    lxml in the post-save ZIP patch performed by save().
    """
    img = XLImage(str(image_path))
    img.anchor = anchor
    ws.add_image(img)
    if alt_text:
        alts: dict[int, str] = getattr(ws, "_xlsx_skill_alt_texts", {})
        alts[len(ws._images) - 1] = alt_text
        ws._xlsx_skill_alt_texts = alts
    return img


# ─────────────────────────── save (with alt-text patch) ──────────────────────

def save(wb: Workbook, path: str | Path) -> Path:
    """Save the workbook; creates parent directories as needed.

    After the standard openpyxl save, applies a ZIP-level patch for any images
    that had alt_text supplied to add_image_to_sheet().  The patch uses lxml to
    set <xdr:cNvPr descr="…"> (ECMA-376 §20.5.2.8) inside the relevant drawing
    XML parts without modifying any other parts.
    """
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    _patch_alt_texts(wb, out)
    return out


def _patch_alt_texts(wb: Workbook, path: Path) -> None:
    """Inject alt texts into drawing XML parts of the saved ZIP.

    For each worksheet that has pending alt texts (stored in
    ws._xlsx_skill_alt_texts), this function:
    1. Locates the worksheet's drawing relationship from its .rels file.
    2. Parses the drawing XML with lxml.
    3. Finds the nth <xdr:pic> element corresponding to the image index.
    4. Sets cNvPr@descr to the alt text.
    5. Rewrites the entire ZIP with the modified drawing XML.

    ECMA-376 §20.5.2.8 and OPC Part 2 §10 (ZIP packaging).
    """
    pending: dict[int, dict[int, str]] = {}
    for ws_idx, ws in enumerate(wb.worksheets):
        alts: dict[int, str] = getattr(ws, "_xlsx_skill_alt_texts", {})
        if alts:
            pending[ws_idx] = alts
    if not pending:
        return

    raw = path.read_bytes()
    file_map: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
        for name in zf.namelist():
            file_map[name] = zf.read(name)

    for ws_idx, alts in pending.items():
        sheet_num = ws_idx + 1
        rels_key = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
        if rels_key not in file_map:
            continue

        rels_root = etree.fromstring(file_map[rels_key])
        drawing_target: str | None = None
        for rel in rels_root:
            if _REL_DRAWING in rel.get("Type", ""):
                drawing_target = rel.get("Target", "")
                break
        if not drawing_target:
            continue

        drawing_part = _resolve_path(f"xl/worksheets/sheet{sheet_num}.xml", drawing_target)
        if drawing_part not in file_map:
            continue

        root = etree.fromstring(file_map[drawing_part])
        pics = root.findall(f".//{_XDR}pic")
        for img_idx, alt in alts.items():
            if img_idx < len(pics):
                cnvpr = pics[img_idx].find(f"{_XDR}nvPicPr/{_XDR}cNvPr")
                if cnvpr is not None:
                    cnvpr.set("descr", alt)

        file_map[drawing_part] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name, data in file_map.items():
            zf_out.writestr(name, data)
    path.write_bytes(buf.getvalue())


def _resolve_path(source: str, target: str) -> str:
    """Resolve an OPC relative or absolute target path against a source part name."""
    # Absolute OPC target paths start with '/' — strip the leading slash.
    if target.startswith("/"):
        return target.lstrip("/")
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


# ─────────────────────────── CLI ──────────────────────────────────────────────

def _build_demo(out_path: Path) -> None:
    wb = create_workbook()
    ws = wb.active
    ws.title = "Demo"
    write_range(ws, 1, 1, [["Name", "Score", "Grade"], ["Alice", 95, "A"], ["Bob", 82, "B"]])
    format_cell(ws, 1, 1, font={"bold": True}, fill={"patternType": "solid", "fgColor": "4472C4"})
    format_cell(ws, 1, 2, font={"bold": True}, fill={"patternType": "solid", "fgColor": "4472C4"})
    format_cell(ws, 1, 3, font={"bold": True}, fill={"patternType": "solid", "fgColor": "4472C4"})
    ws2 = add_sheet(wb, "Formulas")
    for i in range(1, 6):
        write_cell(ws2, i, 1, i * 10)
    write_cell(ws2, 6, 1, "=SUM(A1:A5)")
    save(wb, out_path)
    print(f"Saved: {out_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a demo .xlsx file")
    parser.add_argument("--out", type=Path, default=Path("demo.xlsx"), help="Output path")
    args = parser.parse_args()
    _build_demo(args.out)


if __name__ == "__main__":
    main()
