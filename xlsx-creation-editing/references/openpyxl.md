# openpyxl API Reference (summarised)

Source: https://openpyxl.readthedocs.io/  
Version: 3.1.x

---

## Workbook

```python
import openpyxl

wb = openpyxl.Workbook()           # new workbook with one sheet
wb = openpyxl.load_workbook(path)  # load existing file
wb = openpyxl.load_workbook(path, data_only=True)  # read cached formula results
wb.save(path)
wb.sheetnames                      # list of sheet title strings
wb.worksheets                      # list of Worksheet objects
wb.active                          # active worksheet
```

### Sheet management

```python
ws = wb.create_sheet(title="Name", index=None)  # index=None → append
del wb["Name"]                     # delete sheet
ws.title = "New Name"              # rename
wb.move_sheet(ws, offset=-1)       # reorder (offset = target − current index)
wb.copy_worksheet(ws)              # duplicate
```

---

## Worksheet

```python
ws = wb.active
ws = wb["Sheet1"]
ws.title         # str
ws.dimensions    # e.g. "A1:D10"
ws.max_row       # int
ws.max_column    # int
```

### Accessing cells

```python
cell = ws["A1"]                     # Cell by coordinate string
cell = ws.cell(row=1, column=1)     # Cell by 1-based index
cell.value                          # read/write

# Iterate rows
for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
    for val in row:
        ...

# Slice
cells = ws["A1:C3"]                 # 2-D tuple of Cell objects
```

---

## Cell values

openpyxl maps SpreadsheetML types to Python types:

| SpreadsheetML type | Python type |
|--------------------|------------|
| n (number)         | int / float |
| s (shared string)  | str |
| b (boolean)        | bool |
| d (date)           | datetime.datetime |
| inlineStr          | str |
| formula (`=…`)     | str (data_only=False) |

Formula strings start with `=` and are preserved as-is when `data_only=False`.

---

## Styles

All style objects are **immutable value objects**; assign a new instance to change.

```python
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Font
cell.font = Font(
    name="Calibri", size=11, bold=False, italic=False,
    underline="none",   # "single", "double"
    color="000000",     # ARGB hex string (no #)
)

# Fill
cell.fill = PatternFill(
    patternType="solid",   # "solid", "gray125", etc.
    fgColor="FFFF00",      # ARGB hex — yellow
    bgColor="000000",
)

# Border
cell.border = Border(
    left=Side(style="thin"),    # None, "thin", "medium", "thick", "dashed", etc.
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Number format
cell.number_format = "#,##0.00"     # Excel format string
cell.number_format = "YYYY-MM-DD"

# Alignment
cell.alignment = Alignment(
    horizontal="center",   # "left", "center", "right", "justify"
    vertical="center",     # "top", "center", "bottom"
    wrap_text=True,
)
```

---

## Merged cells

```python
ws.merge_cells("A1:C3")
ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=3)
ws.unmerge_cells("A1:C3")

# Merged ranges
for rng in ws.merged_cells.ranges:
    print(rng)   # MergedCell object with .min_row, .max_row, etc.
```

---

## Images

```python
from openpyxl.drawing.image import Image

img = Image("path/to/file.png")
img.anchor = "B2"          # top-left cell anchor (string)
img.width = 200            # pixels
img.height = 150
ws.add_image(img)

# Loaded images
ws._images          # list of Image objects (may be BytesIO-based after load)
```

**Alt text is not settable through openpyxl.**  The drawing XML's
`<xdr:cNvPr descr="…">` attribute must be written via lxml after saving
(see `create_xlsx.py` → `_patch_alt_texts()`).

---

## Formulas

Write formula strings starting with `=`:
```python
ws["A6"] = "=SUM(A1:A5)"
ws.cell(row=6, column=1, value="=AVERAGE(B1:B10)")
```

When reading with `data_only=False` (default), formula strings are returned.
When reading with `data_only=True`, the cached result is returned (may be None
if the file has never been opened in Excel).

---

## Validation result

`openpyxl.load_workbook()` does not raise on most non-fatal issues; instead
it emits `UserWarning` messages via Python's warnings module.

---

## Utilities

```python
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

get_column_letter(1)            # "A"
column_index_from_string("A")   # 1
range_boundaries("B2:C3")       # (2, 2, 3, 3) → (min_col, min_row, max_col, max_row)
```
