---
name: xlsx-creation-editing
description: |
  Create, edit, and validate Excel (.xlsx) files using openpyxl and direct
  lxml XML manipulation where needed. Covers the three primary use cases:

    1. Workbook & sheets — create workbooks; add, rename, delete, and reorder
       worksheets; write and read cell values including formula strings.
    2. Cells, ranges & formatting — write single cells and 2-D ranges; apply
       Font, PatternFill, Border, Alignment, and number formats; merge and
       unmerge cell ranges.
    3. Images & alt text — embed images with openpyxl.drawing.image.Image;
       inject accessibility alt text into <xdr:cNvPr descr="…"> via a
       post-save ZIP patch using lxml; read alt text back from the drawing XML.

  Includes a validation script that checks ZIP structure, required OPC parts,
  SpreadsheetML content type, and XML well-formedness, plus benchmarks against
  the Apache POI XLSX test corpus.
license: Apache-2.0
version: "1.0.0"
dependencies:
  python: ">=3.10"
  packages:
    - openpyxl>=3.1.0
    - lxml>=4.9.0
    - pillow>=9.0.0
---

# xlsx-creation-editing skill

Open-source (Apache 2.0) skill for creating and editing `.xlsx` Excel files.

## Quick start

```bash
pip install openpyxl lxml pillow

# Create a demo workbook
python scripts/create_xlsx.py --out my_workbook.xlsx

# Read all values from a workbook
python scripts/edit_xlsx.py read my_workbook.xlsx

# Validate a workbook
python scripts/validate_xlsx.py --file my_workbook.xlsx

# Run evals
POI_PATH=/path/to/poi/test-data/spreadsheet python evals/eval_runner.py

# Run corpus tests (120 assertions across 15 fixtures)
POI_PATH=... python evals/corpus_test.py

# Run benchmarks
POI_PATH=... python scripts/benchmark.py --quick
```

## Feature overview

### Workbook & sheet operations

| Operation | API used |
|-----------|----------|
| Create workbook | `openpyxl.Workbook()` |
| Add sheet | `wb.create_sheet(title, index)` |
| Rename sheet | `ws.title = new_name` |
| Delete sheet | `del wb[ws.title]` |
| Reorder sheet | `wb.move_sheet(sheet, offset)` |
| Save | `wb.save(path)` + alt-text ZIP patch |

### Cells, ranges & formatting

| Operation | API used |
|-----------|----------|
| Write cell | `ws.cell(row, column, value=...)` |
| Write range | `ws.cell()` loop over 2-D list |
| Read cell | `ws.cell(row, column).value` |
| Read all values | `ws.iter_rows(values_only=True)` |
| Apply font | `cell.font = Font(bold=True, ...)` |
| Apply fill | `cell.fill = PatternFill(patternType='solid', ...)` |
| Apply border | `cell.border = Border(left=Side(...), ...)` |
| Number format | `cell.number_format = "#,##0.00"` |
| Alignment | `cell.alignment = Alignment(horizontal='center')` |
| Merge cells | `ws.merge_cells(start_row=..., ...)` |
| Unmerge cells | `ws.unmerge_cells(start_row=..., ...)` |

### Images & alt text

| Operation | API used |
|-----------|----------|
| Add image | `openpyxl.drawing.image.Image` + `ws.add_image()` |
| Set alt text | lxml: patch `<xdr:cNvPr descr="…">` post-save |
| Read alt text | lxml: `cnvpr.get("descr", "")` on `<xdr:pic>` element |
| Replace image | swap `ws._images[index]` before save |

## lxml vs openpyxl API

openpyxl covers ~85 % of common operations. Direct lxml is required for:

- **Image alt text**: `descr` attribute on `<xdr:cNvPr>` inside
  `<xdr:nvPicPr>` — no openpyxl setter (ECMA-376 §20.5.2.8)
- **Sheet reorder fallback**: manipulate `<workbook><sheets>` child ordering
  if `wb.move_sheet()` is absent (ECMA-376 §18.2.19)

## Test fixtures (Apache POI corpus)

| Fixture | Sheets | Notes |
|---------|--------|-------|
| `sample.xlsx` | 3 | Basic values and rich text |
| `simple-monthly-budget.xlsx` | 1 | Conditional formatting, merged cells |
| `ExcelTables.xlsx` | 1 | Excel table objects |
| `Formatting.xlsx` | 3 | Cell styles and formatting |
| `ConditionalFormattingSamples.xlsx` | 18 | Conditional formatting, embedded images |
| `123233_charts.xlsx` | 5 | Embedded charts |
| `styles.xlsx` | 3 | Named styles |
| `DateFormatTests.xlsx` | 2 | Date and time formats |
| `AverageTaxRates.xlsx` | 3 | Large dataset |
| `simple-table-named-range.xlsx` | 1 | Named ranges |
| `reordered_sheets.xlsx` | 4 | Sheet reorder stability |
| `Booleans.xlsx` | 3 | Boolean cell values |
| `InlineStrings.xlsx` | 3 | Inline string cells |
| `DataValidations-49244.xlsx` | 1 | Data validation rules |
| `ExcelPivotTableSample.xlsx` | 3 | Pivot table |
