---
name: xlsx-creation-editing
description: >
  Create, edit, inspect, and repair Microsoft Excel `.xlsx` files. Use when the
  user wants to generate or modify an Excel workbook, add or replace
  worksheets, cells, ranges, formulas, formatting, tables, charts, filters,
  statistics, engineering calculations, financial models, images, or alt text,
  generate or execute VBA macros, or work directly with SpreadsheetML package
  parts such as workbook XML, worksheet XML, drawing XML, relationships, and
  content types.
license: Apache-2.0
compatibility: >
  Requires Python with `openpyxl`, `lxml`, and `pillow`. VBA execution requires
  `excel.exe` on PATH (Windows or Wine). COM automation requires `pywin32`
  (pip install pywin32, Windows only). Visual verification (`render_xlsx.py`)
  requires LibreOffice (`soffice`) with the Calc component; page rendering also
  needs Poppler (`pdftoppm`), and page counting falls back to `pypdf` then
  `pdfinfo`. Apache POI spreadsheet fixtures are optional external test data
  fetched automatically by the comprehensive corpus workflow.
metadata:
  version: "1.2.0"
---

# xlsx Creation and Editing Skill

A skill for creating, editing, validating, and inspecting `.xlsx` files for
real workbook tasks such as accounting, finance, engineering, reporting, and
data analysis.
Built from openpyxl documentation, the ECMA-376 OOXML specification, and
curated Excel function references.

## Workflow

1. Open or create the target workbook with `scripts/create_xlsx.py` or `scripts/edit_xlsx.py`.
2. Use `references/openpyxl.md` for high-level API patterns and `references/ooxml-xlsx.md` when editing SpreadsheetML package parts, drawing XML, relationships, or content types directly.
3. When writing native Excel formulas, load the relevant file under `references/functions/` first and verify the exact function name, arguments, and version caveats before writing the formula.
4. Prefer the smallest possible mutation. Preserve existing SpreadsheetML markup unless the user explicitly wants broader cleanup or repair.
5. Save to a target path and run `scripts/validate_xlsx.py` after non-trivial structural edits, package repairs, image changes, or other risky mutations. Validation supports editing work; it is not the primary goal of the skill.
6. **Close the visual feedback loop** when printed/visual layout matters with `scripts/render_xlsx.py`: render the workbook to PDF and pass `--render` to rasterize pages to PNG so you can *see* the result — catch columns cut off at a page edge, a chart that didn't render, or formats that overflow. `validate_xlsx.py` checks structure; `render_xlsx.py` checks appearance. Note: spreadsheet pagination follows each sheet's **print area / page setup / fit-to-page scaling** — set those first (openpyxl `ws.page_setup`) if you care about the printed layout, then render (optionally `--expect-pages N`).
7. For deeper regression checks, use `evals/eval_runner.py` for the small prompt-style suite and `scripts/comprehensive_test.py` for the Apache POI-backed corpus workflow.

## VBA Workflow

Use VBA when the task requires logic that runs **inside Excel** — dynamic cell
population, conditional formatting via code, chart automation, Solver invocation,
UserForm interaction, or batch processing across many workbooks. For structural
edits that can be done from Python (writing cells, formatting, images), prefer
`create_xlsx.py` / `edit_xlsx.py` instead.

1. **Generate VBA code.** Load the narrowest relevant file from
   `references/vba-docs/` (see the VBA Reference Map below) before writing
   macro code to avoid hallucinating object names or method signatures.
2. **Save and execute.** Use `scripts/vba_runner.py` to save the macro and run
   it via one of three modes:
   - **VBScript bridge** (`generate-vbs`): inject code at runtime via COM —
     no macro pre-embedded in the workbook required.
   - **excel.exe CLI** (`invoke`): open a workbook that already contains an
     `Auto_Open` or `Workbook_Open` macro.
   - **pywin32 COM** (`run-com`): direct COM automation (Windows + pywin32).
3. **Verify.** Run `scripts/validate_xlsx.py` on the output workbook.

> Excel has no CLI switch to run a named macro on startup (unlike Word's `/m`).
> Use the VBScript bridge or pywin32 COM for macro execution without embedding.

See `references/excel-cli.md` for the complete excel.exe switch reference.

### When to use VBA vs openpyxl

| Task | Use |
|---|---|
| Write cells, ranges, formatting, images, formulas | openpyxl scripts |
| Solver optimization (SolverAdd, SolverSolve) | VBA — Solver is COM-only |
| UserForms, custom dialogs, input prompts | VBA |
| Chart VBA manipulation, series binding | VBA |
| Conditional formatting via code (complex rules) | VBA |
| Batch processing many workbooks with Excel's calc engine | VBA via Auto_Open |
| Headless server-side generation | openpyxl scripts |

## VBA Reference Map

Load the narrowest relevant file before writing VBA code.

| Task | Reference file |
|---|---|
| Range object fundamentals | `references/vba-docs/Cells-and-Ranges/refer-to-cells-by-using-a-range-object.md` |
| A1 notation and named ranges | `references/vba-docs/Cells-and-Ranges/refer-to-cells-and-ranges-by-using-a1-notation.md`, `refer-to-named-ranges.md` |
| Cell index (row/col) access | `references/vba-docs/Cells-and-Ranges/refer-to-cells-by-using-index-numbers.md` |
| Looping over ranges | `references/vba-docs/Cells-and-Ranges/looping-through-a-range-of-cells.md` |
| Rows and columns | `references/vba-docs/Cells-and-Ranges/refer-to-rows-and-columns.md` |
| Multiple ranges and unions | `references/vba-docs/Cells-and-Ranges/refer-to-multiple-ranges.md` |
| ActiveCell and selection | `references/vba-docs/Cells-and-Ranges/working-with-the-active-cell.md`, `selecting-and-activating-cells.md` |
| 3-D (multi-sheet) ranges | `references/vba-docs/Cells-and-Ranges/working-with-3-d-ranges.md` |
| Formula vs Formula2 | `references/vba-docs/Cells-and-Ranges/range-formula-vs-formula2.md` |
| Cell error constants | `references/vba-docs/Cells-and-Ranges/cell-error-values.md` |
| Create / open workbooks | `references/vba-docs/Workbooks-and-Worksheets/create-a-workbook.md`, `opening-a-workbook.md` |
| Add / replace worksheets | `references/vba-docs/Workbooks-and-Worksheets/create-or-replace-a-worksheet.md` |
| Refer to sheets by name/index | `references/vba-docs/Workbooks-and-Worksheets/refer-to-sheets-by-name.md`, `refer-to-sheets-by-index-number.md` |
| Headers and footers | `references/vba-docs/Workbooks-and-Worksheets/formatting-and-vba-codes-for-headers-and-footers.md` |
| Sort sheets alphabetically | `references/vba-docs/Workbooks-and-Worksheets/sort-worksheets-alphanumerically-by-name.md` |
| Worksheet events | `references/vba-docs/Events-WorksheetFunctions-Shapes/worksheet-object-events.md` |
| Application-level events | `references/vba-docs/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md` |
| Chart object events | `references/vba-docs/Events-WorksheetFunctions-Shapes/chart-object-events.md` |
| Worksheet functions in VBA | `references/vba-docs/Events-WorksheetFunctions-Shapes/using-excel-worksheet-functions-in-visual-basic.md` |
| Available worksheet functions | `references/vba-docs/Events-WorksheetFunctions-Shapes/list-of-worksheet-functions-available-to-visual-basic.md` |
| Shapes and drawing objects | `references/vba-docs/Events-WorksheetFunctions-Shapes/working-with-shapes-drawing-objects.md` |
| Solver functions | `references/vba-docs/Functions/using-the-solver-vba-functions.md` (then load specific SolverXxx files) |
| UserForms and custom dialogs | `references/vba-docs/Controls-DialogBoxes-Forms/create-a-user-form.md`, `create-a-custom-dialog-box.md` |
| ComboBox with unique values | `references/vba-docs/Controls-DialogBoxes-Forms/add-a-unique-list-of-values-to-a-combo-box.md` |
| ActiveX controls on sheets | `references/vba-docs/Controls-DialogBoxes-Forms/using-activex-controls-on-sheets.md` |
| Custom ribbon menu | `references/vba-docs/Controls-DialogBoxes-Forms/create-a-custom-menu-that-calls-a-macro.md` |
| Sparklines | `references/vba-docs/Sparklines/find-all-the-sparklines-on-a-sheet.md` |
| Export to Word | `references/vba-docs/Working-with-Other-Applications/exporting-a-range-to-a-table-in-a-word-document.md` |
| Send email via Outlook | `references/vba-docs/Working-with-Other-Applications/sending-email-to-a-list-of-recipients-using-excel-and-outlook.md` |
| Import Outlook contacts | `references/vba-docs/Working-with-Other-Applications/import-outlook-contacts-to-a-worksheet.md` |
| PowerPivot model | `references/vba-docs/about-the-powerpivot-model-object-in-excel.md` |
| Performance optimization | `references/vba-docs/Excel-Performance/excel-improving-calculation-performance.md` |

## When To Use Which Script

- `scripts/create_xlsx.py`: create new workbooks, add or rename sheets, write cells and ranges, apply formatting, merge cells, add images, and save with post-save alt-text patching.
- `scripts/edit_xlsx.py`: open an existing workbook, read values, inspect image alt text from drawing XML, and replace worksheet images through the Python API.
- `scripts/validate_xlsx.py`: verify package integrity, required workbook parts, XML well-formedness, sheet and image counts, and relationship or content-type integrity after risky edits.
- `scripts/render_xlsx.py`: render the workbook to PDF (via LibreOffice `soffice`) and optionally rasterize pages to PNG (via Poppler `pdftoppm`) to inspect the printed layout — the appearance-level companion to `validate_xlsx.py`. Page count depends on each sheet's print area / page setup / fit-to-page scaling.
- `scripts/benchmark.py`: measure open, iterate, write, and validate timings against a local Apache POI spreadsheet fixture checkout.
- `scripts/comprehensive_test.py`: run the Apache POI-backed spreadsheet regression workflow with automatic baseline acquisition, fixture verification, structured logging, metadata capture, and cleanup.
- `evals/eval_runner.py`: run the small bundled eval set for common workbook, formatting, formula, image, and reorder workflows.
- `evals/corpus_test.py`: compatibility wrapper for the comprehensive Apache POI-backed workflow.

## Formula Authoring Rules

- Use native Excel formulas when they solve the user task more clearly than hard-coded computed values.
- Before writing a formula, consult the matching file under `references/functions/` instead of relying on memory.
- Prefer modern supported function names when compatibility allows; use legacy names only when the workbook or user context requires them.
- Preserve formulas as formulas. Do not replace native workbook calculations with static values unless the user explicitly asks for exported results.
- When building accounting, finance, engineering, or analytics workbooks, favor formulas that keep the workbook auditable by an Excel user after delivery.

## Excel Function References

Use these references to reduce hallucinations in native Excel function names and arguments.
Load the narrowest relevant file before writing formulas.

- `references/functions/financial.md` — financial modeling, loans, bonds, NPV, IRR, depreciation, rates, annuities
- `references/functions/statistical.md` — descriptive statistics, distributions, regression, forecasting, percentiles
- `references/functions/math-trig.md` — arithmetic, rounding, matrix math, trigonometry, sequences, random values
- `references/functions/engineering.md` — unit conversions, complex numbers, bitwise operations, engineering math
- `references/functions/date-time.md` — dates, times, working days, durations, fiscal-period logic
- `references/functions/lookup-reference.md` — XLOOKUP, XMATCH, FILTER, SORT, INDEX/MATCH, dynamic-array reference helpers
- `references/functions/logical.md` — IF, IFS, SWITCH, IFERROR, LET, LAMBDA, MAP, REDUCE, SCAN
- `references/functions/text.md` — text extraction, formatting, joins, cleanup, text-to-number conversion
- `references/functions/database.md` — DSUM, DCOUNT, DAVERAGE, and criteria-range database functions
- `references/functions/information.md` — IS* checks, CELL, TYPE, SHEET/SHEETS, formula introspection
- `references/functions/compatibility.md` — legacy Excel names and their modern replacements
- `references/functions/cube.md` — OLAP / Power Pivot cube functions
- `references/functions/web.md` — WEBSERVICE, FILTERXML, ENCODEURL limitations and patterns

## Quick Start

```bash
pip install openpyxl lxml pillow

# Create a demo workbook
python scripts/create_xlsx.py --out my_workbook.xlsx

# Read all values from a workbook
python scripts/edit_xlsx.py read my_workbook.xlsx

# Validate a workbook
python scripts/validate_xlsx.py --file my_workbook.xlsx

# Run bundled evals
POI_PATH=/path/to/poi/test-data/spreadsheet python evals/eval_runner.py

# Run the full Apache POI-backed corpus workflow
python scripts/comprehensive_test.py

# Preserve cloned fixtures and outputs for debugging
python scripts/comprehensive_test.py --keep-artifacts

# Compatibility entrypoint
python evals/corpus_test.py

# Run quick benchmarks against local POI fixtures
POI_PATH=... python scripts/benchmark.py --quick
```

For corpus runs, the workflow clones Apache POI automatically and uses
`test-data/spreadsheet` as the baseline source.

## Operational Rules

- Treat image replacement as same-format-only unless you also update package metadata such as content types, media part names, and relationship targets.
- Validate after each targeted edit when working on existing workbooks that contain drawings, formulas, or complex formatting.
- If the user gives an ambiguous target like "replace the second logo" or "fix the budget sheet totals", inspect the workbook first and identify the exact sheet, cell range, or image index before editing.
- `scripts/edit_xlsx.py` relies in part on private openpyxl worksheet image state such as `ws._images`; treat image replacement and inspection as version-sensitive operations and prefer conservative edits.
- `scripts/create_xlsx.py` currently patches image alt text after save by editing drawing XML inside the ZIP package. If drawing relationships are missing or malformed, inspect the package before attempting repair.

## Capabilities

### 1. Workbook Creation (`create_xlsx.py`)
- Create new workbooks and add, rename, delete, or reorder worksheets
- Write single cells and 2-D ranges, including formula strings
- Apply `Font`, `PatternFill`, `Border`, `Alignment`, and number formats
- Merge and unmerge rectangular ranges
- Add images and patch accessibility alt text into `<xdr:cNvPr descr="...">`
- Build spreadsheet-native models where formulas remain editable and auditable in Excel

### 2. Workbook Editing (`edit_xlsx.py`)
- Open existing `.xlsx` files and read individual cells or all worksheet values
- Preserve formulas when loaded with `data_only=False`
- Inspect worksheet drawing XML to read image alt text
- Replace worksheet images by index while preserving the original anchor
- Update existing workbooks used for accounting, finance, engineering, dashboards, or analysis without flattening formulas unnecessarily

### 3. Validation (`validate_xlsx.py`)
- Verify the file is a valid ZIP / OOXML package
- Check that required workbook parts exist and identify the workbook part from `[Content_Types].xml`
- Validate XML well-formedness across `.xml` and `.rels` parts
- Count and report worksheets and embedded images
- Check workbook, worksheet, drawing, and media relationship integrity
- Detect broken internal references and content-type mismatches after structural edits

### 4. Benchmark (`benchmark.py`)
- Time open, iterate, write, and validate operations against selected Apache POI workbook fixtures

### 5. Comprehensive Corpus Workflow (`scripts/comprehensive_test.py`)
- Clone `https://github.com/apache/poi.git` into an isolated temp workspace
- Sparse-check out `test-data/spreadsheet` at `trunk` by default
- Verify required fixtures for existence, size, ZIP integrity, validation health, and SHA-256 hashes
- Run structured regression tests covering workbook iteration, value extraction, sheet-count checks, validation, write idempotency, reorder stability, merge round-trip, and image enumeration
- Emit JSON-lines workflow logs plus `run-metadata.json` with baseline commit, fixture hashes, test results, and cleanup status
- Remove clone and output artifacts automatically unless `--keep-artifacts` is used

### 6. VBA Automation (`vba_runner.py`)

Three execution modes for running VBA macros against workbooks:

- **VBScript bridge** (`generate-vbs`): generates a `.vbs` COM-automation wrapper that opens the workbook, injects VBA code at runtime, runs the named Sub, saves, and closes. Requires Windows and Excel Trust Center "Trust access to the VBA project object model".
- **excel.exe CLI** (`invoke`): opens a workbook via excel.exe CLI switches. The workbook must already contain an `Auto_Open` or `Workbook_Open` Sub; Excel fires it automatically on open. See `references/excel-cli.md` for all supported switches.
- **pywin32 COM** (`run-com`): direct COM automation via `win32com.client.Dispatch("Excel.Application")`. Requires `pip install pywin32` and Windows.

Helper utilities: `write_bas_file()` saves VBA code as a `.bas` module file; `build_auto_open_stub()` generates a `Sub Auto_Open()` wrapper; `check_excel()` verifies excel.exe is on PATH.

## Common Workbook Tasks

- Accounting and finance: budgets, forecasts, amortization tables, cash-flow models, KPI sheets, variance analysis
- Engineering: unit conversions, formula-driven calculation sheets, matrix math, lookup-backed design tables
- Data analysis: summaries, filters, statistical calculations, ranking, dynamic-array reports, workbook-native dashboards
- Reporting: formatted worksheets, charts, tables, formulas, and supporting notes intended for Excel users to keep editing

## OOXML Concepts (ECMA-376)

| Feature | XML element | Namespace prefix |
|---|---|---|
| Workbook sheets list | `sheets` / `sheet` | default SpreadsheetML namespace |
| Worksheet cell | `c` | default SpreadsheetML namespace |
| Formula | `f` | default SpreadsheetML namespace |
| Merged range | `mergeCell ref="A1:C3"` | default SpreadsheetML namespace |
| Worksheet drawing relationship | relationship type `.../drawing` | OPC rels namespace |
| Picture | `xdr:pic` | `xdr` |
| Image blob ref | `a:blip r:embed` | `a`, `r` |
| Alt text | `xdr:cNvPr descr=""` | `xdr` |

## openpyxl Limitations

These features require direct XML inspection or package-level handling:

- **Image alt text**: openpyxl does not expose a native setter or getter for `descr` on drawing non-visual properties, so the skill patches or reads drawing XML directly.
- **Image replacement details**: image replacement currently uses private worksheet image state and should be treated as version-sensitive.
- **Relationship and content-type repair**: openpyxl is not a full OOXML repair toolkit, so broken package relationships or content-type mismatches may require direct ZIP/XML inspection.
- **Cached formula values**: openpyxl preserves formulas, but it does not recalculate them; cached values may be stale or absent until Excel recalculates the workbook.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/create_xlsx.py` | Create new workbooks, format cells, add images, and save with alt-text patching |
| `scripts/edit_xlsx.py` | Read workbook values, inspect drawing XML, and replace worksheet images |
| `scripts/validate_xlsx.py` | Validate workbook structure, XML well-formedness, and relationship targets |
| `scripts/render_xlsx.py` | Render to PDF/PNG to verify printed layout (print-area dependent) |
| `scripts/benchmark.py` | Performance benchmarks against local POI spreadsheet fixtures |
| `scripts/comprehensive_test.py` | Optional deep regression tests using external Apache POI spreadsheet fixtures |
| `scripts/vba_runner.py` | VBA macro generation and Excel automation helpers (VBScript bridge, excel.exe CLI, pywin32 COM) |

## Test fixtures (Apache POI corpus)

The comprehensive workflow uses Apache POI spreadsheet fixtures from
`test-data/spreadsheet/`, including:

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

## References

- `references/openpyxl.md` — API summary for workbook, worksheet, cell, style, formula, and image operations
- `references/ooxml-xlsx.md` — SpreadsheetML and DrawingML package structure, relationships, content types, and alt-text XML details
- `references/functions/` — category-based Excel function references used to verify function names, signatures, and compatibility before writing native workbook formulas
- `references/excel-cli.md` — excel.exe command-line switch reference; VBA execution patterns (Auto_Open, VBScript bridge, pywin32 COM)
- `references/vba-docs/` — curated Excel VBA reference docs organized by topic area; use the VBA Reference Map above to locate the narrowest relevant file before writing macro code

## Evals

- `evals/evals.json` — eval prompt set for common workbook, formatting, image, and reorder workflows
- `evals/eval_runner.py` — small bundled runner for the eval prompt set
- `evals/eval_viewer.html` — static viewer snapshot for the eval set
- `evals/corpus_test.py` — compatibility wrapper for the comprehensive Apache POI-backed workflow
