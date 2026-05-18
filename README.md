# open-skills

Open-source (Apache 2.0) agent skills for document creation and editing.

## Getting Started

Install dependencies:

```bash
pip install python-pptx python-docx openpyxl lxml pillow
```

Create demo files:

```bash
python docx-creation-editing/scripts/create_docx.py --out demo.docx
python pptx-creation-editing/scripts/create_pptx.py --out demo.pptx
python xlsx-creation-editing/scripts/create_xlsx.py --out demo.xlsx
```

Validate files:

```bash
python docx-creation-editing/scripts/validate_docx.py --file demo.docx
python pptx-creation-editing/scripts/validate_pptx.py --file demo.pptx
python xlsx-creation-editing/scripts/validate_xlsx.py --file demo.xlsx
```

Run the eval suites (requires POI test corpus):

```bash
POI_PATH=/path/to/poi/test-data/document   python docx-creation-editing/scripts/comprehensive_test.py
POI_PATH=/path/to/poi/test-data/slideshow  python pptx-creation-editing/evals/eval_runner.py
POI_PATH=/path/to/poi/test-data/spreadsheet python xlsx-creation-editing/evals/eval_runner.py
```

## Skills

### `docx-creation-editing`

Create and edit Word (`.docx`) documents.

**Features**
- **Documents** — create, edit, validate, and save `.docx` files
- **Paragraphs** — add, replace, and format text using python-docx and direct XML when needed
- **Tables & images** — build tables, insert pictures, and preserve relationships
- **Validation** — ZIP integrity, required OPC parts, XML well-formedness, broken relationships

See [`docx-creation-editing/`](docx-creation-editing/) for source and [`docx-creation-editing.skill`](docx-creation-editing.skill) for the packaged artifact.

---

### `pptx-creation-editing`

Create, edit, and validate PowerPoint (`.pptx`) files.

**Features**
- **Slides** — add, remove, reorder (lxml `<p:sldIdLst>` manipulation), bulk creation from JSON spec
- **Shapes & images** — text boxes, pictures with alt text, tables; image replacement by swapping OPC relationship targets
- **Speaker notes** — read/write via python-pptx API; direct lxml for multi-paragraph/formatted notes
- **Validation** — ZIP integrity, required OPC parts, XML well-formedness, relationship targets

```bash
# Create a demo deck
python pptx-creation-editing/scripts/create_pptx.py --out demo.pptx

# Edit: set notes on slide 0
python pptx-creation-editing/scripts/edit_pptx.py set-notes deck.pptx 0 "My speaker note"

# Validate
python pptx-creation-editing/scripts/validate_pptx.py --file deck.pptx
```

| Test suite | Result |
|------------|--------|
| Eval tests | 5 / 5 passed |
| POI corpus assertions | 120 / 120 passed |
| Fixtures tested | 15 |

See [`pptx-creation-editing/SKILL.md`](pptx-creation-editing/SKILL.md) for full API reference.

---

### `xlsx-creation-editing`

Create, edit, and validate Excel (`.xlsx`) spreadsheets.

**Features**
- **Workbook & sheets** — create workbooks; add, rename, delete, and reorder worksheets using `wb.move_sheet()` (lxml fallback for edge cases)
- **Cells, ranges & formatting** — write single cells and 2-D ranges; formula strings (starting with `=`); apply Font, PatternFill, Border, Alignment, and number formats; merge/unmerge ranges
- **Images & alt text** — embed images with openpyxl; inject `<xdr:cNvPr descr="…">` alt text via a post-save ZIP patch using lxml (ECMA-376 §20.5.2.8); read alt text back from drawing XML
- **Validation** — ZIP integrity, required OPC parts, SpreadsheetML content type, XML well-formedness, relationship targets

```bash
# Create a demo workbook
python xlsx-creation-editing/scripts/create_xlsx.py --out demo.xlsx

# Read all cell values
python xlsx-creation-editing/scripts/edit_xlsx.py read demo.xlsx

# Validate
python xlsx-creation-editing/scripts/validate_xlsx.py --file demo.xlsx
```

| Test suite | Result |
|------------|--------|
| Eval tests | 5 / 5 passed |
| POI corpus assertions | 120 / 120 passed |
| Fixtures tested | 15 |

See [`xlsx-creation-editing/SKILL.md`](xlsx-creation-editing/SKILL.md) for full API reference.

---

## Structure

```
open-skills/
├── docx-creation-editing/
│   ├── SKILL.md
│   ├── LICENSE.txt
│   ├── references/        # API and OOXML reference docs
│   ├── scripts/           # create_docx.py, edit_docx.py, validate_docx.py, benchmark.py
│   └── evals/             # evals.json, comprehensive_test.py, eval_viewer.html
├── docx-creation-editing.skill
├── pptx-creation-editing/
│   ├── SKILL.md
│   ├── LICENSE.txt
│   ├── references/
│   ├── scripts/           # create_pptx.py, edit_pptx.py, validate_pptx.py, benchmark.py
│   └── evals/             # evals.json, eval_runner.py, corpus_test.py, eval_viewer.html
├── pptx-creation-editing.skill
├── xlsx-creation-editing/
│   ├── SKILL.md
│   ├── LICENSE.txt
│   ├── references/        # openpyxl.md, ooxml-xlsx.md
│   ├── scripts/           # create_xlsx.py, edit_xlsx.py, validate_xlsx.py, benchmark.py
│   └── evals/             # evals.json, eval_runner.py, corpus_test.py, eval_viewer.html
└── xlsx-creation-editing.skill
```

## License

Apache 2.0 — see [LICENSE](LICENSE).
