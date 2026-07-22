---
name: pptx-creation-editing
description: >
  Create, edit, inspect, and repair Microsoft PowerPoint `.pptx` files. Use when
  the user wants to generate or modify a PowerPoint deck, add or replace slides,
  text, tables, pictures, and speaker notes, set image alt text, generate or
  execute VBA macros, or work directly with PresentationML package parts such as
  slide XML, notes XML, relationships, and content types.
license: Apache-2.0
compatibility: >
  Requires Python with `python-pptx`, `lxml`, and `pillow`. VBA execution requires
  `POWERPNT.exe` on PATH (Windows or Wine) and a macro-enabled `.pptm` for the
  `/M` switch. COM automation requires `pywin32` (pip install pywin32, Windows
  only) and `Visible = True`. Visual verification (`render_pptx.py`) requires
  LibreOffice (`soffice`) with the Impress component; slide rendering also needs
  Poppler (`pdftoppm`), and slide counting falls back to `pypdf` then `pdfinfo`.
  Apache POI fixtures are optional external test data fetched automatically by the
  comprehensive corpus workflow.
metadata:
  version: "1.3.0"
---

# pptx Creation and Editing Skill

A skill for creating, editing, validating, and regression-testing `.pptx` files.
Built from python-pptx documentation and the ECMA-376 OOXML specification.

## Workflow

1. Open or create the target presentation with `scripts/create_pptx.py` or `scripts/edit_pptx.py`.
2. Use `references/python-pptx.md` for high-level API patterns and `references/ooxml-pptx.md` when editing XML structures, relationships, slide ordering, notes parts, or content types directly.
3. Prefer the smallest possible mutation. Preserve existing PresentationML markup unless the user explicitly wants broad cleanup or repair.
4. Save to a target path and run `scripts/validate_pptx.py` after every non-trivial mutation.
5. **Close the visual feedback loop** with `scripts/render_pptx.py`: render the deck to PDF (one page per slide) to confirm the **slide count**, and pass `--render` to rasterize slides to PNG so you can *see* them — catch text overflowing a box, an image off the slide, or an unreadable color — before delivering. `validate_pptx.py` checks structure; `render_pptx.py` checks appearance. Use `--expect-slides N` (non-zero exit if it misses) when the deck has a target length.
6. For deep regression checks, run `scripts/comprehensive_test.py`. It clones Apache POI, sparse-checks out `test-data/slideshow`, verifies fixtures, runs the corpus workflow, writes structured logs and metadata, and cleans up unless artifacts are retained.

## VBA Workflow

Use VBA when the task requires logic that runs **inside PowerPoint** — dynamic
slide generation, animation sequencing, shape manipulation via code, UserForm
interaction, batch processing across many presentations, or COM-based
inter-application automation. For structural edits that can be done from Python
(adding slides, text, images, notes), prefer `create_pptx.py` / `edit_pptx.py`.

1. **Generate VBA code.** Load the narrowest relevant file from
   `references/vba-docs/` (see the VBA Reference Map below) before writing macro
   code to avoid hallucinating object names or method signatures.
2. **Save and execute.** Use `scripts/vba_runner.py` to save the macro and run it
   via one of three modes:
   - **VBScript bridge** (`generate-vbs`): inject code at runtime via COM —
     no macro pre-embedded in the presentation required.
   - **POWERPNT.exe /M** (`invoke --macro`): open a `.pptm` that already contains
     the named macro and run it via the CLI.
   - **pywin32 COM** (`run-com`): direct COM automation (Windows + pywin32).
3. **Verify.** Run `scripts/validate_pptx.py` on the output presentation.

> PowerPoint has no `Auto_Open` Sub. Use `Private Sub Presentation_Open()` in
> `ThisPresentation` for code that runs when a presentation opens. PowerPoint COM
> requires `Visible = True`; fully headless operation is unreliable.

See `references/powerpnt-cli.md` for the complete POWERPNT.exe switch reference.

### When to use VBA vs python-pptx

| Task | Use |
|---|---|
| Add slides, text boxes, images, tables, notes | python-pptx scripts |
| Animate shapes, set timing sequences | VBA |
| UserForms, custom dialogs, input prompts | VBA |
| Batch process many presentations with PowerPoint's engine | VBA via Presentation_Open |
| COM-based inter-app automation (send to Word, export PDF) | VBA |
| Headless server-side generation | python-pptx scripts |

## VBA Reference Map

Load the narrowest relevant file before writing VBA code.

| Task | Reference file |
|---|---|
| Core VBA concepts for PowerPoint | `references/vba-docs/Concepts/concepts-powerpoint-vba-reference.md` |
| Auto macros (Presentation_Open, Auto_Open equivalents) | `references/vba-docs/Concepts/auto-macros.md` |
| Control and dialog box events | `references/vba-docs/Concepts/control-and-dialog-box-events-powerpoint.md` |
| Language-specific properties and methods | `references/vba-docs/Concepts/language-specific-properties-and-methods-powerpoint.md` |
| OLE programmatic identifiers (COM ProgIDs) | `references/vba-docs/Concepts/ole-programmatic-identifiers-powerpoint.md` |
| Application-level events | `references/vba-docs/How-to/use-events-with-the-application-object.md` |
| Working with shapes and drawing objects | `references/vba-docs/How-to/work-with-shapes-drawing-objects.md` |
| Working with tables on slides | `references/vba-docs/How-to/work-with-tables.md` |
| Working with panes and views | `references/vba-docs/How-to/work-with-panes-and-views.md` |
| Working with partial documents | `references/vba-docs/How-to/work-with-partial-documents.md` |
| Return objects from collections | `references/vba-docs/How-to/return-objects-from-collections.md` |
| ActiveX controls on slides | `references/vba-docs/How-to/use-activex-controls-on-slides.md` |
| ActiveX controls on documents | `references/vba-docs/How-to/use-activex-controls-on-documents.md` |
| UserForms — create | `references/vba-docs/How-to/create-userforms.md` |
| UserForms — display | `references/vba-docs/How-to/display-custom-dialog-boxes.md` |
| Custom dialog boxes | `references/vba-docs/How-to/create-custom-dialog-boxes.md` |
| Add controls to UserForms | `references/vba-docs/How-to/add-controls-to-userforms.md` |
| Set control properties | `references/vba-docs/How-to/set-control-properties.md` |
| Initialize control properties | `references/vba-docs/How-to/initialize-control-properties.md` |
| Use control values while code runs | `references/vba-docs/How-to/use-control-values-while-code-is-running.md` |
| Control another Office app from PowerPoint | `references/vba-docs/How-to/control-one-microsoft-office-application-from-another.md` |

## When To Use Which Script

- `scripts/create_pptx.py`: create new decks, add title/content/blank slides, add images or tables, and set speaker notes.
- `scripts/edit_pptx.py`: open an existing deck, reorder or remove slides, read or update notes, replace images, and inspect alt text.
- `scripts/validate_pptx.py`: verify package integrity, required parts, XML well-formedness, slide counts, and relationship targets.
- `scripts/render_pptx.py`: render the deck to PDF (via LibreOffice `soffice`, one page per slide) to report the slide count, and optionally rasterize slides to PNG (via Poppler `pdftoppm`) for visual inspection — the appearance-level companion to `validate_pptx.py`.
- `scripts/benchmark.py`: measure create/open/save performance.
- `scripts/comprehensive_test.py`: run the Apache POI-backed slideshow regression workflow with automatic baseline acquisition, fixture verification, structured logging, metadata capture, and cleanup.

## Quick Start

```bash
pip install python-pptx lxml pillow

# Create a new presentation
python scripts/create_pptx.py --out my_deck.pptx

# Edit an existing presentation
python scripts/edit_pptx.py set-notes my_deck.pptx 0 "Speaker note for slide 1"

# Validate a presentation
python scripts/validate_pptx.py --file my_deck.pptx

# Run evals
python evals/eval_runner.py

# Run the full Apache POI-backed corpus workflow
python scripts/comprehensive_test.py

# Preserve cloned fixtures and outputs for debugging
python scripts/comprehensive_test.py --keep-artifacts

# Compatibility entrypoint
python evals/corpus_test.py

# Run benchmarks against a local POI checkout if desired
python scripts/benchmark.py --poi-path /path/to/poi/test-data/slideshow
```

For corpus runs, the workflow clones Apache POI automatically and uses `test-data/slideshow` as the baseline source.

## Operational Rules

- Treat image replacement as same-format-only unless you also update package metadata such as content types and media part names.
- When working with notes, prefer targeted edits and validate after save because some files auto-create missing notes parts or normalize notes structures during round-trip.
- For slide reorder or delete operations, preserve relationship integrity and validate the resulting package before returning it.
- If the user gives an ambiguous target like "replace the second logo" or "update the notes on that slide", inspect the deck first and identify the exact slide index, picture index, or placeholder before editing.

## Capabilities

### 1. Presentation Creation (`create_pptx.py`)
- Create new presentations with title, content, and blank slides
- Add text boxes, tables, and images
- Set image alt text via the non-visual properties XML (`descr` on `cNvPr`)
- Add and update speaker notes

### 2. Presentation Editing (`edit_pptx.py`)
- Reorder and remove slides by editing `<p:sldIdLst>` and relationships directly
- Read and write speaker notes
- Replace existing images by swapping the image relationship target
- Read and set accessibility alt text on shapes

### 3. Validation (`validate_pptx.py`)
- Verify the file is a valid ZIP / OOXML package
- Check that required parts exist and identify the presentation part from `[Content_Types].xml`
- Validate XML well-formedness and relationship target resolution
- Count and report slides, notes slides, and images

### 4. Comprehensive Corpus Workflow (`scripts/comprehensive_test.py`)
- Clone `https://github.com/apache/poi.git` into an isolated temp workspace
- Sparse-check out `test-data/slideshow` at `trunk` by default
- Verify required fixtures for existence, size, ZIP integrity, validation health, and SHA-256 hashes
- Run structured regression tests covering slide iteration, notes behavior, notes round-trip, slide reorder stability, multi-image replacement, no-placeholder handling, and validation
- Emit JSON-lines workflow logs plus `run-metadata.json` with baseline commit, fixture hashes, test results, and cleanup status
- Remove clone and output artifacts automatically unless `--keep-artifacts` is used

### 5. VBA Automation (`vba_runner.py`)

Three execution modes for running VBA macros against presentations:

- **VBScript bridge** (`generate-vbs`): generates a `.vbs` COM-automation wrapper that opens the presentation, injects VBA code at runtime, runs the named Sub, saves, and closes. Requires `Visible = True` and Trust Center "Trust access to the VBA project object model".
- **POWERPNT.exe /M** (`invoke --macro`): opens a `.pptm` via CLI and runs the named macro. Syntax: `POWERPNT.exe /M <file> "<MacroName>"`. The file must be macro-enabled.
- **pywin32 COM** (`run-com`): direct COM automation via `win32com.client.Dispatch("PowerPoint.Application")`. Requires `pip install pywin32`, Windows, and `Visible = True`.

Helper utilities: `write_bas_file()` saves VBA code as a `.bas` module file; `build_presentation_open_stub()` generates a `Presentation_Open` Sub for `ThisPresentation`; `check_powerpnt()` verifies POWERPNT.exe is on PATH.

## OOXML Concepts (ECMA-376)

| Feature | XML element | Namespace prefix |
|---|---|---|
| Slide list | `p:sldIdLst` | `p` |
| Slide relationship | `p:sldId r:id` | `p`, `r` |
| Notes slide | `p:notes` | `p` |
| Notes body placeholder | `p:ph type="body"` | `p` |
| Picture | `p:pic` | `p` |
| Image blob ref | `a:blip r:embed` | `a`, `r` |
| Alt text | `p:cNvPr descr=""` | `p` |
| Table | `a:tbl` inside `p:graphicFrame` | `a`, `p` |

## python-pptx Limitations

These features require direct lxml or low-level package manipulation:

- **Slide reordering/deletion**: no high-level python-pptx API for changing `<p:sldIdLst>` order or dropping slide relationships
- **Alt text on pictures**: no dedicated setter for `descr` on non-visual properties
- **Some notes edge cases**: malformed or partially missing notes structures may require XML-level inspection
- **Image replacement**: the relationship target can be swapped, but format-sensitive replacement still needs care

## Test Fixtures (Apache POI corpus)

The comprehensive workflow uses Apache POI slideshow fixtures from `test-data/slideshow/`, including:

- `shapes.pptx`
- `SampleShow.pptx`
- `aascu.org_hbcu_leadershipsummit_cooper_.pptx`
- `table_test2.pptx`
- `bug65551.pptx`
- `testPPT.pptx`
- `layouts.pptx`
- `45545_Comment.pptx`
- `2411-Performance_Up.pptx`
- `WithMaster.pptx`
- `ArtisticEffectSample.pptx`
- `bar-chart.pptx`
- `table-with-theme.pptx`
- `keyframes.pptx`
- `customGeo.pptx`

## Scripts

| Script | Purpose |
|---|---|
| `scripts/create_pptx.py` | Create new presentations, add slides, text, images, tables, and notes |
| `scripts/edit_pptx.py` | Open existing presentations, reorder/remove slides, read/write notes, replace images |
| `scripts/validate_pptx.py` | Validate package structure, XML well-formedness, and relationship targets |
| `scripts/render_pptx.py` | Render to PDF/PNG to verify slide count and visual layout |
| `scripts/benchmark.py` | Performance benchmarks against local POI slideshow fixtures |
| `scripts/comprehensive_test.py` | Deep regression tests using external Apache POI slideshow fixtures |
| `scripts/vba_runner.py` | VBA macro generation and PowerPoint automation helpers (VBScript bridge, POWERPNT.exe /M, pywin32 COM) |

## References

- `references/python-pptx.md` — API summary from official python-pptx docs
- `references/ooxml-pptx.md` — ECMA-376 concepts for slides, notes, images, and relationships
- `references/powerpnt-cli.md` — POWERPNT.exe command-line switch reference; VBA execution patterns (/M, VBScript bridge, pywin32 COM)
- `references/vba-docs/` — curated PowerPoint VBA reference docs; use the VBA Reference Map above to locate the narrowest relevant file before writing macro code

## Evals

- `evals/evals.json` — eval prompt set for common slide, notes, and image workflows
- `evals/eval_viewer.html` — static viewer snapshot for the eval set
- `evals/corpus_test.py` — compatibility wrapper for the comprehensive Apache POI-backed workflow
