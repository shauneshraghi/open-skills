---
name: docx-creation-editing
description: >
  Create, edit, inspect, and repair Microsoft Word `.docx` files. Use when the
  user wants to generate or modify a Word document, add or replace text, tables,
  or images, set image alt text, insert/read/delete comments, accept or reject
  tracked changes, generate or execute VBA macros, or work directly with OOXML
  package parts such as `document.xml`, `comments.xml`, relationships, and
  content types.
license: Apache-2.0
compatibility: >
  Requires Python with `python-docx` and `lxml`. VBA execution requires
  `winword.exe` on PATH (Windows or Wine). COM automation requires `pywin32`
  (pip install pywin32, Windows only). Visual verification (`render_docx.py`)
  requires LibreOffice (`soffice`); page rendering also needs Poppler
  (`pdftoppm`), and page counting falls back to `pypdf` then `pdfinfo`. Apache
  POI fixtures are optional external test data, not bundled with this skill.
metadata:
  version: "0.4.0"
---

# docx Creation and Editing Skill

A skill for creating and editing `.docx` files (Office Open XML Word documents).
Built from python-docx documentation and the ECMA-376 OOXML specification.

## Working from a User Template or Example

When the user asks for complex formatting — page numbers, headers/footers, merged
table cells, numbered sections, multi-column layout, specific styles — proactively
ask: **"Do you have an example or template document I can match?"**

If the user provides a `.docx`:

1. Run `scripts/inspect_docx.py` to list its parts and extract the relevant OOXML:
   ```bash
   python scripts/inspect_docx.py uploaded.docx --list
   python scripts/inspect_docx.py uploaded.docx word/document.xml word/styles.xml
   python scripts/inspect_docx.py uploaded.docx word/numbering.xml word/header1.xml word/footer1.xml
   ```
2. Match the extracted patterns against `references/ooxml-examples.md` to identify
   what constructs are in use (styles, numbering IDs, sectPr layout, field types).
3. Reproduce those exact patterns in the output document — do not guess at style
   names, numIds, or margin values when the source document has them.

If no template is available, use the curated examples in `references/ooxml-examples.md`
as ground-truth starting points for common patterns.

## Workflow

1. Open or create the target document with `scripts/create_docx.py` or `scripts/edit_docx.py`.
2. Use `references/python-docx.md` for high-level API patterns and `references/ooxml-docx.md` when editing XML structures, relationships, content types, or revision markup directly.
3. For complex formatting (page numbers, headers, tables, numbering), load the relevant section of `references/ooxml-examples.md` before writing lxml code.
4. Prefer the smallest possible mutation. Preserve existing OOXML markup unless the user explicitly wants a broad cleanup such as accepting or rejecting all tracked changes.
5. Save to a target path and run `scripts/validate_docx.py` after every non-trivial mutation.
6. **Close the visual feedback loop** with `scripts/render_docx.py`: render the `.docx` to PDF to confirm the true **page count** and overall layout, and pass `--render` to rasterize pages to PNG so you can *see* the result — catch a document that overflows its target length, a table or figure that breaks awkwardly, or content that blows out a margin — before delivering. `validate_docx.py` checks structure; `render_docx.py` checks appearance. When the user gives a target length, use `--expect-pages N` (non-zero exit if it misses) and iterate.
7. For deep regression checks, use `scripts/comprehensive_test.py` only when the optional external fixture corpus is available.

## VBA Workflow

Use VBA when the task requires logic that runs **inside Word** — conditional
formatting, user prompts, Find/Replace with regex, building block insertion,
content control binding, or batch processing across many documents. For
structural edits that can be done from Python (adding paragraphs, comments,
images), prefer `create_docx.py` / `edit_docx.py` instead.

1. **Generate VBA code.** Load the narrowest relevant file from
   `references/vba-docs/` (see the VBA Reference Map below) before writing
   macro code to avoid hallucinating object names or method signatures.
2. **Save and execute.** Use `scripts/vba_runner.py` to save the macro and run
   it via one of three modes:
   - **VBScript bridge** (`generate-vbs`): inject code at runtime via COM —
     no macro pre-embedded in the document required.
   - **winword.exe /m** (`invoke`): run a named macro already in the .docm
     VBA project via the Word CLI.
   - **pywin32 COM** (`run-com`): direct COM automation (Windows + pywin32).
3. **Verify.** Run `scripts/validate_docx.py` on the output document.

See `references/winword-cli.md` for the complete winword.exe switch reference
and invocation patterns.

### When to use VBA vs python-docx

| Task | Use |
|---|---|
| Structural edits: paragraphs, tables, images, comments, track changes | python-docx / lxml scripts |
| Conditional logic running inside Word, Find/Replace with wildcards | VBA |
| Content controls bound to XML data store | VBA (content control binding API) |
| Batch processing many documents with Word's built-in engine | VBA via winword.exe |
| Headless server-side generation | python-docx / lxml scripts |

## VBA Reference Map

Load the narrowest relevant file before writing VBA code.

| Task | Reference file |
|---|---|
| Document and Application object patterns | `references/vba-docs/Working-with-Word/working-with-document-objects.md` |
| Range and Selection manipulation | `references/vba-docs/Working-with-Word/working-with-range-objects.md`, `working-with-the-selection-object.md` |
| Find and Replace (including wildcards) | `references/vba-docs/Customizing-Word/finding-and-replacing-text-or-formatting.md` |
| Inserting and editing text | `references/vba-docs/Customizing-Word/inserting-text-in-a-document.md`, `editing-text.md` |
| Applying formatting | `references/vba-docs/Customizing-Word/applying-formatting-to-text.md` |
| Tables via VBA | `references/vba-docs/Working-with-Word/working-with-tables.md` |
| Content controls and data binding | `references/vba-docs/Working-with-Word/working-with-content-controls.md`, `references/vba-docs/Objects-Properties-Methods/bind-a-content-control-to-a-node-in-the-data-store.md` |
| Auto macros (AutoOpen, AutoClose, AutoExec) | `references/vba-docs/Customizing-Word/auto-macros.md` |
| Automating common Word tasks | `references/vba-docs/Customizing-Word/automating-common-word-tasks.md` |
| Looping collections | `references/vba-docs/Customizing-Word/looping-through-a-collection.md` |
| ActiveDocument and Application property | `references/vba-docs/Miscellaneous/referring-to-the-active-document-element.md`, `references/vba-docs/Customizing-Word/determining-whether-the-application-property-is-necessary.md` |
| Application-level events | `references/vba-docs/Objects-Properties-Methods/using-events-with-the-application-object-word.md` |
| Document-level events | `references/vba-docs/Objects-Properties-Methods/using-events-with-the-document-object.md` |
| Communicating with other Office apps | `references/vba-docs/Customizing-Word/communicating-with-other-applications.md` |
| UserForms and custom dialogs | `references/vba-docs/Customizing-Word/creating-a-userform.md`, `references/vba-docs/Customizing-Word/creating-a-custom-dialog-box.md` |
| Building blocks | `references/vba-docs/Working-with-Word/working-with-building-blocks.md` |
| Charts via VBA | `references/vba-docs/Working-with-Word/working-with-charts.md` |
| Undorecord (undo groups) | `references/vba-docs/Working-with-Word/working-with-the-undorecord-object.md` |
| Selecting text | `references/vba-docs/Customizing-Word/selecting-text-in-a-document.md` |
| Modifying document sections | `references/vba-docs/Customizing-Word/modifying-a-portion-of-a-document.md` |

## When To Use Which Script

- `scripts/inspect_docx.py`: extract and pretty-print any OOXML part from a user-supplied `.docx` — use this first when matching a template or diagnosing an existing document.
- `scripts/create_docx.py`: create new documents, add headings, tables, inline images, anchored images, or replace images in-place.
- `scripts/edit_docx.py`: open an existing document, read comments, add or remove comments, add tracked changes, or accept/reject tracked changes.
- `scripts/validate_docx.py`: verify package integrity, comment marker consistency, image relationships, and basic revision integrity after edits.
- `scripts/render_docx.py`: render the `.docx` to PDF (via LibreOffice `soffice`) to report the true page count, and optionally rasterize pages to PNG (via Poppler `pdftoppm`) for visual inspection — the appearance-level companion to `validate_docx.py`. Note: Word fields such as a `TOC` field do **not** auto-populate in the LibreOffice render; prefer a manually-built table of contents when you need it visible in the rendered preview.
- `scripts/benchmark.py`: measure create/open/save performance.
- `scripts/comprehensive_test.py`: deep fixture-based regression coverage when Apache POI test data is available.

## Operational Rules

- Treat image replacement as same-format-only unless you also update package metadata such as content types and part names.
- When deleting comments, remove only the comment markers and the `w:comment` body. Do not discard surrounding content runs unless they become empty.
- For existing large or heavily reviewed documents, validate after each targeted edit rather than batching many risky OOXML mutations together.
- If the user gives an ambiguous target like "delete that comment" or "replace the second logo", inspect the document first and identify the exact paragraph, image index, or comment id before editing.

## Capabilities

### 5. VBA Automation (`vba_runner.py`)

- **`check_winword()`** — verify `winword.exe` is on PATH, return its path
- **`write_bas_file(vba_code, path)`** — save VBA code as a `.bas` module file with standard VBA header
- **`build_auto_open_stub(sub_name)`** — generate an `AutoOpen` Sub that calls a named Sub
- **`generate_vbs_runner(doc_path, vba_code, sub_name, out_path)`** — write a `.vbs` COM wrapper that injects the VBA code and runs the Sub (no pre-embedded macro needed)
- **`invoke_winword(doc_path, macro_name, template_path, quiet, disable_addins)`** — invoke `winword.exe` with `/m`, `/t`, `/q`, `/a` switches
- **`run_via_com(doc_path, vba_code, sub_name)`** — inject and run VBA via `win32com` (requires pywin32, Windows only)

### 1. Document Creation (`create_docx.py`)
- Create new documents with headings, paragraphs, tables, and lists
- Add **inline images** via `Document.add_picture()` (python-docx native)
- Add **anchored (floating) images** via direct `wp:anchor` lxml XML manipulation
- Set image alt text via `wp:docPr descr` attribute (ECMA-376 §20.4.2.5)
- Replace existing images by swapping the image part relationship target

### 2. Document Editing (`edit_docx.py`)
- Open existing `.docx` files and modify paragraphs, runs, and styles
- Insert **comments** on runs using `Document.add_comment()` (python-docx 1.x native API)
- Read all comments with author, initials, date, and text
- Delete comments by manipulating `w:commentRangeStart`, `w:commentRangeEnd`,
  `w:commentReference` elements and the `comments.xml` part directly via lxml
- Insert **tracked insertions** (`w:ins`) and **tracked deletions** (`w:del`) via lxml
- Accept all tracked changes (flatten `w:ins` content, remove `w:del` content)
- Reject all tracked changes (remove `w:ins` content, flatten `w:del` content)

### 3. Validation (`validate_docx.py`)
- Verify the file is a valid ZIP / OOXML package
- Check that required parts exist (`word/document.xml`, `[Content_Types].xml`, etc.)
- Validate that all image relationships resolve to present media parts
- Count and report comments, tracked changes, and images

### 4. Benchmark (`benchmark.py`)
- Time create/open/save operations for documents of varying sizes

## OOXML Concepts (ECMA-376)

| Feature | XML element | Namespace prefix |
|---|---|---|
| Paragraph | `w:p` | `w` |
| Run | `w:r` | `w` |
| Inline image | `wp:inline` inside `w:drawing` | `wp` |
| Anchored image | `wp:anchor` inside `w:drawing` | `wp` |
| Image blob ref | `a:blip r:embed` | `a`, `r` |
| Alt text | `wp:docPr descr=""` | `wp` |
| Comment body | `w:comment` in `comments.xml` | `w` |
| Comment anchor start | `w:commentRangeStart w:id` | `w` |
| Comment anchor end | `w:commentRangeEnd w:id` | `w` |
| Comment ref run | `w:commentReference w:id` | `w` |
| Tracked insertion | `w:ins w:id w:author w:date` | `w` |
| Tracked deletion | `w:del w:id w:author w:date` | `w` |
| Deleted run text | `w:delText` (not `w:t`) | `w` |

## python-docx Limitations

These features require **direct lxml manipulation** — python-docx (≤ 1.2) has no native API:

- **Anchored (floating) images** — python-docx only exposes inline shapes
- **Comment deletion** — must remove XML markup triplet + `w:comment` element from part
- **Track changes** — no python-docx API; must construct `w:ins`/`w:del` wrappers via lxml
- **Accept/reject changes** — must walk the XML tree and restructure elements

## Scripts

| Script | Purpose |
|---|---|
| `scripts/inspect_docx.py` | Extract and pretty-print OOXML parts from any `.docx` |
| `scripts/create_docx.py` | Create new documents with images |
| `scripts/edit_docx.py` | Edit existing documents (comments, track changes) |
| `scripts/validate_docx.py` | Validate document structure and content |
| `scripts/render_docx.py` | Render to PDF/PNG to verify page count and visual layout |
| `scripts/vba_runner.py` | VBA code generation, .bas/.vbs file writing, and Word execution via CLI or COM |
| `scripts/benchmark.py` | Performance benchmarks |
| `scripts/comprehensive_test.py` | Optional deep regression tests using external Apache POI fixtures |

## Test Fixtures (Apache POI corpus)

| File | Tests |
|---|---|
| `poi/test-data/document/comment.docx` | Reading comments |
| `poi/test-data/document/testComment.docx` | Comment with images, nested structure |
| `poi/test-data/document/VariousPictures.docx` | Inline and anchored images |
| `poi/test-data/document/bug56075-changeTracking_on.docx` | Track changes enabled |

## References

- `references/ooxml-examples.md` — curated OOXML patterns for page numbers, headers/footers, merged tables, numbering, section properties, and run formatting; load the relevant section before writing lxml code
- `references/python-docx.md` — API summary from official python-docx docs
- `references/ooxml-docx.md` — ECMA-376 concepts for images, comments, track changes
- `references/winword-cli.md` — winword.exe command-line switches and automation invocation patterns
- `references/vba-docs/` — Word VBA reference docs (74 files across 4 directories):
  - `Customizing-Word/` — text editing, formatting, Find/Replace, auto macros, dialogs, loops (30 files)
  - `Working-with-Word/` — Document/Range/Selection/Table/Chart/ContentControl/BuildingBlock objects (9 files)
  - `Objects-Properties-Methods/` — object model, ActiveX, events, content control binding, COM (18 files)
  - `Miscellaneous/` — ActiveDocument patterns, collections, storing values, bookmarks, recording macros (17 files)

## Evals

- `evals/evals.json` — eval prompt set for common image, comment, and track-change workflows
- `evals/eval_viewer.html` — static viewer snapshot for the eval set
