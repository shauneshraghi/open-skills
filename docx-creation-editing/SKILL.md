---
name: docx-creation-editing
description: >
  Create and edit Microsoft Word (.docx) files using python-docx and direct lxml
  XML manipulation. Handles document structure, inline/anchored images with alt text,
  comment insertion/reading/deletion, and track-changes (insert, delete, accept/reject).
  Understands the OOXML ZIP structure and namespace conventions from ECMA-376.
license: Apache-2.0
version: 0.1.0
---

# docx Creation and Editing Skill

A Claude Code skill for creating and editing `.docx` files (Office Open XML Word documents).
Built from python-docx official documentation and the ECMA-376 OOXML specification.

## Capabilities

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
| `scripts/create_docx.py` | Create new documents with images |
| `scripts/edit_docx.py` | Edit existing documents (comments, track changes) |
| `scripts/validate_docx.py` | Validate document structure and content |
| `scripts/benchmark.py` | Performance benchmarks |

## Test Fixtures (Apache POI corpus)

| File | Tests |
|---|---|
| `poi/test-data/document/comment.docx` | Reading comments |
| `poi/test-data/document/testComment.docx` | Comment with images, nested structure |
| `poi/test-data/document/VariousPictures.docx` | Inline and anchored images |
| `poi/test-data/document/bug56075-changeTracking_on.docx` | Track changes enabled |

## References

- `references/python-docx.md` — API summary from official python-docx docs
- `references/ooxml-docx.md` — ECMA-376 concepts for images, comments, track changes
