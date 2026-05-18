---
name: pptx-creation-editing
description: |
  Create, edit, and validate PowerPoint (.pptx) files using python-pptx and direct
  lxml XML manipulation where needed. Covers the three primary use cases:

    1. Slides — add, remove, and reorder slides; set layouts and slide masters;
       bulk-create slides from structured data.
    2. Shapes & images — add text boxes, pictures (with alt text), tables, and charts;
       replace existing images by swapping the embedded relationship target.
    3. Speaker notes — insert, read, and edit per-slide notes; uses the python-pptx
       high-level API where available and direct lxml manipulation for edge cases
       (missing notes master, raw XML access).

  Includes a validation script that checks ZIP structure, required OPC parts, and XML
  well-formedness, plus a benchmark against the Apache POI PPTX test corpus.
license: Apache-2.0
version: "1.0.0"
dependencies:
  python: ">=3.10"
  packages:
    - python-pptx>=1.0.0
    - lxml>=4.9.0
    - pillow>=9.0.0
---

# pptx-creation-editing skill

Open-source (Apache 2.0) skill for creating and editing `.pptx` PowerPoint files.

## Quick start

```bash
pip install python-pptx lxml pillow

# Create a new presentation
python scripts/create_pptx.py --out my_deck.pptx

# Edit an existing presentation
python scripts/edit_pptx.py set-notes my_deck.pptx 0 "Speaker note for slide 1"

# Validate a presentation
python scripts/validate_pptx.py --file my_deck.pptx

# Run evals
POI_PATH=/path/to/poi/test-data/slideshow python evals/eval_runner.py

# Run benchmarks
python scripts/benchmark.py --poi-path /path/to/poi/test-data/slideshow
```

## Feature overview

### Slides

| Operation | API used |
|-----------|----------|
| Add slide with layout | `prs.slides.add_slide(layout)` |
| Iterate slides | `for slide in prs.slides:` |
| Reorder slides | lxml: `sldIdLst.remove(elem); sldIdLst.insert(i, elem)` |
| Delete slide | lxml: `sldIdLst.remove(elem)` + `prs.part.drop_rel(rId)` |
| Get slide count | `len(prs.slides)` |

### Shapes & images

| Operation | API used |
|-----------|----------|
| Add text box | `slide.shapes.add_textbox(left, top, width, height)` |
| Add picture | `slide.shapes.add_picture(path, left, top, width, height)` |
| Set alt text | lxml: `pic._element.nvPicPr.cNvPr.set("descr", alt)` |
| Add table | `slide.shapes.add_table(rows, cols, left, top, w, h)` |
| Replace image | `slide.part.get_or_add_image_part(new_path)` + update `blip@r:embed` |

### Speaker notes

| Operation | API used |
|-----------|----------|
| Create notes slide | `slide.notes_slide` (auto-creates via python-pptx) |
| Set notes text | `slide.notes_slide.notes_text_frame.text = "..."` |
| Read notes text | `slide.notes_slide.notes_text_frame.text` |
| Check existence | `slide.has_notes_slide` |
| Raw XML manipulation | lxml: access `notes._element` → `spTree` → body placeholder |

## lxml vs python-pptx API

python-pptx covers ~80% of common operations. Direct lxml is required for:

- **Slide reordering/deletion**: no python-pptx API; manipulate `<p:sldIdLst>` directly
- **Alt text on pictures**: `descr` attribute on `<p:cNvPr>` inside `<p:nvPicPr>`
- **Notes formatting**: paragraph/run-level XML in the notes text body
- **Replacing images**: update `r:embed` on `<a:blip>` element
- **Notes master absence**: python-pptx auto-creates the notes slide part, but raw XML
  access lets you inspect and repair malformed notes XML from third-party tools

## Test fixtures (Apache POI corpus)

| Fixture | Slides | Notes | Images | Tables | Feature focus |
|---------|--------|-------|--------|--------|---------------|
| `shapes.pptx` | 6 | — | yes | yes | Shape and image operations |
| `SampleShow.pptx` | 2 | 2 | — | — | Notes read/write |
| `aascu.org_hbcu_leadershipsummit_cooper_.pptx` | 16 | 32 | 9 | yes | Full round-trip |
| `bug65551.pptx` | 1 | 1 | — | — | Edge case: single-slide with notes |
| `table_test2.pptx` | 1 | — | — | yes | Table operations |
