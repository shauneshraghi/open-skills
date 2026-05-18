# python-pptx API Reference

Summarized from the python-pptx 1.0.x official documentation
(https://python-pptx.readthedocs.io). Not copied verbatim.

## Core objects

### `Presentation`

```python
from pptx import Presentation

prs = Presentation()               # blank 10×7.5-inch deck
prs = Presentation("file.pptx")   # open existing
prs.save("out.pptx")

prs.slide_width   # Emu; e.g. Inches(10) = 9144000
prs.slide_height  # Emu
prs.slides        # Slides collection
prs.slide_layouts # all layouts across all masters (flattened)
prs.slide_masters # SlideMasters collection
```

`Inches(n)` / `Pt(n)` / `Emu(n)` come from `pptx.util` and convert to EMU
(English Metric Units, 1 inch = 914400 EMU).

### `Slides`

```python
slides = prs.slides
len(slides)                       # count
slide = slides[i]                 # access by index (0-based)
for slide in slides: ...          # iterate
slide = slides.add_slide(layout)  # append new slide
slides.get(sldId)                 # lookup by slide-ID integer
```

**Reordering and deletion** are not exposed by the python-pptx API; they require
direct lxml manipulation of the `<p:sldIdLst>` element (see `edit_pptx.py`).

### `Slide`

```python
slide.shapes          # ShapeCollection
slide.placeholders    # filtered view of shapes with ph element
slide.slide_layout    # SlideMaster → SlideLayout reference
slide.has_notes_slide # bool
slide.notes_slide     # NotesSlide (auto-creates if absent)
slide.name            # optional name attribute
slide.part            # SlidePart (OPC part; gives access to rels)
```

### `ShapeCollection` / shapes

```python
shapes = slide.shapes
shapes.add_textbox(left, top, width, height)            # → TextBox
shapes.add_picture(path, left, top, width, height)      # → Picture
shapes.add_table(rows, cols, left, top, width, height)  # → GraphicFrame
shapes.add_shape(MSO_SHAPE, left, top, width, height)   # → Shape (autoshape)
shapes.add_chart(chart_type, left, top, width, height, chart_data)  # → GraphicFrame
shapes.title          # shortcut to title placeholder or None
```

Shape types (from `pptx.enum.shapes.MSO_SHAPE_TYPE`):
- `AUTO_SHAPE (1)`, `PICTURE (13)`, `PLACEHOLDER (14)`, `TEXT_BOX (17)`, etc.

### `TextFrame` and runs

```python
tf = shape.text_frame
tf.text                    # collapsed plain text (newlines between paras)
tf.text = "replace all"    # replaces all paragraphs
tf.word_wrap = True
para = tf.paragraphs[0]
para.text                  # plain text of paragraph
run = para.add_run()
run.text = "hello"
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0, 0)
para.alignment = PP_ALIGN.CENTER
```

### `Picture` shape

```python
pic = slide.shapes.add_picture(path, left, top, w, h)
pic.image              # Image object (content_type, blob, …)
pic.image.content_type # e.g. "image/png"

# Alt text — no high-level API; use lxml:
pic._element.nvPicPr.cNvPr.set("descr", "alt text string")
# Read back:
pic._element.nvPicPr.cNvPr.get("descr", "")
```

### `Table`

```python
tf = slide.shapes.add_table(rows, cols, left, top, width, height)
tbl = tf.table
cell = tbl.cell(row_idx, col_idx)  # Cell object
cell.text = "value"
cell.text_frame                    # full TextFrame for formatting
tbl.columns[i].width = Inches(2)
tbl.rows[j].height = Inches(0.5)
```

### Speaker Notes

```python
# python-pptx automatically creates the notes slide part if it doesn't exist
if slide.has_notes_slide:
    text = slide.notes_slide.notes_text_frame.text
else:
    pass  # no notes yet

ns = slide.notes_slide  # creates NotesSlidePart on first access if absent
ns.notes_text_frame.text = "My note"

# For formatted notes, use paragraph/run API:
tf = ns.notes_text_frame
tf.clear()
p = tf.paragraphs[0]
run = p.add_run()
run.text = "Bold intro"
run.font.bold = True
```

**How python-pptx creates a notes slide** (relevant for lxml use): calling
`slide.notes_slide` on a slide without an existing notes slide triggers
`NotesSlidePart.new()`. It creates a new XML part from a template that includes
the required `<p:sp>` placeholders (`sldImg` and `body`). The part is linked to
the slide via a relationship of type `notesSlide`. If no notes master exists,
python-pptx also creates one from a built-in template.

### Image replacement

```python
from pptx.oxml.ns import qn

# Find the picture shape
for shape in slide.shapes:
    if shape.shape_type == 13:  # PICTURE
        blip = shape._element.blipFill.blip
        old_rId = blip.get(qn("r:embed"))
        # Add new image part; get its relationship ID
        new_part, new_rId = slide.part.get_or_add_image_part(new_image_path)
        # Swap the reference
        blip.set(qn("r:embed"), new_rId)
```

### Slide layouts and masters

```python
master = prs.slide_masters[0]
master.slide_layouts           # SlideLayouts on this master
layout = prs.slide_layouts[1]  # "Title and Content" (flattened cross-master index)
layout.name                    # e.g. "Title and Content"
layout.placeholders            # placeholder shapes available to slides
```

### OPC relationships (advanced)

Each `Part` has a `rels` dict mapping `rId → _Relationship`.

```python
part = slide.part
for rId, rel in part.rels.items():
    print(rId, rel.reltype, rel.target_part)

new_img_part, rId = part.get_or_add_image_part(path)
part.drop_rel(rId)   # remove a relationship (e.g. when deleting a slide)
```

## Key enumerations

```python
from pptx.enum.text import PP_ALIGN       # LEFT, CENTER, RIGHT, JUSTIFY
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from pptx.enum.chart import XL_CHART_TYPE # BAR_CLUSTERED, LINE, PIE, …
from pptx.enum.dml import MSO_THEME_COLOR
```

## Error handling notes

- `pptx.exc.InvalidXmlError` — malformed XML in the PPTX package
- `KeyError` when accessing `slide.part.rels[rId]` if relationship was dropped
- `ValueError` from `prs.slides.add_slide(layout)` if layout is from a different prs

## Version notes (1.0.x)

- `Presentation.slide_layouts` now returns a flattened list across all masters
- `Slides.__getitem__` and `__iter__` both use `CT_SlideIdList.sldId_lst`
- `slide.notes_slide` auto-creates both the notes slide and notes master parts
