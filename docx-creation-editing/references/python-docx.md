# python-docx API Reference Summary

Summarized from the python-docx source code (v1.2.0) and official documentation.
Not copied verbatim — paraphrased for skill use.

## Core Document API

```python
from docx import Document
from docx.shared import Inches, Pt, Emu, RGBColor

doc = Document()           # new blank document
doc = Document("file.docx")  # open existing
doc.save("out.docx")
```

### Paragraphs and Runs

```python
para = doc.add_paragraph("text", style="Heading 1")
para = doc.add_heading("Title", level=1)   # shortcut
run  = para.add_run("more text")
run.bold = True
run.italic = True
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0xFF, 0, 0)

# Access existing paragraphs
for p in doc.paragraphs:
    print(p.text)
    for r in p.runs:
        print(r.text, r.bold)
```

### Tables

```python
table = doc.add_table(rows=2, cols=3, style="Table Grid")
cell  = table.cell(0, 0)
cell.text = "Hello"
para  = cell.paragraphs[0]
```

### Styles

```python
doc.styles                    # Styles collection
para.style = doc.styles["Normal"]
para.style = "Heading 2"      # by name string
```

## Images

### Inline Images (python-docx native)

```python
# Width/height are optional — omit to use image's native DPI size
pic = doc.add_picture("path/to/image.png", width=Inches(2))
pic = doc.add_picture(io_stream, width=Inches(2), height=Inches(1))
```

`add_picture()` appends the image to the last paragraph in the document body.
To insert into a specific paragraph:

```python
run = para.add_run()
run.add_picture("image.png", width=Inches(2))
```

This produces a `<w:drawing><wp:inline>...</wp:inline></w:drawing>` structure.
The image binary is stored as a relationship part (`word/media/imageN.ext`).

### Inline Image Alt Text

Alt text lives on the `<wp:docPr descr="...">` element inside `<wp:inline>`.
python-docx does not expose a direct setter, so use lxml:

```python
from docx.oxml.ns import qn

drawing = run._r.find(qn("w:drawing"))
inline  = drawing.find(qn("wp:inline"))
docPr   = inline.find(qn("wp:docPr"))
docPr.set("descr", "my alt text")
```

### Anchored (Floating) Images

python-docx has no native API for anchored images.
Construct a `<wp:anchor>` element directly and inject it into a `<w:drawing>`:

```python
from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import parse_xml
from docx.shared import Emu

# The anchor XML template (simplified; see scripts/create_docx.py for full form)
anchor_xml = f"""<wp:anchor ... >
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="column"><wp:align>left</wp:align></wp:positionH>
  <wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>
  <wp:extent cx="{cx}" cy="{cy}"/>
  <wp:docPr id="{shape_id}" name="{name}" descr="{alt_text}"/>
  <wp:cNvGraphicFramePr/>
  <a:graphic>...</a:graphic>
</wp:anchor>"""
```

### Image Replacement

Replace an existing image by updating the relationship target in the image part:

```python
# Find the rId of the image to replace
blip = drawing.find(".//" + qn("a:blip"))
rid  = blip.get(qn("r:embed"))

# Swap the image bytes in the part's package
image_part = doc.part.related_parts[rid]
image_part._blob = new_image_bytes
```

## Comments (python-docx 1.x native API)

### Insert Comment

```python
# Attach a comment to one or more runs
run     = para.add_run("text to comment on")
comment = doc.add_comment(run, text="My comment", author="Alice", initials="A")

# Or attach to a span of runs
comment = doc.add_comment([run1, run2], text="Spans multiple runs", author="Bob")
```

This inserts `<w:commentRangeStart>`, `<w:commentRangeEnd>`, and a
`<w:commentReference>` run around the target runs, and adds a `<w:comment>`
element to `word/comments.xml`.

### Read Comments

```python
for comment in doc.comments:
    print(comment.comment_id)
    print(comment.author)
    print(comment.initials)
    print(comment.timestamp)   # datetime | None
    print(comment.text)        # plain text, newline-separated paragraphs
```

### Delete a Comment (lxml required)

No native API. Must:
1. Find the `w:commentRangeStart`, `w:commentRangeEnd`, `w:commentReference`
   elements in `document.xml` with matching `w:id`, and remove them.
2. Remove the `w:comment` element from `comments.xml` with matching `w:id`.

```python
from docx.oxml.ns import qn

cid = str(comment.comment_id)
body = doc.element.body

for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
    for el in body.iter(qn(tag)):
        if el.get(qn("w:id")) == cid:
            el.getparent().remove(el)

# commentReference lives inside a run
for el in body.iter(qn("w:commentReference")):
    if el.get(qn("w:id")) == cid:
        run_el = el.getparent()
        # remove the wrapper run if it only contained the reference
        run_el.getparent().remove(run_el)

# Remove from comments part (python-docx 1.x uses _comments_part, a private attr)
try:
    comments_elm = doc.part._comments_part.element
except AttributeError:
    comments_elm = doc.comments._comments_elm
for c in comments_elm.findall(qn("w:comment")):
    if c.get(qn("w:id")) == cid:
        comments_elm.remove(c)
```

## Track Changes (lxml only — no python-docx API)

### Insert a Tracked Insertion

Wrap runs in a `<w:ins>` element:

```python
from lxml import etree
from docx.oxml.ns import qn
import datetime

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ins = etree.SubElement(para._p, f"{{{W}}}ins")
ins.set(f"{{{W}}}id", str(next_revision_id))
ins.set(f"{{{W}}}author", author)
ins.set(f"{{{W}}}date", datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))

run = etree.SubElement(ins, f"{{{W}}}r")
t   = etree.SubElement(run, f"{{{W}}}t")
t.text = "inserted text"
t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
```

### Insert a Tracked Deletion

Wrap runs in a `<w:del>` element; use `<w:delText>` instead of `<w:t>`:

```python
del_el = etree.SubElement(para._p, f"{{{W}}}del")
del_el.set(f"{{{W}}}id", str(next_revision_id))
del_el.set(f"{{{W}}}author", author)
del_el.set(f"{{{W}}}date", datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))

run = etree.SubElement(del_el, f"{{{W}}}r")
dt  = etree.SubElement(run, f"{{{W}}}delText")
dt.text = "deleted text"
dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
```

### Accept All Changes

- Replace each `<w:ins>` with its children (keep the inserted content).
- Remove each `<w:del>` and all its children (discard the deleted content).

### Reject All Changes

- Remove each `<w:ins>` and all its children (discard the inserted content).
- Replace each `<w:del>` with its `<w:r>` children, converting `<w:delText>`
  back to `<w:t>` (restore the deleted content).

## Core Properties

```python
props = doc.core_properties
props.author  = "Alice"
props.title   = "My Document"
props.subject = "Testing"
```

## Sections and Page Layout

```python
section = doc.sections[0]
section.page_width  = Inches(8.5)
section.page_height = Inches(11)
section.left_margin = section.right_margin = Inches(1)
```
