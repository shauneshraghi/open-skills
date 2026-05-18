# OOXML / ECMA-376 Concepts for .docx Files

Relevant concepts from the ECMA-376 4th Edition specification, Part 1 (WordprocessingML).
Focused on images, comments, and revision tracking — the hardest features.

## Package Structure

A `.docx` file is a ZIP archive. Key parts:

```
[Content_Types].xml          — MIME types for each part
_rels/.rels                  — package-level relationships
word/document.xml            — main document story
word/_rels/document.xml.rels — relationships from the document part
word/styles.xml              — paragraph/character/table styles
word/settings.xml            — document settings (track changes flag)
word/comments.xml            — comment bodies (present when comments exist)
word/media/imageN.ext        — embedded image blobs
```

### Content Types

Every part must appear in `[Content_Types].xml`:

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
```

### Relationship Types (ECMA-376 §15.2)

| Short name | Relationship URI |
|---|---|
| `officeDocument` | `.../officeDocument/2006/relationships/officeDocument` |
| `image` | `.../officeDocument/2006/relationships/image` |
| `comments` | `.../officeDocument/2006/relationships/comments` |
| `styles` | `.../officeDocument/2006/relationships/styles` |
| `settings` | `.../officeDocument/2006/relationships/settings` |

## Namespace Prefixes

| Prefix | URI | Usage |
|---|---|---|
| `w` | `.../wordprocessingml/2006/main` | All Word elements |
| `r` | `.../officeDocument/2006/relationships` | Relationship IDs |
| `wp` | `.../drawingml/2006/wordprocessingDrawing` | Image placement |
| `a` | `.../drawingml/2006/main` | DrawingML geometry/blip |
| `pic` | `.../drawingml/2006/picture` | Picture element |
| `xml` | `http://www.w3.org/XML/1998/namespace` | xml:space |

## Images (ECMA-376 §20.4)

### Inline Image (`wp:inline`)

An inline image flows with the text, like a large character.

```xml
<w:r>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="2743200" cy="2057400"/>  <!-- EMUs: 1 inch = 914400 -->
      <wp:docPr id="1" name="Picture 1" descr="alt text here"/>
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks noChangeAspect="1"/>
      </wp:cNvGraphicFramePr>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="image.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="rId5"/>  <!-- relationship ID to word/media/imageN.png -->
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="2057400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
```

**Key attributes:**
- `wp:extent cx/cy` — size in EMUs (English Metric Units). 1 inch = 914,400 EMU.
- `wp:docPr descr` — alt text (accessibility). ECMA-376 §20.4.2.5.
- `a:blip r:embed` — relationship ID linking to the image part.

### Anchored (Floating) Image (`wp:anchor`)

A floating image is positioned absolutely on the page or relative to a margin/column.

```xml
<w:drawing>
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
             simplePos="0" relativeHeight="251658240" behindDoc="0"
             locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column">
      <wp:align>left</wp:align>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:align>top</wp:align>
    </wp:positionV>
    <wp:extent cx="2743200" cy="2057400"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapSquare wrapText="bothSides"/>
    <wp:docPr id="2" name="Picture 2" descr="alt text"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>...</a:graphic>  <!-- same pic:pic structure as inline -->
  </wp:anchor>
</w:drawing>
```

**Key differences from inline:**
- `wp:positionH` / `wp:positionV` — control floating position
- `wp:wrapSquare` (or `wp:wrapThrough`, `wp:wrapNone`, etc.) — text wrapping mode
- `relativeHeight` — z-order (higher = in front)
- `behindDoc="1"` — place behind text

### EMU Conversion

```python
from docx.shared import Inches, Cm, Pt, Emu

cx = Inches(3)      # 2743200 EMU
cy = Cm(5)          # 1814100 EMU
cx = Emu(914400)    # exactly 1 inch
```

### Image Replacement

To swap an image blob in place:
1. Find the `<a:blip r:embed="rId...">` element.
2. Look up the relationship in `word/_rels/document.xml.rels`.
3. Update the part's `_blob` bytes (in python-docx via `part.related_parts[rId]._blob`).

## Comments (ECMA-376 §17.13.4)

Comments span a range of content via three marker elements plus a body in `comments.xml`.

### Markers in document.xml

```xml
<w:commentRangeStart w:id="0"/>  <!-- before the commented content -->
<w:r><w:t>commented text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>    <!-- after the commented content -->
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="0"/>  <!-- reference mark (balloon anchor) -->
</w:r>
```

### Comment body in comments.xml

```xml
<w:comments xmlns:w="...">
  <w:comment w:id="0" w:author="Alice" w:initials="A"
             w:date="2024-01-15T10:30:00Z">
    <w:p>
      <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
      <w:r>
        <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r><w:t>Comment text goes here.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
```

### Comment ID uniqueness

`w:id` must be a non-negative integer unique within the document. The standard
practice is to use `max(existing_ids) + 1`, falling back to gap-filling when
the value would exceed 2^31 − 1.

### Deleting a Comment

Remove all four markers (`commentRangeStart`, `commentRangeEnd`, the
`commentReference` run, and the `w:comment` element in `comments.xml`).
Removing only some of these causes Word to show a corrupt document warning.

## Track Changes / Revisions (ECMA-376 §17.13.5)

Track changes record insertions and deletions as XML wrappers around runs,
preserving both the old and new text simultaneously. A rendering app shows
or hides the changes based on user settings.

### Tracked Insertion (`w:ins`)

```xml
<w:ins w:id="1" w:author="Bob" w:date="2024-01-15T10:30:00Z">
  <w:r>
    <w:t xml:space="preserve">inserted text</w:t>
  </w:r>
</w:ins>
```

The `w:ins` element wraps one or more `w:r` runs. Content appears in the
document if changes are accepted; discarded if rejected.

### Tracked Deletion (`w:del`)

```xml
<w:del w:id="2" w:author="Bob" w:date="2024-01-15T10:30:00Z">
  <w:r>
    <w:delText xml:space="preserve">deleted text</w:delText>
  </w:r>
</w:del>
```

**Critical:** deleted runs use `<w:delText>` not `<w:t>`. Content is discarded
if changes are accepted; restored if rejected.

### Revision ID (`w:id`)

Each `w:ins`/`w:del` carries a `w:id` that must be unique within the document.
Use `max(all existing revision IDs) + 1`.

### Accepting All Changes

```
For each w:ins in body:
    Replace the w:ins element with its w:r children (keep content)
For each w:del in body:
    Remove the w:del element and all its children (discard content)
```

### Rejecting All Changes

```
For each w:ins in body:
    Remove the w:ins element and all its children (discard inserted content)
For each w:del in body:
    For each w:r child: convert w:delText children to w:t
    Replace the w:del element with its (converted) w:r children (restore content)
```

### Enabling Track Changes Mode

Set `<w:trackChanges/>` in `word/settings.xml` to make Word continue tracking
after the document is opened:

```xml
<w:settings>
  <w:trackChanges/>
  ...
</w:settings>
```

## Revision ID Pool

Both comments and track-change elements use `w:id`. They share the same
integer space within a document. When generating new IDs, scan all existing
`w:commentRangeStart`, `w:ins`, `w:del` elements for the maximum used id.
