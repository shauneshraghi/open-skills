# OOXML PresentationML — Concepts for Shapes, Images, Notes, and Animations

Summarized from ECMA-376 Part 1 (5th edition), DrawingML and PresentationML schemas.
Namespace prefixes used throughout: `p:` = PresentationML, `a:` = DrawingML, `r:` = relationships.

## Package structure (OPC — ECMA-376 Part 2)

A `.pptx` file is a ZIP archive conforming to the Open Packaging Convention.

```
[Content_Types].xml
_rels/.rels
ppt/
  presentation.xml          ← root part (p:presentation)
  _rels/presentation.xml.rels
  slides/
    slide1.xml              ← each slide (p:sld)
    _rels/slide1.xml.rels
  slideLayouts/
    slideLayout1.xml        ← layout (p:sldLayout)
  slideMasters/
    slideMaster1.xml        ← master (p:sldMaster)
  notesSlides/
    notesSlide1.xml         ← notes (p:notes)
    _rels/notesSlide1.xml.rels
  notesMasters/
    notesMaster1.xml        ← notes master (p:notesMaster)
  media/
    image1.png              ← embedded media
  theme/
    theme1.xml              ← DrawingML theme
```

Content types (registered in `[Content_Types].xml`):
- Presentation: `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml`
- Slide: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
- Notes slide: `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml`
- Notes master: `application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml`
- Image (PNG): `image/png`

## Presentation root (`p:presentation`)

```xml
<p:presentation xmlns:p="…" xmlns:r="…">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:notesMasterIdLst>
    <p:notesMasterId r:id="rId2"/>
  </p:notesMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId3"/>   <!-- first slide -->
    <p:sldId id="257" r:id="rId4"/>   <!-- second slide -->
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
```

**Slide ordering** is determined purely by the document order of `<p:sldId>` children
inside `<p:sldIdLst>`. Moving a `<p:sldId>` element to a new position reorders that
slide — no other XML needs to change. (ECMA-376 §19.2.1.38)

**Slide IDs** must be unique integers ≥ 256 within the presentation. The `r:id`
attribute references the slide's OPC relationship in `presentation.xml.rels`.

## Slide (`p:sld`)

```xml
<p:sld xmlns:p="…" xmlns:a="…" xmlns:r="…">
  <p:cSld>                        <!-- common slide data -->
    <p:spTree>                    <!-- shape tree -->
      <p:nvGrpSpPr>…</p:nvGrpSpPr>
      <p:grpSpPr>…</p:grpSpPr>
      <!-- shapes here -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>…</p:clrMapOvr>   <!-- optional color map override -->
</p:sld>
```

## Shapes

### Text box / autoshape (`p:sp`)

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="TextBox 1" descr="alt text"/>  <!-- §20.1.2.2.8 -->
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="914400" y="685800"/>   <!-- position in EMU -->
      <a:ext cx="2743200" cy="457200"/> <!-- size in EMU -->
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:r><a:rPr lang="en-US" b="1"/><a:t>Hello</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

`descr` on `<p:cNvPr>` / `<a:cNvPr>` is the OOXML alt-text field used by screen
readers. The title equivalent is the `title` attribute on the same element.

### Picture (`p:pic`)

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="3" name="Picture 2" descr="A red square"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId5"/>          <!-- rId → media/image1.png -->
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>…</a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

**Image replacement**: change `r:embed` on `<a:blip>` to point to a different
`rId` in the slide's `.rels` file. The old image part is unreferenced but still
present in the ZIP until the file is saved by a conforming producer.

### Table (`p:graphicFrame` + `a:tbl`)

Tables are wrapped in a `<p:graphicFrame>` that contains a `<a:graphic>` element:

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="4" name="Table 3"/>
    …
  </p:nvGraphicFramePr>
  <p:xfrm>…</p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="1828800"/>  <!-- column widths in EMU -->
        </a:tblGrid>
        <a:tr h="457200">          <!-- row height in EMU -->
          <a:tc>
            <a:txBody>…</a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

### Placeholder (`p:ph`)

Placeholders inherit position and formatting from the slide layout. They carry a
`type` and `idx` that must match an entry in the corresponding `<p:sldLayout>`:

```xml
<p:nvSpPr>
  <p:nvPr>
    <p:ph type="title"/>         <!-- title placeholder -->
    <p:ph type="body" idx="1"/> <!-- body/content placeholder -->
    <p:ph type="sldImg"/>       <!-- slide image (notes slide only) -->
  </p:nvPr>
</p:nvSpPr>
```

## Notes slide (`p:notes`)

```xml
<p:notes xmlns:p="…" xmlns:a="…" xmlns:r="…">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>…</p:nvGrpSpPr>
      <p:grpSpPr>…</p:grpSpPr>

      <!-- Slide thumbnail placeholder (required) -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image Placeholder 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="sldImg"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>

      <!-- Notes text body placeholder (required) -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr><a:normAutofit/></a:bodyPr>
          <a:lstStyle/>
          <a:p><a:r><a:t>Speaker note text here</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>
```

**Relationships** (in `notesSlide1.xml.rels`):
```xml
<Relationships>
  <Relationship Id="rId1"
    Type=".../notesMaster"  Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2"
    Type=".../slide"        Target="../slides/slide1.xml"/>
</Relationships>
```

The slide also has a corresponding relationship back to the notes slide:
```xml
<!-- In slide1.xml.rels -->
<Relationship Id="rId6"
  Type=".../notesSlide"  Target="../notesSlides/notesSlide1.xml"/>
```

## Animations (`p:timing`)

Basic animation structure on a slide (informational; not implemented in scripts):

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">
        <p:childTnLst>
          <!-- animate shape with id="3" using a fly-in entrance -->
          <p:par>
            <p:cTn id="2" fill="hold" nodeType="mainSeq">…</p:cTn>
          </p:par>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>
```

Animations are currently read-only in python-pptx (no write API); access via
`slide._element.timing` if present.

## EMU (English Metric Units)

| Unit | EMU value |
|------|----------|
| 1 inch | 914400 |
| 1 cm | 360000 |
| 1 pt | 12700 |
| 1 mm | 36000 |

Standard slide dimensions (widescreen 16:9): `cx=9144000, cy=5143500` (10×5.63 in).

## Namespace URIs

| Prefix | URI |
|--------|-----|
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` |
