# ECMA-376 SpreadsheetML & DrawingML Reference (xlsx)

Source: ECMA-376 5th Edition Part 1 (SpreadsheetML) and Part 1 Chapter 20 (DrawingML)

---

## SpreadsheetML package structure (ECMA-376 Part 1 §12)

A conforming `.xlsx` OPC package contains at minimum:

```
[Content_Types].xml
_rels/.rels
xl/workbook.xml                        (workbook part)
xl/_rels/workbook.xml.rels
xl/worksheets/sheet{n}.xml             (one per sheet)
xl/worksheets/_rels/sheet{n}.xml.rels
xl/sharedStrings.xml                   (optional; shared string table)
xl/styles.xml                          (styles part)
xl/theme/theme1.xml                    (optional)
xl/media/image{n}.{ext}               (embedded media)
xl/drawings/drawing{n}.xml            (worksheet drawings)
xl/drawings/_rels/drawing{n}.xml.rels
```

---

## Workbook element (§18.2.27)

```xml
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <!-- §18.2.19: document order of <sheet> children defines tab order -->
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
```

**§18.2.19 `<sheets>`**: The order of `<sheet>` child elements is normative for
the user-visible sheet tab order.  Reordering sheets means reordering these
children, which is what `wb.move_sheet()` and the lxml fallback in
`create_xlsx.py` do.

---

## Worksheet element (§18.3.1.99)

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" t="s"><v>0</v></c>        <!-- shared string index 0 -->
      <c r="B1" t="n"><v>42</v></c>       <!-- number -->
      <c r="C1" t="b"><v>1</v></c>        <!-- boolean true -->
      <c r="D1"><f>SUM(A1:C1)</f></c>     <!-- formula; <v> holds cached result -->
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A2:C3"/>             <!-- §18.3.1.55 -->
  </mergeCells>
</worksheet>
```

### Cell types (§18.18.11)

| `t` attribute | Python type returned by openpyxl |
|---------------|----------------------------------|
| `n` or absent | int / float |
| `s` | str (via shared strings) |
| `b` | bool |
| `d` | datetime |
| `inlineStr` | str |
| formula (f element) | str starting with `=` (data_only=False) |

---

## Styles part (§18.8)

`xl/styles.xml` contains indexed arrays of font, fill, border, and alignment
definitions.  Each cell references a style index via `s=` attribute.
openpyxl's `Font`, `PatternFill`, `Border`, and `Alignment` objects correspond
directly to these definitions.

---

## SpreadsheetML Drawing part (§20.5)

Worksheet images live in a separate drawing part
(`xl/drawings/drawing{n}.xml`) referenced from the worksheet via an OPC
relationship of type:

```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing
```

### Drawing namespace

```
xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
```

### Image (picture) anchor structure (§20.5.2)

```xml
<xdr:wsDr xmlns:xdr="...spreadsheetDrawing">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>     <!-- 0-based column index -->
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>     <!-- 0-based row index -->
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="…" cy="…"/>   <!-- size in EMU -->
    <xdr:pic>
      <xdr:nvPicPr>
        <!-- §20.5.2.8: cNvPr carries id, name, and accessibility description -->
        <xdr:cNvPr id="1" name="Image 1" descr="Alt text here"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>    <!-- relationship id pointing to xl/media/ -->
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"/>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
```

### §20.5.2.8 `<xdr:cNvPr>` — non-visual drawing properties

Relevant attributes:

| Attribute | Type | Description |
|-----------|------|-------------|
| `id` | uint | unique drawing object id |
| `name` | string | display name |
| `descr` | string | **accessibility alt text** |
| `hidden` | boolean | whether the shape is hidden |

openpyxl sets `descr="Picture"` by default.  The `xlsx-creation-editing`
skill overwrites this via a post-save ZIP patch using lxml (see
`create_xlsx.py` → `_patch_alt_texts()`).

---

## Content types (ECMA-376 Part 2 §15)

The `[Content_Types].xml` part lists the ContentType for every part.  For
SpreadsheetML workbooks the workbook part uses:

| File extension | ContentType |
|----------------|------------|
| `.xlsx` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| `.xlsm` | `application/vnd.ms-excel.sheet.macroEnabled.main+xml` |
| `.xltx` | `application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml` |

---

## OPC relationships (ECMA-376 Part 2 §10)

Relationship target paths may be absolute (`/xl/drawings/drawing1.xml`) or
relative (`../drawings/drawing1.xml`).  The `_resolve_path()` helper in
`create_xlsx.py` and `edit_xlsx.py` handles both forms.
