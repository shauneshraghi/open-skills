# OOXML Examples

Curated, annotated patterns extracted from real Word documents. Each example is
namespace-clean and rsid-stripped. Use these as ground-truth templates — paste
the relevant block into lxml construction code or compare against `inspect_docx.py`
output to understand an uploaded document's structure.

Load only the section(s) relevant to the current task.

---

## Page Numbers

### PAGE field — right-aligned footer (10pt)

Standard `PAGE` field using the three-run fldChar pattern. The cached value (`<w:t>`)
is a display artifact; Word replaces it at render time.

```xml
<!-- word/footer1.xml — single paragraph, right-aligned page number -->
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:jc w:val="right"/>
      <w:rPr>
        <w:sz w:val="20"/>     <!-- 10pt = half-points -->
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:pPr>
    <!-- Run 1: open field -->
    <w:r>
      <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <!-- Run 2: field instruction -->
    <w:r>
      <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText>
    </w:r>
    <!-- Run 3: separator + cached display value -->
    <w:r>
      <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr><w:noProof/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:t>1</w:t>   <!-- cached; Word overwrites at render -->
    </w:r>
    <!-- Run 4: close field -->
    <w:r>
      <w:rPr><w:noProof/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>
```

### PAGE of NUMPAGES — "Page X of Y" footer

Replace the single `PAGE` instruction with two fields separated by literal text.

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Footer"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <!-- "Page " -->
  <w:r><w:t xml:space="preserve">Page </w:t></w:r>
  <!-- PAGE field -->
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <!-- " of " -->
  <w:r><w:t xml:space="preserve"> of </w:t></w:r>
  <!-- NUMPAGES field -->
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:rPr><w:noProof/></w:rPr><w:t>10</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

### Start page numbering at a specific number

Set in `sectPr`. Use `w:pgNumType` to override the start value.

```xml
<w:sectPr>
  <!-- ... headerReference / footerReference ... -->
  <w:pgNumType w:start="1"/>
  <w:pgSz w:w="12240" w:h="15840"/>   <!-- US Letter, portrait -->
</w:sectPr>
```

---

## Headers and Footers

### sectPr — wiring headers and footers

`sectPr` references headers/footers by relationship ID. `type="default"` = odd
pages (and all pages when `<w:titlePg/>` is absent). `type="first"` = first page
only (requires `<w:titlePg/>`). `type="even"` = even pages.

```xml
<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:headerReference w:type="default" r:id="rId5"/>
  <w:headerReference w:type="even"    r:id="rId6"/>
  <w:headerReference w:type="first"   r:id="rId7"/>
  <w:footerReference w:type="default" r:id="rId8"/>
  <w:footerReference w:type="even"    r:id="rId9"/>
  <w:footerReference w:type="first"   r:id="rId10"/>
  <w:titlePg/>   <!-- required for first-page header/footer to activate -->
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
```

### Header with two-column layout (logo left, title right)

Pattern from journals template: table-based header with two cells.

```xml
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tbl>
    <w:tblPr>
      <w:tblW w:w="0" w:type="auto"/>
      <w:tblBorders>
        <w:insideH w:val="single" w:sz="24" w:space="0" w:color="auto"/>
      </w:tblBorders>
      <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"
                 w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="4420"/>
      <w:gridCol w:w="4436"/>
    </w:tblGrid>
    <w:tr>
      <!-- Left cell: author/org name -->
      <w:tc>
        <w:tcPr><w:tcW w:w="4644" w:type="dxa"/></w:tcPr>
        <w:p>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Book Antiqua" w:hAnsi="Book Antiqua"/>
              <w:sz w:val="16"/>
            </w:rPr>
            <w:t>Author Name</w:t>
          </w:r>
        </w:p>
      </w:tc>
      <!-- Right cell: right-aligned journal/document title -->
      <w:tc>
        <w:tcPr><w:tcW w:w="4644" w:type="dxa"/></w:tcPr>
        <w:p>
          <w:pPr><w:jc w:val="right"/></w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Book Antiqua" w:hAnsi="Book Antiqua"/>
              <w:sz w:val="16"/>
            </w:rPr>
            <w:t>Journal Title</w:t>
          </w:r>
        </w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
</w:hdr>
```

### Footer with top border line

Pattern from government report templates: thin rule above footer text.

```xml
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:pBdr>
        <w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/>
      </w:pBdr>
    </w:pPr>
    <w:r><w:t>Document Title — Confidential Draft</w:t></w:r>
  </w:p>
</w:ftr>
```

---

## Tables

### Basic table with column-span (horizontal merge)

`w:gridSpan` merges cells horizontally across N grid columns.

```xml
<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>   <!-- 100% of page width -->
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"
               w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
    <w:gridCol w:w="3000"/>
  </w:tblGrid>

  <!-- Header row: single cell spanning all 3 columns -->
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9000" w:type="dxa"/>
        <w:gridSpan w:val="3"/>
        <w:shd w:val="clear" w:color="auto" w:fill="4472C4"/>  <!-- blue bg -->
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/></w:pPr>
        <w:r>
          <w:rPr><w:color w:val="FFFFFF"/><w:b/></w:rPr>
          <w:t>Section Header</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>

  <!-- Data row: three equal cells -->
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell A</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell B</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
      <w:p><w:r><w:t>Cell C</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

### Vertical merge (rowspan)

`w:vMerge` with no attribute on the first cell, `w:vMerge w:val="restart"` is
wrong — the *first* cell in the span gets `w:vMerge w:val="restart"`, continuation
cells get `<w:vMerge/>` (no val attribute).

```xml
<!-- Row 1: start of 2-row span in column 1 -->
<w:tr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="2000" w:type="dxa"/>
      <w:vMerge w:val="restart"/>
    </w:tcPr>
    <w:p><w:r><w:t>Spans 2 rows</w:t></w:r></w:p>
  </w:tc>
  <w:tc>
    <w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t>Row 1, Col 2</w:t></w:r></w:p>
  </w:tc>
</w:tr>

<!-- Row 2: continuation — MUST include an empty paragraph -->
<w:tr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="2000" w:type="dxa"/>
      <w:vMerge/>   <!-- no val = continuation -->
    </w:tcPr>
    <w:p/>   <!-- required empty paragraph -->
  </w:tc>
  <w:tc>
    <w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t>Row 2, Col 2</w:t></w:r></w:p>
  </w:tc>
</w:tr>
```

### Table with full borders

```xml
<w:tblPr>
  <w:tblW w:w="5000" w:type="pct"/>
  <w:tblBorders>
    <w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  </w:tblBorders>
</w:tblPr>
```

### Cell padding override

```xml
<w:tcPr>
  <w:tcW w:w="3000" w:type="dxa"/>
  <w:tcMar>
    <w:top    w:w="72"  w:type="dxa"/>   <!-- 72 twentieths = 0.05 inch -->
    <w:left   w:w="108" w:type="dxa"/>
    <w:bottom w:w="72"  w:type="dxa"/>
    <w:right  w:w="108" w:type="dxa"/>
  </w:tcMar>
</w:tcPr>
```

---

## Numbering and Lists

### Outline numbered list (I. A. 1.) — multilevel

From government report templates. Attach to Heading styles via `w:pStyle`.
Reference this `abstractNumId` from a `<w:num>` element, then reference the
`numId` from paragraph `<w:numPr>`.

```xml
<!-- word/numbering.xml -->
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="upperRoman"/>
      <w:pStyle w:val="Heading1"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="upperLetter"/>
      <w:pStyle w:val="Heading2"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:firstLine="0"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:pStyle w:val="Heading3"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:firstLine="0"/></w:pPr>
    </w:lvl>
  </w:abstractNum>

  <!-- Concrete num instance — referenced from paragraphs -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>

</w:numbering>
```

Paragraph reference:

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="1"/>
    </w:numPr>
  </w:pPr>
  <w:r><w:t>Section Title</w:t></w:r>
</w:p>
```

### Bullet list (•) — hybridMultilevel

```xml
<w:abstractNum w:abstractNumId="1">
  <w:multiLevelType w:val="hybridMultilevel"/>
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="bullet"/>
    <w:lvlText w:val="&#x2022;"/>   <!-- bullet character -->
    <w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
  </w:lvl>
  <w:lvl w:ilvl="1">
    <w:start w:val="1"/>
    <w:numFmt w:val="bullet"/>
    <w:lvlText w:val="o"/>
    <w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>
  </w:lvl>
</w:abstractNum>
```

### Lettered list (a. b. c.) — single level

```xml
<w:abstractNum w:abstractNumId="2">
  <w:multiLevelType w:val="hybridMultilevel"/>
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="lowerLetter"/>
    <w:lvlText w:val="%1."/>
    <w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="1A1A1A"/></w:rPr>
  </w:lvl>
</w:abstractNum>
```

---

## Section Properties

### US Letter portrait, standard margins

```xml
<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pgSz w:w="12240" w:h="15840"/>   <!-- 8.5 x 11 in @ 1440 twips/inch -->
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720"/>
  <w:docGrid w:linePitch="360"/>
</w:sectPr>
```

### A4 portrait

```xml
<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pgSz w:w="11909" w:h="16834" w:code="9"/>
  <w:pgMar w:top="1843" w:right="1411" w:bottom="1843" w:left="1469"
           w:header="1138" w:footer="837" w:gutter="173"/>
  <w:cols w:space="720"/>
  <w:docGrid w:linePitch="360"/>
</w:sectPr>
```

### Two-column layout (section break)

Use a `sectPr` inside a paragraph's `pPr` to start a new section mid-document.

```xml
<!-- Paragraph that triggers the section change -->
<w:p>
  <w:pPr>
    <w:sectPr>
      <w:cols w:num="2" w:space="720" w:equalWidth="1"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720"/>
      <w:type w:val="continuous"/>   <!-- no page break between sections -->
    </w:sectPr>
  </w:pPr>
</w:p>
```

### Continuous section break (no page break)

```xml
<w:sectPr>
  <w:type w:val="continuous"/>
  <!-- page size and margins... -->
</w:sectPr>
```

### Line numbering (academic / journal style)

From journals template — adds marginal line numbers.

```xml
<w:sectPr>
  <w:lnNumType w:countBy="1" w:restart="continuous"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
```

---

## Paragraph Formatting

### Spacing (before/after in twips, line spacing)

```xml
<w:pPr>
  <w:spacing w:before="240" w:after="120" w:line="276" w:lineRule="auto"/>
  <!-- before=12pt, after=6pt, line=1.15x (276/240) -->
</w:pPr>
```

### Indentation

```xml
<w:pPr>
  <w:ind w:left="720" w:right="720" w:firstLine="720"/>
  <!-- left=0.5in, right=0.5in, firstLine=0.5in hanging indent: use w:hanging -->
</w:pPr>
```

### Keep with next / keep lines together (headings, table headers)

```xml
<w:pPr>
  <w:keepNext/>       <!-- keep with following paragraph -->
  <w:keepLines/>      <!-- don't split this paragraph across pages -->
  <w:pageBreakBefore/> <!-- force page break before this paragraph -->
</w:pPr>
```

---

## Run Formatting

### Common character properties

```xml
<w:rPr>
  <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
  <w:b/>                          <!-- bold -->
  <w:i/>                          <!-- italic -->
  <w:u w:val="single"/>           <!-- underline -->
  <w:strike/>                     <!-- strikethrough -->
  <w:color w:val="FF0000"/>       <!-- red (hex RGB) -->
  <w:highlight w:val="yellow"/>   <!-- highlight -->
  <w:sz w:val="24"/>              <!-- 12pt = 24 half-points -->
  <w:szCs w:val="24"/>            <!-- complex script size -->
  <w:vertAlign w:val="superscript"/>  <!-- or "subscript" -->
</w:rPr>
```

### Shading a run (background color)

```xml
<w:rPr>
  <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>  <!-- yellow bg -->
</w:rPr>
```

---

## Working from inspect_docx.py Output

When the user uploads a .docx for reference, run:

```bash
python scripts/inspect_docx.py uploaded.docx --list
python scripts/inspect_docx.py uploaded.docx word/document.xml word/styles.xml
python scripts/inspect_docx.py uploaded.docx word/numbering.xml word/header1.xml word/footer1.xml
```

Match the output patterns against sections in this file to identify what OOXML
constructs are in use, then reproduce them in `create_docx.py` / `edit_docx.py`
or via direct lxml manipulation.

Key things to check in an uploaded template:
- `sectPr` for page size, margins, columns, header/footer wiring
- `word/styles.xml` for named styles to reference in `w:pStyle`
- `word/numbering.xml` for `abstractNumId` → `numId` mapping
- `word/header*.xml` / `word/footer*.xml` for layout patterns
