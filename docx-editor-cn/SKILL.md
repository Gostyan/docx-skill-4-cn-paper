---
name: docx-editor-cn
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of \"Word doc\", \"word document\", \".docx\", or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a \"report\", \"memo\", \"letter\", \"template\", or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
license: Proprietary. LICENSE.txt has complete terms
---

# DOCX creation, editing, and analysis

## Overview

A .docx file is a ZIP archive containing XML files.

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` or unpack for raw XML |
| Create new document | `node scripts/new_doc.js` (edit CONTENT section first) |
| Edit existing document | Unpack → edit XML → repack - see Editing Existing Documents below |
| Insert 三线表 (XML editing) | `python scripts/table.py unpacked/ "1-1" "标题" --headers … --rows …` |
| Insert block formula (XML editing) | `python scripts/formula.py unpacked/ "LaTeX" 1 --anchor "锚文本"` |

### Converting .doc to .docx

Legacy `.doc` files must be converted before editing:

```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content

```bash
# Text extraction with tracked changes
pandoc --track-changes=all document.docx -o output.md

# Raw XML access
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes

To produce a clean document with all tracked changes accepted (requires LibreOffice):

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## Creating New Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Setup
```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation
After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.
```bash
python scripts/office/validate.py doc.docx
```

### Page Size

```javascript
// Chinese academic papers use A4 with 2.5cm margins on all sides
sections: [{
  properties: {
    page: {
      size: {
        width: 11906,   // A4 width in DXA
        height: 16838   // A4 height in DXA
      },
      margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } // 2.5cm = 1418 DXA
    }
  },
  children: [/* content */]
}]
```

**Common page sizes (DXA units, 1440 DXA = 1 inch, 567 DXA = 1 cm):**

| Paper | Width | Height | Content Width (2.5cm margins) |
|-------|-------|--------|-------------------------------|
| A4 (Chinese standard) | 11,906 | 16,838 | 9,070 |
| US Letter | 12,240 | 15,840 | 9,404 |

**Landscape orientation:** docx-js swaps width/height internally, so pass portrait dimensions and let it handle the swap:
```javascript
size: {
  width: 12240,   // Pass SHORT edge as width
  height: 15840,  // Pass LONG edge as height
  orientation: PageOrientation.LANDSCAPE  // docx-js swaps them in the XML
},
// Content width = 15840 - left margin - right margin (uses the long edge)
```

### Styles (Academic Thesis Formatting)

**CRITICAL**: When generating academic papers, course designs, or mathematical documents, strictly adhere to the following Chinese academic formatting standards.

**1. Global Page & Normal Text (【全局页面与正文规范】)**
- **Page**: A4 (11906 × 16838 DXA), margins all 2.5cm (`1418` DXA).
- **Fonts**: Chinese = SimSun (宋体), English/Math/Code = Cambria Math (default); Times New Roman is an acceptable alternative for English text.
- **Size**: 12pt (小四) -> `size: 24` in docx-js (half-points).
- **Paragraph**: First-line indent 2 characters (`firstLine: 480` DXA), Line spacing **single** (`line: 240, lineRule: LineRuleType.AUTO`) by default; alternatives: fixed 20pt (`line: 400, lineRule: LineRuleType.EXACT`) or 1.5x (`line: 360, lineRule: LineRuleType.AUTO`). Before/After spacing 0pt.

**2. Headings (【标题规范】)**
- **Heading 1 (一级标题)**: SimHei (黑体), 16pt (三号, `size: 32`), Bold, Centered, 1.2x line spacing (`line: 288`). Auto-numbering: Arabic `1`, `2`, `3` (displayed in heading text or prefixed manually as 一、二、). See Heading & Reference Auto-Numbering below.
- **Heading 2 (二级标题)**: SimHei (黑体), 14pt (四号, `size: 28`), Bold, Left-aligned, 1.5x line spacing (`line: 360`). Auto-numbering: `1.1`, `1.2`, `2.1` (chapter-tracking Arabic decimal).
- **Heading 3 (三级标题)**: SimHei (黑体), 12pt (小四, `size: 24`), Bold, Left-aligned, 1.1x line spacing (`line: 264`). Auto-numbering: `1.1.1`, `2.4.1`.

**⚠️ Heading Numbering Rule**: Use a **single all-decimal multi-level numbering config** for all three heading levels. This is the ONLY way to get correct cross-chapter tracking (1.1, 2.1, 2.4 etc.) while keeping H2/H3 numbers purely Arabic. If CHINESE_COUNTING is used for H1 level, H2 will render as `二.4` instead of `2.4` — this is a known OOXML rendering issue.

**3. Figures, Tables & Math (【图表与公式规范】)**
- **Figure Captions (图标题)**: SimSun (宋体), 11pt (`size: 22`), Centered, Bold. Spacing: Single spacing, 0.5 lines before (`before: 120`), 3pt after (`after: 60`). Position: Below the figure. Format: "图 章-图序" e.g. "图 1-1".
- **Table Captions (表标题)**: SimSun (宋体), 11pt (`size: 22`), Centered, **no indent**, Bold. Spacing: Single spacing, 0.5 lines before (`before: 120`), 3pt after (`after: 60`). Position: Above the table. Format: "表 章-表序" e.g. "表 1-1". Tables MUST use **三线表 (three-line table)** style (thick top/bottom, thin after header, no other borders).
- **Block Math Formulas**: Use the formula table layout — a 3-column borderless table: [1cm spacer | formula centered | equation number right-aligned 1cm]. Use `scripts/formula.py` to generate the XML. Inline math: write as plain text with italic formatting.

**Creating a new document — use the template script:**
```bash
# Copy the template, edit the CONTENT section, then run:
node scripts/new_doc.js
# Outputs output.docx with all styles, numbering, and page setup pre-configured.
# Change OUTPUT_PATH inside the script if needed.
```

`scripts/new_doc.js` provides ready-to-use helper functions: `h1/h2/h3(text)`, `body(text)`, `tableCaption/figCaption(label)`, `threeLineTable(headers, rows, colWidths)`, `ref(text)`, `blank()`. The STYLES / NUMBERING / page constants at the top of the file are the canonical implementation of all specs below; refer to that file for exact docx-js values.

### Lists (NEVER use unicode bullets)

```javascript
// ❌ WRONG - never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })  // BAD
new Paragraph({ children: [new TextRun("\u2022 Item")] })  // BAD

// ✅ CORRECT - use numbering config with LevelFormat.BULLET
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Numbered item")] }),
    ]
  }]
});

// ⚠️ Each reference creates INDEPENDENT numbering
// Same reference = continues (1,2,3 then 4,5,6)
// Different reference = restarts (1,2,3 then 1,2,3)
```

### Tables (三线表 / Three-Line Table)

**Chinese academic papers MUST use 三线表**: thick top/bottom (1.5pt), thin after header (0.75pt), no other borders. Caption goes **above** the table.

**For new documents — use `threeLineTable()` helper in `scripts/new_doc.js`:**
```javascript
// Inside the CONTENT array:
tableCaption('表 1-1 符号说明'),
threeLineTable(
  ['符号', '说明'],                       // headers
  [['S', '状态空间'], ['A', '动作空间']], // data rows
  [1800, 7270]                            // DXA column widths; must sum to 9070
),
```

**For editing existing documents (XML) — use `scripts/table.py`:**
```bash
# Insert after paragraph containing anchor text:
python scripts/table.py unpacked/ "1-1" "符号说明" \\
    --headers "符号,说明" \\
    --rows '[["S","状态空间"],["A","动作空间"]]' \\
    --anchor "以下是符号说明"

# Custom column widths (must sum to 9070):
python scripts/table.py unpacked/ "2-1" "参数" \\
    --headers "参数,取值,说明" \\
    --rows '[["lr","0.001","学习率"]]' \\
    --widths "1800,1500,5770"

# Print XML snippet only (no file modification):
python scripts/table.py --caption "表 1-1 示例" \\
    --headers "列1,列2" --rows '[["a","b"]]'
```

**Key rules:**
- Content width = 9070 DXA (A4 2.5cm margins); `columnWidths` must sum exactly to 9070
- Always use `WidthType.DXA` — never `WidthType.PERCENTAGE`
- Table width = sum of `columnWidths`; set matching `width` on each cell too

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  children: [new ImageRun({
    type: "png", // Required: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" } // All three required
  })]
})
```

### Page Breaks

```javascript
// CRITICAL: PageBreak must be inside a Paragraph
new Paragraph({ children: [new PageBreak()] })

// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Table of Contents

```javascript
// The overridden Heading1/2/3 styles (defined in paragraphStyles above) include outlineLevel,
// which TableOfContents uses to build the TOC. Using heading: HeadingLevel.HEADING_X on
// each paragraph applies that overridden style automatically — giving both TOC support AND
// the custom SimHei/size formatting. Do NOT skip paragraphStyles and rely on built-in styles.
new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-3" })

// Apply to heading paragraphs — heading property applies the overridden style:
new Paragraph({
  heading: HeadingLevel.HEADING_1, // → w:pStyle "Heading1" (SimHei 16pt, outlineLevel:0)
  children: [new TextRun("一、引言")]
})
new Paragraph({
  heading: HeadingLevel.HEADING_2, // → w:pStyle "Heading2" (SimHei 14pt, outlineLevel:1)
  children: [new TextRun("1.1 研究背景")]
})
```

### Headers/Footers

```javascript
sections: [{
  properties: {
    page: { margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } } // 2.5cm = 1418 DXA
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Heading & Reference Auto-Numbering

The numbering config is pre-built in `scripts/new_doc.js`. Use the `h1/h2/h3()` helpers — numbering is applied automatically.

**CRITICAL — numbering format rule:**

| Scheme | H1 | H2 | H3 | Chapter tracking |
|--------|----|----|-----|------------------|
| ✅ All-decimal (used in `new_doc.js`) | `1` `2` `3` | `1.1` `2.4` | `1.1.1` | ✔ automatic |
| ❌ Mixed Chinese+Arabic | `一` `二` | `一.1` `二.4` ← **broken** | `二.4.1` ← **broken** | ✘ |

NEVER use `LevelFormat.CHINESE_COUNTING` for H1 in a multi-level config — `%1` in H2 text expands to "二", giving `二.4`. Always use `LevelFormat.DECIMAL` for every level.

To display `一、二、三` on H1, write the Chinese character **in the paragraph text** and call `h1()` without a numbering override. H2/H3 auto-number then shows `1`, `1.1` etc. independently per chapter (not cross-chapter `2.4`). Use Option A (`h1('引言')`) if full cross-chapter tracking like `2.4` is needed.

```javascript
// Option A — full auto-numbering (H1: 1 2 3, H2: 1.1 2.4, H3: 1.1.1)
h1('引言')         // → numbered "1"
h2('研究背景')     // → numbered "1.1"
h3('研究现状')     // → numbered "1.1.1"
h1('方法')         // → numbered "2"
h2('算法设计')     // → numbered "2.1"

// Option B — Chinese H1 text, Arabic H2/H3 (only if cross-chapter prefix not needed)
new Paragraph({ heading: HeadingLevel.HEADING_1, indent: { firstLine: 0 },
  children: [new TextRun('一、引言')] })  // no numbering on H1
h2('研究背景')   // → 1  (resets each chapter, no "1.1")
h3('研究现状')   // → 1.1
```

### Block Formula Layout

Block formulas use a **3-column borderless table** so the equation number can be right-aligned while the formula is centered. Use `scripts/formula.py` to generate and insert formula blocks automatically.

```bash
# Generate a formula block and insert after paragraph containing anchor text:
python scripts/formula.py unpacked/ "Q_n(x,a) = r + \\gamma \\max_{a'} Q_{n-1}" 1 --anchor "由此可得"

# Output XML snippet only (no insertion):
python scripts/formula.py --latex "E=mc^2" --number 2
```

**Layout (A4, 9070 DXA content width):**
```
┌──────────┬──────────────────────────┬──────────┐
│  1cm gap │   formula   (centered)   │   (n)    │
│ 567 DXA  │        7936 DXA          │ 567 DXA  │
└──────────┴──────────────────────────┴──────────┘
  no border throughout; single line spacing
```

If direct XML editing is preferred, the skeleton is:
```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9070" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="none"/><w:left w:val="none"/>
      <w:bottom w:val="none"/><w:right w:val="none"/>
      <w:insideH w:val="none"/><w:insideV w:val="none"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="567"/><w:gridCol w:w="7936"/><w:gridCol w:w="567"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr>
      <w:p/></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="7936" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="center"/>
        <w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <!-- OMML formula here -->
        </m:oMath>
      </w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/>
        <w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
        <w:r><w:t>(%NUMBER%)</w:t></w:r>
      </w:p></w:tc>
  </w:tr>
</w:tbl>
```

### References (GB/T 7714-2015)

Use the `"references"` numbering reference defined above. Format text manually per GB/T 7714-2015 entry type.

```javascript
// Reference list style (add to paragraphStyles)
{ id: "Reference", name: "Reference", basedOn: "Normal",
  run: { font: { ascii: "Cambria Math", eastAsia: "SimSun" }, size: 24 },
  paragraph: { spacing: { line: 240, lineRule: LineRuleType.AUTO },
               indent: { left: 480, hanging: 480 } } }

// Usage — one Paragraph per entry:
new Paragraph({
  style: "Reference",
  numbering: { reference: "references", level: 0 },
  children: [new TextRun("Author. "),
             new TextRun({ text: "Title", italics: true }),
             new TextRun("[J]. Journal, Year, Vol(Issue): Pages.")]
})
```

**GB/T 7714-2015 common format templates (fill in fields):**

| Type | Tag | Format |
|------|-----|--------|
| 期刊论文 | [J] | 作者. 题名[J]. 刊名, 年, 卷(期): 起止页. |
| 专著 | [M] | 作者. 书名[M]. 出版地: 出版者, 出版年: 起止页. |
| 学位论文 | [D] | 作者. 题名[D]. 保存地: 保存单位, 年份. |
| 会议论文 | [C] | 作者. 题名[C]//会议名. 出版地: 出版者, 年: 起止页. |
| 网络资源 | [EB/OL] | 作者. 题名[EB/OL]. (发布日期)[引用日期]. URL. |

### Footer (Page Numbers)

Center-aligned page number in the footer. Uses A4 2.5cm margins.

```javascript
const { Footer, Paragraph, TextRun, PageNumber, AlignmentType } = require('docx');

sections: [{
  properties: {
    page: {
      size: { width: 11906, height: 16838 },           // A4
      margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } // 2.5cm
    }
  },
  footers: {
    default: new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT] })]
        })
      ]
    })
  },
  children: [/* content */]
}]

// For "第 X 页 共 Y 页" style:
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [
    new TextRun("第 "),
    new TextRun({ children: [PageNumber.CURRENT] }),
    new TextRun(" 页  共 "),
    new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
    new TextRun(" 页")
  ]
})
```

---

### Critical Rules for docx-js

- **Set page size explicitly** - always use A4 (11906 × 16838 DXA) with 2.5cm margins (1418 DXA) for Chinese academic documents
- **Landscape: pass portrait dimensions** - docx-js swaps width/height internally; pass short edge as `width`, long edge as `height`, and set `orientation: PageOrientation.LANDSCAPE`
- **Never use `\n`** - use separate Paragraph elements
- **Never use unicode bullets** - use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** - standalone creates invalid XML
- **ImageRun requires `type`** - always specify png/jpg/etc
- **Heading numbering: all-decimal only** - NEVER use `LevelFormat.CHINESE_COUNTING` in a multi-level config; it causes H2/H3 to render as `二.4` instead of `2.4`. Use `LevelFormat.DECIMAL` for every level; write Chinese characters (一、二、) manually in H1 paragraph text if needed.
- **Tables MUST be 三线表** - use thick top/bottom borders (size:12, 1.5pt) and thin header-bottom border (size:6, 0.75pt); set all other borders to `BorderStyle.NONE`; no fill color on header cells
- **Table caption: no first-line indent** - override Normal style's `firstLine` with `indent: { firstLine: 0 }` on the caption paragraph
- **Block formulas use 3-column table** - [567 DXA spacer | 7936 DXA formula centered | 567 DXA number right-aligned]; use `scripts/formula.py` to generate; never use plain `$$` LaTeX in the final docx
- **Always set table `width` with DXA** - never use `WidthType.PERCENTAGE` (breaks in Google Docs)
- **Tables need dual widths** - `columnWidths` array AND cell `width`, both must match
- **Table width = sum of columnWidths** - for DXA, ensure they add up exactly
- **Always add cell margins** - use `margins: { top: 80, bottom: 80, left: 120, right: 120 }` for readable padding
- **Use `ShadingType.CLEAR`** - never SOLID for table shading
- **TOC: use `heading: HeadingLevel.HEADING_X`** - applies the overridden Heading style (with `outlineLevel`), giving TOC recognition AND custom formatting simultaneously
- **Override built-in styles** - use exact IDs: "Heading1", "Heading2", etc.; always use `font: { ascii, eastAsia, hAnsi }` object form — `font: "SimHei"` string shorthand only sets the ASCII slot and leaves Chinese characters falling back to Normal font
- **Include `outlineLevel`** - required for TOC (0 for H1, 1 for H2, etc.)

---

## Editing Existing Documents

**Follow all 3 steps in order.**

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```
Extracts XML, pretty-prints, merges adjacent runs, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing. Use `--merge-runs false` to skip run merging.

### Step 2: Edit XML

Edit files in `unpacked/word/`. See XML Reference below for patterns.

**Use "Claude" as the author** for tracked changes and comments, unless the user explicitly requests use of a different name.

**Use the Edit tool directly for string replacement. Do not write Python scripts.** Scripts introduce unnecessary complexity. The Edit tool shows exactly what is being replaced.

**CRITICAL: Use smart quotes for new content.** When adding text with apostrophes or quotes, use XML entities to produce smart quotes:
```xml
<!-- Use these entities for professional typography -->
<w:t>Here&#x2019;s a quote: &#x201C;Hello&#x201D;</w:t>
```
| Entity | Character |
|--------|-----------|
| `&#x2018;` | ‘ (left single) |
| `&#x2019;` | ’ (right single / apostrophe) |
| `&#x201C;` | “ (left double) |
| `&#x201D;` | ” (right double) |

**Adding comments:** Use `comment.py` to handle boilerplate across multiple XML files (text must be pre-escaped XML):
```bash
python scripts/comment.py unpacked/ 0 "Comment text with &amp; and &#x2019;"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0  # reply to comment 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"  # custom author name
```
Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```
Validates with auto-repair, condenses XML, and creates DOCX. Use `--validate false` to skip.

**Auto-repair will fix:**
- `durableId` >= 0x7FFFFFFF (regenerates valid ID)
- Missing `xml:space="preserve"` on `<w:t>` with whitespace

**Auto-repair won't fix:**
- Malformed XML, invalid element nesting, missing relationships, schema violations

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings. Don't inject tracked change tags inside a run.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.
- **Never regenerate `styles.xml`**: When editing an existing document, all template styles live in `word/styles.xml`. Edit only `document.xml` content; do not overwrite or recreate `styles.xml` unless explicitly asked — doing so erases the user's template.
- **Preserve `<w:pStyle>` references**: When inserting new paragraphs, copy the `<w:pStyle w:val="..."/>` from an adjacent paragraph of the same type. Omitting `<w:pStyle>` silently falls back to the document default style, losing all heading/body formatting from the template.
- **Smart quotes are XML-encoded after unpack**: The unpack step converts `"` / `"` to `&#x201C;` / `&#x201D;` entities. When searching `document.xml` for heading text that contains Chinese quotation marks, search for the entity form (`&#x201C;`), not the raw Unicode character.

---

## XML Reference

### Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: Add `xml:space="preserve"` to `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)

### Tracked Changes

**Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Inside `<w:del>`**: Use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

**Minimal edits** - only mark what changes:
```xml
<!-- Change "30 days" to "60 days" -->
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting entire paragraphs/list items** - when removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Add `<w:del/>` inside `<w:pPr><w:rPr>`:
```xml
<w:p>
  <w:pPr>
    <w:numPr>...</w:numPr>  <!-- list numbering if present -->
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```
Without the `<w:del/>` in `<w:pPr><w:rPr>`, accepting changes leaves an empty paragraph/list item.

**Rejecting another author's insertion** - nest deletion inside their insertion:
```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion** - add insertion after (don't modify their deletion):
```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py` (see Step 2), add markers to document.xml. For replies, use `--parent` flag and nest markers inside the parent's.

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment markers are direct children of w:p, never inside w:r -->
<w:commentRangeStart w:id="0"/>
<w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted</w:delText></w:r>
</w:del>
<w:r><w:t> more text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment 0 with reply 1 nested inside -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

### Images

1. Add image file to `word/media/`
2. Add relationship to `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId5" Type=".../image" Target="media/image1.png"/>
```
3. Add content type to `[Content_Types].xml`:
```xml
<Default Extension="png" ContentType="image/png"/>
```
4. Reference in document.xml:
```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="914400" cy="914400"/>  <!-- EMUs: 914400 = 1 inch -->
    <a:graphic>
      <a:graphicData uri=".../picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

---

## Dependencies

- **pandoc**: Text extraction AND LaTeX→OMML formula conversion (`scripts/formula.py` requires pandoc ≥ 2.0)
- **docx**: `npm install docx` (new documents)
- **temml**: `npm install temml` (LaTeX → MathML conversion for Word native math)
- **fast-xml-parser**: `npm install fast-xml-parser` (MathML parsing for docx conversion)
- **LibreOffice**: PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images

---

## Markdown to Word Conversion (Chinese Academic Papers)

This section documents comprehensive solutions for converting Markdown papers with LaTeX formulas, tables, and citations to properly formatted Word documents following Chinese academic standards.

### Quick Start

```bash
# Install dependencies
npm install docx temml fast-xml-parser

# Run conversion
node scripts/new_doc.js
```

### Critical Issues & Solutions

The following 8 issues were identified and solved during production usage. All solutions are implemented in `scripts/new_doc.js` and `scripts/mathml-to-docx.js`.

#### Issue 1: Three-Line Table Middle Borders Visible

**Problem**: Body row borders in 三线表 appeared visible instead of invisible.

**Solution**: Set ALL body row borders to `NONE`, only keep:
- Header top: `THICK` (1.5pt)
- Header bottom: `THIN` (0.75pt)
- Last row bottom: `THICK` (1.5pt)

```javascript
const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// Header row
cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)

// Body rows - ALL borders NONE except last row bottom
cellOf(cell, colWidths[i], {
  top: NONE,
  bottom: isLastRow ? THICK : NONE,
  left: NONE,
  right: NONE,
})
```

#### Issue 2: Formula Table Borders Visible

**Problem**: Block formula tables (3-column layout) showed visible borders.

**Solution**: Set ALL borders including `insideHorizontal` and `insideVertical` to `NONE`:

```javascript
return new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [567, 7936, 567],
  borders: {
    top: NONE,
    bottom: NONE,
    left: NONE,
    right: NONE,
    insideHorizontal: NONE,  // CRITICAL: Must include these
    insideVertical: NONE,    // CRITICAL: Must include these
  },
  rows: [new TableRow({ children: [leftCell, formulaCell, numberCell] })],
});
```

#### Issue 3: Heading Numbering Out of Sync

**Problem**: Multi-level heading numbers (1.1, 1.2, 2.1) were not synchronized across chapters.

**Solution**: Use manual counters instead of Word's auto-numbering:

```javascript
let currentChapter = 0;
let currentSection = 0;
let currentSubsection = 0;

function h1(text) {
  currentChapter++;
  currentSection = 0;      // Reset on chapter change
  currentSubsection = 0;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun(`${currentChapter} ${text}`)],
  });
}

function h2(text) {
  currentSection++;
  currentSubsection = 0;   // Reset on section change
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun(`${currentChapter}.${currentSection} ${text}`)],
  });
}

function h3(text) {
  currentSubsection++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun(`${currentChapter}.${currentSection}.${currentSubsection} ${text}`)],
  });
}
```

#### Issue 4: Heading English/Numbers in Wrong Font (SimSun instead of Cambria Math)

**Problem**: English text and numbers in headings, figure captions, and table captions displayed in SimSun (宋体) instead of Cambria Math.

**Solution**: Use mixed font configuration with `ascii`, `eastAsia`, and `hAnsi` properties:

```javascript
// Style definition for headings
{
  id: 'Heading1', name: 'Heading 1', basedOn: 'Normal',
  run: {
    font: {
      ascii: 'Cambria Math',    // English letters, numbers
      eastAsia: 'SimHei',        // Chinese characters (黑体)
      hAnsi: 'Cambria Math',     // Western European characters
    },
    size: 32, bold: true,
  },
  // ... paragraph settings
}

// For captions (SimSun body font for Chinese)
font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }
```

#### Issue 5: Block Formulas Not Using Word Equation Editor

**Problem**: LaTeX formulas were rendered as plain text or images instead of native Word equations.

**Solution**: Use `temml` (LaTeX→MathML) + `mathml-to-docx.js` (MathML→OMML) pipeline:

```javascript
const temml = require('temml');
const { mathmlToDocxChildren } = require('./mathml-to-docx');

function latexToMath(latex) {
  const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
  const children = mathmlToDocxChildren(mathml);
  if (children && children.length) {
    return new Math({ children });
  }
  // Fallback
  return new Math({ children: [new MathRun(latex)] });
}

// Usage in formula table
const mathObj = latexToMath('Q_n(x, a) = r + \\gamma V_{n-1}(y)');
```

#### Issue 6: Main Title English in Wrong Font

**Problem**: Document title's English text showed in SimSun instead of Cambria Math.

**Solution**: Apply same mixed font to title paragraph:

```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: '论文标题 Paper Title',
    bold: true,
    size: 36,
    font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
  })],
})
```

#### Issue 7: Inline Math Detection Too Aggressive

**Problem**: Inline math regex matched plain numbers (like "1992") and English words (like "Agent", "Watkins"), incorrectly converting them to formula objects.

**Solution**: Use strict regex that ONLY matches actual mathematical content:

```javascript
// Detection function - returns true only for real math content
function containsMath(text) {
  // Greek letters
  if (/[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text)) return true;
  // Unicode subscript/superscript characters
  if (/[₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]/.test(text)) return true;
  // Math operators and special symbols
  if (/[∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩]/.test(text)) return true;
  // Starred symbols like π*, Q*
  if (/[A-Z]\*/.test(text)) return true;
  // Explicit $...$ LaTeX
  if (/\$[^$]+\$/.test(text)) return true;
  return false;
}

// Parsing regex - strict matching for inline formulas
const mathPattern = /\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;

// What it matches:
// - $...$ explicit LaTeX
// - Qₙ(x,a) - function with subscripts
// - V(αₙ) - function with Greek params
// - αₙ - Greek with subscript
// - xₙ - variable with subscript
// - Q* - starred variable

// What it does NOT match:
// - Plain numbers: 1992, 500
// - English words: Agent, Watkins, Dayan
// - Plain parentheses: (1992), (optional)
```

#### Issue 8: Citations Not in Superscript

**Problem**: Reference citations like [1], [2] appeared as normal text instead of superscript.

**Solution**: Detect citation pattern and apply `superScript: true`:

```javascript
function containsCitation(text) {
  return /\[\d+\]/.test(text);
}

// In parseInlineContentWithCitations():
const combinedPattern = /(\[\d+\])|...; // Citation first in alternation

if (match[1]) {
  // Citation [n] - convert to superscript
  children.push(new TextRun({
    text: match[1],
    superScript: true,
  }));
}
```

#### Issue 10: Page Break After Abstract and Before References

**Problem**: Abstract page should end after keywords, and references should start on a new page.

**Solution**: Add `pageBreak()` helper function:

```javascript
// ISSUE 10 FIX: Page break helper
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Usage in CONTENT:
// After keywords (end of abstract)
body('关键词：强化学习；Q-learning'),
pageBreak(),  // <- Start new page after abstract

// Before references
pageBreak(),  // <- Start new page for references
new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun('参考文献')],
}),
```

**Important**: `PageBreak` MUST be wrapped in a `Paragraph` - it cannot be used standalone.

#### Issue 11: Formula Table Cell Vertical Alignment

**Problem**: In the 3-column borderless formula table, the formula and equation number `(n)` are not vertically centered, causing misalignment.

**Solution**: Add `verticalAlign: VerticalAlign.CENTER` to all three TableCell definitions:

```javascript
// First, add VerticalAlign to imports:
const {
  // ... other imports
  VerticalAlign,  // <-- Add this
} = require('docx');

// Then apply to each cell in the formula() function:
const leftCell = new TableCell({
  width: { size: 567, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
});

const formulaCell = new TableCell({
  width: { size: 7936, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [mathObj],
  })],
});

const numberCell = new TableCell({
  width: { size: 567, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    indent: { firstLine: 0 },
    children: [new TextRun(`(${number})`)],
  })],
});
```

**Result**: Formula and equation number now align horizontally on the same baseline.

### File Structure

```
scripts/
├── new_doc.js          # Main template with Issues 1-8, 10-11 fixes (~820 lines)
├── mathml-to-docx.js   # MathML→docx Math converter (~250 lines)
├── formula.py          # Block formula XML insertion (legacy)
├── table.py            # Table XML insertion (legacy)
└── office/
    ├── unpack.py       # DOCX→XML extraction
    ├── pack.py         # XML→DOCX assembly
    └── validate.py     # Document validation
```

### Unicode to LaTeX Mapping

The script includes comprehensive Unicode math symbol conversion:

```javascript
const UNICODE_TO_LATEX = {
  // Greek letters
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', // ... etc
  // Subscripts
  '₀': '_0', '₁': '_1', '₂': '_2', // ... etc
  // Superscripts
  '²': '^2', '³': '^3', // ... etc
  // Special symbols
  '∞': '\\infty', '∑': '\\sum', '∫': '\\int', // ... etc
};
```

### MathML to DOCX Conversion

`scripts/mathml-to-docx.js` converts MathML (from temml) to docx Math components:

- Fractions: `<mfrac>` → `MathFraction`
- Subscripts: `<msub>` → `MathSubScript`
- Superscripts: `<msup>` → `MathSuperScript`
- Combined: `<msubsup>` → `MathSubSuperScript`
- Radicals: `<msqrt>`, `<mroot>` → `MathRadical`
- Summation: `<munderover>` with ∑ → `MathSum`
- Integrals: `<munderover>` with ∫ → `MathIntegral`
- Matrices: `<mtable>` → `MathMatrix`

### Workflow for New Conversions

1. **Copy `new_doc.js` template** to your project
2. **Install dependencies**: `npm install docx temml fast-xml-parser`
3. **Edit CONTENT section** with your document structure
4. **Use helper functions**:
   - `h1/h2/h3(text)` - Headings with auto-numbering
   - `body(text)` - Body paragraph (auto-detects math/citations)
   - `formula(latex, number)` - Block formula
   - `threeLineTable(headers, rows, colWidths)` - Three-line table
   - `tableCaption/figCaption(label)` - Captions
   - `pageBreak()` - Page break (Issue 10)
   - `ref(text)` - Reference entry
5. **Run**: `node new_doc.js`
