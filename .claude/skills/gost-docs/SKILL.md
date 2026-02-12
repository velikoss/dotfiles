---
name: gost-docx
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files) formatted according to Russian GOST academic standards (ГОСТ) for курсовая работа, ВКР, дипломная работа, or any Russian university paper. Triggers include: any mention of \"ГОСТ\", \"курсовая\", \"ВКР\", \"дипломная\", \"оформление по ГОСТу\", \"оформление работы\", \"Times New Roman 14\", \"полуторный интервал\", or requests to produce academic documents with Russian formatting standards. Also triggers when the user asks for a document with margins 30/10/20/20 mm, Times New Roman font, 1.5 line spacing, or section numbering like \"1.1\", \"1.1.1\". Use this skill even if the user just says \"оформи документ\" or \"сделай по ГОСТу\" without specifying exact standards. Do NOT use for non-Russian academic formatting (APA, MLA, Chicago) or general business documents."
license: Proprietary. LICENSE.txt has complete terms
---

# GOST-Compliant DOCX Creation for Russian Academic Papers

## Overview

This skill creates .docx files formatted according to Russian GOST standards for курсовая работа (КР), курсовой проект (КП), ВКР, and other academic papers. A .docx file is a ZIP archive containing XML files.

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` or unpack for raw XML |
| Create new GOST document | Use `docx-js` with GOST parameters below |
| Edit existing document | Unpack → edit XML → repack |

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

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## GOST Formatting Parameters

These are the exact formatting rules extracted from the GOST standard. Follow them precisely when creating documents.

### Page Setup (A4 format, 210×297 mm)

```javascript
// GOST requires A4 paper with specific margins
sections: [{
  properties: {
    page: {
      size: {
        width: 11906,   // A4 width in DXA (210 mm)
        height: 16838    // A4 height in DXA (297 mm)
      },
      margin: {
        top: 1134,     // 20 mm
        bottom: 1134,  // 20 mm
        left: 1701,    // 30 mm
        right: 567     // 10 mm
      }
    }
  },
  children: [/* content */]
}]
```

**GOST Margins (DXA units, 567 DXA ≈ 1 cm):**

| Side | mm | DXA |
|------|-----|------|
| Left | 30 | 1701 |
| Right | 10 | 567 |
| Top | 20 | 1134 |
| Bottom | 20 | 1134 |

**Content width**: 11906 - 1701 - 567 = **9638 DXA** (170 mm)

### Page Numbering

Pages are numbered with Arabic numerals, centered at the bottom of the page, continuous numbering including appendices.

```javascript
footers: {
  default: new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 28 })]
    })]
  })
}
```

### Body Text Parameters (Основной текст)

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Style | Regular (Обычный) |
| Size | 14 pt (size: 28 in docx-js, which uses half-points) |
| Color | Black (auto) |
| Alignment | Justified (По ширине) |
| First line indent | 1.25 cm (709 DXA) |
| Left/Right indent | 0 |
| Spacing before | 0 |
| Spacing after | 0 |
| Line spacing | 1.5 (Полуторный) |
| Widow/orphan control | Yes (Запрет висящих строк) |

```javascript
// GOST body text style
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 28 }, // 14pt
        paragraph: {
          spacing: { line: 360 }, // 1.5 line spacing (240 * 1.5)
          indent: { firstLine: 709 }, // 1.25 cm first line indent
          alignment: AlignmentType.JUSTIFIED
        }
      }
    },
    paragraphStyles: [/* heading styles below */]
  }
});
```

### Heading Styles

#### Heading Level 1 (Заголовок первого уровня) — Section titles

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Style | Bold (Полужирный) |
| Size | 18 pt (size: 36) |
| Case | All Caps (Все прописные) |
| Alignment | Justified (По ширине) |
| First line indent | 1.25 cm (709 DXA) |
| Spacing before | 0 pt |
| Spacing after | 12 pt (240 half-points) |
| Line spacing | 1.5 |
| Page break | New page (С новой страницы) |
| Keep with next | Yes |

**Special sections** (СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ, ПРИЛОЖЕНИЯ) are formatted with Heading 1 style but **centered** and printed in ALL CAPS.

```javascript
{
  id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { size: 36, bold: true, font: "Times New Roman", allCaps: true },
  paragraph: {
    spacing: { before: 0, after: 240, line: 360 },
    indent: { firstLine: 709 },
    alignment: AlignmentType.JUSTIFIED,
    outlineLevel: 0,
    pageBreakBefore: true,
    keepNext: true
  }
}
```

#### Heading Level 2 (Заголовок второго уровня) — Subsection titles

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Style | Bold |
| Size | 16 pt (size: 32) |
| Alignment | Justified |
| First line indent | 1.25 cm |
| Spacing before | 24 pt (480) |
| Spacing after | 12 pt (240) |
| Line spacing | 1.5 |
| Keep with next | Yes |
| Keep lines together | Yes |
| Widow/orphan control | Yes |

```javascript
{
  id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { size: 32, bold: true, font: "Times New Roman" },
  paragraph: {
    spacing: { before: 480, after: 240, line: 360 },
    indent: { firstLine: 709 },
    alignment: AlignmentType.JUSTIFIED,
    outlineLevel: 1,
    keepNext: true,
    keepLines: true
  }
}
```

#### Heading Level 3+ (Заголовок третьего и последующих уровней) — Point titles

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Style | Bold |
| Size | 14 pt (size: 28) |
| Alignment | Justified |
| First line indent | 1.25 cm |
| Spacing before | 24 pt (480) |
| Spacing after | 12 pt (240) |
| Line spacing | 1.5 |
| Keep with next | Yes |
| Keep lines together | Yes |

```javascript
{
  id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
  run: { size: 28, bold: true, font: "Times New Roman" },
  paragraph: {
    spacing: { before: 480, after: 240, line: 360 },
    indent: { firstLine: 709 },
    alignment: AlignmentType.JUSTIFIED,
    outlineLevel: 2,
    keepNext: true,
    keepLines: true
  }
}
```

### Section Numbering Rules

Sections, subsections, and points are numbered with Arabic numerals:
- Section: `1`, `2`, `3`...
- Subsection: `1.1`, `1.2`, `2.1`...
- Point: `1.1.1`, `1.1.2`...
- No period after the number, just a space before the title
- Sections "Содержание", "Введение", "Заключение", "Список использованных источников" are NOT numbered
- All heading titles start with a capital letter
- No punctuation at the end of heading titles (if two sentences, separate with a period)

### Table Formatting

**Table caption** (Надпись таблицы):

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman, Italic (Курсив) |
| Size | 12 pt (size: 24) |
| Alignment | Justified |
| No first line indent | |
| Spacing before | 6 pt |
| Line spacing | Single (Одинарный) |
| Keep with next | Yes |

**Table content** (Содержание таблицы):

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Size | 12 pt (size: 24) |
| Alignment | Justified or Center |
| Line spacing | Single |

**Table rules:**
- Title format: `Таблица X.Y — Название` (dash is em-dash `—`)
- No period at the end of the title
- Column headers centered, row headers left-aligned
- Headers in singular, capitalized
- No diagonal cell division
- No "№ п/п" column
- For table continuation: `Продолжение Таблицы X.Y` above continued parts
- Line after table has 6pt spacing before

```javascript
// Table caption paragraph
new Paragraph({
  spacing: { before: 120, line: 240 }, // 6pt before, single spacing
  indent: { firstLine: 0 },
  alignment: AlignmentType.JUSTIFIED,
  keepNext: true,
  children: [new TextRun({
    text: "Таблица 1.1 — Исходные данные",
    font: "Times New Roman", size: 24, italics: true
  })]
})

// Table cells use 12pt Times New Roman, single spacing
new Table({
  width: { size: 9638, type: WidthType.DXA }, // Full GOST content width
  // ... rows and cells with font size 24 (12pt)
})
```

### Figure/Illustration Formatting (Надпись к иллюстрации)

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman, Bold |
| Size | 12 pt (size: 24) |
| Alignment | Center |
| No first line indent | |
| Spacing after | 6 pt |
| Line spacing | Single |

**Rules:**
- Caption goes BELOW the illustration
- Format: `Рисунок X.Y — Название` (no abbreviation of "Рисунок")
- No period at end
- Numbered within sections: `Рисунок 2.7`
- References in text: `(Рисунок 3.1)` or `на Рисунке 3.1`

```javascript
// Figure caption
new Paragraph({
  spacing: { after: 120, line: 240 },
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  children: [new TextRun({
    text: "Рисунок 1.4 — Детали прибора",
    font: "Times New Roman", size: 24, bold: true
  })]
})
```

### Code Listing Formatting (Листинг)

**Listing caption** (Надпись листинга):

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman, Italic |
| Size | 12 pt |
| Alignment | Left |
| Spacing before | 6 pt |
| Line spacing | Single |
| Keep with next | Yes |

**Listing content** (Содержание листинга):

| Parameter | Value |
|-----------|-------|
| Font | Courier New |
| Size | 10 pt (size: 20) |
| Alignment | Left |
| Line spacing | Single |

**Rules:**
- Caption format: `Листинг X.Y — Название`
- Listing is placed in a border/frame
- Line references use step 10: `{30, 50, 140}` or `{20-80}`
- 6pt spacing after listing

### Lists (Маркированные и нумерованные списки)

Lists use the same body text style (14pt Times New Roman, 1.5 spacing).

**Marker options:** `–`, `―`, `●`, `■`, `○`, `1.`, `а)`

**List positioning:**
- Marker indent: 1.25 cm
- Text tabulation after marker
- Text indent: 1.9 cm (approximately)

**Rules:**
- Use only ONE marker type throughout the entire document
- Bulleted list items start with lowercase, end with semicolon (`;`), last item ends with period (`.`)
- Numbered list items start with uppercase, end with period (`.`)

```javascript
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "gost-bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2013", // en-dash –
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: 1077, hanging: 368 } // ~1.9cm left, marker at 1.25cm
            },
            run: { font: "Times New Roman" }
          }
        }]
      },
      {
        reference: "gost-numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: 1077, hanging: 368 }
            }
          }
        }]
      }
    ]
  }
});
```

### References / Bibliography (Список использованных источников)

**Title:** "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" — Heading 1 style, centered, ALL CAPS.

**Section title style in bibliography** (from Table 9.1):

| Parameter | Value |
|-----------|-------|
| Font | Times New Roman |
| Style | Regular |
| Size | 14 pt |
| Case | All Caps |
| Alignment | Center |
| No first line indent | |
| Spacing before/after | 6 pt |
| Line spacing | 1.5 |
| Keep with next | Yes |

**Rules:**
- Sources ordered by first mention in text
- Minimum 7 sources
- In-text citations: `[1]`, `[2]`, `[4, c.21-25]`
- City abbreviations: М. (Москва), Л. (Ленинград), СПб. (Санкт-Петербург), К. (Киев)
- No space after abbreviations like `М.:`, `СПб.:`
- No space after period in author initials: `Громов Г.Р.`

**Book format:** `Автор И.О. Название. — М.: Издательство, 2015. — 336 с.`
**Article format:** `Автор И.О. Название // Журнал. — Год. — Т. X. — Вып. Y. — С. XX–YY.`
**Web format:** `Автор И.О. Название ― URL: http://... (Дата обращения: ДД.ММ.ГГГГ)`

### Typography Rules

**Quotes:** Russian text uses `«…»`, English text uses `"…"`

**Dashes:**
- Hyphen `-` — no spaces, for compound words (физико-математический)
- Minus `–` (en-dash) — for subtraction
- Em-dash `—` — with spaces, for punctuation (Экономика — это…)

**Numbers:**
- Arabic numerals only (except quarters/semesters use Roman)
- No case endings on Roman numerals or dates
- Unit of measurement after the last number in a series
- Math signs (+, –, =, >, <) only in formulas; spell out in text

**No word hyphenation** (перенос слов не допускается).

**No blank space** on pages — content should fill pages.

---

## Creating New GOST Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Full Setup Example

```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 28 }, // 14pt
        paragraph: {
          spacing: { line: 360 }, // 1.5 line spacing
          indent: { firstLine: 709 } // 1.25 cm
        }
      }
    },
    paragraphStyles: [
      // Heading 1: 18pt bold, ALL CAPS, new page
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Times New Roman", allCaps: true },
        paragraph: {
          spacing: { before: 0, after: 240, line: 360 },
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 0,
          pageBreakBefore: true,
          keepNext: true
        }
      },
      // Heading 2: 16pt bold
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Times New Roman" },
        paragraph: {
          spacing: { before: 480, after: 240, line: 360 },
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 1,
          keepNext: true,
          keepLines: true
        }
      },
      // Heading 3: 14pt bold
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman" },
        paragraph: {
          spacing: { before: 480, after: 240, line: 360 },
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 2,
          keepNext: true,
          keepLines: true
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, bottom: 1134, left: 1701, right: 567 } // 20/20/30/10 mm
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 28 })]
        })]
      })
    },
    children: [/* content */]
  }]
});

Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation

After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.
```bash
python scripts/office/validate.py doc.docx
```

### Lists (NEVER use unicode bullets)

```javascript
// ❌ WRONG - never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })

// ✅ CORRECT - use GOST-compliant list markers
const doc = new Document({
  numbering: {
    config: [
      { reference: "gost-bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2013", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1077, hanging: 368 } } } }] },
      { reference: "gost-numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1077, hanging: 368 } } } }] },
    ]
  },
  sections: [{
    children: [
      // Bulleted: lowercase start, semicolon end
      new Paragraph({ numbering: { reference: "gost-bullets", level: 0 },
        children: [new TextRun("элемент списка;")] }),
      // Numbered: uppercase start, period end
      new Paragraph({ numbering: { reference: "gost-numbers", level: 0 },
        children: [new TextRun("Элемент списка.")] }),
    ]
  }]
});
```

### Tables

**CRITICAL: Tables need dual widths** — set both `columnWidths` on the table AND `width` on each cell.

```javascript
const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

// Table caption (italic, 12pt, 6pt before)
new Paragraph({
  spacing: { before: 120, line: 240 },
  indent: { firstLine: 0 },
  alignment: AlignmentType.JUSTIFIED,
  keepNext: true,
  children: [new TextRun({
    text: "Таблица 1.1 \u2014 Исходные данные", // em-dash
    font: "Times New Roman", size: 24, italics: true
  })]
})

// Table content (12pt, single spacing)
new Table({
  width: { size: 9638, type: WidthType.DXA }, // GOST content width
  columnWidths: [4819, 4819],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: 4819, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({
            spacing: { line: 240 }, // single spacing in tables
            alignment: AlignmentType.CENTER, // headers centered
            children: [new TextRun({ text: "Заголовок", font: "Times New Roman", size: 24 })]
          })]
        })
      ]
    })
  ]
})
```

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new ImageRun({
    type: "png",
    data: fs.readFileSync("image.png"),
    transformation: { width: 400, height: 300 },
    altText: { title: "Title", description: "Desc", name: "Name" }
  })]
})
// Caption BELOW the image (bold, 12pt, centered)
new Paragraph({
  spacing: { after: 120, line: 240 },
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  children: [new TextRun({
    text: "Рисунок 1.4 \u2014 Детали прибора",
    font: "Times New Roman", size: 24, bold: true
  })]
})
```

### Page Breaks

```javascript
new Paragraph({ children: [new PageBreak()] })
// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### Table of Contents

```javascript
new TableOfContents("Содержание", { hyperlink: true, headingStyleRange: "1-3" })
```

### Headers/Footers

```javascript
sections: [{
  properties: {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1134, bottom: 1134, left: 1701, right: 567 }
    }
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun("Стр. "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Critical Rules for GOST docx-js

- **Always use A4 paper** — 11906 × 16838 DXA (never US Letter)
- **GOST margins** — left 30mm (1701), right 10mm (567), top/bottom 20mm (1134)
- **Font: Times New Roman only** — 14pt body, 18pt H1, 16pt H2, 14pt H3, 12pt tables/captions
- **1.5 line spacing** for body text and headings (line: 360); single spacing for tables and captions (line: 240)
- **First line indent 1.25 cm** (709 DXA) for body text and headings
- **Justified alignment** for body text and most headings
- **Page numbering** centered at bottom, Arabic numerals, continuous
- **No word hyphenation** — disable hyphenation
- **Russian quotes «…»** — use `\u00AB` and `\u00BB` or XML entities
- **Em-dash with spaces** for punctuation: `\u2014` (`—`)
- **Never use `\n`** — use separate Paragraph elements
- **Never use unicode bullets** — use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** — standalone creates invalid XML
- **ImageRun requires `type`** — always specify png/jpg/etc
- **Always set table `width` with DXA** — never use `WidthType.PERCENTAGE`
- **Tables need dual widths** — `columnWidths` array AND cell `width`, both must match
- **Use `ShadingType.CLEAR`** — never SOLID for table shading
- **TOC requires HeadingLevel only** — no custom styles on heading paragraphs
- **Override built-in styles** — use exact IDs: "Heading1", "Heading2", etc.
- **Include `outlineLevel`** — required for TOC (0 for H1, 1 for H2, etc.)

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

**Use the Edit tool directly for string replacement. Do not write Python scripts.**

**CRITICAL: Use Russian quotes for new content.** When adding Russian text, use proper quote marks:
```xml
<!-- Russian quotes -->
<w:t>Пример &#x00AB;цитата&#x00BB;</w:t>
<!-- Em-dash with spaces -->
<w:t>Экономика &#x2014; это наука</w:t>
```

**Adding comments:**
```bash
python scripts/comment.py unpacked/ 0 "Comment text with &amp; and &#x2019;"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"
```
Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs.

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

**Minimal edits** — only mark what changes:
```xml
<w:r><w:t>Срок составляет </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> дней.</w:t></w:r>
```

**Deleting entire paragraphs** — mark the paragraph mark as deleted:
```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Удаляемый абзац...</w:delText></w:r>
  </w:del>
</w:p>
```

**Rejecting another author's insertion:**
```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion:**
```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py`, add markers to document.xml:

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<w:commentRangeStart w:id="0"/>
<w:r><w:t>commented text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
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

- **pandoc**: Text extraction
- **docx**: `npm install -g docx` (new documents)
- **LibreOffice**: PDF conversion (auto-configured via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images
