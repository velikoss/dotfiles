---
name: gost-docx
description: "Create, edit, and rewrite Microsoft Word documents (.docx) formatted to Russian academic GOST standards (ГОСТ) — курсовая работа, курсовой проект, ВКР, дипломная, магистерская диссертация, реферат, отчёт по практике, любая письменная работа российского вуза. Triggers on any mention of «ГОСТ», «оформление по ГОСТу», «курсовая», «ВКР», «дипломная», «реферат», «магистерская», «оформи документ», «сделай по госту», «оформи работу»; on margins 30/10/20/20 mm, Times New Roman 14, полуторный интервал; on numbering like «1.1.1»; on centered titles СОДЕРЖАНИЕ / ВВЕДЕНИЕ / ЗАКЛЮЧЕНИЕ / СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ / ПРИЛОЖЕНИЯ; on academic captions («Таблица 1.1 — …», «Рисунок 1.1 — …»). Use this skill even when the user only says «оформи» or «по ГОСТу» without naming exact standards — Russian academic formatting breaks in subtle ways without this guidance. Do NOT use for APA/MLA/Chicago citation styles, business documents, or non-Russian academic papers."
license: Free to use. Created by Velikiy (velikoss) Kirill
---

# GOST DOCX — Russian academic paper formatting

## Overview

Build .docx files that pass formal ГОСТ-style review at Russian universities. The skill covers two modes:

1. **Create** a new document from scratch with `docx-js` (Node.js).
2. **Edit** an existing .docx by unpacking, modifying XML, and repacking using the public `docx` skill scripts.

Almost every formatting failure on GOST work comes from the same handful of subtle rules. They are listed first below — read them before writing any code.

---

## Critical rules — read before writing code

These are the issues that most often produce visibly wrong output. Internalize all nine before generating a document.

1. **Centered "special" sections override Heading 1's indent.** Section titles `СОДЕРЖАНИЕ`, `ВВЕДЕНИЕ`, `ЗАКЛЮЧЕНИЕ`, `СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ`, `ПРИЛОЖЕНИЯ`, `ТЕОРЕТИЧЕСКАЯ ЧАСТЬ`, `ПРАКТИЧЕСКАЯ ЧАСТЬ` use Heading 1 style **but** are centered with **no first-line indent**. If you forget to override `indent: { firstLine: 0 }`, the title shifts ≈1 cm to the right and looks off-center. Always set both `alignment: AlignmentType.CENTER` AND `indent: { firstLine: 0 }` on these. Use the `Heading1Centered` style provided in the template below — never apply `Heading1` directly to these titles.

2. **List indents are exact.** Marker is at **1.25 cm**, text at **2.25 cm**. In docx-js: `indent: { left: 1276, hanging: 567 }` (1276 DXA = 2.25 cm; 567 DXA = 1 cm hang back to marker). Any other values violate Table 4.1 of the standard.

3. **Bibliography section heading is NOT bold.** "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" uses **regular weight 14 pt**, all caps, centered, with `before: 120, after: 120` (6 pt). It is a separate style from regular Heading 1. Do not just reuse Heading 1.

4. **Title page has no visible page number.** Use `titlePage: true` in the section properties AND an empty `first` footer. Numbering still counts the title (next page is "2"), but no number renders on page 1.

5. **Em-dash with spaces, hyphen without.** Title format is `Таблица 1.1 — Название` (em-dash U+2014 `—`, with spaces). Hyphen `-` joins compound words (физико-математический), no spaces. Never use `--` or the en-dash U+2013 for these titles.

6. **Russian quotes are «…», not "…" or "…".** Use U+00AB `«` and U+00BB `»`. The smart-quote characters that some editors auto-insert are wrong.

7. **Headings 2 and 3 need `before: 480` (24 pt), Heading 1 needs `before: 0`.** Heading 1 starts a new page, so it doesn't need extra spacing before; H2 and H3 do. Mixing these up makes section breaks look cramped or sparse.

8. **Tables use single line spacing and 12 pt; body text uses 1.5 spacing and 14 pt.** Table cells must explicitly set `spacing: { line: 240 }` and `size: 24` on every paragraph inside cells, otherwise they inherit body defaults and the table looks too tall.

9. **A listing that crosses a page break gets «Продолжение листинга X.Y» — placed by measuring, never by guessing.** Neither docx-js nor XML editing knows where a page ends, so any document with a code listing must go through `scripts/listing_continuation.py` as the last step (see «Перенос листинга» below). It splits the frame exactly where the page fills up, puts the continuation caption on the first line of the next page and verifies the result. Never insert a continuation caption or a page break into a listing by hand.

---

## Decide: create or edit?

| Scenario | Mode |
|---|---|
| New document from scratch, you control all formatting | **Create** (use docx-js, see below) |
| User uploaded a .docx and wants edits, fixes, or content added | **Edit** (unpack → modify XML → repack, see Mode 2) |
| User uploaded a `.doc` (legacy Word) | Convert first: `python scripts/office/soffice.py --headless --convert-to docx file.doc`, then choose Create or Edit |

For deeper XML editing reference, see `references/xml-editing.md`. For the full GOST parameter tables verbatim from the standard, see `references/deep-spec.md`.

---

## Mode 1 — Creating a new GOST document

### Setup

```bash
npm install -g docx
```

The public `docx` skill at `/mnt/skills/public/docx/` provides packing, unpacking, and validation scripts. Read its SKILL.md first if you haven't.

### Page parameters (A4, GOST margins)

| Side | mm | DXA |
|---|---|---|
| Left | 30 | 1701 |
| Right | 10 | 567 |
| Top | 20 | 1134 |
| Bottom | 20 | 1134 |

Page size: `width: 11906, height: 16838` (A4 in DXA). Content width: 9638 DXA (170 mm).

### Complete minimal template — start here, then add content

This template includes every style and config a typical курсовая needs, with the correct values. Copy it, then fill in the children of the section.

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, AlignmentType, LevelFormat, PageNumber, PageBreak,
  WidthType, BorderStyle, HeightRule, ShadingType, TabStopType
} = require('docx');
const fs = require('fs');

const doc = new Document({
  // ---- DEFAULT STYLE: 14 pt Times New Roman, justified, 1.5 spacing, 1.25 cm first line ----
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 28 }, // 28 = 14 pt (half-points)
        paragraph: {
          spacing: { line: 360 },         // 1.5 line spacing
          indent: { firstLine: 709 },     // 1.25 cm
          alignment: AlignmentType.JUSTIFIED
        }
      }
    },
    paragraphStyles: [
      // ---- Heading 1 (numbered sections like "1 ТЕОРЕТИЧЕСКАЯ ЧАСТЬ") ----
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:  { font: "Times New Roman", size: 36, bold: true, allCaps: true }, // 18 pt
        paragraph: {
          spacing: { before: 0, after: 240, line: 360 }, // after 12 pt
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 0,
          pageBreakBefore: true,
          keepNext: true
        }
      },
      // ---- Heading 1 CENTERED (for СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, etc.) ----
      // CRITICAL: alignment CENTER + firstLine 0 to keep visually centered
      {
        id: "Heading1Centered", name: "Heading 1 Centered", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:  { font: "Times New Roman", size: 36, bold: true, allCaps: true },
        paragraph: {
          spacing: { before: 0, after: 240, line: 360 },
          indent: { firstLine: 0 },                      // <-- override the body default
          alignment: AlignmentType.CENTER,
          outlineLevel: 0,
          pageBreakBefore: true,
          keepNext: true
        }
      },
      // ---- Heading 2 (subsections like "1.1 Концепции") ----
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:  { font: "Times New Roman", size: 32, bold: true }, // 16 pt
        paragraph: {
          spacing: { before: 480, after: 240, line: 360 }, // before 24 pt, after 12 pt
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 1,
          keepNext: true,
          keepLines: true
        }
      },
      // ---- Heading 3 (points like "1.1.1 Эволюционная концепция") ----
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:  { font: "Times New Roman", size: 28, bold: true }, // 14 pt
        paragraph: {
          spacing: { before: 480, after: 240, line: 360 },
          indent: { firstLine: 709 },
          alignment: AlignmentType.JUSTIFIED,
          outlineLevel: 2,
          keepNext: true,
          keepLines: true
        }
      },
      // ---- Bibliography section heading: regular weight, 14 pt, all caps, centered ----
      // (This is its own style per Table 9.1 — do NOT reuse Heading1 for it.)
      {
        id: "BiblioHeading", name: "Bibliography Heading", basedOn: "Normal", next: "Normal",
        run:  { font: "Times New Roman", size: 28, allCaps: true }, // regular weight
        paragraph: {
          spacing: { before: 120, after: 120, line: 360 }, // 6 pt before/after
          indent: { left: 709, firstLine: 0 },             // 1.25 cm left, no first-line
          alignment: AlignmentType.CENTER,
          keepNext: true
        }
      }
    ]
  },

  // ---- LIST CONFIGURATION ----
  // Marker at 1.25 cm, text at 2.25 cm (per Table 4.1)
  numbering: {
    config: [
      {
        reference: "gost-bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "—", // em-dash marker
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: { left: 1276, hanging: 567 } }, // CORRECT GOST indents
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
            paragraph: { indent: { left: 1276, hanging: 567 } }
          }
        }]
      }
    ]
  },

  sections: [{
    properties: {
      page: {
        size:   { width: 11906, height: 16838 },                    // A4
        margin: { top: 1134, bottom: 1134, left: 1701, right: 567 } // 20/20/30/10 mm
      },
      titlePage: true   // makes first-page footer separate
    },
    footers: {
      // Page numbers on all pages except title
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          indent: { firstLine: 0 },
          children: [new TextRun({
            children: [PageNumber.CURRENT],
            font: "Times New Roman", size: 28
          })]
        })]
      }),
      first: new Footer({ children: [new Paragraph({ children: [] })] }) // empty: title page no number
    },
    children: [
      // === TITLE PAGE content goes here (centered text runs, etc.) ===
      // ...

      // === СОДЕРЖАНИЕ — use Heading1Centered ===
      new Paragraph({ style: "Heading1Centered", children: [new TextRun("Содержание")] }),
      // ...TOC entries here...

      // === ВВЕДЕНИЕ — use Heading1Centered ===
      new Paragraph({ style: "Heading1Centered", children: [new TextRun("Введение")] }),
      new Paragraph({ children: [new TextRun("Текст введения...")] }),

      // === Numbered section — use Heading1 (justified, indented, ALL CAPS via style) ===
      new Paragraph({ style: "Heading1", children: [new TextRun("1 Теоретическая часть")] }),
      new Paragraph({ style: "Heading2", children: [new TextRun("1.1 Концепции возникновения денег")] }),
      new Paragraph({ style: "Heading3", children: [new TextRun("1.1.1 Эволюционная концепция")] }),
      new Paragraph({ children: [new TextRun("Основной текст пункта...")] }),

      // === Bibliography — use BiblioHeading ===
      new Paragraph({ style: "BiblioHeading", children: [new TextRun("Список использованных источников")] }),
      // ...numbered references...
    ]
  }]
});

Packer.toBuffer(doc).then(buf => fs.writeFileSync("output.docx", buf));
```

After writing the file, validate:

```bash
python /mnt/skills/public/docx/scripts/office/validate.py output.docx
```

If validation fails, unpack with `unpack.py`, fix the XML, repack with `pack.py`.

### Style cookbook — element-by-element

Each subsection lists the GOST values from the source standard, then the docx-js snippet that produces them. Use the template above as the host for these snippets.

#### Body text (основной текст)

Times New Roman 14 pt, justified, 1.25 cm first-line indent, 1.5 line spacing, no spacing before/after. Set as document default; no override needed in normal paragraphs.

```javascript
new Paragraph({ children: [new TextRun("Обычный текст абзаца.")] })
```

#### Heading 1 — numbered sections

`1 ТЕОРЕТИЧЕСКАЯ ЧАСТЬ` — bold, 18 pt, ALL CAPS (via `allCaps`), justified, 1.25 cm first-line indent, 12 pt after, page break before. Defined in template as `Heading1`.

```javascript
new Paragraph({ style: "Heading1", children: [new TextRun("1 Теоретическая часть")] })
```

Always supply title-cased text — `allCaps: true` upper-cases at render time without converting the underlying text.

#### Heading 1 — special centered sections

`СОДЕРЖАНИЕ`, `ВВЕДЕНИЕ`, `ЗАКЛЮЧЕНИЕ`, `СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ`, `ПРИЛОЖЕНИЯ`, `ТЕОРЕТИЧЕСКАЯ ЧАСТЬ`, `ПРАКТИЧЕСКАЯ ЧАСТЬ`. Defined as `Heading1Centered` in the template.

```javascript
new Paragraph({ style: "Heading1Centered", children: [new TextRun("Введение")] })
```

The bibliography heading uses `BiblioHeading` instead (regular weight, see below).

#### Heading 2 / Heading 3

```javascript
new Paragraph({ style: "Heading2", children: [new TextRun("1.1 Подраздел")] })
new Paragraph({ style: "Heading3", children: [new TextRun("1.1.1 Пункт")] })
```

#### Lists

Bulleted: lowercase start, ends with `;`, last item ends with `.`. Numbered: capitalized start, ends with `.`. Use ONE marker style throughout the document.

```javascript
// bulleted
new Paragraph({
  numbering: { reference: "gost-bullets", level: 0 },
  children: [new TextRun("первый пункт списка;")]
})
new Paragraph({
  numbering: { reference: "gost-bullets", level: 0 },
  children: [new TextRun("последний пункт списка.")]
})

// numbered
new Paragraph({
  numbering: { reference: "gost-numbers", level: 0 },
  children: [new TextRun("Первый пункт.")]
})
```

Never insert bullet glyphs (•, ●, –) by hand into a TextRun — they will be flat text without proper indent or hanging.

#### Tables — caption + content

**Caption** (above table): italic, 12 pt, justified, no first-line indent, 6 pt before, single spacing, keep-with-next. Title format `Таблица X.Y — Название` (em-dash).

**Content**: 12 pt Times New Roman, single line spacing, no first-line indent in cells. Headers are typically bold and centered; body cells justified or centered. No `№ п/п` column.

For continuation across pages: the second part is preceded by a paragraph `Продолжение Таблицы X.Y` formatted as a table caption. To make the header row repeat on each page, set `tableHeader: true` on the header row — this is the docx-js way to suppress the bottom border on the first part automatically.

```javascript
const border  = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

// Caption
new Paragraph({
  spacing: { before: 120, after: 0, line: 240 },
  indent: { firstLine: 0 },
  alignment: AlignmentType.JUSTIFIED,
  keepNext: true,
  children: [new TextRun({
    text: "Таблица 1.1 — Исходные данные",
    font: "Times New Roman", size: 24, italics: true
  })]
})

// Table — note dual width specification (table.width AND each cell.width)
new Table({
  width: { size: 9638, type: WidthType.DXA },     // full GOST content width
  columnWidths: [4819, 4819],
  rows: [
    new TableRow({
      tableHeader: true,                           // repeats on next page if continued
      children: [
        new TableCell({
          borders, width: { size: 4819, type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({
            spacing: { line: 240 }, alignment: AlignmentType.CENTER, indent: { firstLine: 0 },
            children: [new TextRun({ text: "Заголовок", font: "Times New Roman", size: 24, bold: true })]
          })]
        }),
        new TableCell({
          borders, width: { size: 4819, type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({
            spacing: { line: 240 }, alignment: AlignmentType.CENTER, indent: { firstLine: 0 },
            children: [new TextRun({ text: "Значение", font: "Times New Roman", size: 24, bold: true })]
          })]
        })
      ]
    }),
    // ...data rows...
  ]
})

// Add a 6 pt-before paragraph immediately after the table (per §6.6)
new Paragraph({ spacing: { before: 120 }, children: [] })
```

#### Figures (illustrations)

Caption goes **below** the figure: bold, 12 pt, centered, no first-line indent, 6 pt after, single spacing. Format `Рисунок X.Y — Название` (em-dash, no period).

```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  children: [new ImageRun({
    type: "png",
    data: fs.readFileSync("diagram.png"),
    transformation: { width: 450, height: 300 },
    altText: { title: "Диаграмма", description: "Описание", name: "diagram" }
  })]
})

new Paragraph({
  spacing: { before: 0, after: 120, line: 240 }, // 6 pt after
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  children: [new TextRun({
    text: "Рисунок 1.1 — Описание иллюстрации",
    font: "Times New Roman", size: 24, bold: true
  })]
})

// Empty body-text paragraph after the caption to restore vertical rhythm
new Paragraph({ children: [] })
```

#### Code listings

Caption: italic, 12 pt, **left-aligned** (not justified), 6 pt before, single spacing, keep-with-next. Format `Листинг X.Y — Название`.

Content: Courier New 10 pt, left-aligned, single spacing, no first-line indent. The listing is wrapped in a frame border per §8.4 — apply paragraph borders to each line of the listing.

```javascript
const codeBorder = { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 4 };

// Caption
new Paragraph({
  spacing: { before: 120, after: 0, line: 240 },
  alignment: AlignmentType.LEFT,
  indent: { firstLine: 0 },
  keepNext: true,
  children: [new TextRun({
    text: "Листинг 1.1 — Процедура открытия программы",
    font: "Times New Roman", size: 24, italics: true
  })]
})

// Listing body — apply borders to each paragraph; left/right on every line, top on first, bottom on last
new Paragraph({
  spacing: { line: 240 },
  alignment: AlignmentType.LEFT,
  indent: { firstLine: 0 },
  border: { top: codeBorder, left: codeBorder, right: codeBorder },
  children: [new TextRun({ text: "procedure OpenForm(Sender: TObject);", font: "Courier New", size: 20 })]
})
new Paragraph({
  spacing: { line: 240 },
  alignment: AlignmentType.LEFT,
  indent: { firstLine: 0 },
  border: { left: codeBorder, right: codeBorder },
  children: [new TextRun({ text: "begin", font: "Courier New", size: 20 })]
})
new Paragraph({
  spacing: { line: 240 },
  alignment: AlignmentType.LEFT,
  indent: { firstLine: 0 },
  border: { left: codeBorder, right: codeBorder, bottom: codeBorder },
  children: [new TextRun({ text: "end;", font: "Courier New", size: 20 })]
})

new Paragraph({ spacing: { before: 120 }, children: [] }) // 6 pt before the next thing
```

Write every listing as ONE continuous frame, however long, and do not set `keepNext`/`keepLines` on the code lines or `cantSplit` on a frame table — the page split is done afterwards by the script below. A single-cell table (1×1, `Table Grid` borders, code paragraphs inside) is an equally valid frame («в табличной рамке» — this is how the department template does it); the script handles both forms.

#### Перенос листинга — «Продолжение листинга X.Y»

When a listing doesn't fit on the page, its frame is closed at the bottom of the page, and the next page starts with the caption `Продолжение листинга X.Y` followed by a new frame with the remaining lines. What reviewers reject, and what the script checks:

- a continuation caption on a page where the listing's code isn't (the caption at the bottom of one page, the code on the next one);
- a continuation caption that isn't the very first line of its page, or that sits on the wrong page (not the page right after the previous part);
- a break before the page is full — the frame ends halfway down the page and the rest moves on. A split is allowed only when the listing really reaches the bottom margin; a listing that fits on its page must stay whole, with no «Продолжение»;
- a caption formatted unlike a listing caption. GOST format: Times New Roman 12 pt, italic, not bold, left-aligned, no indent, single spacing, keep-with-next, no period at the end; text exactly `Продолжение листинга 1.8` (lowercase «листинга», the same number as the original caption, including appendix numbers like `Б.2`).

Run it as the final step after the document is generated or edited (it needs LibreOffice, `python-docx` and `pdfplumber`). The script lives at `scripts/listing_continuation.py` next to this SKILL.md (use its absolute path):

```bash
python scripts/listing_continuation.py output.docx output_final.docx
# only check, change nothing (e.g. a document the user formatted by hand):
python scripts/listing_continuation.py output.docx --check-only
```

What it does: first it glues any existing parts of each listing back into one frame (so a re-run after edits is safe, and wrong manual splits are fixed), then it lays out the document in LibreOffice, finds the line where each frame crosses the page boundary and cuts it there, inserting the caption with «page break before». After every cut it lays out the document again (a cut shifts everything below it) until all listings pass, then runs the full check and prints a per-listing report: `Листинг 1.2: стр. 3, стр. 4 — 1 продолж.` Exit code 0 means everything passed; 1 means problems remain — they are listed with the listing number and page, so fix them, don't hand over the file with failures. Blank code lines at the cut point are dropped so a frame never starts or ends with an empty line. It also writes `output_final.render.pdf`; look at the pages where listings split before handing the file over.

If the user's department convention labels the last part `Окончание листинга X.Y`, pass `--last-label Окончание`; by default every part is `Продолжение листинга`.

Tell the user that the split points were measured in LibreOffice with metric-compatible fonts: Word normally paginates the same way, but if they change text above a listing, the split must be re-run (just run the script again on the edited file).

#### Bibliography

Use `BiblioHeading` for the title (regular weight, NOT bold — see Critical Rule #3). Entries are body text style, numbered manually (the numbers are part of the citation key referenced in `[1]`, `[4, c.21-25]`).

```javascript
new Paragraph({ style: "BiblioHeading", children: [new TextRun("Список использованных источников")] })

new Paragraph({ children: [new TextRun(
  "1. Громов Г.Р. Очерки информационной технологии. — М.: ИнфоАрт, 2024. — 336 с."
)]})
new Paragraph({ children: [new TextRun(
  "2. Иванов И.И. Заголовок статьи // Журнал. — 2023. — № 4. — С. 12–18."
)]})
```

Entry formats from §9.10:

- **Book**: `Автор И.О. Название. — М.: Издательство, ГГГГ. — NNN с.`
- **Article**: `Автор И.О. Название // Журнал. — ГГГГ. — Т. X. — Вып. Y. — С. NN–MM.`
- **Web**: `Автор И.О. Название ― URL: http://… (Дата обращения: ДД.ММ.ГГГГ)`

City abbreviations: `М.` (Москва), `Л.` (Ленинград), `СПб.` (Санкт-Петербург), `К.` (Киев). No space after `М.:` or `СПб.:`. No space after period in initials: `Громов Г.Р.` (not `Громов Г. Р.`).

Minimum 7 sources; for non-natural-science fields, none older than 5 years.

#### Appendices (приложения)

Each appendix starts on a new page. Letter labeling skips Ё, З, Й, О, Ч, Ь, Ы, Ъ → А, Б, В, Г, Д, Е, Ж, И, К, Л, М, Н, П, Р, С, Т, У, Ф, Х, Ц, Ш, Щ, Э, Ю, Я. Sub-appendices use dot-letter: `Приложение А.3`. Tables and figures inside appendices use letter prefix: `Рисунок Б.5`.

The first page of the section lists all appendices as plain body text. Each appendix's own first page has the word `Приложение А` (Heading 3 style, centered) on the first line, and the appendix title on the second line as plain body text, centered, no spacing before/after.

```javascript
new Paragraph({
  style: "Heading3",
  alignment: AlignmentType.CENTER,            // override Heading3 default
  indent: { firstLine: 0 },
  spacing: { before: 0, after: 0, line: 360 },
  children: [new TextRun("Приложение А")]
})
new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  spacing: { before: 0, after: 0, line: 360 },
  children: [new TextRun("Графический материал")]
})
```

---

## Mode 2 — Editing an existing GOST document

Use the public docx skill's scripts for unpack/edit/pack. The public skill is at `/mnt/skills/public/docx/`.

### Workflow

```bash
# Step 1: Unpack
python /mnt/skills/public/docx/scripts/office/unpack.py document.docx unpacked/

# Step 2: Edit XML files in unpacked/word/ — see references/xml-editing.md for patterns

# Step 3: Repack
python /mnt/skills/public/docx/scripts/office/pack.py unpacked/ output.docx --original document.docx
```

### Common XML edits at a glance

- **Add a tracked-change insertion**: wrap new run in `<w:ins w:author="Claude" w:date="…">`.
- **Add a tracked-change deletion**: replace `<w:r><w:t>…</w:t></w:r>` with `<w:del w:author="Claude" w:date="…"><w:r><w:delText>…</w:delText></w:r></w:del>`.
- **Russian quotes inside XML**: use `&#x00AB;` and `&#x00BB;`.
- **Em-dash with spaces in XML**: `&#x2014;` between spaces, with `xml:space="preserve"` on the `<w:t>` if leading or trailing whitespace.

For full editing patterns including comments, image embedding, paragraph deletion, and rejecting/restoring others' changes, read `references/xml-editing.md`.

---

## Russian typography rules

| Element | Rule | Unicode |
|---|---|---|
| Quotes | `«…»` (always for Russian text) | U+00AB, U+00BB |
| Em-dash (тире) | With spaces, for punctuation: `Слово — это…` | U+2014 |
| Hyphen (дефис) | No spaces, compound words: `физико-математический` | `-` |
| En-dash (минус) | Math context only, e.g., `5 – 2` | U+2013 |
| Numbers | Arabic only (Roman only for quarters / семестры), no case endings on dates or Roman numerals | — |
| Math signs | `+`, `–`, `=`, `<`, `>` only inside formulas; spell out in body text ("плюс", "равно") | — |
| No hyphenation | `перенос слов` not allowed; docx-js default is correct (no auto-hyphenation in settings.xml) | — |

---

## Common docx-js pitfalls

- **Tables need dual widths.** Set `width` on the Table AND on every cell. Use `WidthType.DXA`, never `WidthType.PERCENTAGE`.
- **`ImageRun` requires `type`.** Always pass `type: "png"` (or jpg/gif/etc.).
- **`PageBreak` must live inside a `Paragraph`** — never standalone, or you get invalid XML.
- **Never use `\n`** in text — produce separate `Paragraph` elements.
- **Use `ShadingType.CLEAR`** for table cell shading, not `SOLID`.
- **`TableOfContents` requires built-in HeadingLevel styles** — don't apply custom styles to the heading paragraphs you want included.
- **`outlineLevel` is required on heading styles** for the TOC to find them (0 for H1, 1 for H2, 2 for H3).
- **Override built-in style IDs**: use exactly `"Heading1"`, `"Heading2"`, `"Heading3"`. Custom IDs like `"Heading1Centered"` work for non-built-in styles but won't be picked up by automatic TOCs.
- **Page numbers on title page**: always combine `titlePage: true` in section properties with an empty `first` footer.
- **Empty paragraph after each figure caption**: insert `new Paragraph({ children: [] })` to restore the body text vertical rhythm.
- **Use Russian quotes in code constants too**: `"«...»"`, not `'"..."'`.
- **docx-js writes `<w:pBdr>` children in the wrong order** (top, bottom, left, right; the schema wants top, left, bottom, right), so bordered listing lines fail `validate.py`. `scripts/listing_continuation.py` rewrites every listing frame in schema order, so run it before validating — no separate fix script needed.
- **Never hard-code where a page ends.** Long listings are split only by `scripts/listing_continuation.py` after the document is built.

---

## Reference files

When you need to dig deeper than this SKILL.md:

- **`references/deep-spec.md`** — Full GOST parameter tables verbatim from the standard (every value in every Table 1.1 / 2.1 / etc.), plus rules for numbering, typography edge cases, and the source ГОСТ rules document for cross-checking.
- **`references/xml-editing.md`** — Complete XML reference for editing existing .docx files: tracked changes (insert/delete/reject/restore), comments (with `comment.py`), embedded images, and OOXML schema gotchas.

Read the relevant reference whenever you're about to make a non-trivial choice and want to confirm against the source standard rather than rely on memory.

---

## Dependencies

- `docx` (npm, install globally): generating new documents
- `pandoc`: text extraction with `--track-changes=all`
- `LibreOffice` via `scripts/office/soffice.py`: PDF and `.doc` conversion
- `pdftoppm` (Poppler): converting generated PDFs to images for visual inspection
- `python-docx`, `pdfplumber` (pip): used by `scripts/listing_continuation.py`

