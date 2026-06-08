# DOCX Diploma Formatter Checker — Design Spec
Date: 2026-06-09

## Overview

A purely client-side web application that checks DOCX diploma papers for formatting compliance. Students upload their file, the browser parses the XML locally, and errors are shown as yellow highlights with hover tooltips directly on a rendered view of the document. No server, no data leaves the student's computer.

---

## Users

- **Students** — upload their own diploma DOCX, see formatting errors, fix and re-upload
- **Teacher** — hosts the static site once; no ongoing maintenance needed

---

## Formatting Rules

### Page Margins (checked once per document)

| Margin | Required |
|--------|----------|
| Left   | 3.0 cm   |
| Right  | 1.0 cm   |
| Top    | 2.0 cm   |
| Bottom | 2.0 cm   |

### Body Text Paragraphs

| Parameter       | Required value          |
|-----------------|-------------------------|
| Font            | Times New Roman          |
| Size            | 14 pt                    |
| Bold            | No                       |
| Alignment       | Justify (both)           |
| First line indent | 1.25 cm               |
| Line spacing    | Single (1.0 / 240 twips) |

### Heading Paragraphs

| Parameter       | Required value |
|-----------------|----------------|
| Font            | Times New Roman |
| Size            | 20 pt           |
| Bold            | Yes             |
| Alignment       | Left            |
| First line indent | 0 (none)     |
| Line spacing    | Not checked     |

**Heading detection:** A paragraph is treated as a heading if its Word style name starts with `Heading` or `Заголовок`, OR if it is bold AND its font size is 20 pt.

**Note:** More rules will be added in future iterations. The checker is designed so that adding a new rule means adding one entry to a rules object — no structural changes needed.

---

## Architecture

Pure browser app — no backend, no build tool, no npm.

```
DocumentationChecker/
├── index.html        ← upload UI, document renderer, tooltip markup
├── checker.js        ← DOCX parsing, rule checking, error generation
├── styles.css        ← A4 page layout, yellow highlight, tooltip styling
└── lib/
    └── jszip.min.js  ← only external dependency (vendored, no CDN)
```

### Data Flow

```
Student drops DOCX
      ↓
JSZip unpacks the ZIP in browser memory
      ↓
Parse word/styles.xml  → build style inheritance map (Normal, Heading1, ...)
Parse word/document.xml → extract paragraphs with all run properties
Parse word/settings.xml → extract page margin values
      ↓
Checker iterates paragraphs:
  1. Classify: heading or body
  2. Resolve effective formatting (own value → style default → Normal default)
  3. Compare against rules
  4. Collect errors per paragraph
      ↓
Renderer builds HTML:
  - A4-sized white page containers
  - Each paragraph as a <p> element
  - Paragraphs with errors wrapped in .highlight-error
  - Tooltip div injected after each highlighted paragraph
  - Page margin errors shown in a banner above page 1
      ↓
User hovers yellow paragraph → CSS/JS shows tooltip with errors + recommendations
```

---

## UI States

### State 1 — Upload
- Centered drop zone: drag-and-drop or click to browse
- Accepts `.docx` files only
- Simple, no distractions

### State 2 — Processing
- Spinner with "Анализируем файл..." message
- Runs synchronously (DOCX parsing is fast, < 1 second for typical files)

### State 3 — Results
- Summary bar at top: filename, error count, "Проверить другой файл" button
- If page margin errors exist: red banner above the document
- Document rendered as A4 pages in sequence
- Paragraphs with errors: yellow background + left border
- Hover → tooltip appears above the paragraph showing:
  - Parameter name
  - Current value → Required value
  - Short recommendation ("Измените шрифт на Times New Roman")
- Paragraphs without errors: no styling

### Error tooltip format (per error inside tooltip)

```
Шрифт
Текущее: Arial  →  Требуемое: Times New Roman
Рекомендация: Выделите текст и смените шрифт.
```

---

## DOCX XML Mapping

All values sourced from Open XML spec:

| Parameter        | XML location                                  | Unit / notes                        |
|------------------|-----------------------------------------------|--------------------------------------|
| Font name        | `<w:rFonts w:ascii="...">` in `<w:rPr>`       | String                               |
| Font size        | `<w:sz w:val="...">` in `<w:rPr>`             | Half-points → divide by 2 for pt    |
| Bold             | `<w:b/>` presence in `<w:rPr>`                | Presence = true                      |
| Alignment        | `<w:jc w:val="...">` in `<w:pPr>`             | "both" = justify, "left" = left     |
| First line indent| `<w:ind w:firstLine="...">` in `<w:pPr>`      | Twips → divide by 567 for cm        |
| Line spacing     | `<w:spacing w:line="..." w:lineRule="auto">`   | 240 twips + auto = single           |
| Page margins     | `<w:pgMar>` in `<w:sectPr>`                   | Twips → divide by 567 for cm        |
| Paragraph style  | `<w:pStyle w:val="...">` in `<w:pPr>`         | String (e.g. "Heading1", "Normal")  |

**Style inheritance:** if a run has no explicit font/size, look up the paragraph's named style in `styles.xml`, then fall back to the `Normal` style defaults.

---

## Tolerance

Numeric comparisons use a small tolerance to avoid false positives from floating-point conversion:

- Font size: ± 0 pt (exact match required)
- Margins: ± 0.05 cm
- First line indent: ± 0.05 cm
- Line spacing: exact match on raw twip value (240 for single)

---

## Deployment

1. Push project folder to a GitHub repository
2. Connect repository to Netlify (free tier, one click)
3. Netlify serves the static files — students visit the URL

No server, no database, no ongoing cost.

---

## Constraints & Limitations

- Page numbers are not available in DOCX XML — the rendered HTML pagination will differ from Word's pagination. Paragraphs are identified by their position and first ~60 characters of text.
- The rendered document is not a pixel-perfect replica of Word. Layout is approximate but sufficient to locate errors.
- Heavily macro-driven or password-protected DOCX files may fail to parse.
- Line spacing check only applies to body text paragraphs (headings excluded).
