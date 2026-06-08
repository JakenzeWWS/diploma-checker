# DOCX Diploma Formatter Checker — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a pure client-side web app where students upload a DOCX diploma and see formatting errors highlighted inline on a rendered view of their document.

**Architecture:** One HTML page, zero backend. JSZip unpacks the DOCX in-browser, vanilla JS parses the XML and checks formatting rules, the document is rendered as HTML with yellow-highlighted error paragraphs and hover tooltips.

**Tech Stack:** HTML5, CSS3, vanilla JavaScript (ES2020), JSZip 3.x (vendored in `lib/`)

---

## File Structure

| File | Responsibility |
|---|---|
| `lib/jszip.min.js` | JSZip library — unzips DOCX in browser |
| `checker.js` | All DOCX parsing + rule checking. Exports `DocxChecker` global. Pure functions only, no DOM. |
| `index.html` | Single page: upload zone → loading → results view + renderer + event wiring |
| `styles.css` | A4 page layout, upload zone styling, yellow highlight, tooltip |
| `test.html` | Lightweight unit test runner — open in browser to verify `checker.js` |

---

## DOCX XML Cheat-Sheet (reference for all tasks)

```
work.docx (ZIP)
├── word/document.xml   — paragraphs, runs, paragraph/run properties, sectPr (margins)
└── word/styles.xml     — style definitions (Normal, Heading1, ...)

Key XML elements (namespace W = http://schemas.openxmlformats.org/wordprocessingml/2006/main):
  <w:p>              paragraph
  <w:pPr>            paragraph properties
    <w:pStyle w:val="Normal|Heading1|...">  style id
    <w:jc w:val="both|left|center|right">  alignment (both = justify)
    <w:ind w:firstLine="709">              first-line indent in twips
    <w:spacing w:line="240" w:lineRule="auto">  line spacing
  <w:r>              run (text with uniform formatting)
  <w:rPr>            run properties
    <w:rFonts w:ascii="Times New Roman">   font
    <w:sz w:val="28">                      font size in HALF-POINTS (÷2 = pt)
    <w:b/>                                 bold (presence = true)
  <w:t>              text content
  <w:pgMar w:left="1701" w:right="567" w:top="1134" w:bottom="1134">  page margins in twips

Unit conversions:
  1 cm = 566.929 twips  (use CM_PER_TWIP = 2.54 / 1440)
  14 pt = <w:sz w:val="28">  (half-points)
  20 pt = <w:sz w:val="40">
  single spacing = <w:spacing w:line="240" w:lineRule="auto">
  1.25 cm indent = ~709 twips
  left 3 cm = ~1701 twips
```

---

## Task 1: Project Setup

**Files:**
- Create: `lib/jszip.min.js`
- Create: `checker.js` (empty)
- Create: `index.html` (empty)
- Create: `styles.css` (empty)
- Create: `test.html` (empty)

- [ ] **Step 1: Create directory structure**

```bash
cd /Users/Zhangerkhan_1/Desktop/DocumentationChecker
mkdir -p lib
```

- [ ] **Step 2: Download JSZip**

```bash
curl -L https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js -o lib/jszip.min.js
```

Verify it downloaded (should be ~40KB):
```bash
ls -lh lib/jszip.min.js
```
Expected: file exists, size ~40-50KB.

- [ ] **Step 3: Create empty placeholder files**

```bash
touch checker.js index.html styles.css test.html
```

- [ ] **Step 4: Initialize git**

```bash
git init
echo "node_modules/" > .gitignore
git add .
git commit -m "chore: project skeleton with jszip"
```

---

## Task 2: checker.js — Skeleton, RULES, XML Helpers

**Files:**
- Modify: `checker.js`

- [ ] **Step 1: Write the full skeleton**

Write this to `checker.js`:

```javascript
const DocxChecker = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const CM_PER_TWIP = 2.54 / 1440;

  const RULES = {
    body: {
      font: 'Times New Roman',
      sizePt: 14,
      bold: false,
      alignment: 'both',
      firstLineIndentCm: 1.25,
      lineSpacingTwips: 240,
      lineSpacingRule: 'auto',
    },
    heading: {
      font: 'Times New Roman',
      sizePt: 20,
      bold: true,
      alignment: 'left',
      firstLineIndentCm: 0,
    },
    margins: {
      leftCm: 3.0,
      rightCm: 1.0,
      topCm: 2.0,
      bottomCm: 2.0,
      toleranceCm: 0.05,
    },
  };

  function parseXml(xmlString) {
    return new DOMParser().parseFromString(xmlString, 'application/xml');
  }

  // Get first child element by tag in W namespace
  function el(parent, tag) {
    return parent.getElementsByTagNameNS(W, tag)[0] || null;
  }

  // Get all child elements by tag in W namespace
  function els(parent, tag) {
    return Array.from(parent.getElementsByTagNameNS(W, tag));
  }

  // Get attribute — tries W namespace first, then unprefixed
  function attr(element, name) {
    if (!element) return null;
    return element.getAttributeNS(W, name) || element.getAttribute(name) || null;
  }

  function twipsToCm(twips) {
    if (twips === null || twips === undefined) return null;
    return parseFloat((parseInt(twips, 10) * CM_PER_TWIP).toFixed(3));
  }

  // --- Public API (stubs for now, implemented in later tasks) ---

  function parseStyles(xmlString) { return new Map(); }
  function parseMargins(xmlString) { return null; }
  function parseParagraphs(xmlString) { return []; }
  function resolveFormatting(paragraph, styleMap) { return {}; }
  function classifyParagraph(paragraph, styleMap) { return 'body'; }
  function checkParagraph(paragraph, type, styleMap) { return []; }
  function checkMargins(margins) { return []; }
  async function checkDocument(zip) { return { marginErrors: [], paragraphResults: [] }; }

  return {
    RULES, parseStyles, parseMargins, parseParagraphs,
    resolveFormatting, classifyParagraph, checkParagraph,
    checkMargins, checkDocument,
  };
})();
```

- [ ] **Step 2: Open `test.html` in browser to confirm no syntax errors**

Write this to `test.html`:

```html
<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8"><title>Tests</title></head>
<body>
<h2>DocxChecker Tests</h2>
<pre id="out"></pre>
<script src="lib/jszip.min.js"></script>
<script src="checker.js"></script>
<script>
const out = document.getElementById('out');
const results = [];

function assert(desc, condition) {
  results.push({ desc, pass: !!condition });
}

function run() {
  // Task 2: basic structure
  assert('DocxChecker exists', typeof DocxChecker === 'object');
  assert('RULES.body.font is Times New Roman', DocxChecker.RULES.body.font === 'Times New Roman');
  assert('RULES.body.sizePt is 14', DocxChecker.RULES.body.sizePt === 14);
  assert('RULES.heading.sizePt is 20', DocxChecker.RULES.heading.sizePt === 20);
  assert('RULES.margins.leftCm is 3', DocxChecker.RULES.margins.leftCm === 3.0);

  const passed = results.filter(r => r.pass).length;
  out.textContent = results.map(r => (r.pass ? '✓' : '✗') + ' ' + r.desc).join('\n');
  out.textContent += `\n\n${passed}/${results.length} passed`;
}

run();
</script>
</body>
</html>
```

Open `test.html` in a browser. Expected: all 5 checks show ✓.

- [ ] **Step 3: Commit**

```bash
git add checker.js test.html
git commit -m "feat: checker.js skeleton with RULES and XML helpers"
```

---

## Task 3: checker.js — parseStyles()

**Files:**
- Modify: `checker.js`
- Modify: `test.html`

- [ ] **Step 1: Replace the `parseStyles` stub with the real implementation**

In `checker.js`, replace:
```javascript
function parseStyles(xmlString) { return new Map(); }
```

With:
```javascript
function parseStyles(xmlString) {
  const doc = parseXml(xmlString);
  const styleMap = new Map();

  for (const styleEl of els(doc, 'style')) {
    const styleId = attr(styleEl, 'styleId');
    if (!styleId) continue;

    const nameEl = el(styleEl, 'name');
    const normalizedName = (attr(nameEl, 'val') || '').toLowerCase();

    const entry = { normalizedName };

    // Run-level defaults (font, size, bold)
    const rPr = el(styleEl, 'rPr');
    if (rPr) {
      const fonts = el(rPr, 'rFonts');
      if (fonts) {
        entry.font = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || null;
      }

      const sz = el(rPr, 'sz');
      if (sz) {
        const val = attr(sz, 'val');
        if (val) entry.sizePt = parseInt(val, 10) / 2;
      }

      const bEl = el(rPr, 'b');
      if (bEl) {
        const val = attr(bEl, 'val');
        entry.bold = val === null || (val !== '0' && val !== 'false' && val !== 'off');
      }
    }

    // Paragraph-level defaults (alignment, indent, spacing)
    const pPr = el(styleEl, 'pPr');
    if (pPr) {
      const jc = el(pPr, 'jc');
      if (jc) entry.alignment = attr(jc, 'val');

      const ind = el(pPr, 'ind');
      if (ind) {
        const fl = attr(ind, 'firstLine');
        if (fl !== null) entry.firstLineIndentCm = twipsToCm(fl);
      }

      const spacing = el(pPr, 'spacing');
      if (spacing) {
        const line = attr(spacing, 'line');
        const lineRule = attr(spacing, 'lineRule');
        if (line) entry.lineSpacingTwips = parseInt(line, 10);
        if (lineRule) entry.lineSpacingRule = lineRule;
      }
    }

    styleMap.set(styleId, entry);
  }

  return styleMap;
}
```

- [ ] **Step 2: Add tests for parseStyles in test.html**

In `test.html`, inside the `run()` function, after the existing asserts, add:

```javascript
// Task 3: parseStyles
const FAKE_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="28"/>
    </w:rPr>
    <w:pPr>
      <w:jc w:val="both"/>
      <w:spacing w:line="240" w:lineRule="auto"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman"/>
      <w:sz w:val="40"/>
      <w:b/>
    </w:rPr>
    <w:pPr>
      <w:jc w:val="left"/>
    </w:pPr>
  </w:style>
</w:styles>`;

const styleMap = DocxChecker.parseStyles(FAKE_STYLES_XML);
assert('parseStyles returns a Map', styleMap instanceof Map);
assert('Normal style exists', styleMap.has('Normal'));
assert('Normal font is Times New Roman', styleMap.get('Normal').font === 'Times New Roman');
assert('Normal sizePt is 14', styleMap.get('Normal').sizePt === 14);
assert('Normal alignment is both', styleMap.get('Normal').alignment === 'both');
assert('Normal lineSpacingTwips is 240', styleMap.get('Normal').lineSpacingTwips === 240);
assert('Heading1 exists', styleMap.has('Heading1'));
assert('Heading1 normalizedName is heading 1', styleMap.get('Heading1').normalizedName === 'heading 1');
assert('Heading1 sizePt is 20', styleMap.get('Heading1').sizePt === 20);
assert('Heading1 bold is true', styleMap.get('Heading1').bold === true);
```

- [ ] **Step 3: Open test.html in browser**

Expected: all new asserts show ✓. If any fail, check the `attr()` helper — try logging `element.outerHTML` to inspect the parsed XML.

- [ ] **Step 4: Commit**

```bash
git add checker.js test.html
git commit -m "feat: parseStyles — build style inheritance map from styles.xml"
```

---

## Task 4: checker.js — parseMargins() and parseParagraphs()

**Files:**
- Modify: `checker.js`
- Modify: `test.html`

- [ ] **Step 1: Replace the `parseMargins` stub**

In `checker.js`, replace:
```javascript
function parseMargins(xmlString) { return null; }
```

With:
```javascript
function parseMargins(xmlString) {
  const doc = parseXml(xmlString);
  const pgMar = el(doc, 'pgMar');
  if (!pgMar) return null;

  return {
    leftCm:   twipsToCm(attr(pgMar, 'left')),
    rightCm:  twipsToCm(attr(pgMar, 'right')),
    topCm:    twipsToCm(attr(pgMar, 'top')),
    bottomCm: twipsToCm(attr(pgMar, 'bottom')),
  };
}
```

- [ ] **Step 2: Replace the `parseParagraphs` stub**

In `checker.js`, replace:
```javascript
function parseParagraphs(xmlString) { return []; }
```

With:
```javascript
function parseParagraphs(xmlString) {
  const doc = parseXml(xmlString);
  const paragraphs = [];

  for (const pEl of els(doc, 'p')) {
    const pPr = el(pEl, 'pPr');

    // Paragraph style ID
    let styleId = 'Normal';
    if (pPr) {
      const pStyle = el(pPr, 'pStyle');
      if (pStyle) styleId = attr(pStyle, 'val') || 'Normal';
    }

    // Paragraph-level property overrides
    const paraOverride = {};
    if (pPr) {
      const jc = el(pPr, 'jc');
      if (jc) paraOverride.alignment = attr(jc, 'val');

      const ind = el(pPr, 'ind');
      if (ind) {
        const fl = attr(ind, 'firstLine');
        if (fl !== null) paraOverride.firstLineIndentCm = twipsToCm(fl);
      }

      const spacing = el(pPr, 'spacing');
      if (spacing) {
        const line = attr(spacing, 'line');
        const lineRule = attr(spacing, 'lineRule');
        if (line) paraOverride.lineSpacingTwips = parseInt(line, 10);
        if (lineRule) paraOverride.lineSpacingRule = lineRule;
      }
    }

    // Run-level formatting — collect from all runs
    const runFormats = [];
    for (const rEl of els(pEl, 'r')) {
      const rPr = el(rEl, 'rPr');
      const runFmt = {};
      if (rPr) {
        const fonts = el(rPr, 'rFonts');
        if (fonts) runFmt.font = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || null;

        const sz = el(rPr, 'sz');
        if (sz) {
          const val = attr(sz, 'val');
          if (val) runFmt.sizePt = parseInt(val, 10) / 2;
        }

        const bEl = el(rPr, 'b');
        if (bEl) {
          const val = attr(bEl, 'val');
          runFmt.bold = val === null || (val !== '0' && val !== 'false' && val !== 'off');
        }
      }
      runFormats.push(runFmt);
    }

    const text = els(pEl, 't').map(t => t.textContent).join('');
    paragraphs.push({ styleId, paraOverride, runFormats, text });
  }

  return paragraphs;
}
```

- [ ] **Step 3: Add tests in test.html**

```javascript
// Task 4: parseMargins
const FAKE_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:jc w:val="both"/>
        <w:spacing w:line="240" w:lineRule="auto"/>
        <w:ind w:firstLine="709"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman"/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>Основной текст диплома.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:sz w:val="40"/>
          <w:b/>
        </w:rPr>
        <w:t>ВВЕДЕНИЕ</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgMar w:left="1701" w:right="567" w:top="1134" w:bottom="1134"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const margins = DocxChecker.parseMargins(FAKE_DOC_XML);
assert('parseMargins returns object', margins !== null);
assert('left margin ~3.0 cm', Math.abs(margins.leftCm - 3.0) < 0.05);
assert('right margin ~1.0 cm', Math.abs(margins.rightCm - 1.0) < 0.05);
assert('top margin ~2.0 cm', Math.abs(margins.topCm - 2.0) < 0.05);
assert('bottom margin ~2.0 cm', Math.abs(margins.bottomCm - 2.0) < 0.05);

// Task 4: parseParagraphs
const paras = DocxChecker.parseParagraphs(FAKE_DOC_XML);
assert('parseParagraphs returns 2 paragraphs', paras.length === 2);
assert('first para styleId is Normal', paras[0].styleId === 'Normal');
assert('first para text contains текст', paras[0].text.includes('текст'));
assert('first para alignment override is both', paras[0].paraOverride.alignment === 'both');
assert('first para indent ~1.25 cm', Math.abs(paras[0].paraOverride.firstLineIndentCm - 1.25) < 0.05);
assert('first para run font is Times New Roman', paras[0].runFormats[0].font === 'Times New Roman');
assert('first para run sizePt is 14', paras[0].runFormats[0].sizePt === 14);
assert('second para styleId is Heading1', paras[1].styleId === 'Heading1');
assert('second para text is ВВЕДЕНИЕ', paras[1].text === 'ВВЕДЕНИЕ');
assert('second para run bold is true', paras[1].runFormats[0].bold === true);
```

- [ ] **Step 4: Run tests in browser — expect all ✓**

- [ ] **Step 5: Commit**

```bash
git add checker.js test.html
git commit -m "feat: parseMargins and parseParagraphs"
```

---

## Task 5: checker.js — resolveFormatting() and classifyParagraph()

**Files:**
- Modify: `checker.js`
- Modify: `test.html`

- [ ] **Step 1: Replace `resolveFormatting` stub**

```javascript
function resolveFormatting(paragraph, styleMap) {
  const normalStyle = styleMap.get('Normal') || {};
  const paraStyle = styleMap.get(paragraph.styleId) || {};

  // For run-level props, collect all non-null values across runs
  const runFonts = paragraph.runFormats.map(r => r.font).filter(Boolean);
  const runSizes = paragraph.runFormats.map(r => r.sizePt).filter(v => v !== undefined);
  const runBolds = paragraph.runFormats.map(r => r.bold).filter(v => v !== undefined);

  return {
    // Run-level: first explicit value from runs → style default → Normal default
    font:   runFonts[0]  ?? paraStyle.font  ?? normalStyle.font  ?? null,
    sizePt: runSizes[0]  ?? paraStyle.sizePt ?? normalStyle.sizePt ?? null,
    bold:   runBolds[0]  ?? paraStyle.bold  ?? normalStyle.bold  ?? false,
    // Collect ALL fonts/sizes to detect mixed formatting
    allFonts: [...new Set(runFonts)],
    allSizes: [...new Set(runSizes)],

    // Paragraph-level: paragraph override → style → Normal
    alignment: paragraph.paraOverride.alignment
               ?? paraStyle.alignment
               ?? normalStyle.alignment
               ?? 'left',
    firstLineIndentCm: paragraph.paraOverride.firstLineIndentCm
                       ?? paraStyle.firstLineIndentCm
                       ?? normalStyle.firstLineIndentCm
                       ?? 0,
    lineSpacingTwips: paragraph.paraOverride.lineSpacingTwips
                      ?? paraStyle.lineSpacingTwips
                      ?? normalStyle.lineSpacingTwips
                      ?? null,
    lineSpacingRule: paragraph.paraOverride.lineSpacingRule
                     ?? paraStyle.lineSpacingRule
                     ?? normalStyle.lineSpacingRule
                     ?? null,
  };
}
```

- [ ] **Step 2: Replace `classifyParagraph` stub**

```javascript
function classifyParagraph(paragraph, styleMap) {
  const styleEntry = styleMap.get(paragraph.styleId) || {};
  const normalizedName = styleEntry.normalizedName || '';

  // Check by Word style name (works with both English and Russian Word)
  if (normalizedName.startsWith('heading') || normalizedName.startsWith('заголовок')) {
    return 'heading';
  }

  // Fallback: bold + size 20 (for students who formatted manually)
  const fmt = resolveFormatting(paragraph, styleMap);
  if (fmt.bold === true && fmt.sizePt === 20) return 'heading';

  return 'body';
}
```

- [ ] **Step 3: Add tests in test.html**

```javascript
// Task 5: resolveFormatting and classifyParagraph
// reuse styleMap and paras from Task 3/4 tests above
const fmt0 = DocxChecker.resolveFormatting(paras[0], styleMap);
assert('resolveFormatting: font is Times New Roman', fmt0.font === 'Times New Roman');
assert('resolveFormatting: sizePt is 14', fmt0.sizePt === 14);
assert('resolveFormatting: alignment is both', fmt0.alignment === 'both');
assert('resolveFormatting: indent ~1.25 cm', Math.abs(fmt0.firstLineIndentCm - 1.25) < 0.05);
assert('resolveFormatting: lineSpacingTwips is 240', fmt0.lineSpacingTwips === 240);

const fmt1 = DocxChecker.resolveFormatting(paras[1], styleMap);
assert('resolveFormatting heading: sizePt is 20', fmt1.sizePt === 20);
assert('resolveFormatting heading: bold is true', fmt1.bold === true);

assert('classifyParagraph: Normal para is body', DocxChecker.classifyParagraph(paras[0], styleMap) === 'body');
assert('classifyParagraph: Heading1 para is heading', DocxChecker.classifyParagraph(paras[1], styleMap) === 'heading');
```

- [ ] **Step 4: Run tests — expect all ✓**

- [ ] **Step 5: Commit**

```bash
git add checker.js test.html
git commit -m "feat: resolveFormatting and classifyParagraph"
```

---

## Task 6: checker.js — checkParagraph() and checkMargins()

**Files:**
- Modify: `checker.js`
- Modify: `test.html`

- [ ] **Step 1: Replace `checkParagraph` stub**

```javascript
function checkParagraph(paragraph, type, styleMap) {
  if (!paragraph.text.trim()) return []; // skip empty lines

  const fmt = resolveFormatting(paragraph, styleMap);
  const rules = RULES[type];
  const errors = [];

  // Font — check all unique fonts found in runs
  const badFonts = fmt.allFonts.filter(f => f !== rules.font);
  if (badFonts.length > 0) {
    errors.push({
      param: 'Шрифт',
      current: [...new Set(badFonts)].join(', '),
      required: rules.font,
      recommendation: `Выделите текст и смените шрифт на «${rules.font}».`,
    });
  }

  // Size — check all unique sizes found in runs
  const badSizes = fmt.allSizes.filter(s => s !== rules.sizePt);
  if (badSizes.length > 0) {
    errors.push({
      param: 'Размер шрифта',
      current: badSizes.map(s => s + ' pt').join(', '),
      required: rules.sizePt + ' pt',
      recommendation: `Выделите текст и установите размер ${rules.sizePt} pt.`,
    });
  }

  // Bold (body only — body text must not be bold)
  if (type === 'body' && fmt.bold === true) {
    errors.push({
      param: 'Жирный шрифт',
      current: 'Да',
      required: 'Нет',
      recommendation: 'Уберите жирное начертание в основном тексте.',
    });
  }

  // Alignment
  const alignment = fmt.alignment || 'left';
  if (alignment !== rules.alignment) {
    const labels = { both: 'По ширине', left: 'По левому краю', center: 'По центру', right: 'По правому краю' };
    errors.push({
      param: 'Выравнивание',
      current: labels[alignment] || alignment,
      required: labels[rules.alignment] || rules.alignment,
      recommendation: `Установите выравнивание: «${labels[rules.alignment]}».`,
    });
  }

  // First line indent
  const indent = fmt.firstLineIndentCm ?? 0;
  if (Math.abs(indent - rules.firstLineIndentCm) > 0.05) {
    errors.push({
      param: 'Отступ первой строки',
      current: indent.toFixed(2) + ' см',
      required: rules.firstLineIndentCm + ' см',
      recommendation: `Установите отступ первой строки ${rules.firstLineIndentCm} см.`,
    });
  }

  // Line spacing (body only)
  if (type === 'body') {
    const spacingOk =
      fmt.lineSpacingTwips === rules.lineSpacingTwips &&
      fmt.lineSpacingRule === rules.lineSpacingRule;
    if (!spacingOk && fmt.lineSpacingTwips !== null) {
      const currentVal = fmt.lineSpacingTwips
        ? (fmt.lineSpacingTwips / 240).toFixed(1)
        : 'не определён';
      errors.push({
        param: 'Межстрочный интервал',
        current: currentVal,
        required: '1.0 (одинарный)',
        recommendation: 'Выделите текст, откройте «Межстрочный интервал» и выберите «Одинарный».',
      });
    }
  }

  return errors;
}
```

- [ ] **Step 2: Replace `checkMargins` stub**

```javascript
function checkMargins(margins) {
  if (!margins) return [];

  const tol = RULES.margins.toleranceCm;
  const checks = [
    { key: 'leftCm',   label: 'Левое поле',   required: RULES.margins.leftCm },
    { key: 'rightCm',  label: 'Правое поле',  required: RULES.margins.rightCm },
    { key: 'topCm',    label: 'Верхнее поле', required: RULES.margins.topCm },
    { key: 'bottomCm', label: 'Нижнее поле',  required: RULES.margins.bottomCm },
  ];

  return checks
    .filter(({ key, required }) => Math.abs(margins[key] - required) > tol)
    .map(({ label, key, required }) => ({
      param: label,
      current: margins[key].toFixed(2) + ' см',
      required: required.toFixed(1) + ' см',
      recommendation: `Установите ${label.toLowerCase()} ${required} см в настройках полей документа.`,
    }));
}
```

- [ ] **Step 3: Add tests in test.html**

```javascript
// Task 6: checkParagraph and checkMargins

// Body paragraph with correct formatting → no errors
const correctBodyErrors = DocxChecker.checkParagraph(paras[0], 'body', styleMap);
assert('correct body paragraph has no errors', correctBodyErrors.length === 0);

// Heading paragraph with correct formatting → no errors
const correctHeadingErrors = DocxChecker.checkParagraph(paras[1], 'heading', styleMap);
assert('correct heading paragraph has no errors', correctHeadingErrors.length === 0);

// Body paragraph with wrong font
const wrongFontPara = {
  styleId: 'Normal', text: 'Текст с неверным шрифтом.',
  paraOverride: { alignment: 'both', firstLineIndentCm: 1.25, lineSpacingTwips: 240, lineSpacingRule: 'auto' },
  runFormats: [{ font: 'Arial', sizePt: 14, bold: false }],
};
const wrongFontErrors = DocxChecker.checkParagraph(wrongFontPara, 'body', styleMap);
assert('wrong font produces error', wrongFontErrors.some(e => e.param === 'Шрифт'));
assert('wrong font error shows Arial as current', wrongFontErrors.find(e => e.param === 'Шрифт').current === 'Arial');

// Wrong size
const wrongSizePara = {
  styleId: 'Normal', text: 'Текст с неверным размером.',
  paraOverride: { alignment: 'both', firstLineIndentCm: 1.25, lineSpacingTwips: 240, lineSpacingRule: 'auto' },
  runFormats: [{ font: 'Times New Roman', sizePt: 12 }],
};
const wrongSizeErrors = DocxChecker.checkParagraph(wrongSizePara, 'body', styleMap);
assert('wrong size produces error', wrongSizeErrors.some(e => e.param === 'Размер шрифта'));

// Correct margins → no errors
const correctMarginErrors = DocxChecker.checkMargins(margins);
assert('correct margins have no errors', correctMarginErrors.length === 0);

// Wrong left margin
const wrongMargins = { leftCm: 2.5, rightCm: 1.0, topCm: 2.0, bottomCm: 2.0 };
const wrongMarginErrors = DocxChecker.checkMargins(wrongMargins);
assert('wrong left margin produces error', wrongMarginErrors.some(e => e.param === 'Левое поле'));
```

- [ ] **Step 4: Run tests — expect all ✓**

- [ ] **Step 5: Commit**

```bash
git add checker.js test.html
git commit -m "feat: checkParagraph and checkMargins rule checkers"
```

---

## Task 7: checker.js — checkDocument() Orchestrator

**Files:**
- Modify: `checker.js`

- [ ] **Step 1: Replace `checkDocument` stub**

```javascript
async function checkDocument(zip) {
  const docXml    = await zip.file('word/document.xml').async('string');
  const stylesXml = await zip.file('word/styles.xml').async('string');

  const styleMap = parseStyles(stylesXml);
  const margins  = parseMargins(docXml);
  const rawParagraphs = parseParagraphs(docXml);

  const marginErrors = checkMargins(margins);

  const paragraphResults = rawParagraphs.map((para, index) => {
    const type   = classifyParagraph(para, styleMap);
    const errors = checkParagraph(para, type, styleMap);
    return {
      index,
      text:   para.text,
      type,
      errors,
      hasErrors: errors.length > 0,
    };
  });

  const totalErrors = marginErrors.length
    + paragraphResults.filter(p => p.hasErrors).length;

  return { marginErrors, paragraphResults, totalErrors };
}
```

- [ ] **Step 2: Commit**

```bash
git add checker.js
git commit -m "feat: checkDocument async orchestrator"
```

---

## Task 8: index.html — HTML Skeleton (3 States)

**Files:**
- Modify: `index.html`

- [ ] **Step 1: Write the full HTML**

```html
<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Проверка оформления диплома</title>
  <link rel="stylesheet" href="styles.css" />
</head>
<body>

  <!-- STATE 1: Upload -->
  <div id="state-upload">
    <div class="upload-container">
      <h1>Проверка оформления дипломной работы</h1>
      <p class="subtitle">Загрузите файл .docx — приложение проверит шрифт, поля, отступы и межстрочный интервал.</p>
      <label id="drop-zone" for="file-input">
        <div class="drop-icon">📄</div>
        <div class="drop-primary">Перетащите файл сюда</div>
        <div class="drop-secondary">или нажмите для выбора</div>
        <div class="drop-hint">Принимается только .docx</div>
      </label>
      <input type="file" id="file-input" accept=".docx" />
      <p id="upload-error" class="upload-error" hidden></p>
    </div>
  </div>

  <!-- STATE 2: Loading -->
  <div id="state-loading" hidden>
    <div class="loading-container">
      <div class="spinner"></div>
      <p>Анализируем файл...</p>
    </div>
  </div>

  <!-- STATE 3: Results -->
  <div id="state-results" hidden>
    <div id="summary-bar">
      <span id="summary-filename"></span>
      <span id="summary-errors"></span>
      <button id="btn-reset">Проверить другой файл</button>
    </div>
    <div id="margin-banner" hidden>
      <strong>Поля страницы:</strong>
      <ul id="margin-error-list"></ul>
    </div>
    <div id="document-view"></div>
  </div>

  <script src="lib/jszip.min.js"></script>
  <script src="checker.js"></script>
  <script src="app.js"></script>
</body>
</html>
```

Note: the wiring logic will go in `app.js` (filled in Tasks 10 and 11). Using a separate file keeps `index.html` clean.

- [ ] **Step 2: Create empty app.js**

```bash
touch app.js
```

- [ ] **Step 3: Open index.html in browser**

Expected: blank page, no console errors.

- [ ] **Step 4: Commit**

```bash
git add index.html app.js
git commit -m "feat: index.html HTML skeleton with 3 states"
```

---

## Task 9: styles.css — Full Styling

**Files:**
- Modify: `styles.css`

- [ ] **Step 1: Write the full CSS**

```css
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: #f0f0f0;
  color: #222;
  min-height: 100vh;
}

/* ── Upload State ─────────────────────────────────────── */

#state-upload {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
  padding: 2rem;
}

.upload-container {
  text-align: center;
  max-width: 520px;
  width: 100%;
}

.upload-container h1 {
  font-size: 1.6rem;
  margin-bottom: 0.5rem;
  color: #1a1a2e;
}

.subtitle {
  color: #555;
  margin-bottom: 2rem;
  line-height: 1.5;
}

#drop-zone {
  display: block;
  border: 2px dashed #aaa;
  border-radius: 12px;
  padding: 3rem 2rem;
  cursor: pointer;
  transition: border-color 0.2s, background 0.2s;
  background: #fff;
}

#drop-zone:hover, #drop-zone.drag-over {
  border-color: #3b82f6;
  background: #eff6ff;
}

.drop-icon { font-size: 3rem; margin-bottom: 0.75rem; }
.drop-primary { font-size: 1.1rem; font-weight: 600; margin-bottom: 0.25rem; }
.drop-secondary { color: #666; margin-bottom: 0.5rem; }
.drop-hint { font-size: 0.8rem; color: #999; }

#file-input { display: none; }

.upload-error {
  margin-top: 1rem;
  color: #dc2626;
  font-size: 0.9rem;
}

/* ── Loading State ────────────────────────────────────── */

#state-loading {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
  flex-direction: column;
  gap: 1rem;
  color: #555;
}

.spinner {
  width: 48px;
  height: 48px;
  border: 4px solid #e5e7eb;
  border-top-color: #3b82f6;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin { to { transform: rotate(360deg); } }

/* ── Summary Bar ──────────────────────────────────────── */

#summary-bar {
  position: sticky;
  top: 0;
  z-index: 100;
  display: flex;
  align-items: center;
  gap: 1rem;
  padding: 0.75rem 1.5rem;
  background: #1a1a2e;
  color: #fff;
  flex-wrap: wrap;
}

#summary-filename {
  font-weight: 600;
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

#summary-errors {
  background: #ef4444;
  color: #fff;
  padding: 0.2rem 0.7rem;
  border-radius: 999px;
  font-size: 0.85rem;
  white-space: nowrap;
}

#summary-errors.no-errors { background: #22c55e; }

#btn-reset {
  background: transparent;
  border: 1px solid #fff;
  color: #fff;
  padding: 0.4rem 1rem;
  border-radius: 6px;
  cursor: pointer;
  font-size: 0.85rem;
  white-space: nowrap;
}

#btn-reset:hover { background: rgba(255,255,255,0.15); }

/* ── Margin Error Banner ──────────────────────────────── */

#margin-banner {
  background: #fef2f2;
  border-left: 4px solid #ef4444;
  padding: 0.75rem 1.5rem;
  margin: 1rem auto;
  max-width: 794px;
  border-radius: 0 6px 6px 0;
}

#margin-banner strong { color: #b91c1c; }
#margin-error-list { margin-top: 0.4rem; padding-left: 1.2rem; color: #7f1d1d; font-size: 0.9rem; }
#margin-error-list li { margin-bottom: 0.2rem; }

/* ── Document View (A4 Pages) ─────────────────────────── */

#document-view {
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 2rem 1rem 4rem;
  gap: 2rem;
}

.page {
  background: #fff;
  width: 794px;       /* A4 at 96dpi */
  min-height: 1123px;
  padding: 96px 76px 96px 114px; /* top, right, bottom, left — matches 2,1,2,3 cm at 96dpi */
  box-shadow: 0 4px 20px rgba(0,0,0,0.12);
  position: relative;
}

@media (max-width: 860px) {
  .page {
    width: 100%;
    min-height: unset;
    padding: 2rem 1.5rem;
  }
}

/* ── Paragraph Rendering ──────────────────────────────── */

.para {
  font-family: 'Times New Roman', serif;
  font-size: 14pt;
  line-height: 1;
  margin-bottom: 0;
  white-space: pre-wrap;
}

.para-body { text-align: justify; text-indent: 1.25cm; }
.para-heading { text-align: left; text-indent: 0; font-weight: bold; font-size: 20pt; }
.para-empty { min-height: 1em; }

/* ── Error Highlight + Tooltip ────────────────────────── */

.para-error {
  background: #fef08a;  /* yellow */
  border-left: 3px solid #ca8a04;
  padding-left: 4px;
  margin-left: -7px;
  position: relative;
  cursor: help;
}

.tooltip {
  display: none;
  position: absolute;
  left: 0;
  bottom: calc(100% + 6px);
  z-index: 200;
  background: #1e293b;
  color: #f8fafc;
  border-radius: 8px;
  padding: 0.75rem 1rem;
  min-width: 280px;
  max-width: 400px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.25);
  font-family: -apple-system, BlinkMacSystemFont, sans-serif;
  font-size: 0.8rem;
  line-height: 1.5;
  pointer-events: none;
}

.tooltip::after {
  content: '';
  position: absolute;
  top: 100%;
  left: 16px;
  border: 6px solid transparent;
  border-top-color: #1e293b;
}

.para-error:hover .tooltip { display: block; }

.tooltip-error + .tooltip-error { margin-top: 0.6rem; border-top: 1px solid #334155; padding-top: 0.6rem; }

.tooltip-param { font-weight: 700; color: #fbbf24; margin-bottom: 0.2rem; }
.tooltip-values { color: #cbd5e1; }
.tooltip-values .current { color: #f87171; }
.tooltip-values .required { color: #86efac; }
.tooltip-rec { color: #94a3b8; font-style: italic; margin-top: 0.2rem; }
```

- [ ] **Step 2: Open index.html in browser**

Expected: the page shows the upload drop zone, styled nicely.

- [ ] **Step 3: Commit**

```bash
git add styles.css
git commit -m "feat: styles.css — A4 layout, upload zone, highlight, tooltip"
```

---

## Task 10: app.js — renderDocument()

**Files:**
- Modify: `app.js`

- [ ] **Step 1: Write the renderDocument function**

```javascript
function renderDocument(results, filename) {
  // Summary bar
  document.getElementById('summary-filename').textContent = filename;
  const errCount = results.totalErrors;
  const errEl = document.getElementById('summary-errors');
  if (errCount === 0) {
    errEl.textContent = 'Ошибок не найдено ✓';
    errEl.classList.add('no-errors');
  } else {
    errEl.textContent = `Ошибок: ${errCount}`;
    errEl.classList.remove('no-errors');
  }

  // Margin errors banner
  const marginBanner = document.getElementById('margin-banner');
  const marginList = document.getElementById('margin-error-list');
  marginList.innerHTML = '';
  if (results.marginErrors.length > 0) {
    results.marginErrors.forEach(err => {
      const li = document.createElement('li');
      li.textContent = `${err.param}: ${err.current} → требуется ${err.required}. ${err.recommendation}`;
      marginList.appendChild(li);
    });
    marginBanner.hidden = false;
  } else {
    marginBanner.hidden = true;
  }

  // Document view
  const docView = document.getElementById('document-view');
  docView.innerHTML = '';

  // Split paragraphs into fake "pages" of ~40 paragraphs each
  const PAGE_SIZE = 40;
  const allParas = results.paragraphResults;
  const pageCount = Math.max(1, Math.ceil(allParas.length / PAGE_SIZE));

  for (let p = 0; p < pageCount; p++) {
    const pageEl = document.createElement('div');
    pageEl.className = 'page';

    const slice = allParas.slice(p * PAGE_SIZE, (p + 1) * PAGE_SIZE);

    slice.forEach(para => {
      const wrapper = document.createElement('div');

      if (!para.text.trim()) {
        wrapper.className = 'para para-empty';
        pageEl.appendChild(wrapper);
        return;
      }

      const typeClass = para.type === 'heading' ? 'para-heading' : 'para-body';
      wrapper.className = `para ${typeClass}`;
      wrapper.textContent = para.text;

      if (para.hasErrors) {
        wrapper.classList.add('para-error');
        const tooltip = buildTooltip(para.errors);
        wrapper.appendChild(tooltip);
      }

      pageEl.appendChild(wrapper);
    });

    docView.appendChild(pageEl);
  }
}

function buildTooltip(errors) {
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip';

  errors.forEach((err, i) => {
    const block = document.createElement('div');
    block.className = 'tooltip-error';

    block.innerHTML = `
      <div class="tooltip-param">${escHtml(err.param)}</div>
      <div class="tooltip-values">
        Текущее: <span class="current">${escHtml(err.current)}</span> →
        Требуемое: <span class="required">${escHtml(err.required)}</span>
      </div>
      <div class="tooltip-rec">${escHtml(err.recommendation)}</div>
    `;

    tooltip.appendChild(block);
  });

  return tooltip;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
```

- [ ] **Step 2: Commit**

```bash
git add app.js
git commit -m "feat: renderDocument — build document view with highlights and tooltips"
```

---

## Task 11: app.js — File Input Wiring (Full Pipeline)

**Files:**
- Modify: `app.js`

- [ ] **Step 1: Add the full event wiring at the top of app.js**

Prepend this to `app.js` (before `renderDocument`):

```javascript
// ── State helpers ──────────────────────────────────────
function showState(name) {
  ['upload', 'loading', 'results'].forEach(s => {
    document.getElementById(`state-${s}`).hidden = s !== name;
  });
}

// ── File processing pipeline ───────────────────────────
async function processFile(file) {
  if (!file || !file.name.endsWith('.docx')) {
    showUploadError('Пожалуйста, загрузите файл в формате .docx');
    return;
  }

  showState('loading');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    if (!zip.file('word/document.xml') || !zip.file('word/styles.xml')) {
      throw new Error('Файл повреждён или не является корректным .docx документом.');
    }

    const results = await DocxChecker.checkDocument(zip);
    showState('results');
    renderDocument(results, file.name);
  } catch (err) {
    showState('upload');
    showUploadError('Ошибка при анализе файла: ' + err.message);
  }
}

function showUploadError(msg) {
  const el = document.getElementById('upload-error');
  el.textContent = msg;
  el.hidden = false;
}

// ── Event listeners ────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('file-input');
  const dropZone  = document.getElementById('drop-zone');
  const btnReset  = document.getElementById('btn-reset');

  // File input change
  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) processFile(fileInput.files[0]);
  });

  // Drag and drop
  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  });

  // Reset button
  btnReset.addEventListener('click', () => {
    fileInput.value = '';
    document.getElementById('upload-error').hidden = true;
    document.getElementById('document-view').innerHTML = '';
    showState('upload');
  });
});
```

- [ ] **Step 2: Open index.html in browser and verify all three states work**

1. Upload state shows on load ✓
2. Drop zone highlights on drag-over ✓
3. Clicking drop zone opens file picker ✓
4. "Reset" button is wired (will test fully in Task 12)

- [ ] **Step 3: Commit**

```bash
git add app.js
git commit -m "feat: app.js — file input wiring and pipeline"
```

---

## Task 12: End-to-End Manual Test

**Files:** (none changed — this is a manual test)

You need a sample DOCX file with known formatting issues for this test. Create one in Microsoft Word or LibreOffice:

1. Open Word, create a new document
2. Set page margins: left 2.5 cm (wrong, should be 3.0), right 1.0, top 2.0, bottom 2.0
3. Type a heading line: "ВВЕДЕНИЕ" — make it Arial, 20pt, bold, left-aligned, no indent
4. Type a body paragraph: change font to Arial 12pt, justify, indent 1.25cm
5. Type another body paragraph: Times New Roman 14pt, justify, indent 1.25cm (correct)
6. Save as `test-sample.docx`

- [ ] **Step 1: Open index.html in browser, upload `test-sample.docx`**

Expected results:
- Summary bar shows filename and error count > 0
- Red margin banner appears (left margin is 2.5 not 3.0)
- Heading paragraph "ВВЕДЕНИЕ" shows yellow highlight (Arial font error)
- Body paragraph with Arial 12pt shows yellow highlight (font + size errors)
- Correct paragraph has no highlight
- Hovering a highlighted paragraph shows tooltip with errors

- [ ] **Step 2: Verify tooltip content**

Hover over the Arial 12pt paragraph. Tooltip should show:
```
Шрифт
Текущее: Arial → Требуемое: Times New Roman
Рекомендация: Выделите текст и смените шрифт на «Times New Roman».

Размер шрифта
Текущее: 12 pt → Требуемое: 14 pt
Рекомендация: Выделите текст и установите размер 14 pt.
```

- [ ] **Step 3: Test "Проверить другой файл" button**

Click the button. Expected: returns to upload state, drop zone appears.

- [ ] **Step 4: Test with a non-docx file**

Drag a .pdf or .txt file onto the drop zone.
Expected: error message "Пожалуйста, загрузите файл в формате .docx" appears.

- [ ] **Step 5: Commit test sample (optional)**

```bash
git add test-sample.docx
git commit -m "test: add sample DOCX with known formatting errors"
```

---

## Task 13: Deploy to GitHub + Netlify

**Files:** (none changed)

- [ ] **Step 1: Create a GitHub repository**

Go to https://github.com/new — create a new public repository named `diploma-checker` (or any name).

- [ ] **Step 2: Push the project**

```bash
git remote add origin https://github.com/YOUR_USERNAME/diploma-checker.git
git branch -M main
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

- [ ] **Step 3: Connect to Netlify**

1. Go to https://netlify.com — sign up with your GitHub account (free)
2. Click "Add new site" → "Import an existing project"
3. Choose GitHub → select `diploma-checker`
4. Build settings: leave everything blank (no build command, publish directory = `.`)
5. Click "Deploy site"

- [ ] **Step 4: Verify deployment**

Netlify gives you a URL like `https://random-name-123.netlify.app`. Open it, upload a DOCX. Everything should work exactly as locally.

- [ ] **Step 5: Share the URL with students**

The URL is permanent and free. Every time you push to GitHub, Netlify auto-deploys the update.

---

## Self-Review Notes

**Spec coverage check:**
- ✓ Font (Times New Roman) — checked in `checkParagraph`
- ✓ Size (14pt body, 20pt heading) — checked in `checkParagraph`
- ✓ Bold (no for body, yes for heading) — checked in `checkParagraph`
- ✓ Line spacing (1.0 / 240 twips) — checked in `checkParagraph`
- ✓ Page margins (left 3, right 1, top 2, bottom 2 cm) — checked in `checkMargins`
- ✓ First line indent (1.25 cm body, 0 heading) — checked in `checkParagraph`
- ✓ Alignment (justify body, left heading) — checked in `checkParagraph`
- ✓ Heading detection (by style name OR bold+20pt) — `classifyParagraph`
- ✓ Style inheritance (Normal → paragraph style → run override) — `resolveFormatting`
- ✓ Upload UI with drag-and-drop — `app.js`
- ✓ Loading state — `showState('loading')`
- ✓ Yellow highlight + hover tooltip — `styles.css` + `buildTooltip`
- ✓ Margin error banner — `renderDocument`
- ✓ Reset button — `app.js`
- ✓ Error message for wrong file type — `processFile`
- ✓ Deployment to Netlify — Task 13
