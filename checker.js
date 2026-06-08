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

  function el(parent, tag) {
    return parent.getElementsByTagNameNS(W, tag)[0] || null;
  }

  function els(parent, tag) {
    return Array.from(parent.getElementsByTagNameNS(W, tag));
  }

  function attr(element, name) {
    if (!element) return null;
    return element.getAttributeNS(W, name) || element.getAttribute(name) || null;
  }

  function twipsToCm(twips) {
    if (twips === null || twips === undefined) return null;
    return parseFloat((parseInt(twips, 10) * CM_PER_TWIP).toFixed(3));
  }

  function parseStyles(xmlString) {
    const doc = parseXml(xmlString);
    if (doc.documentElement.tagName === 'parsererror') return new Map();
    const styleMap = new Map();

    for (const styleEl of els(doc, 'style')) {
      const styleId = attr(styleEl, 'styleId');
      if (!styleId) continue;

      const nameEl = el(styleEl, 'name');
      const normalizedName = (attr(nameEl, 'val') || '').toLowerCase();

      const entry = { normalizedName };

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
  function parseMargins(xmlString) {
    const doc = parseXml(xmlString);
    if (doc.documentElement.tagName === 'parsererror') return null;
    const pgMar = el(doc, 'pgMar');
    if (!pgMar) return null;

    return {
      leftCm:   twipsToCm(attr(pgMar, 'left')),
      rightCm:  twipsToCm(attr(pgMar, 'right')),
      topCm:    twipsToCm(attr(pgMar, 'top')),
      bottomCm: twipsToCm(attr(pgMar, 'bottom')),
    };
  }
  function parseParagraphs(xmlString) {
    const doc = parseXml(xmlString);
    if (doc.documentElement.tagName === 'parsererror') return [];
    const paragraphs = [];

    for (const pEl of els(doc, 'p')) {
      const pPr = el(pEl, 'pPr');

      let styleId = 'Normal';
      if (pPr) {
        const pStyle = el(pPr, 'pStyle');
        if (pStyle) styleId = attr(pStyle, 'val') || 'Normal';
      }

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
  function resolveFormatting(paragraph, styleMap) { return {}; }
  function classifyParagraph(paragraph, styleMap) { return 'body'; }
  function checkParagraph(paragraph, type, styleMap) { return []; }
  function checkMargins(margins) { return []; }
  async function checkDocument(zip) { return { marginErrors: [], paragraphResults: [], totalErrors: 0 }; }

  return {
    RULES, parseStyles, parseMargins, parseParagraphs,
    resolveFormatting, classifyParagraph, checkParagraph,
    checkMargins, checkDocument,
  };
})();
