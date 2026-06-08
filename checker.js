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
  function resolveFormatting(paragraph, styleMap) {
    const normalStyle = styleMap.get('Normal') || {};
    const paraStyle = styleMap.get(paragraph.styleId) || {};

    const runFonts = paragraph.runFormats.map(r => r.font).filter(Boolean);
    const runSizes = paragraph.runFormats.map(r => r.sizePt).filter(v => v !== undefined && v !== null);
    const runBolds = paragraph.runFormats.map(r => r.bold).filter(v => v !== undefined);

    return {
      font:   runFonts[0]  ?? paraStyle.font  ?? normalStyle.font  ?? null,
      sizePt: runSizes[0]  ?? paraStyle.sizePt ?? normalStyle.sizePt ?? null,
      bold:   runBolds[0]  ?? paraStyle.bold  ?? normalStyle.bold  ?? false,
      allFonts: [...new Set(runFonts)],
      allSizes: [...new Set(runSizes)],

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
  function classifyParagraph(paragraph, styleMap) {
    const styleEntry = styleMap.get(paragraph.styleId) || {};
    const normalizedName = styleEntry.normalizedName || '';

    if (normalizedName.startsWith('heading') || normalizedName.startsWith('заголовок')) {
      return 'heading';
    }

    const fmt = resolveFormatting(paragraph, styleMap);
    if (fmt.bold === true && fmt.sizePt === 20) return 'heading';

    return 'body';
  }
  function checkParagraph(paragraph, type, styleMap) {
    if (!paragraph.text.trim()) return [];

    const fmt = resolveFormatting(paragraph, styleMap);
    const rules = RULES[type];
    const errors = [];

    // Font — flag any run with wrong font
    const badFonts = fmt.allFonts.filter(f => f !== rules.font);
    if (badFonts.length > 0) {
      errors.push({
        param: 'Шрифт',
        current: [...new Set(badFonts)].join(', '),
        required: rules.font,
        recommendation: `Выделите текст и смените шрифт на «${rules.font}».`,
      });
    }

    // Size — flag any run with wrong size
    const badSizes = fmt.allSizes.filter(s => s !== rules.sizePt);
    if (badSizes.length > 0) {
      errors.push({
        param: 'Размер шрифта',
        current: badSizes.map(s => s + ' pt').join(', '),
        required: rules.sizePt + ' pt',
        recommendation: `Выделите текст и установите размер ${rules.sizePt} pt.`,
      });
    }

    // Bold — body text must not be bold
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
        text:      para.text,
        type,
        errors,
        hasErrors: errors.length > 0,
      };
    });

    const totalErrors = marginErrors.length
      + paragraphResults.filter(p => p.hasErrors).length;

    return { marginErrors, paragraphResults, totalErrors };
  }

  return {
    RULES, parseStyles, parseMargins, parseParagraphs,
    resolveFormatting, classifyParagraph, checkParagraph,
    checkMargins, checkDocument,
  };
})();
