const DocxChecker = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
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
        if (fonts) entry.font = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || null;

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

  // Parse a single <w:p> element into raw paragraph data
  function parseParagraphEl(pEl) {
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
    const imageRels = [];

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

      // Detect inline images: find <a:blip r:embed="rId..."/> inside this run
      for (const blip of Array.from(rEl.getElementsByTagNameNS(A_NS, 'blip'))) {
        const rEmbed = blip.getAttributeNS(R_NS, 'embed');
        if (rEmbed) imageRels.push(rEmbed);
      }
    }

    const text = els(pEl, 't').map(t => t.textContent).join('');
    const hasPageBreakRun = els(pEl, 'br').some(br => attr(br, 'type') === 'page');
    const hasLastRenderedBreak = els(pEl, 'lastRenderedPageBreak').length > 0;
    const hasPageBreakBefore = pPr ? !!el(pPr, 'pageBreakBefore') : false;

    return { styleId, paraOverride, runFormats, text, imageRels,
      hasPageBreakRun, hasLastRenderedBreak, hasPageBreakBefore };
  }

  // Parse a <w:tbl>: only direct <w:p> children of each <w:tc> (no nested tables)
  function parseTableEl(tblEl) {
    const rows = [];
    for (const trEl of Array.from(tblEl.children)) {
      if (trEl.localName !== 'tr') continue;
      const cells = [];
      for (const tcEl of Array.from(trEl.children)) {
        if (tcEl.localName !== 'tc') continue;
        const paragraphs = [];
        for (const child of Array.from(tcEl.children)) {
          if (child.localName === 'p') paragraphs.push(parseParagraphEl(child));
        }
        cells.push({
          paragraphs,
          text: paragraphs.map(p => p.text).join('\n').trim(),
        });
      }
      if (cells.length > 0) rows.push({ cells });
    }
    return { rows };
  }

  // Iterate direct <w:body> children to produce ordered blocks (paragraphs + tables)
  function parseBodyBlocks(xmlString) {
    const doc = parseXml(xmlString);
    if (doc.documentElement.tagName === 'parsererror') return [];

    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body) return [];

    const blocks = [];
    for (const child of Array.from(body.children)) {
      if (child.localName === 'p') {
        blocks.push({ kind: 'p', ...parseParagraphEl(child) });
      } else if (child.localName === 'tbl') {
        blocks.push({ kind: 'tbl', ...parseTableEl(child) });
      }
      // sectPr and other elements are skipped
    }

    // Assign startsNewPage to each paragraph block
    let isFirst = true;
    let prevHadHardBreak = false;
    for (const block of blocks) {
      if (block.kind === 'p') {
        block.startsNewPage = isFirst
          || block.hasPageBreakBefore
          || block.hasLastRenderedBreak
          || prevHadHardBreak;
        isFirst = false;
        prevHadHardBreak = block.hasPageBreakRun;
      } else {
        block.startsNewPage = false;
        // preserve prevHadHardBreak across tables
      }
    }

    return blocks;
  }

  // Backward-compatible wrapper — keeps tests passing
  function parseParagraphs(xmlString) {
    return parseBodyBlocks(xmlString)
      .filter(b => b.kind === 'p')
      .map(b => ({
        styleId: b.styleId,
        paraOverride: b.paraOverride,
        runFormats: b.runFormats,
        text: b.text,
        startsNewPage: b.startsNewPage,
        inTable: false,
      }));
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
      allBolds: [...new Set(runBolds)],

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

    const badFonts = fmt.allFonts.filter(f => f !== rules.font);
    if (badFonts.length > 0) {
      errors.push({
        param: 'Шрифт',
        current: [...new Set(badFonts)].join(', '),
        required: rules.font,
        recommendation: `Выделите текст и смените шрифт на «${rules.font}».`,
      });
    } else if (fmt.allFonts.length === 0 && fmt.font !== null && fmt.font !== rules.font) {
      errors.push({
        param: 'Шрифт',
        current: fmt.font,
        required: rules.font,
        recommendation: `Выделите текст и смените шрифт на «${rules.font}».`,
      });
    }

    const badSizes = fmt.allSizes.filter(s => s !== rules.sizePt);
    if (badSizes.length > 0) {
      errors.push({
        param: 'Размер шрифта',
        current: badSizes.map(s => s + ' pt').join(', '),
        required: rules.sizePt + ' pt',
        recommendation: `Выделите текст и установите размер ${rules.sizePt} pt.`,
      });
    }

    if (type === 'body' && fmt.allBolds.includes(true)) {
      errors.push({
        param: 'Жирный шрифт',
        current: 'Да',
        required: 'Нет',
        recommendation: 'Уберите жирное начертание в основном тексте.',
      });
    }

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

    const indent = fmt.firstLineIndentCm ?? 0;
    if (Math.abs(indent - rules.firstLineIndentCm) > 0.05) {
      errors.push({
        param: 'Отступ первой строки',
        current: indent.toFixed(2) + ' см',
        required: rules.firstLineIndentCm + ' см',
        recommendation: `Установите отступ первой строки ${rules.firstLineIndentCm} см.`,
      });
    }

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
      .filter(({ key, required }) => margins[key] === null || Math.abs(margins[key] - required) > tol)
      .map(({ label, key, required }) => ({
        param: label,
        current: margins[key] !== null ? margins[key].toFixed(2) + ' см' : 'не определено',
        required: required.toFixed(1) + ' см',
        recommendation: `Установите ${label.toLowerCase()} ${required} см в настройках полей документа.`,
      }));
  }

  async function checkDocument(zip) {
    const docXml = await zip.file('word/document.xml').async('string');
    const stylesFile = zip.file('word/styles.xml');
    const styleMap = stylesFile
      ? parseStyles(await stylesFile.async('string'))
      : new Map();

    const margins = parseMargins(docXml);
    const marginErrors = checkMargins(margins);

    // Build relationship id → image target path map
    const relMap = {};
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (relsFile) {
      const relDoc = parseXml(await relsFile.async('string'));
      for (const rel of Array.from(relDoc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship'))) {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        const type = rel.getAttribute('Type') || '';
        if (id && target && type.includes('image')) relMap[id] = target;
      }
    }

    const rawBlocks = parseBodyBlocks(docXml);

    let pageNum = 1;
    let firstBlock = true;
    const blockResults = [];

    for (const block of rawBlocks) {
      if (block.kind === 'p') {
        if (block.startsNewPage && !firstBlock) pageNum++;
        firstBlock = false;

        const type = classifyParagraph(block, styleMap);
        const errors = checkParagraph(block, type, styleMap);

        // Resolve image relationship IDs to base64 data URLs
        const imageSrcs = [];
        for (const relId of block.imageRels) {
          const target = relMap[relId];
          if (!target) continue;
          const ext = target.split('.').pop().toLowerCase();
          if (ext === 'emf' || ext === 'wmf') continue; // not renderable in browsers
          const imgFile = zip.file('word/' + target);
          if (!imgFile) continue;
          const mimes = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
          const mime = mimes[ext] || ('image/' + ext);
          const base64 = await imgFile.async('base64');
          imageSrcs.push(`data:${mime};base64,${base64}`);
        }

        blockResults.push({
          kind: 'p',
          text: block.text,
          type,
          errors,
          hasErrors: errors.length > 0,
          pageNum,
          startsNewPage: block.startsNewPage,
          imageSrcs,
        });
      } else if (block.kind === 'tbl') {
        firstBlock = false;
        blockResults.push({ kind: 'tbl', pageNum, rows: block.rows });
      }
    }

    const totalErrors = marginErrors.length
      + blockResults.filter(b => b.kind === 'p' && b.hasErrors).length;

    return { marginErrors, blockResults, totalErrors };
  }

  return {
    RULES, parseStyles, parseMargins, parseParagraphs, parseBodyBlocks,
    resolveFormatting, classifyParagraph, checkParagraph,
    checkMargins, checkDocument,
  };
})();
