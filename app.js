// ── State helpers ──────────────────────────────────────
function showState(name) {
  ['upload', 'loading', 'results'].forEach(s => {
    document.getElementById(`state-${s}`).hidden = s !== name;
  });
}

function showUploadError(msg) {
  const el = document.getElementById('upload-error');
  el.textContent = msg;
  el.hidden = false;
}

// ── Stored results for re-render when skipPages changes ─
let _lastResults  = null;
let _lastFilename = '';

// ── File processing pipeline ───────────────────────────
async function processFile(file) {
  if (!file || !file.name.toLowerCase().endsWith('.docx')) {
    showUploadError('Пожалуйста, загрузите файл в формате .docx');
    return;
  }

  showState('loading');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    if (!zip.file('word/document.xml')) {
      throw new Error('Файл повреждён или не является корректным .docx документом.');
    }

    const results = await DocxChecker.checkDocument(zip);
    _lastResults  = results;
    _lastFilename = file.name;

    showState('results');
    renderDocument(results, file.name, getSkipPages());
  } catch (err) {
    showState('upload');
    showUploadError('Ошибка при анализе файла: ' + err.message);
  }
}

function getSkipPages() {
  const input = document.getElementById('skip-pages');
  return input ? Math.max(0, parseInt(input.value, 10) || 0) : 0;
}

// ── Document renderer ──────────────────────────────────
function renderDocument(results, filename, skipPages) {
  skipPages = skipPages || 0;

  document.getElementById('summary-filename').textContent = filename;

  const activeErrors = results.blockResults
    .filter(b => b.kind === 'p' && b.hasErrors && b.pageNum > skipPages).length;
  const totalActive = results.marginErrors.length + activeErrors;

  const errEl = document.getElementById('summary-errors');
  if (totalActive === 0) {
    errEl.textContent = 'Ошибок не найдено ✓';
    errEl.classList.add('no-errors');
  } else {
    errEl.textContent = 'Ошибок: ' + totalActive;
    errEl.classList.remove('no-errors');
  }

  // Margin errors banner
  const marginBanner = document.getElementById('margin-banner');
  const marginList   = document.getElementById('margin-error-list');
  marginList.innerHTML = '';
  if (results.marginErrors.length > 0) {
    results.marginErrors.forEach(err => {
      const li = document.createElement('li');
      li.textContent = err.param + ': ' + err.current + ' → требуется ' + err.required + '. ' + err.recommendation;
      marginList.appendChild(li);
    });
    marginBanner.hidden = false;
  } else {
    marginBanner.hidden = true;
  }

  // Group blocks by page number
  const pages = [];
  let lastPageNum = -1;
  results.blockResults.forEach(block => {
    if (block.pageNum !== lastPageNum) {
      pages.push([]);
      lastPageNum = block.pageNum;
    }
    pages[pages.length - 1].push(block);
  });

  const docView = document.getElementById('document-view');
  docView.innerHTML = '';

  pages.forEach(pageBlocks => {
    const pageNum   = pageBlocks[0].pageNum;
    const isSkipped = pageNum <= skipPages;

    const pageEl = document.createElement('div');
    pageEl.className = 'page' + (isSkipped ? ' page-skipped' : '');

    if (isSkipped) {
      const notice = document.createElement('div');
      notice.className = 'page-skip-notice';
      notice.textContent = 'Страница ' + pageNum + ' — исключена из проверки (титульная/содержание)';
      pageEl.appendChild(notice);
    }

    pageBlocks.forEach(block => {
      if (block.kind === 'p') {
        renderParagraph(block, pageEl, isSkipped);
      } else if (block.kind === 'tbl') {
        renderTable(block, pageEl);
      }
    });

    docView.appendChild(pageEl);
  });
}

function renderParagraph(para, container, isSkipped) {
  const wrapper = document.createElement('div');

  const hasImages = para.imageSrcs && para.imageSrcs.length > 0;

  if (!para.text.trim() && !hasImages) {
    wrapper.className = 'para para-empty';
    container.appendChild(wrapper);
    return;
  }

  const typeClass = para.type === 'heading' ? 'para-heading' : 'para-body';
  wrapper.className = 'para ' + typeClass;

  if (hasImages) {
    para.imageSrcs.forEach(src => {
      const img = document.createElement('img');
      img.src = src;
      img.className = 'doc-image';
      wrapper.appendChild(img);
    });
  }

  if (para.text.trim()) {
    const span = document.createElement('span');
    span.textContent = para.text;
    wrapper.appendChild(span);
  }

  if (para.hasErrors && !isSkipped) {
    wrapper.classList.add('para-error');
    wrapper.appendChild(buildTooltip(para.errors));
  }

  container.appendChild(wrapper);
}

function renderTable(tbl, container) {
  const table = document.createElement('table');
  table.className = 'doc-table';

  tbl.rows.forEach(row => {
    const tr = document.createElement('tr');
    row.cells.forEach(cell => {
      const td = document.createElement('td');
      if (cell.text) td.textContent = cell.text;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  container.appendChild(table);
}

function buildTooltip(errors) {
  const tooltip = document.createElement('div');
  tooltip.className = 'tooltip';

  errors.forEach(err => {
    const block = document.createElement('div');
    block.className = 'tooltip-error';
    block.innerHTML =
      '<div class="tooltip-param">' + escHtml(err.param) + '</div>' +
      '<div class="tooltip-values">Текущее: <span class="current">' + escHtml(err.current) + '</span>' +
      ' → Требуемое: <span class="required">' + escHtml(err.required) + '</span></div>' +
      '<div class="tooltip-rec">' + escHtml(err.recommendation) + '</div>';
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

// ── Event listeners ────────────────────────────────────
// ── Tab switching ──────────────────────────────────────
function showSection(name) {
  document.getElementById('section-checker').hidden  = name !== 'checker';
  document.getElementById('section-template').hidden = name !== 'template';
  document.getElementById('tab-checker').classList.toggle('tab-active',  name === 'checker');
  document.getElementById('tab-template').classList.toggle('tab-active', name === 'template');
}

document.addEventListener('DOMContentLoaded', () => {
  const fileInput  = document.getElementById('file-input');
  const dropZone   = document.getElementById('drop-zone');
  const btnReset   = document.getElementById('btn-reset');
  const skipInput  = document.getElementById('skip-pages');

  document.getElementById('tab-checker').addEventListener('click', () => showSection('checker'));
  document.getElementById('tab-template').addEventListener('click', () => {
    showSection('template');
    TemplatePage.loadAndRender();
  });

  TemplatePage.init();

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) processFile(fileInput.files[0]);
  });

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

  btnReset.addEventListener('click', () => {
    fileInput.value = '';
    document.getElementById('upload-error').hidden = true;
    document.getElementById('document-view').innerHTML = '';
    _lastResults  = null;
    _lastFilename = '';
    showState('upload');
  });

  if (skipInput) {
    skipInput.addEventListener('change', () => {
      if (_lastResults) renderDocument(_lastResults, _lastFilename, getSkipPages());
    });
  }
});
