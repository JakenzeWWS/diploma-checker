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

// ── Document renderer ──────────────────────────────────
function renderDocument(results, filename) {
  document.getElementById('summary-filename').textContent = filename;

  const errEl = document.getElementById('summary-errors');
  if (results.totalErrors === 0) {
    errEl.textContent = 'Ошибок не найдено ✓';
    errEl.classList.add('no-errors');
  } else {
    errEl.textContent = 'Ошибок: ' + results.totalErrors;
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

  // Document view
  const docView = document.getElementById('document-view');
  docView.innerHTML = '';

  const PAGE_SIZE = 40;
  const allParas  = results.paragraphResults;
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
      wrapper.className = 'para ' + typeClass;
      wrapper.textContent = para.text;

      if (para.hasErrors) {
        wrapper.classList.add('para-error');
        wrapper.appendChild(buildTooltip(para.errors));
      }

      pageEl.appendChild(wrapper);
    });

    docView.appendChild(pageEl);
  }
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
document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('file-input');
  const dropZone  = document.getElementById('drop-zone');
  const btnReset  = document.getElementById('btn-reset');

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
    showState('upload');
  });
});
