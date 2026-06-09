const TemplatePage = (() => {
  let _rendered = false;

  function splitPlaceholders(text) {
    return text.split(/(\[[^\]]+\])/);
  }

  function hasBracket(text) {
    return /\[[^\]]+\]/.test(text);
  }

  // Inline input — used for short placeholders (name, title, date)
  function makeInput(label) {
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'tpl-input';
    input.placeholder = label;
    input.setAttribute('aria-label', label);
    input.addEventListener('input', () => sizeInput(input));
    sizeInput(input);
    return input;
  }

  // Block textarea — used for long placeholders (abstract paragraph)
  function makeTextarea(label) {
    const ta = document.createElement('textarea');
    ta.className = 'tpl-textarea';
    ta.placeholder = label;
    ta.setAttribute('aria-label', label);
    ta.rows = 4;
    ta.addEventListener('input', () => autoResize(ta));
    return ta;
  }

  function autoResize(ta) {
    ta.style.height = 'auto';
    ta.style.height = ta.scrollHeight + 'px';
  }

  function sizeInput(input) {
    const text = input.value || input.placeholder || '';
    const ruler = document.getElementById('tpl-ruler');
    ruler.style.font = getComputedStyle(input).font;
    ruler.textContent = text;
    input.style.width = Math.max(ruler.offsetWidth + 8, 40) + 'px';
  }

  // Render text with [placeholders] into a container.
  // If the entire text is one long placeholder, render a textarea.
  function renderText(text, container, blockLevel) {
    const parts = splitPlaceholders(text);
    const isOnlyOnePlaceholder = parts.length === 3 && parts[0] === '' && parts[2] === '';

    if (blockLevel && isOnlyOnePlaceholder) {
      const label = parts[1].slice(1, -1);
      // Long placeholder → textarea; short → single-line input that fills width
      const field = label.length > 60 ? makeTextarea(label) : makeInput(label);
      if (field.tagName === 'INPUT') field.style.width = '100%';
      container.appendChild(field);
      return;
    }

    parts.forEach(part => {
      if (!part) return;
      if (/^\[[^\]]+\]$/.test(part)) {
        const label = part.slice(1, -1);
        container.appendChild(label.length > 60 ? makeTextarea(label) : makeInput(label));
      } else {
        container.appendChild(document.createTextNode(part));
      }
    });
  }

  function renderParagraph(block, container) {
    const wrapper = document.createElement('div');
    const hasImages = block.imageSrcs && block.imageSrcs.length > 0;

    if (!block.text.trim() && !hasImages) {
      wrapper.className = 'para para-empty';
      container.appendChild(wrapper);
      return;
    }

    wrapper.className = 'para ' + (block.type === 'heading' ? 'para-heading' : 'para-body');

    if (hasImages) {
      block.imageSrcs.forEach(src => {
        const img = document.createElement('img');
        img.src = src;
        img.className = 'doc-image';
        wrapper.appendChild(img);
      });
    }

    if (block.text.trim()) {
      if (hasBracket(block.text)) {
        renderText(block.text, wrapper, true);
      } else {
        const span = document.createElement('span');
        span.textContent = block.text;
        wrapper.appendChild(span);
      }
    }

    container.appendChild(wrapper);
  }

  function renderTable(block, container) {
    const table = document.createElement('table');
    table.className = 'doc-table';
    block.rows.forEach(row => {
      const tr = document.createElement('tr');
      row.cells.forEach(cell => {
        const td = document.createElement('td');
        if (hasBracket(cell.text)) {
          renderText(cell.text, td, false);
        } else if (cell.text) {
          td.textContent = cell.text;
        }
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    container.appendChild(table);
  }

  async function loadAndRender() {
    if (_rendered) return;
    _rendered = true;

    const view = document.getElementById('template-view');
    view.textContent = '';
    const statusEl = document.createElement('div');
    statusEl.className = 'tpl-status';
    statusEl.textContent = 'Загружаем шаблон…';
    view.appendChild(statusEl);

    try {
      const resp = await fetch('template.docx');
      if (!resp.ok) throw new Error('template.docx не найден (' + resp.status + ')');
      const buf = await resp.arrayBuffer();
      const zip = await JSZip.loadAsync(buf);
      const results = await DocxChecker.checkDocument(zip);

      view.textContent = '';

      const pages = [];
      let lastPage = -1;
      results.blockResults.forEach(block => {
        if (block.pageNum !== lastPage) { pages.push([]); lastPage = block.pageNum; }
        pages[pages.length - 1].push(block);
      });

      pages.forEach(pageBlocks => {
        const pageEl = document.createElement('div');
        pageEl.className = 'page';
        pageBlocks.forEach(block => {
          if (block.kind === 'p') renderParagraph(block, pageEl);
          else if (block.kind === 'tbl') renderTable(block, pageEl);
        });
        view.appendChild(pageEl);
      });

      view.querySelectorAll('.tpl-input').forEach(sizeInput);
      view.querySelectorAll('.tpl-textarea').forEach(autoResize);

    } catch (err) {
      view.textContent = '';
      const errEl = document.createElement('div');
      errEl.className = 'tpl-status tpl-status-error';
      errEl.textContent = 'Ошибка: ' + err.message;
      view.appendChild(errEl);
    }
  }

  function init() {
    const ruler = document.createElement('span');
    ruler.id = 'tpl-ruler';
    ruler.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;pointer-events:none';
    document.body.appendChild(ruler);

    document.getElementById('btn-print').addEventListener('click', () => window.print());
  }

  return { loadAndRender, init };
})();
