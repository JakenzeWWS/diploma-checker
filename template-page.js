const TemplatePage = (() => {
  let _rendered = false;

  // Split text like "hello [name] world" into ["hello ", "[name]", " world"]
  function splitPlaceholders(text) {
    return text.split(/(\[[^\]]+\])/);
  }

  function hasBracket(text) {
    return /\[[^\]]+\]/.test(text);
  }

  // Render text (possibly with [placeholders]) into a container element
  function renderText(text, container) {
    splitPlaceholders(text).forEach(part => {
      if (!part) return;
      if (/^\[[^\]]+\]$/.test(part)) {
        const label = part.slice(1, -1);
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'tpl-input';
        input.placeholder = label;
        input.setAttribute('aria-label', label);
        // Grow input width as user types
        input.addEventListener('input', () => sizeInput(input));
        sizeInput(input);
        container.appendChild(input);
      } else {
        container.appendChild(document.createTextNode(part));
      }
    });
  }

  // Mirror element: keep input width synced to its content / placeholder
  function sizeInput(input) {
    const text = input.value || input.placeholder || '';
    const ruler = document.getElementById('tpl-ruler');
    ruler.style.font = getComputedStyle(input).font;
    ruler.textContent = text;
    input.style.width = Math.max(ruler.offsetWidth + 8, 40) + 'px';
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
        renderText(block.text, wrapper);
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
          renderText(cell.text, td);
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

      // Group blocks by page
      const pages = [];
      let lastPage = -1;
      results.blockResults.forEach(block => {
        if (block.pageNum !== lastPage) { pages.push([]); lastPage = block.pageNum; }
        pages[pages.length - 1].push(block);
      });

      pages.forEach((pageBlocks, i) => {
        const pageEl = document.createElement('div');
        pageEl.className = 'page';
        pageBlocks.forEach(block => {
          if (block.kind === 'p') renderParagraph(block, pageEl);
          else if (block.kind === 'tbl') renderTable(block, pageEl);
        });
        view.appendChild(pageEl);
      });

      // After render, resize all inputs
      view.querySelectorAll('.tpl-input').forEach(sizeInput);

    } catch (err) {
      view.textContent = '';
      const errEl = document.createElement('div');
      errEl.className = 'tpl-status tpl-status-error';
      errEl.textContent = 'Ошибка: ' + err.message;
      view.appendChild(errEl);
    }
  }

  function init() {
    // Invisible text ruler for measuring input widths
    const ruler = document.createElement('span');
    ruler.id = 'tpl-ruler';
    ruler.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;pointer-events:none';
    document.body.appendChild(ruler);

    document.getElementById('btn-print').addEventListener('click', () => window.print());
  }

  return { loadAndRender, init };
})();
