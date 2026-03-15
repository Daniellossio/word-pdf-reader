// ─────────────────────────────────────────────────────────────────────────────
// AnnotationManager — Highlights e notas sobrepostas ao PDF
// ─────────────────────────────────────────────────────────────────────────────

export class AnnotationManager {
  constructor(layer, canvas) {
    this.layer       = layer;
    this.canvas      = canvas;
    this.annotations = [];   // { id, type, page, x, y, width, height, text?, note? }
    this._nextId     = 1;
  }

  addHighlight({ page, x, y, width, height, text }) {
    this.annotations.push({ id: this._nextId++, type: 'highlight', page, x, y, width, height, text });
  }

  addNote({ page, x, y, width, height, note }) {
    this.annotations.push({ id: this._nextId++, type: 'note', page, x, y, width, height, note });
  }

  removeById(id) {
    this.annotations = this.annotations.filter(a => a.id !== id);
  }

  clear() {
    this.annotations = [];
    this.layer.innerHTML = '';
  }

  renderForPage(page) {
    this.layer.innerHTML = '';

    // Position layer over canvas
    const rect = this.canvas.getBoundingClientRect();
    const parentRect = this.canvas.parentElement.getBoundingClientRect();
    this.layer.style.left   = `${rect.left - parentRect.left}px`;
    this.layer.style.top    = `${rect.top  - parentRect.top}px`;
    this.layer.style.width  = `${rect.width}px`;
    this.layer.style.height = `${rect.height}px`;

    this.annotations
      .filter(a => a.page === page)
      .forEach(a => this._renderAnnotation(a));
  }

  _renderAnnotation(a) {
    const el = document.createElement('div');
    el.className = a.type === 'highlight' ? 'highlight-mark' : 'note-mark';
    el.style.left   = `${a.x}px`;
    el.style.top    = `${a.y}px`;
    el.style.width  = `${a.width}px`;
    el.style.height = `${a.height}px`;

    if (a.type === 'note') {
      el.title = a.note;
    }

    // Right-click to remove
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      this.removeById(a.id);
      this.renderForPage(a.page);
    });

    // Click note to show content
    if (a.type === 'note') {
      el.addEventListener('click', () => {
        alert(`📝 Nota:\n\n${a.note}`);
      });
    }

    this.layer.appendChild(el);
  }

  // ── Export / Import ──────────────────────────────────────────────────────
  exportJSON() {
    return JSON.stringify(this.annotations, null, 2);
  }

  importJSON(json) {
    try {
      const data = JSON.parse(json);
      this.annotations = data;
      this._nextId = Math.max(...data.map(a => a.id), 0) + 1;
    } catch (err) {
      console.error('Failed to import annotations:', err);
    }
  }
}
