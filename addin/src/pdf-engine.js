// ─────────────────────────────────────────────────────────────────────────────
// PDFEngine — Renderização nativa via PDF.js
// ─────────────────────────────────────────────────────────────────────────────

const PDFJS_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs';
const WORKER_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

export class PDFEngine {
  constructor(canvas, textLayer, onPageChange) {
    this.canvas      = canvas;
    this.textLayer   = textLayer;
    this.onPageChange = onPageChange;
    this.ctx         = canvas.getContext('2d');
    this.pdf         = null;
    this.currentPage = 1;
    this.totalPages  = 0;
    this.scale       = 1.0;
    this.renderTask  = null;
    this._pdfjs      = null;
    this._searchResults = [];
    this._searchIndex   = -1;
  }

  async _getPdfjs() {
    if (!this._pdfjs) {
      const mod = await import(PDFJS_CDN);
      mod.GlobalWorkerOptions.workerSrc = WORKER_CDN;
      this._pdfjs = mod;
    }
    return this._pdfjs;
  }

  async load(arrayBuffer) {
    const pdfjsLib = await this._getPdfjs();
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    this.pdf = await loadingTask.promise;
    this.totalPages  = this.pdf.numPages;
    this.currentPage = 1;
    await this.render();
  }

  async render() {
    if (!this.pdf) return;

    // Cancel any in-progress render
    if (this.renderTask) {
      this.renderTask.cancel();
    }

    const page = await this.pdf.getPage(this.currentPage);
    const viewport = page.getViewport({ scale: this.scale });

    this.canvas.width  = viewport.width;
    this.canvas.height = viewport.height;

    const renderContext = {
      canvasContext: this.ctx,
      viewport,
    };

    this.renderTask = page.render(renderContext);

    try {
      await this.renderTask.promise;
    } catch (err) {
      if (err?.name !== 'RenderingCancelledException') throw err;
    }

    // Render text layer for selection
    await this._renderTextLayer(page, viewport);

    this.onPageChange?.();
  }

  async _renderTextLayer(page, viewport) {
    this.textLayer.innerHTML = '';
    this.textLayer.style.width  = `${viewport.width}px`;
    this.textLayer.style.height = `${viewport.height}px`;
    this.textLayer.style.left   = this.canvas.offsetLeft + 'px';
    this.textLayer.style.top    = this.canvas.offsetTop  + 'px';

    const textContent = await page.getTextContent();
    const pdfjsLib = await this._getPdfjs();

    pdfjsLib.renderTextLayer({
      textContentSource: textContent,
      container: this.textLayer,
      viewport,
    });
  }

  async nextPage() {
    if (this.currentPage < this.totalPages) {
      this.currentPage++;
      await this.render();
    }
  }

  async prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      await this.render();
    }
  }

  async goToPage(n) {
    const page = Math.max(1, Math.min(n, this.totalPages));
    if (page !== this.currentPage) {
      this.currentPage = page;
      await this.render();
    }
  }

  zoomIn()  { this.scale = Math.min(this.scale + 0.25, 4.0); this.render(); }
  zoomOut() { this.scale = Math.max(this.scale - 0.25, 0.25); this.render(); }

  fitWidth(containerWidth) {
    if (!this.pdf) return;
    this.pdf.getPage(this.currentPage).then(page => {
      const naturalViewport = page.getViewport({ scale: 1.0 });
      this.scale = containerWidth / naturalViewport.width;
      this.render();
    });
  }

  // ── Search ──────────────────────────────────────────────────────────────
  async search(term) {
    if (!this.pdf) return 0;
    this._searchResults = [];
    this._searchIndex   = -1;

    for (let p = 1; p <= this.totalPages; p++) {
      const page = await this.pdf.getPage(p);
      const content = await page.getTextContent();
      const text = content.items.map(i => i.str).join(' ');
      const lowerText = text.toLowerCase();
      const lowerTerm = term.toLowerCase();

      let idx = 0;
      while ((idx = lowerText.indexOf(lowerTerm, idx)) !== -1) {
        this._searchResults.push({ page: p, charIndex: idx });
        idx += lowerTerm.length;
      }
    }

    if (this._searchResults.length > 0) {
      this._searchIndex = 0;
      await this.goToPage(this._searchResults[0].page);
    }

    return this._searchResults.length;
  }

  async navigateSearch(direction) {
    if (this._searchResults.length === 0) return;
    this._searchIndex = (this._searchIndex + direction + this._searchResults.length)
                        % this._searchResults.length;
    const result = this._searchResults[this._searchIndex];
    await this.goToPage(result.page);
  }

  clearSearch() {
    this._searchResults = [];
    this._searchIndex   = -1;
  }
}
