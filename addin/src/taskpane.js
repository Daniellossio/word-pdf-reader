// ─────────────────────────────────────────────────────────────────────────────
// PDF Reader Add-in — Main taskpane script
// Uses PDF.js for native rendering + Office.js for Word integration
// ─────────────────────────────────────────────────────────────────────────────

import { PDFEngine } from './pdf-engine.js';
import { AnnotationManager } from './annotations.js';
import { OneDrivePicker } from './onedrive.js';
import { WordBridge } from './word-bridge.js';

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  currentTool: null,   // 'highlight' | 'note' | null
  selectedText: '',
  selectedRect: null,
};

// ── DOM refs ───────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const canvas      = $('pdf-canvas');
const dropZone    = $('drop-zone');
const viewerContainer = $('viewer-container');
const pageInfo    = $('page-info');
const zoomLevel   = $('zoom-level');
const searchBar   = $('search-bar');
const searchInput = $('search-input');
const searchCount = $('search-count');
const notePopup   = $('note-popup');
const noteText    = $('note-text');

// ── Init ───────────────────────────────────────────────────────────────────
const engine      = new PDFEngine(canvas, $('text-layer'), updatePageInfo);
const annotations = new AnnotationManager($('annotation-layer'), canvas);
const onedrive    = new OneDrivePicker();
const wordBridge  = new WordBridge();

// Inicializa imediatamente — não depende do Office.onReady para a UI funcionar
async function init() {
  setupToolbar();
  setupDragDrop();
  setupSearch();
  setupTextSelection();
  setupNotePopup();
  setupKeyboard();

  // Se aberto via companion app (PDF passado via query string)
  const params = new URLSearchParams(window.location.search);
  const pdfPath = params.get('pdf');
  if (pdfPath) await loadFromCompanion(pdfPath);

  // Poll para mensagens do companion app
  pollCompanionMessage();
}

// Inicializa agora — com ou sem Office
if (typeof Office !== 'undefined') {
  Office.onReady(init);
} else {
  document.addEventListener('DOMContentLoaded', init);
}

// ── Toolbar setup ──────────────────────────────────────────────────────────
function setupToolbar() {
  $('btn-open-local').onclick = () => $('file-input').click();
  $('file-input').onchange = async (e) => {
    const file = e.target.files[0];
    if (file) await loadFromFile(file);
    e.target.value = '';
  };

  $('btn-open-onedrive').onclick = () => onedrive.open(loadFromUrl);

  $('btn-prev').onclick = () => { engine.prevPage(); annotations.renderForPage(engine.currentPage); };
  $('btn-next').onclick = () => { engine.nextPage(); annotations.renderForPage(engine.currentPage); };

  $('btn-zoom-in').onclick  = () => { engine.zoomIn();  updateZoom(); };
  $('btn-zoom-out').onclick = () => { engine.zoomOut(); updateZoom(); };
  $('btn-zoom-fit').onclick = () => { engine.fitWidth(viewerContainer.clientWidth - 48); updateZoom(); };

  $('btn-highlight').onclick = () => toggleTool('highlight');
  $('btn-note').onclick      = () => toggleTool('note');

  $('btn-insert-word').onclick = () => {
    if (state.selectedText) {
      wordBridge.insertText(state.selectedText);
    }
  };
}

function toggleTool(tool) {
  state.currentTool = state.currentTool === tool ? null : tool;
  $('btn-highlight').classList.toggle('active', state.currentTool === 'highlight');
  $('btn-note').classList.toggle('active', state.currentTool === 'note');
  canvas.style.cursor = state.currentTool ? 'crosshair' : 'default';
}

function updatePageInfo() {
  pageInfo.textContent = `${engine.currentPage} / ${engine.totalPages}`;
  $('btn-prev').disabled = engine.currentPage <= 1;
  $('btn-next').disabled = engine.currentPage >= engine.totalPages;
  annotations.renderForPage(engine.currentPage);
}

function updateZoom() {
  zoomLevel.textContent = `${Math.round(engine.scale * 100)}%`;
}

// ── Drag & Drop ────────────────────────────────────────────────────────────
function setupDragDrop() {
  viewerContainer.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  viewerContainer.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  viewerContainer.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file?.type === 'application/pdf') {
      await loadFromFile(file);
    }
  });
}

// ── Load PDF ───────────────────────────────────────────────────────────────
async function loadFromFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  await engine.load(arrayBuffer);
  annotations.clear();
  showViewer();
  updatePageInfo();
  updateZoom();
}

async function loadFromUrl(url, token) {
  try {
    const headers = token ? { Authorization: `Bearer ${token}` } : {};
    const res = await fetch(url, { headers });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buffer = await res.arrayBuffer();
    await engine.load(buffer);
    annotations.clear();
    showViewer();
    updatePageInfo();
    updateZoom();
  } catch (err) {
    console.error('Failed to load PDF from URL:', err);
    alert('Erro ao carregar o PDF. Verifique a conexão.');
  }
}

async function loadFromCompanion(filePath) {
  try {
    // Companion app serves the PDF via localhost:3001/file?path=...
    const url = `http://localhost:3001/file?path=${encodeURIComponent(filePath)}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buffer = await res.arrayBuffer();
    await engine.load(buffer);
    annotations.clear();
    showViewer();
    updatePageInfo();
    updateZoom();
  } catch (err) {
    console.error('Companion load failed:', err);
  }
}

// Poll for messages from companion app (PDF path sent after Word is open)
async function pollCompanionMessage() {
  try {
    const res = await fetch('http://localhost:3001/pending-pdf', { signal: AbortSignal.timeout(2000) });
    if (res.ok) {
      const { path } = await res.json();
      if (path) await loadFromCompanion(path);
    }
  } catch {
    // Companion not running — that's fine
  }
  // Poll every 3s
  setTimeout(pollCompanionMessage, 3000);
}

function showViewer() {
  dropZone.classList.add('hidden');
  canvas.classList.remove('hidden');
}

// ── Text selection ─────────────────────────────────────────────────────────
function setupTextSelection() {
  document.addEventListener('selectionchange', () => {
    const selection = window.getSelection();
    const text = selection?.toString().trim();
    if (text) {
      state.selectedText = text;
      state.selectedRect = selection.getRangeAt(0).getBoundingClientRect();
      $('btn-insert-word').disabled = false;

      if (state.currentTool === 'highlight') {
        addHighlight(selection);
      } else if (state.currentTool === 'note') {
        openNotePopup(selection);
      }
    } else {
      $('btn-insert-word').disabled = true;
    }
  });
}

// ── Annotations ────────────────────────────────────────────────────────────
function addHighlight(selection) {
  const range = selection.getRangeAt(0);
  const rects = Array.from(range.getClientRects());
  const containerRect = canvas.getBoundingClientRect();

  rects.forEach(rect => {
    annotations.addHighlight({
      page: engine.currentPage,
      x: rect.left - containerRect.left,
      y: rect.top - containerRect.top,
      width: rect.width,
      height: rect.height,
      text: selection.toString(),
    });
  });

  annotations.renderForPage(engine.currentPage);
  selection.removeAllRanges();
}

// ── Note popup ─────────────────────────────────────────────────────────────
function setupNotePopup() {
  $('btn-note-save').onclick = () => {
    const text = noteText.value.trim();
    if (text && state.selectedRect) {
      const containerRect = canvas.getBoundingClientRect();
      annotations.addNote({
        page: engine.currentPage,
        x: state.selectedRect.left - containerRect.left,
        y: state.selectedRect.top - containerRect.top,
        width: state.selectedRect.width,
        height: state.selectedRect.height,
        note: text,
      });
      annotations.renderForPage(engine.currentPage);
    }
    notePopup.style.display = 'none';
    noteText.value = '';
  };

  $('btn-note-cancel').onclick = () => {
    notePopup.style.display = 'none';
    noteText.value = '';
  };
}

function openNotePopup(selection) {
  const rect = selection.getRangeAt(0).getBoundingClientRect();
  state.selectedRect = rect;
  notePopup.style.display = 'block';
  notePopup.style.left = `${rect.left}px`;
  notePopup.style.top  = `${rect.bottom + 6}px`;
  noteText.focus();
}

// ── Search ─────────────────────────────────────────────────────────────────
function setupSearch() {
  // Ctrl+F opens search
  $('btn-search-close').onclick = () => {
    searchBar.style.display = 'none';
    engine.clearSearch();
    searchCount.textContent = '';
  };

  $('btn-search-next').onclick = () => navigateSearch(1);
  $('btn-search-prev').onclick = () => navigateSearch(-1);

  searchInput.addEventListener('input', async () => {
    const term = searchInput.value.trim();
    if (term.length < 2) { searchCount.textContent = ''; return; }
    const results = await engine.search(term);
    searchCount.textContent = `${results} resultado${results !== 1 ? 's' : ''}`;
  });

  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') navigateSearch(e.shiftKey ? -1 : 1);
    if (e.key === 'Escape') $('btn-search-close').click();
  });
}

function navigateSearch(direction) {
  engine.navigateSearch(direction);
}

// ── Keyboard shortcuts ─────────────────────────────────────────────────────
function setupKeyboard() {
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'f') { e.preventDefault(); toggleSearch(); }
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') engine.nextPage();
    if (e.key === 'ArrowLeft'  || e.key === 'ArrowUp')   engine.prevPage();
    if (e.ctrlKey && e.key === '+') { engine.zoomIn();  updateZoom(); }
    if (e.ctrlKey && e.key === '-') { engine.zoomOut(); updateZoom(); }
    if (e.key === 'Escape') { state.currentTool = null; toggleTool(null); }
  });
}

function toggleSearch() {
  const visible = searchBar.style.display !== 'none';
  searchBar.style.display = visible ? 'none' : 'flex';
  if (!visible) searchInput.focus();
}
