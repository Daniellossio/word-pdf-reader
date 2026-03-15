// ─────────────────────────────────────────────────────────────────────────────
// WordBridge — Inserção de texto selecionado no documento Word ativo
// ─────────────────────────────────────────────────────────────────────────────

export class WordBridge {
  get _inWord() { return typeof Word !== 'undefined'; }

  // Insere texto no cursor atual do documento Word
  async insertText(text) {
    try {
      if (!this._inWord) { await this.copyToClipboard(text); this._notify('Copiado! (Cole com Ctrl+V)'); return; }
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (err) {
      console.error('WordBridge.insertText failed:', err);
      // Fallback: copiar para clipboard
      await this.copyToClipboard(text);
      this._notify('Texto copiado para a área de transferência.');
    }
  }

  // Insere texto como parágrafo novo no final do documento
  async appendParagraph(text) {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.insertParagraph(text, Word.InsertLocation.end);
        await context.sync();
      });
    } catch (err) {
      console.error('WordBridge.appendParagraph failed:', err);
      await this.copyToClipboard(text);
    }
  }

  // Insere texto como citação formatada (estilo Quote)
  async insertAsQuote(text, source) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const para = selection.insertParagraph(text, Word.InsertLocation.after);
        para.styleBuiltIn = Word.BuiltInStyleName.quote;
        if (source) {
          const sourcePara = para.insertParagraph(`— ${source}`, Word.InsertLocation.after);
          sourcePara.styleBuiltIn = Word.BuiltInStyleName.quoteAttribution;
        }
        await context.sync();
      });
    } catch (err) {
      console.error('WordBridge.insertAsQuote failed:', err);
      await this.insertText(text);
    }
  }

  async copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // Fallback para execCommand (deprecated mas funcional em Office WebView)
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
  }

  _notify(msg) {
    // Toast simples
    const toast = document.createElement('div');
    toast.textContent = msg;
    toast.style.cssText = `
      position:fixed; bottom:16px; left:50%; transform:translateX(-50%);
      background:#333; color:#fff; padding:8px 16px; border-radius:6px;
      font-size:12px; z-index:9999; pointer-events:none;
      animation: fadeout 2.5s forwards;
    `;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 2600);
  }
}
