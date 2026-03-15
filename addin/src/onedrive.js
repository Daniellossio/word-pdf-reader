// ─────────────────────────────────────────────────────────────────────────────
// OneDrivePicker — Navega e seleciona PDFs do OneDrive via Microsoft Graph API
// ─────────────────────────────────────────────────────────────────────────────

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export class OneDrivePicker {
  constructor() {
    this.modal        = document.getElementById('onedrive-modal');
    this.list         = document.getElementById('onedrive-list');
    this.breadcrumb   = document.getElementById('onedrive-breadcrumb');
    this.onSelect     = null;
    this._token       = null;
    this._path        = [];   // breadcrumb stack: [{ name, id }]

    document.getElementById('btn-onedrive-close').onclick = () => this.close();
  }

  async open(onSelectCallback) {
    this.onSelect = onSelectCallback;
    this._token = await this._getToken();
    if (!this._token) {
      alert('Não foi possível autenticar com o OneDrive. Verifique se está logado no Word com sua conta Microsoft.');
      return;
    }
    this._path = [{ name: 'OneDrive', id: 'root' }];
    this.modal.style.display = 'flex';
    await this._loadFolder('root');
  }

  close() {
    this.modal.style.display = 'none';
  }

  async _getToken() {
    try {
      // Office.js provides the token via SSO (same account as Word/Microsoft 365)
      return await new Promise((resolve, reject) => {
        Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        }).then(resolve).catch(reject);
      });
    } catch (err) {
      console.error('Failed to get Office token:', err);
      return null;
    }
  }

  async _loadFolder(folderId) {
    this.list.innerHTML = '<div style="padding:16px;color:#888">Carregando...</div>';
    this._renderBreadcrumb();

    try {
      const endpoint = folderId === 'root'
        ? `${GRAPH_BASE}/me/drive/root/children`
        : `${GRAPH_BASE}/me/drive/items/${folderId}/children`;

      const res = await fetch(`${endpoint}?$select=id,name,size,file,folder,@microsoft.graph.downloadUrl&$orderby=name`, {
        headers: { Authorization: `Bearer ${this._token}` },
      });

      if (!res.ok) throw new Error(`Graph API: ${res.status}`);
      const data = await res.json();

      this._renderItems(data.value || []);
    } catch (err) {
      this.list.innerHTML = `<div style="padding:16px;color:#e44">Erro ao carregar: ${err.message}</div>`;
    }
  }

  _renderItems(items) {
    this.list.innerHTML = '';

    // Only show folders and PDFs
    const filtered = items.filter(item =>
      item.folder || (item.file && item.name.toLowerCase().endsWith('.pdf'))
    );

    if (filtered.length === 0) {
      this.list.innerHTML = '<div style="padding:16px;color:#888">Nenhum PDF encontrado aqui.</div>';
      return;
    }

    filtered.forEach(item => {
      const el = document.createElement('div');
      el.className = 'onedrive-item';

      const icon = item.folder ? '📁' : '📄';
      const size = item.file ? this._formatSize(item.size) : '';

      el.innerHTML = `
        <span class="item-icon">${icon}</span>
        <span class="item-name">${item.name}</span>
        <span class="item-size">${size}</span>
      `;

      el.onclick = async () => {
        if (item.folder) {
          this._path.push({ name: item.name, id: item.id });
          await this._loadFolder(item.id);
        } else {
          // PDF selected — get download URL
          const downloadUrl = await this._getDownloadUrl(item.id);
          this.close();
          this.onSelect?.(downloadUrl, this._token);
        }
      };

      this.list.appendChild(el);
    });
  }

  async _getDownloadUrl(itemId) {
    try {
      const res = await fetch(`${GRAPH_BASE}/me/drive/items/${itemId}?$select=@microsoft.graph.downloadUrl`, {
        headers: { Authorization: `Bearer ${this._token}` },
      });
      const data = await res.json();
      return data['@microsoft.graph.downloadUrl'];
    } catch {
      return `${GRAPH_BASE}/me/drive/items/${itemId}/content`;
    }
  }

  _renderBreadcrumb() {
    this.breadcrumb.innerHTML = this._path
      .map((p, i) => {
        if (i < this._path.length - 1) {
          return `<span style="cursor:pointer;color:#4a9eff" data-idx="${i}">${p.name}</span> / `;
        }
        return `<span>${p.name}</span>`;
      })
      .join('');

    // Navigate back via breadcrumb
    this.breadcrumb.querySelectorAll('[data-idx]').forEach(el => {
      el.onclick = async () => {
        const idx = parseInt(el.dataset.idx);
        this._path = this._path.slice(0, idx + 1);
        await this._loadFolder(this._path[idx].id);
      };
    });
  }

  _formatSize(bytes) {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }
}
