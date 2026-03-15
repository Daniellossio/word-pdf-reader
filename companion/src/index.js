#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────────────────────
// Companion App — Handler padrão de .pdf no Windows
// • Recebe o caminho do PDF como argumento (quando o usuário abre um PDF)
// • Abre o Word com o add-in ativo
// • Serve o arquivo PDF para o add-in via HTTP local
// • Mantém um servidor HTTP leve enquanto o Word estiver aberto
// ─────────────────────────────────────────────────────────────────────────────

const http    = require('http');
const path    = require('path');
const fs      = require('fs');
const { exec, execSync } = require('child_process');

const PORT = 3001;

// ── PDF pendente ─────────────────────────────────────────────────────────────
let pendingPdf = null;

// Argumento de linha de comando: caminho do PDF
const pdfPath = process.argv[2];
if (pdfPath && fs.existsSync(pdfPath)) {
  pendingPdf = path.resolve(pdfPath);
  console.log(`[PDF Reader] Abrindo: ${pendingPdf}`);
}

// ── Servidor HTTP local ───────────────────────────────────────────────────────
const server = http.createServer((req, res) => {
  // CORS para o add-in (localhost:3000)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://localhost:${PORT}`);

  // GET /pending-pdf — add-in pergunta se tem PDF para abrir
  if (url.pathname === '/pending-pdf') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ path: pendingPdf }));
    pendingPdf = null;  // Consome o pending
    return;
  }

  // GET /file?path=... — add-in solicita o conteúdo do PDF
  if (url.pathname === '/file') {
    const filePath = url.searchParams.get('path');
    if (!filePath || !fs.existsSync(filePath)) {
      res.writeHead(404);
      res.end('File not found');
      return;
    }

    // Segurança: só serve arquivos .pdf
    if (!filePath.toLowerCase().endsWith('.pdf')) {
      res.writeHead(403);
      res.end('Only PDF files are allowed');
      return;
    }

    const stat = fs.statSync(filePath);
    res.writeHead(200, {
      'Content-Type': 'application/pdf',
      'Content-Length': stat.size,
    });
    fs.createReadStream(filePath).pipe(res);
    return;
  }

  // GET /health
  if (url.pathname === '/health') {
    res.writeHead(200);
    res.end('ok');
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`[PDF Reader] Servidor local: http://localhost:${PORT}`);
  openWord();
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    // Servidor já rodando — envia o PDF para a instância existente
    console.log('[PDF Reader] Instância já ativa, enviando PDF...');
    notifyExistingInstance();
  } else {
    console.error('[PDF Reader] Erro no servidor:', err);
    process.exit(1);
  }
});

// ── Abrir Word ────────────────────────────────────────────────────────────────
function openWord() {
  // Abre o Word (se não estiver aberto) com um documento em branco
  // O add-in é carregado automaticamente pois está no manifesto sideloaded
  const wordCmd = getWordCommand();
  if (!wordCmd) {
    console.error('[PDF Reader] Word não encontrado. Instale o Microsoft Word.');
    process.exit(1);
  }

  exec(wordCmd, (err) => {
    if (err) console.warn('[PDF Reader] Word já aberto ou erro ao abrir:', err.message);
  });
}

function getWordCommand() {
  // Procura o Word nos caminhos padrão do Windows (via WSL ou nativo)
  const candidates = [
    // Windows nativo
    'start winword.exe',
    // Caminhos comuns
    '"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"',
    '"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE"',
    '"C:\\Program Files\\Microsoft Office\\Office16\\WINWORD.EXE"',
  ];

  // No Windows, usa o primeiro disponível
  if (process.platform === 'win32') {
    return candidates[0];
  }

  // No WSL, usa powershell/cmd para abrir
  return 'cmd.exe /c start winword.exe';
}

// ── Notificar instância existente ─────────────────────────────────────────────
function notifyExistingInstance() {
  if (!pendingPdf) return;
  // A instância existente vai pegar via /pending-pdf poll
  // Mas podemos também fazer um POST para sinalizar
  const body = JSON.stringify({ path: pendingPdf });
  const req = http.request({
    hostname: '127.0.0.1',
    port: PORT,
    path: '/pending-pdf',
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Content-Length': body.length },
  });
  req.write(body);
  req.end();
  console.log('[PDF Reader] PDF enviado para instância existente.');
  process.exit(0);
}

// ── Manter vivo enquanto Word aberto ──────────────────────────────────────────
// O servidor fica ativo; quando o usuário fecha o Word, pode fechar também
process.on('SIGINT', () => {
  console.log('[PDF Reader] Encerrando...');
  server.close();
  process.exit(0);
});
