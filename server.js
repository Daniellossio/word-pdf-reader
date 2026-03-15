// ─────────────────────────────────────────────────────────────────────────────
// Dev server HTTPS — Office.js exige HTTPS mesmo em localhost
// Usa certificado self-signed gerado pelo office-addin-dev-certs
// ─────────────────────────────────────────────────────────────────────────────

const https   = require('https');
const http    = require('http');
const fs      = require('fs');
const path    = require('path');
const { execSync } = require('child_process');

const PORT = 3002;
const ADDIN_DIR = path.join(__dirname, 'addin');

// ── MIME types ────────────────────────────────────────────────────────────────
const MIME = {
  '.html': 'text/html',
  '.css':  'text/css',
  '.js':   'application/javascript',
  '.mjs':  'application/javascript',
  '.json': 'application/json',
  '.xml':  'application/xml',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.pdf':  'application/pdf',
};

// ── Gerar certificado se necessário ──────────────────────────────────────────
function getCerts() {
  const certDir = path.join(__dirname, '.certs');
  const certFile = path.join(certDir, 'localhost.crt');
  const keyFile  = path.join(certDir, 'localhost.key');

  if (!fs.existsSync(certFile) || !fs.existsSync(keyFile)) {
    console.log('Gerando certificado self-signed...');
    fs.mkdirSync(certDir, { recursive: true });
    try {
      execSync(
        `openssl req -x509 -newkey rsa:2048 -keyout ${keyFile} -out ${certFile} ` +
        `-days 3650 -nodes -subj "/CN=localhost" ` +
        `-addext "subjectAltName=DNS:localhost,IP:127.0.0.1"`,
        { stdio: 'inherit' }
      );
    } catch {
      console.error('OpenSSL não encontrado. Instale com: sudo apt install openssl');
      process.exit(1);
    }
  }

  return {
    cert: fs.readFileSync(certFile),
    key:  fs.readFileSync(keyFile),
  };
}

// ── Request handler ───────────────────────────────────────────────────────────
function handler(req, res) {
  let urlPath = new URL(req.url, `https://localhost:${PORT}`).pathname;

  // Route: /addin/* → serve addin files
  // Route: /manifest/* → serve manifest
  let filePath;
  if (urlPath === '/' || urlPath === '/addin' || urlPath === '/addin/') {
    filePath = path.join(ADDIN_DIR, 'taskpane.html');
  } else if (urlPath.startsWith('/addin/')) {
    filePath = path.join(ADDIN_DIR, urlPath.slice('/addin/'.length));
  } else if (urlPath.startsWith('/manifest/')) {
    filePath = path.join(__dirname, 'manifest', urlPath.slice('/manifest/'.length));
  } else if (urlPath === '/assets/icon-16.png' ||
             urlPath === '/assets/icon-32.png' ||
             urlPath === '/assets/icon-64.png' ||
             urlPath === '/assets/icon-80.png') {
    // Serve placeholder icon
    filePath = path.join(ADDIN_DIR, 'assets', path.basename(urlPath));
  } else {
    res.writeHead(404);
    res.end('Not found');
    return;
  }

  if (!fs.existsSync(filePath)) {
    res.writeHead(404);
    res.end(`File not found: ${filePath}`);
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME[ext] || 'application/octet-stream';

  res.writeHead(200, {
    'Content-Type': mime,
    'Cache-Control': 'no-cache',
    // Permissões necessárias para Office.js
    'X-Content-Type-Options': 'nosniff',
  });

  fs.createReadStream(filePath).pipe(res);
}

// ── Start ──────────────────────────────────────────────────────────────────────
const certs = getCerts();
const server = https.createServer(certs, handler);

server.listen(PORT, () => {
  console.log(`\n✅ PDF Reader Add-in Server`);
  console.log(`   HTTPS: https://localhost:${PORT}/addin/taskpane.html`);
  console.log(`   Manifest: https://localhost:${PORT}/manifest/manifest.xml`);
  console.log('\n📋 Para instalar o add-in no Word:');
  console.log('   1. Abra o Word');
  console.log('   2. Inserir → Suplementos → Meus Suplementos → Carregar Suplemento');
  console.log('   3. Selecione: manifest/manifest.xml');
  console.log('\n⚠️  Certificado self-signed — aceite no navegador em: https://localhost:3002');
  console.log('   (Isso só precisa ser feito uma vez)\n');
});
