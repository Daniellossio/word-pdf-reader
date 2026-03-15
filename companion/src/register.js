// ─────────────────────────────────────────────────────────────────────────────
// register.js — Registra o companion app como handler padrão de .pdf no Windows
// Execute como Administrador: node register.js install | uninstall
// ─────────────────────────────────────────────────────────────────────────────

const { execSync } = require('child_process');
const path = require('path');
const fs   = require('fs');

const EXE_NAME  = 'pdf-reader-companion.exe';
const APP_NAME  = 'PDFReaderForWord';
const APP_DESC  = 'PDF Reader for Word';

// Caminho do executável (após pkg build)
const exePath = path.resolve(__dirname, '..', 'dist', EXE_NAME);

function install() {
  if (!fs.existsSync(exePath)) {
    console.error(`Executável não encontrado: ${exePath}`);
    console.error('Execute primeiro: npm run build');
    process.exit(1);
  }

  console.log(`Registrando ${APP_NAME} como leitor de PDF...`);
  console.log(`Executável: ${exePath}`);

  // Escapa o caminho para o registro
  const escaped = exePath.replace(/\\/g, '\\\\');

  // Comandos REG para registrar o handler de .pdf
  const commands = [
    // Registra o app
    `reg add "HKCU\\Software\\Classes\\${APP_NAME}" /ve /d "${APP_DESC}" /f`,
    `reg add "HKCU\\Software\\Classes\\${APP_NAME}\\shell\\open\\command" /ve /d "\\"${escaped}\\" \\"%1\\"" /f`,

    // Registra a extensão .pdf
    `reg add "HKCU\\Software\\Classes\\.pdf" /ve /d "${APP_NAME}" /f`,
    `reg add "HKCU\\Software\\Classes\\.pdf\\OpenWithProgids" /v "${APP_NAME}" /d "" /f`,

    // Registra como aplicativo capaz de abrir PDFs
    `reg add "HKCU\\Software\\Classes\\Applications\\${EXE_NAME}\\shell\\open\\command" /ve /d "\\"${escaped}\\" \\"%1\\"" /f`,
    `reg add "HKCU\\Software\\Classes\\Applications\\${EXE_NAME}\\SupportedTypes" /v ".pdf" /d "" /f`,

    // Registra no OpenWithList para aparecer nas opções "Abrir com"
    `reg add "HKCU\\Software\\Classes\\.pdf\\OpenWithList\\${EXE_NAME}" /f`,
  ];

  commands.forEach(cmd => {
    try {
      execSync(cmd, { stdio: 'ignore' });
    } catch (err) {
      console.warn(`Aviso: ${cmd}\n  → ${err.message}`);
    }
  });

  // Notifica o Windows para atualizar as associações
  try {
    execSync('ie4uinit.exe -show', { stdio: 'ignore' });
  } catch {}

  console.log('\n✅ Registrado com sucesso!');
  console.log('Para definir como padrão: clique direito em qualquer PDF → "Abrir com" → "PDF Reader for Word" → "Sempre usar este aplicativo"');
  console.log('Ou vá em: Configurações → Aplicativos → Aplicativos padrão → Procurar um padrão por tipo de arquivo → .pdf');
}

function uninstall() {
  console.log(`Removendo ${APP_NAME}...`);

  const commands = [
    `reg delete "HKCU\\Software\\Classes\\${APP_NAME}" /f`,
    `reg delete "HKCU\\Software\\Classes\\Applications\\${EXE_NAME}" /f`,
    `reg delete "HKCU\\Software\\Classes\\.pdf\\OpenWithList\\${EXE_NAME}" /f`,
    `reg delete "HKCU\\Software\\Classes\\.pdf\\OpenWithProgids" /v "${APP_NAME}" /f`,
  ];

  commands.forEach(cmd => {
    try {
      execSync(cmd, { stdio: 'ignore' });
    } catch {}
  });

  console.log('✅ Removido com sucesso.');
}

const action = process.argv[2];
if (action === 'install')   install();
else if (action === 'uninstall') uninstall();
else {
  console.log('Uso: node register.js install | uninstall');
  process.exit(1);
}
