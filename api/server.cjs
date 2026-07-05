const http = require('http');
const { readFileSync, unlinkSync, existsSync } = require('fs');
const { join, extname } = require('path');
const { tmpdir } = require('os');

// Load generators (local files — works in any deployment)
const { generatePPTX } = require('./generate_pptx.cjs');
const { generateDOCX }  = require('./generate_docx.cjs');

if (typeof generatePPTX !== 'function') throw new Error('generatePPTX not loaded');
if (typeof generateDOCX  !== 'function') throw new Error('generateDOCX not loaded');
console.log('✅ Générateurs chargés (PPTX + DOCX)');

const PORT   = parseInt(process.env.PORT || '3737');
const DIST   = join(__dirname, '..', 'dist');
const IS_PROD = existsSync(DIST);

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.json': 'application/json',
  '.png':  'image/png',
  '.svg':  'image/svg+xml',
  '.ico':  'image/x-icon',
  '.woff2':'font/woff2',
};

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => { try { resolve(JSON.parse(body)); } catch { reject(new Error('JSON invalide')); } });
    req.on('error', reject);
  });
}

function serveStatic(res, filePath) {
  try {
    const buf = readFileSync(filePath);
    const ext = extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Cache-Control': ext === '.html' ? 'no-cache' : 'public, max-age=31536000' });
    res.end(buf);
  } catch {
    try {
      const html = readFileSync(join(DIST, 'index.html'));
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      res.end(html);
    } catch { res.writeHead(404); res.end('Not found'); }
  }
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') { res.writeHead(204, CORS); res.end(); return; }

  if (req.method === 'GET' && req.url === '/health') {
    res.writeHead(200, { ...CORS, 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, version: '1.0.0', prod: IS_PROD }));
    return;
  }

  if (req.method === 'POST' && req.url === '/generate-pptx') {
    try {
      const data = await readBody(req);
      const tmp  = join(tmpdir(), `pptx_${Date.now()}_${Math.random().toString(36).slice(2)}.pptx`);
      await generatePPTX(data, tmp);
      const buf  = readFileSync(tmp);
      try { unlinkSync(tmp); } catch {}
      const fname = `PitchDeck_${(data.projectTitle || 'startup').replace(/[^\w\-]+/g, '_')}.pptx`;
      res.writeHead(200, { ...CORS,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${fname}"`, 'Content-Length': buf.length });
      res.end(buf);
    } catch (e) {
      console.error('[PPTX]', e.message);
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/generate-docx') {
    try {
      const data = await readBody(req);
      const tmp  = join(tmpdir(), `docx_${Date.now()}_${Math.random().toString(36).slice(2)}.docx`);
      await generateDOCX(data, tmp);
      const buf  = readFileSync(tmp);
      try { unlinkSync(tmp); } catch {}
      const fname = `Guide_Projet_${(data.projectTitle || 'startup').replace(/[^\w\-]+/g, '_')}.docx`;
      res.writeHead(200, { ...CORS,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${fname}"`, 'Content-Length': buf.length });
      res.end(buf);
    } catch (e) {
      console.error('[DOCX]', e.message);
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (IS_PROD) {
    let p = req.url.split('?')[0];
    if (p === '/') p = '/index.html';
    serveStatic(res, join(DIST, p));
    return;
  }

  res.writeHead(404, CORS); res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('\n╔════════════════════════════════════════════╗');
  console.log('║  🚀 StartupAI — Diplôme/Startup 1275       ║');
  console.log('║     CNCSIIU • MHESR • الجزائر              ║');
  console.log('╚════════════════════════════════════════════╝');
  console.log(`\n   URL  → http://localhost:${PORT}`);
  console.log(`   Mode → ${IS_PROD ? 'Production ✅ (serving dist/)' : 'Development 🟡'}\n`);
});

server.on('error', e => { console.error('Server error:', e.message); process.exit(1); });
