'use strict';
const http  = require('http');
const https = require('https');
const { readFileSync, unlinkSync, existsSync } = require('fs');
const { join, extname } = require('path');
const { tmpdir } = require('os');

// ── Generators ─────────────────────────────────────────────────────────────
const { generatePPTX } = require('./generate_pptx.cjs');
const { generateDOCX }  = require('./generate_docx.cjs');
console.log('✅ Generators loaded');

// ── Config ──────────────────────────────────────────────────────────────────
const PORT    = parseInt(process.env.PORT || '3737');
const API_KEY = process.env.ANTHROPIC_API_KEY || '';
const DIST    = join(__dirname, '..', 'dist');
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

// ── Helpers ─────────────────────────────────────────────────────────────────
function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); }
      catch { reject(new Error('JSON invalide')); }
    });
    req.on('error', reject);
  });
}

function serveStatic(res, filePath) {
  try {
    const buf = readFileSync(filePath);
    const ext = extname(filePath);
    res.writeHead(200, {
      'Content-Type': MIME[ext] || 'application/octet-stream',
      'Cache-Control': ext === '.html' ? 'no-cache' : 'public, max-age=31536000',
    });
    res.end(buf);
  } catch {
    try {
      const html = readFileSync(join(DIST, 'index.html'));
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      res.end(html);
    } catch { res.writeHead(404); res.end('Not found'); }
  }
}

// ── Claude proxy helper ──────────────────────────────────────────────────────
function proxyToClaude(body, stream) {
  return new Promise((resolve, reject) => {
    body.stream = stream;
    const payload = JSON.stringify(body);
    const req = https.request({
      hostname: 'api.anthropic.com',
      path:     '/v1/messages',
      method:   'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Length':    Buffer.byteLength(payload),
      },
    }, resolve);
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

// ── Request handler ──────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {

  // CORS preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, CORS); res.end(); return;
  }

  // Health check
  if (req.method === 'GET' && req.url === '/health') {
    res.writeHead(200, { ...CORS, 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, prod: IS_PROD, hasKey: !!API_KEY }));
    return;
  }

  // ── Claude non-streaming proxy ─────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/claude') {
    try {
      const body    = await readBody(req);
      const apiRes  = await proxyToClaude(body, false);
      const chunks  = [];
      apiRes.on('data', c => chunks.push(c));
      apiRes.on('end', () => {
        res.writeHead(apiRes.statusCode, { ...CORS, 'Content-Type': 'application/json' });
        res.end(Buffer.concat(chunks));
      });
    } catch (e) {
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // ── Claude streaming proxy (SSE) ───────────────────────────────────────
  if (req.method === 'POST' && req.url === '/claude/stream') {
    try {
      const body   = await readBody(req);
      const apiRes = await proxyToClaude(body, true);
      res.writeHead(200, {
        ...CORS,
        'Content-Type':  'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection':    'keep-alive',
      });
      apiRes.pipe(res);
      req.on('close', () => apiRes.destroy());
    } catch (e) {
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // ── Generate PPTX ──────────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/generate-pptx') {
    try {
      const data = await readBody(req);
      const tmp  = join(tmpdir(), `pptx_${Date.now()}.pptx`);
      await generatePPTX(data, tmp);
      const buf  = readFileSync(tmp);
      try { unlinkSync(tmp); } catch {}
      const fname = `PitchDeck_${(data.projectTitle || 'startup').replace(/[^\w\-]+/g, '_')}.pptx`;
      res.writeHead(200, {
        ...CORS,
        'Content-Type':        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${fname}"`,
        'Content-Length':      buf.length,
      });
      res.end(buf);
    } catch (e) {
      console.error('[PPTX]', e.message);
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // ── Generate DOCX ──────────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/generate-docx') {
    try {
      const data = await readBody(req);
      const tmp  = join(tmpdir(), `docx_${Date.now()}.docx`);
      await generateDOCX(data, tmp);
      const buf  = readFileSync(tmp);
      try { unlinkSync(tmp); } catch {}
      const fname = `Guide_Projet_${(data.projectTitle || 'startup').replace(/[^\w\-]+/g, '_')}.docx`;
      res.writeHead(200, {
        ...CORS,
        'Content-Type':        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${fname}"`,
        'Content-Length':      buf.length,
      });
      res.end(buf);
    } catch (e) {
      console.error('[DOCX]', e.message);
      res.writeHead(500, { ...CORS, 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  // ── Serve static frontend ──────────────────────────────────────────────
  if (IS_PROD) {
    let p = req.url.split('?')[0];
    if (p === '/') p = '/index.html';
    serveStatic(res, join(DIST, p));
    return;
  }

  res.writeHead(404, CORS); res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n🚀 StartupAI → http://localhost:${PORT}`);
  console.log(`   Claude API key: ${API_KEY ? '✅ configured' : '❌ MISSING — set ANTHROPIC_API_KEY'}`);
  console.log(`   Mode: ${IS_PROD ? 'Production ✅' : 'Development 🟡'}\n`);
});

server.on('error', e => { console.error('Server error:', e.message); process.exit(1); });
