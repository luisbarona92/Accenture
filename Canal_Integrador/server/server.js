/**
 * Dashboard Update Server — Canal Integrador
 * Local Node.js server that parses Excel (Partners_BBDD_v2.xlsx) and updates
 * ONLY the "05 · Detalle Leads" section (S5) of the HTML dashboard.
 *
 * PPT parsing was removed: S0/S1/S2/S3/S4 (Resumen, Captación, Integradores,
 * Funnel ventas, Tareas) are maintained manually based on the weekly PMO
 * slide deck — the server never touches those sections.
 *
 * Port: 3001
 */

'use strict';

const http = require('http');
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

let JSZip;
try {
  JSZip = require('jszip');
} catch (e) {
  console.error('jszip not found. Run: npm install');
  process.exit(1);
}

// ── Fixed paths ──────────────────────────────────────────────────────────────
const EXCEL_PATH = 'C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/Partners_BBDD_v2.xlsx';
const HTML_SRC   = 'C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/Canal_Integrador/Seguimiento_dashboard.html';
const HTML_DST   = 'C:/Users/luis.barona.arroyo/tmpAccenture/Canal_Integrador/Seguimiento_dashboard.html';
const GIT_REPO   = 'C:/Users/luis.barona.arroyo/tmpAccenture';
const PORT       = 3001;
const SERVER_VERSION = '2.0-excel-only';

// ── Integrador name normalisation ────────────────────────────────────────────

const INT_NAME_MAP = {
  'MAXEN TECHNOLOGIES SL':'Maxen','MAXEN':'Maxen',
  'OX-ONE NETWORKS SL':'Ox-One','OX-ONE':'Ox-One',
  'COVERNET':'Covernet','COVER NET':'Covernet',
  'DUOIT':'DuoIT','DuoIT':'DuoIT',
  'COBITRATEL':'Cobitratel','Cobitratel':'Cobitratel',
  'DATAQU':'Dataqu','Dataqu':'Dataqu',
  'SISYTEC':'Sisytec','Sisytec':'Sisytec',
  'SICTEMAS':'Sictemas','Sictemas':'Sictemas',
  'MARISKALNET':'Mariskalnet','Mariskalnet':'Mariskalnet',
  'PAPERLABS':'Paperlabs','Paperlabs':'Paperlabs',
  'LISOT':'Lisot','Lisot':'Lisot',
  'AIDEEA':'Aideea','Aideea':'Aideea',
  'AMIRITMO':'Amiritmo','Amiritmo':'Amiritmo',
  'MICRO SIP':'Microsip','MICROSIP':'Microsip','Microsip':'Microsip',
  'BINFOR SOLUTIONS':'Binfor Solutions','BINFOR':'Binfor Solutions',
  'ELPIS INFORMATICA':'Elpis Informática','ELPIS':'Elpis Informática',
};

function normalizeIntegrador(raw) {
  const key = (raw || '').trim();
  if (INT_NAME_MAP[key]) return INT_NAME_MAP[key];
  const stripped = key.replace(/\s+(SL|SA|SLU|SCP|CB|SLL)$/i,'').trim();
  if (INT_NAME_MAP[stripped]) return INT_NAME_MAP[stripped];
  return key.replace(/\b\w/g, c => c.toUpperCase()).slice(0, 25);
}

function normalizeEstado(raw) {
  const r = (raw || '').toLowerCase().trim();
  if (r.includes('venta')) return 'Venta realizada';
  if (r.includes('en curso')) return 'En curso';
  if (r === 'ko') return 'KO';
  return 'Pendiente';
}

// ── XML helpers ──────────────────────────────────────────────────────────────

function getAttr(str, attr) {
  const re = new RegExp(attr + '="([^"]*)"');
  const m = str.match(re);
  return m ? m[1] : null;
}

// Parse sharedStrings.xml → array of strings
// Excel uses plain <t> tags (NOT <a:t> like PPTX)
function parseSharedStrings(xml) {
  const strings = [];
  const siRe = /<si>([\s\S]*?)<\/si>/g;
  let si;
  while ((si = siRe.exec(xml)) !== null) {
    const texts = [];
    const tRe = /<t[^>]*>([^<]*)<\/t>/g;
    let t;
    while ((t = tRe.exec(si[1])) !== null) {
      const txt = t[1]
        .replace(/&amp;/g,'&').replace(/&lt;/g,'<')
        .replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'");
      if (txt) texts.push(txt);
    }
    strings.push(texts.join(''));
  }
  return strings;
}

// Parse a cell value from sheet XML given shared strings
function cellVal(cellXml, shared) {
  if (!cellXml) return '';
  const tAttr = getAttr(cellXml, 't');
  const vMatch = cellXml.match(/<v>([\s\S]*?)<\/v>/);
  if (!vMatch) return '';
  const v = vMatch[1].trim();
  if (tAttr === 's') {
    const idx = parseInt(v, 10);
    return shared[idx] || '';
  }
  return v;
}

// ── Excel parser ─────────────────────────────────────────────────────────────

async function parseExcel() {
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error('Excel no encontrado: ' + EXCEL_PATH);
  }
  const buf = fs.readFileSync(EXCEL_PATH);
  const zip = await JSZip.loadAsync(buf);

  const sheetFile = zip.file('xl/worksheets/sheet4.xml');
  if (!sheetFile) throw new Error('Hoja 4 (Leads) no encontrada en el Excel');
  const sheetXml = await sheetFile.async('string');
  const ssXml    = await zip.file('xl/sharedStrings.xml').async('string');
  const shared   = parseSharedStrings(ssXml);

  // Single-pass: collect row5 (totals) and all lead rows (11+)
  const rowRe = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let row5xml = '';
  const leadRowsMap = {};
  let m;
  while ((m = rowRe.exec(sheetXml)) !== null) {
    const rNum = parseInt(m[1]);
    if (rNum === 5)       row5xml = m[2];
    else if (rNum >= 11)  leadRowsMap[rNum] = m[2];
  }

  function getCell(col, rowNum, rowContent) {
    const startTag = '<c r="' + col + rowNum + '"';
    const startIdx = rowContent.indexOf(startTag);
    if (startIdx === -1) return '';
    const endIdx = rowContent.indexOf('</c>', startIdx);
    if (endIdx === -1) return '';
    const cellXml = rowContent.slice(startIdx, endIdx + 4);
    return cellVal(cellXml, shared);
  }

  // Row 5 — totals
  const total     = parseInt(getCell('C', 5, row5xml), 10) || 0;
  const ventas    = parseInt(getCell('F', 5, row5xml), 10) || 0;
  const enCurso   = parseInt(getCell('G', 5, row5xml), 10) || 0;
  const pendiente = parseInt(getCell('H', 5, row5xml), 10) || 0;
  const ko        = parseInt(getCell('I', 5, row5xml), 10) || 0;

  // Rows 11+ — individual leads
  const leads = [];
  const sortedNums = Object.keys(leadRowsMap).map(Number).sort((a,b) => a-b);
  for (const rowNum of sortedNums) {
    const rc = leadRowsMap[rowNum];
    const g  = col => getCell(col, rowNum, rc);

    const id = parseInt(g('B'), 10);
    if (!id || isNaN(id)) continue;

    leads.push({
      id,
      nombre:               g('C') || 'Lead ' + id,
      cif:                  g('D') || '',
      integrador:           normalizeIntegrador(g('E')),
      modelo:               (g('F') || '').includes('2') ? 'Tier 2' : 'Tier 1',
      fecGen:               parseFloat(g('G')) || 0,
      comercial:            g('H') || '',
      estado:               normalizeEstado(g('I')),
      segmento:             g('K') || '',
      sigPasoEq:            g('L') || '',
      fecSigPasoEq:         parseFloat(g('M')) || 0,
      sigPasoComercial:     g('N') || '',
      fecSigPasoComercial:  parseFloat(g('O')) || 0,
      venta:                g('P') ? (parseFloat(g('P')) || null) : null,
      productos:            g('Q') || ''
    });
  }

  return { total, ventas, enCurso, pendiente, ko, leads };
}

// ── HTML patcher — ONLY S5 ───────────────────────────────────────────────────

/**
 * Update data-t attribute AND inner text for a data-cnt element.
 */
function replaceCnt(h, cntKey, val) {
  h = h.replace(
    new RegExp(`(data-cnt="${cntKey}"[^>]*data-t=")[^"]+(")`),
    `$1${val}$2`
  );
  h = h.replace(
    new RegExp(`(data-cnt="${cntKey}"[^>]*>)\\d+(<)`),
    `$1${val}$2`
  );
  return h;
}

/**
 * Patch HTML with S5-only data.
 * This function is intentionally limited to the "05 · Detalle Leads" section
 * and its KPI counters. Other sections are maintained manually (see header).
 */
function patchHTMLOnlyS5(html, s5) {
  if (!s5) return html;
  let h = html;

  const { total, ventas, enCurso, pendiente, ko } = s5;

  // S5 KPI counters
  h = replaceCnt(h, 's5-total',     total);
  h = replaceCnt(h, 's5-ventas',    ventas);
  h = replaceCnt(h, 's5-encurso',   enCurso);
  h = replaceCnt(h, 's5-pendiente', pendiente);
  h = replaceCnt(h, 's5-ko',        ko);

  // Subtitle ("N leads potenciales")
  h = h.replace(
    /(data-cnt="s5-subtitle"[^>]*>)\d+ leads potenciales(<)/,
    `$1${total} leads potenciales$2`
  );
  // Leads chip ("N leads")
  h = h.replace(
    /(data-cnt="s5-leads-count"[^>]*>)\d+ leads(<)/,
    `$1${total} leads$2`
  );

  // Rebuild LEADS JS array from Excel rows
  if (s5.leads && s5.leads.length > 0) {
    const leadsJson = JSON.stringify(s5.leads, null, 2)
      .replace(/</g, '\\u003C')
      .replace(/>/g, '\\u003E');
    const newLeads = 'const LEADS = ' + leadsJson + ';';
    const leadsStart = h.indexOf('const LEADS = [');
    if (leadsStart !== -1) {
      const leadsEnd = h.indexOf('\n];', leadsStart);
      if (leadsEnd !== -1) {
        h = h.slice(0, leadsStart) + newLeads + h.slice(leadsEnd + 3);
      }
    }
  }

  return h;
}

// ── Core parse function (Excel only) ─────────────────────────────────────────

async function doParse() {
  const log = [];

  log.push({ type:'info', msg:'Modo: solo Excel (PPT ignorado por diseño).' });
  log.push({ type:'info', msg:'Parseando Excel (Partners_BBDD_v2.xlsx)...' });

  let s5;
  try {
    s5 = await parseExcel();
    log.push({ type:'ok', msg:`Excel OK: total=${s5.total} ventas=${s5.ventas} enCurso=${s5.enCurso} pendiente=${s5.pendiente} ko=${s5.ko} (${s5.leads.length} filas de detalle).` });
  } catch (e) {
    log.push({ type:'err', msg:`Error parseando Excel: ${e.message}` });
    throw e;
  }

  const data = {
    meta: {
      source: 'excel-only',
      scope:  's5',
      parsedAt: new Date().toISOString()
    },
    s5
  };

  log.push({ type:'ok', msg:'Parseo completado. Listo para aplicar solo a 05 · Detalle Leads.' });

  return { log, data };
}

// ── HTTP server ──────────────────────────────────────────────────────────────

function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function json(res, data, code) {
  cors(res);
  res.writeHead(code || 200, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(data));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => { try { resolve(JSON.parse(body || '{}')); } catch(e) { resolve({}); } });
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  // OPTIONS preflight
  if (req.method === 'OPTIONS') {
    cors(res);
    res.writeHead(204);
    res.end();
    return;
  }

  const url = req.url.split('?')[0];

  // GET /api/ping
  if (req.method === 'GET' && url === '/api/ping') {
    json(res, { ok: true, version: SERVER_VERSION, scope: 's5-only' });
    return;
  }

  // POST /api/parse  (Excel only)
  if (req.method === 'POST' && url === '/api/parse') {
    try {
      // Body is accepted but ignored — scope is fixed to s5/excel-only.
      await readBody(req);
      const result = await doParse();
      json(res, { ok: true, ...result });
    } catch (e) {
      json(res, { ok: false, error: e.message }, 500);
    }
    return;
  }

  // POST /api/apply  (Excel only → only S5 section of HTML)
  if (req.method === 'POST' && url === '/api/apply') {
    try {
      await readBody(req);
      const parseResult = await doParse();
      const log = [...parseResult.log];
      const data = parseResult.data;

      log.push({ type:'info', msg:'Leyendo HTML fuente...' });
      let html = fs.readFileSync(HTML_SRC, 'utf8');

      log.push({ type:'info', msg:'Aplicando cambios SOLO a 05 · Detalle Leads...' });
      const htmlBefore = html;
      html = patchHTMLOnlyS5(html, data.s5);

      if (html === htmlBefore) {
        log.push({ type:'info', msg:'HTML sin cambios (Excel ya sincronizado).' });
        json(res, { ok: true, log, data });
        return;
      }

      log.push({ type:'info', msg:'Escribiendo HTML fuente...' });
      fs.writeFileSync(HTML_SRC, html, 'utf8');

      log.push({ type:'info', msg:'Copiando HTML al repo git...' });
      const dstDir = path.dirname(HTML_DST);
      if (!fs.existsSync(dstDir)) fs.mkdirSync(dstDir, { recursive: true });
      fs.writeFileSync(HTML_DST, html, 'utf8');
      log.push({ type:'ok', msg:`HTML copiado a ${HTML_DST}` });

      // Git operations
      log.push({ type:'info', msg:'Ejecutando git add + commit + push...' });
      try {
        execSync('git add Canal_Integrador/Seguimiento_dashboard.html', { cwd: GIT_REPO, stdio: 'pipe' });
        const stamp = new Date().toLocaleDateString('es-ES', { day:'2-digit', month:'2-digit', year:'numeric' });
        const commitMsg = `chore: auto-update S5 Detalle Leads desde Excel (${stamp})`;
        execSync(`git commit -m "${commitMsg}"`, { cwd: GIT_REPO, stdio: 'pipe' });
        execSync('git push origin main', { cwd: GIT_REPO, stdio: 'pipe' });
        log.push({ type:'ok', msg:'Git push completado correctamente.' });
      } catch (gitErr) {
        const errMsg = gitErr.stderr ? gitErr.stderr.toString() : gitErr.message;
        if (/nothing to commit/i.test(errMsg)) {
          log.push({ type:'info', msg:'Git: nada que commitear (sin cambios).' });
        } else {
          log.push({ type:'err', msg:`Git error: ${errMsg.slice(0,200)}` });
        }
      }

      json(res, { ok: true, log, data });
    } catch (e) {
      json(res, { ok: false, error: e.message }, 500);
    }
    return;
  }

  // 404
  cors(res);
  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ ok: false, error: 'Not found' }));
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`✅ Dashboard update server v${SERVER_VERSION} running at http://localhost:${PORT}`);
  console.log(`   Scope: Excel → 05 · Detalle Leads (S5) únicamente`);
  console.log(`   Endpoints: GET /api/ping  POST /api/parse  POST /api/apply`);
});

server.on('error', e => {
  if (e.code === 'EADDRINUSE') {
    console.error(`❌ Port ${PORT} already in use. Stop the existing server first.`);
  } else {
    console.error('Server error:', e.message);
  }
  process.exit(1);
});
