/**
 * Dashboard Update Server — Canal Integrador
 * Local Node.js server that parses Excel + PPTX and applies changes to HTML dashboard.
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
const EXCEL_PATH   = 'C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/Partners_BBDD_v1.xlsx';
const PPT_DIR      = 'C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE';
const HTML_SRC     = 'C:/Users/luis.barona.arroyo/OneDrive - Accenture/Documents/MAS ORANGE/Canal_Integrador/Seguimiento_dashboard.html';
const HTML_DST     = 'C:/Users/luis.barona.arroyo/tmpAccenture/Canal_Integrador/Seguimiento_dashboard.html';
const GIT_REPO     = 'C:/Users/luis.barona.arroyo/tmpAccenture';
const PORT         = 3001;

// ── Helpers ──────────────────────────────────────────────────────────────────

function findLatestPPT() {
  const files = fs.readdirSync(PPT_DIR)
    .filter(f => /Weekly PMO.*\.pptx$/i.test(f))
    .map(f => ({ name: f, mtime: fs.statSync(path.join(PPT_DIR, f)).mtime }))
    .sort((a, b) => b.mtime - a.mtime);
  if (!files.length) throw new Error('No PPT file matching /Weekly PMO.*\\.pptx$/i found in ' + PPT_DIR);
  return path.join(PPT_DIR, files[0].name);
}

function colorToGroup(hex) {
  if (!hex) return 'posible';
  try {
    const r = parseInt(hex.slice(0,2), 16);
    const g = parseInt(hex.slice(2,4), 16);
    const b = parseInt(hex.slice(4,6), 16);
    // green: high g, low r, low b
    if (g > 120 && r < 100 && b < 100) return 'verde';
    // red: high r, low g, low b
    if (r > 150 && g < 80 && b < 80) return 'churn';
    // gray: similar r,g,b all medium
    if (Math.abs(r-g) < 30 && Math.abs(g-b) < 30 && r > 100 && r < 200) return 'hold';
    // orange/amber: high r, medium g, low b
    if (r > 150 && g > 80 && b < 80) return 'posible';
    return 'posible';
  } catch (e) {
    return 'posible';
  }
}

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
  if (r.includes('venta')) return 'Venta';
  if (r.includes('en curso')) return 'En curso';
  if (r === 'ko') return 'KO';
  return 'Pendiente';
}

// ── XML helpers ──────────────────────────────────────────────────────────────

function extractAllText(xml) {
  const matches = [];
  const re = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const txt = m[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'").trim();
    if (txt) matches.push(txt);
  }
  return matches;
}

function getAttr(str, attr) {
  const re = new RegExp(attr + '="([^"]*)"');
  const m = str.match(re);
  return m ? m[1] : null;
}

// Parse sharedStrings.xml → array of strings
// Excel uses plain <t> tags (NOT <a:t> like PPTX), so we use a dedicated regex
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
  const buf = fs.readFileSync(EXCEL_PATH);
  const zip = await JSZip.loadAsync(buf);

  const sheetXml = await zip.file('xl/worksheets/sheet4.xml').async('string');
  const ssXml    = await zip.file('xl/sharedStrings.xml').async('string');
  const shared   = parseSharedStrings(ssXml);

  // Single-pass: collect row5 and all lead rows (11+)
  const rowRe = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let row5xml = '';
  const leadRowsMap = {};
  let m;
  while ((m = rowRe.exec(sheetXml)) !== null) {
    const rNum = parseInt(m[1]);
    if (rNum === 5)       row5xml = m[2];
    else if (rNum >= 11)  leadRowsMap[rNum] = m[2];
  }

  // Get a cell from given row content (indexOf avoids regex backslash escaping issues)
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

// ── PPT parser ───────────────────────────────────────────────────────────────

async function parsePPT(pptPath) {
  const buf = fs.readFileSync(pptPath);
  const zip = await JSZip.loadAsync(buf);

  const slides = {};
  for (let n = 1; n <= 6; n++) {
    const key = `ppt/slides/slide${n}.xml`;
    if (zip.file(key)) {
      slides[n] = await zip.file(key).async('string');
    } else {
      slides[n] = '';
    }
  }

  // ── Slide 1: date ──────────────────────────────────────────────────────────
  let date = '';
  try {
    const allTexts = extractAllText(slides[1]);
    for (const t of allTexts) {
      if (/\d+\s+\w+\s+de\s+\d{4}/i.test(t) || /\d{1,2}\s+\w+\s+\d{4}/i.test(t)) {
        date = t.trim();
        break;
      }
    }
    if (!date) {
      // fallback: look for date-like patterns
      for (const t of allTexts) {
        if (/\d{1,2}[\.\-\/]\d{1,2}[\.\-\/]\d{2,4}/.test(t)) {
          date = t.trim();
          break;
        }
      }
    }
  } catch (e) { date = ''; }

  // ── Slide 3: captación / funnel ────────────────────────────────────────────
  let s1 = { h1:0, h2:0, h3:0, h4:0, h4pct:'', h3pct:'', h2pct:'',
              firmados: {count:0, names:[]}, enProceso: {count:0, names:[]}, churn: {count:0, names:[]} };
  try {
    const xml3 = slides[3];

    // Extract shapes with single-number text + y position
    const spRe = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    const numShapes = [];
    let sp;
    while ((sp = spRe.exec(xml3)) !== null) {
      const spContent = sp[1];
      const texts = extractAllText(spContent);
      const joined = texts.join('').trim();
      if (/^\d+$/.test(joined) && parseInt(joined) > 0) {
        // Get y position
        const yMatch = spContent.match(/<a:off[^>]*y="(\d+)"/);
        const y = yMatch ? parseInt(yMatch[1]) : 999999999;
        numShapes.push({ val: parseInt(joined), y });
      }
    }

    // Sort by y ascending
    numShapes.sort((a, b) => a.y - b.y);

    // The 5 numbers from top to bottom: Pool(~5000+), H1(~800+), H2(~70), H3(~48), H4(~24)
    // Filter to only funnel candidates (large enough)
    const funnelNums = numShapes.filter(s => s.val >= 5);
    if (funnelNums.length >= 5) {
      s1.h1 = funnelNums[1].val;
      s1.h2 = funnelNums[2].val;
      s1.h3 = funnelNums[3].val;
      s1.h4 = funnelNums[4].val;
    } else if (funnelNums.length >= 4) {
      s1.h1 = funnelNums[0].val;
      s1.h2 = funnelNums[1].val;
      s1.h3 = funnelNums[2].val;
      s1.h4 = funnelNums[3].val;
    }

    // Compute percentages
    if (s1.h1 > 0) s1.h2pct = Math.round(s1.h2 / s1.h1 * 100) + '%';
    if (s1.h2 > 0) s1.h3pct = Math.round(s1.h3 / s1.h2 * 100) + '%';
    if (s1.h3 > 0) s1.h4pct = Math.round(s1.h4 / s1.h3 * 100) + '%';

    // Firmados: "contratos firmados"
    const allTexts3 = extractAllText(xml3);
    for (let i = 0; i < allTexts3.length; i++) {
      const t = allTexts3[i];
      if (/contrato[s]?\s+firmado[s]?/i.test(t)) {
        // Number before
        const numBefore = allTexts3[i-1] ? parseInt(allTexts3[i-1]) : 0;
        if (!isNaN(numBefore) && numBefore > 0) s1.firmados.count = numBefore;
        // Names after colon
        const colonIdx = t.indexOf(':');
        if (colonIdx > -1) {
          s1.firmados.names = t.slice(colonIdx+1).split(',').map(s => s.trim()).filter(Boolean);
        } else if (allTexts3[i+1]) {
          s1.firmados.names = allTexts3[i+1].split(',').map(s => s.trim()).filter(Boolean);
        }
      }
      if (/en proceso de firma/i.test(t)) {
        const numBefore = allTexts3[i-1] ? parseInt(allTexts3[i-1]) : 0;
        if (!isNaN(numBefore) && numBefore > 0) s1.enProceso.count = numBefore;
        const colonIdx = t.indexOf(':');
        if (colonIdx > -1) {
          s1.enProceso.names = t.slice(colonIdx+1).split(',').map(s => s.trim()).filter(Boolean);
        } else if (allTexts3[i+1]) {
          s1.enProceso.names = allTexts3[i+1].split(',').map(s => s.trim()).filter(Boolean);
        }
      }
      if (/rechazos/i.test(t)) {
        const numBefore = allTexts3[i-1] ? parseInt(allTexts3[i-1]) : 0;
        if (!isNaN(numBefore) && numBefore > 0) s1.churn.count = numBefore;
      }
    }
  } catch (e) { console.error('Slide 3 parse error:', e.message); }

  // ── Slide 4: integradores ──────────────────────────────────────────────────
  let s2 = { total: 0, integradores: [], counts: {verde:0, posible:0, hold:0, churn:0} };
  try {
    const xml4 = slides[4];

    // Find all tables
    const tblRe = /<a:tbl>([\s\S]*?)<\/a:tbl>/g;
    let tbl;
    const tables = [];
    while ((tbl = tblRe.exec(xml4)) !== null) {
      tables.push(tbl[1]);
    }

    for (const tblXml of tables) {
      const rowRe2 = /<a:tr[\s\S]*?>([\s\S]*?)<\/a:tr>/g;
      let rowMatch;
      let rowIdx = 0;
      while ((rowMatch = rowRe2.exec(tblXml)) !== null) {
        rowIdx++;
        if (rowIdx === 1) continue; // skip header
        const rowXml = rowMatch[1];

        // Extract cells
        const cellRe = /<a:tc>([\s\S]*?)<\/a:tc>/g;
        const cells = [];
        let cellMatch;
        while ((cellMatch = cellRe.exec(rowXml)) !== null) {
          const cellTexts = extractAllText(cellMatch[1]);
          cells.push({ text: cellTexts.join(' ').trim(), xml: cellMatch[1] });
        }

        if (cells.length < 2) continue;
        const nombre = cells[0] ? cells[0].text : '';
        if (!nombre || nombre.length < 2) continue;

        const observacion = cells[1] ? cells[1].text : '';

        // Detect group from dot cell (last cell or any cell with ●)
        let grupo = 'posible';
        for (const cell of cells) {
          if (cell.text.includes('●') || cell.text.includes('\u25CF')) {
            // Get color
            const srgbMatch = cell.xml.match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/);
            if (srgbMatch) {
              grupo = colorToGroup(srgbMatch[1]);
            } else {
              const schemeMatch = cell.xml.match(/<a:schemeClr val="([^"]+)"/);
              if (schemeMatch) {
                const sc = schemeMatch[1];
                if (sc === 'accent6' || sc === 'accent3') grupo = 'verde';
                else if (sc === 'accent2') grupo = 'churn';
                else if (sc === 'accent4' || sc === 'accent5') grupo = 'posible';
                else grupo = 'hold';
              }
            }
            break;
          }
        }

        // Guess modelo from observacion
        let modelo = 'LeadGen';
        if (/e2e|end.to.end/i.test(observacion)) modelo = 'e2e';
        else if (/leadgen|lead gen/i.test(observacion)) modelo = 'LeadGen';

        s2.integradores.push({ nombre, observacion, grupo, modelo, estado: grupo === 'verde' ? 'Activo' : grupo === 'churn' ? 'Churn' : 'En proceso' });
      }
    }

    // Count by group
    s2.integradores.forEach(i => { if (s2.counts[i.grupo] !== undefined) s2.counts[i.grupo]++; });
    s2.total = s2.integradores.length;

    // Try to extract total from subtitle
    const allTexts4 = extractAllText(xml4);
    for (const t of allTexts4) {
      const m = t.match(/Detalle de los (\d+) integrador/i);
      if (m) { s2.total = parseInt(m[1]); break; }
    }
  } catch (e) { console.error('Slide 4 parse error:', e.message); }

  // ── Slide 5: funnel ventas ─────────────────────────────────────────────────
  let s3 = { total:0, leads:0, ofertas:0, ventas:0, leadsPct:'', ofertasPct:'', ventasPct:'',
              tabla: [], eop: [] };
  try {
    const xml5 = slides[5];

    // Same funnel extraction as slide 3
    const spRe2 = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    const numShapes2 = [];
    let sp2;
    while ((sp2 = spRe2.exec(xml5)) !== null) {
      const spContent = sp2[1];
      const texts = extractAllText(spContent);
      const joined = texts.join('').trim();
      if (/^\d+$/.test(joined) && parseInt(joined) > 0) {
        const yMatch = spContent.match(/<a:off[^>]*y="(\d+)"/);
        const y = yMatch ? parseInt(yMatch[1]) : 999999999;
        numShapes2.push({ val: parseInt(joined), y });
      }
    }
    numShapes2.sort((a, b) => a.y - b.y);
    const funnelNums2 = numShapes2.filter(s => s.val >= 1);
    if (funnelNums2.length >= 4) {
      s3.total  = funnelNums2[0].val;
      s3.leads  = funnelNums2[1].val;
      s3.ofertas = funnelNums2[2].val;
      s3.ventas = funnelNums2[3].val;
    }

    if (s3.total > 0) {
      s3.leadsPct   = Math.round(s3.leads  / s3.total * 100) + '%';
      s3.ofertasPct = Math.round(s3.ofertas / s3.leads * 100) + '%';
      s3.ventasPct  = Math.round(s3.ventas / s3.total * 100) + '%';
    }

    // Extract subtitle with "X oportunidades"
    const allTexts5 = extractAllText(xml5);
    for (const t of allTexts5) {
      const m = t.match(/(\d+)\s+oportunidades?/i);
      if (m) { s3.total = parseInt(m[1]); break; }
    }

    // Parse integrador table
    const tblRe2 = /<a:tbl>([\s\S]*?)<\/a:tbl>/g;
    let tbl2;
    const tables5 = [];
    while ((tbl2 = tblRe2.exec(xml5)) !== null) tables5.push(tbl2[1]);

    for (const tblXml of tables5) {
      const rowRe3 = /<a:tr[\s\S]*?>([\s\S]*?)<\/a:tr>/g;
      let rowMatch3;
      let rowIdx3 = 0;
      while ((rowMatch3 = rowRe3.exec(tblXml)) !== null) {
        rowIdx3++;
        if (rowIdx3 === 1) continue;
        const rowXml = rowMatch3[1];
        const cellRe2 = /<a:tc>([\s\S]*?)<\/a:tc>/g;
        const cells = [];
        let cm;
        while ((cm = cellRe2.exec(rowXml)) !== null) {
          cells.push(extractAllText(cm[1]).join(' ').trim());
        }
        if (cells.length >= 5 && cells[0]) {
          s3.tabla.push({
            nombre:    cells[0],
            modelo:    cells[1] || '',
            oport:     parseInt(cells[2]) || 0,
            ofertas:   parseInt(cells[3]) || 0,
            ventas:    parseInt(cells[4]) || 0,
            situacion: cells[5] || '',
            proximo:   cells[6] || ''
          });
        }
      }
    }
  } catch (e) { console.error('Slide 5 parse error:', e.message); }

  // ── Slide 6: tareas ────────────────────────────────────────────────────────
  let s4 = { tareas: [], counts: {retrasadas:0, enCurso:0, finalizado:0, onHold:0} };
  try {
    const xml6 = slides[6];

    const tblRe3 = /<a:tbl>([\s\S]*?)<\/a:tbl>/g;
    let tbl3;
    while ((tbl3 = tblRe3.exec(xml6)) !== null) {
      const tblXml = tbl3[1];
      const rowRe4 = /<a:tr[\s\S]*?>([\s\S]*?)<\/a:tr>/g;
      let rowMatch4;
      let rowIdx4 = 0;
      while ((rowMatch4 = rowRe4.exec(tblXml)) !== null) {
        rowIdx4++;
        if (rowIdx4 === 1) continue;
        const rowXml = rowMatch4[1];
        const cellRe3 = /<a:tc>([\s\S]*?)<\/a:tc>/g;
        const cells = [];
        let cm3;
        while ((cm3 = cellRe3.exec(rowXml)) !== null) {
          cells.push(extractAllText(cm3[1]).join(' ').trim());
        }
        if (cells.length >= 3 && cells[0]) {
          const tarea = cells[0];
          const equipo = cells[1] || '';
          const estadoRaw = cells[2] || '';
          const fecha = cells[3] || '';

          let estado = 'En curso';
          if (/retras/i.test(estadoRaw)) estado = 'Retrasado';
          else if (/finaliz/i.test(estadoRaw) || /complet/i.test(estadoRaw)) estado = 'Finalizado';
          else if (/hold/i.test(estadoRaw)) estado = 'On hold';
          else if (/en curso/i.test(estadoRaw) || /curso/i.test(estadoRaw)) estado = 'En curso';

          s4.tareas.push({ tarea, equipo, estado, fecha });

          if (estado === 'Retrasado') s4.counts.retrasadas++;
          else if (estado === 'En curso') s4.counts.enCurso++;
          else if (estado === 'Finalizado') s4.counts.finalizado++;
          else if (estado === 'On hold') s4.counts.onHold++;
        }
      }
    }
  } catch (e) { console.error('Slide 6 parse error:', e.message); }

  return { s1, s2, s3, s4 };
}

// ── HTML patcher ─────────────────────────────────────────────────────────────

/**
 * Update data-t attribute AND inner text for a data-cnt element.
 * HTML structure: data-cnt="KEY" data-t="OLD">OLD<
 */
function replaceCnt(h, cntKey, val) {
  // Update data-t attribute value
  h = h.replace(
    new RegExp(`(data-cnt="${cntKey}"[^>]*data-t=")[^"]+(")`),
    `$1${val}$2`
  );
  // Update inner text (digits only)
  h = h.replace(
    new RegExp(`(data-cnt="${cntKey}"[^>]*>)\\d+(<)`),
    `$1${val}$2`
  );
  return h;
}

function patchHTML(html, data) {
  let h = html;

  // ── Date ──────────────────────────────────────────────────────────────────
  if (data.meta && data.meta.date) {
    const d = data.meta.date;
    h = h.replace(
      /<div class="hero-badge"><b><\/b>[^<]*<\/div>/,
      `<div class="hero-badge"><b></b> Weekly PMO · ${d}</div>`
    );
    h = h.replace(
      /<span class="date-lbl">[^<]*<\/span>/,
      `<span class="date-lbl">${d.toLowerCase()}</span>`
    );
  }

  // ── S1 KPI counters (captación funnel) ────────────────────────────────────
  if (data.s1) {
    const { h1, h2, h3, h4 } = data.s1;
    if (h1) h = replaceCnt(h, 's1-h1', h1);
    if (h2) h = replaceCnt(h, 's1-h2', h2);
    if (h3) h = replaceCnt(h, 's1-h3', h3);
    if (h4) {
      h = replaceCnt(h, 's1-h4', h4);
      // Hero activados mirror
      h = replaceCnt(h, 's0-activados', h4);
    }
    if (h3) h = replaceCnt(h, 's0-captados', h3);
  }

  // ── S3 funnel ventas counters ─────────────────────────────────────────────
  if (data.s3) {
    const { total, leads, ofertas, ventas } = data.s3;
    if (total) {
      h = replaceCnt(h, 's3-total', total);
      h = replaceCnt(h, 's0-oportunidades', total);
    }
    if (leads)   h = replaceCnt(h, 's3-leads', leads);
    if (ofertas) {
      h = replaceCnt(h, 's3-ofertas', ofertas);
      h = replaceCnt(h, 's0-ofertas', ofertas);
    }
    if (ventas) {
      h = replaceCnt(h, 's3-ventas', ventas);
      h = replaceCnt(h, 's0-ventas', ventas);
    }
  }

  // ── S5 KPI counters (leads Excel) ─────────────────────────────────────────
  if (data.s5) {
    const { total, ventas, enCurso, pendiente, ko } = data.s5;
    h = replaceCnt(h, 's5-total',     total);
    h = replaceCnt(h, 's5-ventas',    ventas);
    h = replaceCnt(h, 's5-encurso',   enCurso);
    h = replaceCnt(h, 's5-pendiente', pendiente);
    h = replaceCnt(h, 's5-ko',        ko);

    // Sync S0 resumen + S3 title with Excel total (source of truth for leads)
    h = replaceCnt(h, 's0-oportunidades', total);
    h = replaceCnt(h, 's3-total',         total);
    h = h.replace(
      /Funnel de ventas · \d+ oportunidades identificadas/,
      `Funnel de ventas · ${total} oportunidades identificadas`
    );

    // Subtitle text
    h = h.replace(
      /(data-cnt="s5-subtitle"[^>]*>)\d+ leads potenciales(<)/,
      `$1${total} leads potenciales$2`
    );
    // Leads chip
    h = h.replace(
      /(data-cnt="s5-leads-count"[^>]*>)\d+ leads(<)/,
      `$1${total} leads$2`
    );

    // ── Rebuild LEADS JS array from Excel rows ────────────────────────────
    if (data.s5.leads && data.s5.leads.length > 0) {
      const leadsJson = JSON.stringify(data.s5.leads, null, 2)
        .replace(/</g, '\\u003C')
        .replace(/>/g, '\\u003E');
      h = h.replace(
        /const LEADS\s*=\s*\[[\s\S]*?\n\];/,
        () => 'const LEADS = ' + leadsJson + ';'
      );
    }
  }

  return h;
}

// ── Core parse function ───────────────────────────────────────────────────────

async function doParse() {
  const log = [];
  const result = {};

  log.push({ type:'info', msg:'Buscando fichero PPT...' });
  let pptPath;
  try {
    pptPath = findLatestPPT();
    log.push({ type:'ok', msg:`PPT encontrado: ${path.basename(pptPath)}` });
  } catch (e) {
    log.push({ type:'err', msg:`Error buscando PPT: ${e.message}` });
    throw e;
  }

  log.push({ type:'info', msg:'Parseando Excel (S5)...' });
  let s5;
  try {
    s5 = await parseExcel();
    log.push({ type:'ok', msg:`Excel: total=${s5.total} ventas=${s5.ventas} enCurso=${s5.enCurso} pendiente=${s5.pendiente} ko=${s5.ko}` });
  } catch (e) {
    log.push({ type:'err', msg:`Error parseando Excel: ${e.message}` });
    s5 = { total:0, ventas:0, enCurso:0, pendiente:0, ko:0 };
  }

  log.push({ type:'info', msg:'Parseando PPT (slides 1-6)...' });
  let pptData;
  try {
    pptData = await parsePPT(pptPath);
    log.push({ type:'ok', msg:`PPT: fecha="${pptData.s1?'slide1 ok':'?'}" integradores=${pptData.s2 ? pptData.s2.integradores.length : 0} tareas=${pptData.s4 ? pptData.s4.tareas.length : 0}` });
  } catch (e) {
    log.push({ type:'err', msg:`Error parseando PPT: ${e.message}` });
    pptData = { s1:{}, s2:{ integradores:[], counts:{} }, s3:{}, s4:{ tareas:[], counts:{} } };
  }

  // Re-run slide 1 for date
  let date = '';
  try {
    const buf = fs.readFileSync(pptPath);
    const zip = await JSZip.loadAsync(buf);
    const slide1xml = await zip.file('ppt/slides/slide1.xml').async('string');
    const allTexts1 = extractAllText(slide1xml);
    for (const t of allTexts1) {
      if (/\d+\s+\w+\s+de\s+\d{4}/i.test(t) || /\d{1,2}\s+\w+\s+\d{4}/i.test(t)) {
        date = t.trim(); break;
      }
    }
  } catch (e) { /* ignore */ }

  result.meta = {
    date: date || new Date().toLocaleDateString('es-ES', { day:'numeric', month:'long', year:'numeric' }),
    pptFile: path.basename(pptPath),
    parsedAt: new Date().toISOString()
  };
  result.s1 = pptData.s1 || {};
  result.s2 = pptData.s2 || { integradores:[], counts:{} };
  result.s3 = pptData.s3 || {};
  result.s4 = pptData.s4 || { tareas:[], counts:{} };
  result.s5 = s5;

  log.push({ type:'ok', msg:'Parseo completado. Listo para aplicar.' });

  return { log, data: result };
}

// ── HTTP server ───────────────────────────────────────────────────────────────

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
    json(res, { ok: true, version: '1.0' });
    return;
  }

  // POST /api/parse
  if (req.method === 'POST' && url === '/api/parse') {
    try {
      const result = await doParse();
      json(res, { ok: true, ...result });
    } catch (e) {
      json(res, { ok: false, error: e.message }, 500);
    }
    return;
  }

  // POST /api/apply
  if (req.method === 'POST' && url === '/api/apply') {
    try {
      const parseResult = await doParse();
      const log = [...parseResult.log];
      const data = parseResult.data;

      log.push({ type:'info', msg:'Leyendo HTML fuente...' });
      let html = fs.readFileSync(HTML_SRC, 'utf8');

      log.push({ type:'info', msg:'Aplicando cambios al HTML...' });
      html = patchHTML(html, data);

      log.push({ type:'info', msg:'Escribiendo HTML fuente...' });
      fs.writeFileSync(HTML_SRC, html, 'utf8');

      log.push({ type:'info', msg:'Copiando HTML al repo git...' });
      // Ensure dest directory exists
      const dstDir = path.dirname(HTML_DST);
      if (!fs.existsSync(dstDir)) fs.mkdirSync(dstDir, { recursive: true });
      fs.writeFileSync(HTML_DST, html, 'utf8');
      log.push({ type:'ok', msg:`HTML copiado a ${HTML_DST}` });

      // Git operations
      log.push({ type:'info', msg:'Ejecutando git add + commit + push...' });
      try {
        execSync('git add Canal_Integrador/', { cwd: GIT_REPO, stdio: 'pipe' });
        const commitMsg = `feat: auto-update dashboard - ${data.meta.date} (${data.meta.pptFile})`;
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
  console.log(`✅ Dashboard update server running at http://localhost:${PORT}`);
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
