/**
 * Dashboard inteligencia comercial
 * - Lee Excel (SheetJS), visualiza (Plotly)
 * - Cumplimiento % = facturación real / facturación esperada
 * - Margen % = margen bruto / valor facturado
 * - Potencial = facturación esperada − facturación real
 */
(function () {
  'use strict';

  // --- constantes: sinónimos para mapeo heurístico (columna detectada por subcadena) ---
  var CC_KEYS = {
    codigo: ['codigo', 'código', 'id cliente', 'idcliente', 'cod cliente', 'code', 'n° cliente', 'nº cliente', 'cód', 'nro codigo'],
    cliente: ['cliente', 'razon', 'razón', 'customer', 'nombre', 'deudor', 'empresa', 'nombre cliente', 'razon social', 'ruc'],
    vendedor: ['vendedor', 'asesor', 'comercial', 'seller', 'ejecutiv', 'representant'],
    publico: ['publico', 'privado', 'sector', 'público', 'tipo', 'gob', 'gobierno', 'público/priv', 'público/privado'],
    licitacion: ['licit', 'concurso', 'lic tacion', 'remate'],
    fact_esp: ['esperad', 'meta', 'objetiv', 'target', 'budget', 'proyectad', 'planead', 'presup', 'monto esper'],
    fact_real: ['real', 'neta', 'neto', 'facturado', 'cobr', 'vend', 'monto real', 'fact. real', 'fact neta', 'reali'],
    contratos: ['contrat', 'n contrat', 'num contrat', 'contracts', 'cte', 'no contrat', 'cant contrat', 'n° contrat']
  };

  var MB_KEYS = {
    codigo: ['codigo', 'código', 'id cliente', 'idcliente', 'cod cliente', 'code', 'nro codigo', 'cód', 'cód.'],
    cliente: ['cliente', 'razon', 'razón', 'customer', 'nombre', 'deudor', 'empresa', 'nombre cliente'],
    vendedor: ['vendedor', 'asesor', 'comercial', 'seller', 'ejecutiv'],
    valor: ['valor factur', 'ventas', 'facturación', 'facturacion', 'monto', 'importe', 'ingreso', 'revenue', 'billing', 'vta'],
    margen: ['margen bruto', 'm.b.', 'margen  bruto', 'gross', 'm bruto', 'm/b', 'margen'],
    unidades: ['unidad', 'qty', 'cant', 'pzs', 'unidades', 'pzs.']
  };

  var state = {
    rawCC: null,
    rawMB: null,
    ccMeta: null,
    mbMeta: null,
    rows: null,
    sortMain: { col: 'cliente', dir: 1 },
    cumpCut: 'median',
    margCut: 'median',
    cumpThresh: null,
    margThresh: null
  };

  // --- normalización y números ---
  function normKey(s) {
    if (s == null || s === undefined) return '';
    return String(s)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9%/.]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function parseNumber(x) {
    if (x == null || x === '' || (typeof x === 'object' && !(x instanceof Date))) return NaN;
    if (typeof x === 'number' && !isNaN(x)) return x;
    if (x instanceof Date) return x.getTime();
    var s = String(x).trim();
    s = s.replace(/[$€£\s\u00A0]/g, '');
    if (s === '' || s === '-') return NaN;
    var c = s.replace(/[.,]/g, function (d, i, st) { return d; });
    if (c.indexOf(',') >= 0 && c.indexOf('.') >= 0) {
      if (c.lastIndexOf(',') > c.lastIndexOf('.')) s = c.replace(/\./g, '').replace(',', '.');
      else s = c.replace(/,/g, '');
    } else if (s.indexOf(',') >= 0 && s.indexOf('.') < 0) {
      var p = s.split(',');
      if (p.length === 2 && p[1].length <= 2) s = p[0].replace(/[^\d-]/g, '') + '.' + p[1];
      else s = s.replace(/,/g, '');
    } else s = s.replace(/,/g, '');
    return parseFloat(s);
  }

  function scoreHeader(headerNorm, keywords) {
    if (!headerNorm) return 0;
    var t = 0;
    for (var i = 0; i < keywords.length; i++) {
      var w = normKey(keywords[i]);
      if (headerNorm === w) t += 5;
      else if (headerNorm.indexOf(w) >= 0 && w.length > 1) t += 3;
      else if (w.length > 3 && w.indexOf(headerNorm) >= 0) t += 1;
    }
    return t;
  }

  function findBestIndex(headers, keywords, avoidFn) {
    var bestI = -1;
    var bestS = 0.5;
    for (var c = 0; c < headers.length; c++) {
      if (avoidFn && avoidFn(c)) continue;
      var sc = scoreHeader(headers[c], keywords);
      if (sc > bestS) { bestS = sc; bestI = c; }
    }
    return { index: bestI, score: bestS };
  }

  function mapColumnsCC(headers) {
    var h = headers.map(function (x) { return normKey(x == null ? '' : String(x)); });
    var m = {};
    for (var k in CC_KEYS) {
      if (!CC_KEYS.hasOwnProperty(k)) continue;
      var r = findBestIndex(h, CC_KEYS[k], null);
      m[k] = r.index;
    }
    if (m.fact_esp === m.fact_real && m.fact_esp >= 0) m.fact_real = -1;
    return { map: m, h: h, headers: headers };
  }

  function mapColumnsMB(headers) {
    var h = headers.map(function (x) { return normKey(x == null ? '' : String(x)); });
    var m = {};
    for (var k2 in MB_KEYS) {
      if (!MB_KEYS.hasOwnProperty(k2)) continue;
      var r2 = findBestIndex(h, MB_KEYS[k2], null);
      m[k2] = r2.index;
    }
    return { map: m, h: h, headers: headers };
  }

  function readSheetToMatrix(file, cb) {
    var fr = new FileReader();
    fr.onload = function (e) {
      var data = new Uint8Array(e.target.result);
      var wb = XLSX.read(data, { type: 'array' });
      var sn = wb.SheetNames[0];
      var ws = wb.Sheets[sn];
      var matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      cb(null, { sheetName: sn, matrix: matrix });
    };
    fr.onerror = function () { cb(new Error('No se pudo leer el archivo')); };
    fr.readAsArrayBuffer(file);
  }

  function matrixToObjects(matrix) {
    if (!matrix || !matrix.length) return { headers: [], rows: [] };
    var headers = matrix[0].map(function (c) { return c == null ? '' : String(c).trim(); });
    var rows = [];
    for (var r = 1; r < matrix.length; r++) {
      var line = matrix[r];
      if (!line || !line.length) continue;
      var o = [];
      for (var c2 = 0; c2 < Math.max(line.length, headers.length); c2++) o.push(line[c2] !== undefined ? line[c2] : '');
      if (o.every(function (c3) { return c3 === '' || c3 == null; })) continue;
      rows.push(o);
    }
    return { headers: headers, rows: rows };
  }

  function clientKey(row, map, headers) {
    var ci = map.codigo;
    var cj = map.cliente;
    var code = (ci >= 0 && row[ci] != null) ? String(row[ci]).trim() : '';
    if (code) return 'C:' + normKey(code);
    var name = (cj >= 0 && row[cj] != null) ? String(row[cj]).trim() : '';
    return 'N:' + normKey(name);
  }

  function toBoolLicit(s) {
    if (s == null || s === '') return '—';
    var t = String(s).toLowerCase();
    if (t === '1' || t === 's' || t === 'si' || t === 'sí' || t === 'y' || t === 'true' || t === 'x') return 'Sí';
    if (t === '0' || t === 'n' || t === 'no' || t === 'false' || t === '—' || t === '-') return 'No';
    if (t.indexOf('sí') >= 0 || t.indexOf('si') >= 0 || t.indexOf('licit') >= 0) return 'Sí';
    if (t.indexOf('no') >= 0) return 'No';
    return s;
  }

  function toPublico(s) {
    if (s == null || s === '') return '—';
    var t = normKey(String(s));
    if (t.indexOf('publico') >= 0 || t.indexOf('gobierno') >= 0) return 'Público';
    if (t.indexOf('priv') >= 0) return 'Privado';
    return String(s);
  }

  function aggregateCC(rows, map) {
    var g = {};
    for (var i = 0; i < rows.length; i++) {
      var line = rows[i];
      var k = clientKey(line, map, null);
      if (k === 'N:') continue;
      if (!g[k]) {
        g[k] = { key: k, codigo: map.codigo >= 0 ? String(line[map.codigo] || '') : '', cliente: map.cliente >= 0 ? String(line[map.cliente] || '') : '' };
        if (!g[k].codigo && map.codigo < 0) g[k].codigo = '—';
        g[k].vendedor = map.vendedor >= 0 ? String(line[map.vendedor] || '—') : '—';
        g[k].publico = map.publico >= 0 ? toPublico(line[map.publico]) : '—';
        g[k].licitacion = map.licitacion >= 0 ? toBoolLicit(line[map.licitacion]) : '—';
        g[k].fact_esp = 0;
        g[k].fact_real = 0;
        g[k].contratos = 0;
        g[k]._n = 0;
      } else {
        g[k].vendedor = g[k].vendedor || (map.vendedor >= 0 ? String(line[map.vendedor] || '—') : g[k].vendedor);
      }
      var fesp = map.fact_esp >= 0 ? parseNumber(line[map.fact_esp]) : 0;
      var frea = map.fact_real >= 0 ? parseNumber(line[map.fact_real]) : 0;
      g[k].fact_esp += isNaN(fesp) ? 0 : fesp;
      g[k].fact_real += isNaN(frea) ? 0 : frea;
      var cts = map.contratos >= 0 ? parseNumber(line[map.contratos]) : 1;
      g[k].contratos += (isNaN(cts) ? 0 : cts) || 1;
      g[k]._n += 1;
    }
    var out = [];
    for (var key in g) { if (g.hasOwnProperty(key)) out.push(g[key]); }
    return out;
  }

  function aggregateMB(rows, map) {
    var g2 = {};
    for (var i = 0; i < rows.length; i++) {
      var line2 = rows[i];
      var k2 = clientKey(line2, map, null);
      if (k2 === 'N:') continue;
      if (!g2[k2]) {
        g2[k2] = {
          key: k2,
          codigo: map.codigo >= 0 ? String(line2[map.codigo] || '') : '',
          cliente: map.cliente >= 0 ? String(line2[map.cliente] || '') : '',
          vendedor: map.vendedor >= 0 ? String(line2[map.vendedor] || '—') : '—',
          valor: 0,
          margen: 0,
          unidades: 0
        };
      }
      var v = map.valor >= 0 ? parseNumber(line2[map.valor]) : 0;
      var m2 = map.margen >= 0 ? parseNumber(line2[map.margen]) : 0;
      var u = map.unidades >= 0 ? parseNumber(line2[map.unidades]) : 0;
      g2[k2].valor += isNaN(v) ? 0 : v;
      g2[k2].margen += isNaN(m2) ? 0 : m2;
      g2[k2].unidades += isNaN(u) ? 0 : u;
    }
    var out2 = [];
    for (var k3 in g2) { if (g2.hasOwnProperty(k3)) out2.push(g2[k3]); }
    return out2;
  }

  function mergeData(cc, mb) {
    var m = {};
    for (var i = 0; i < cc.length; i++) { m[cc[i].key] = { cc: cc[i] }; }
    for (var j = 0; j < mb.length; j++) {
      var b = mb[j];
      if (!m[b.key]) m[b.key] = {};
      m[b.key].mb = b;
    }
    for (var k4 in m) {
      if (!m[k4].cc) {
        var b2 = m[k4].mb;
        m[k4].cc = { key: k4, codigo: b2.codigo || '—', cliente: b2.cliente || '—', vendedor: b2.vendedor, publico: '—', licitacion: '—', fact_esp: 0, fact_real: 0, contratos: 0, _n: 0, _soloMB: true };
      }
      if (!m[k4].mb) {
        var a = m[k4].cc;
        m[k4].mb = { key: k4, codigo: a.codigo, cliente: a.cliente, vendedor: a.vendedor, valor: 0, margen: 0, unidades: 0, _soloCC: true };
      }
    }
    var final = [];
    for (var k5 in m) { if (m.hasOwnProperty(k5)) { var rec = m[k5]; var row = buildJoinedRow(rec.cc, rec.mb); if (row) final.push(row); } }
    return final;
  }

  function buildJoinedRow(cc, mb) {
    var fe = +cc.fact_esp;
    var fr = +cc.fact_real;
    var vf = +mb.valor;
    var mbr = +mb.margen;
    var cump;
    if (fe > 0) cump = fr / fe;
    else cump = NaN;
    var mPct;
    if (vf > 0) mPct = mbr / vf;
    else mPct = NaN;
    var vendedor = '—';
    if (cc.vendedor && cc.vendedor !== '—') vendedor = cc.vendedor;
    else if (mb.vendedor && mb.vendedor !== '—') vendedor = mb.vendedor;
    var dif = isNaN(fr) || isNaN(fe) ? NaN : (fr - fe);
    var pot = isNaN(fe) || isNaN(fr) ? NaN : (fe - fr);
    return {
      key: cc.key,
      codigo: cc.codigo && cc.codigo !== '—' ? cc.codigo : (mb.codigo || '—'),
      cliente: (cc.cliente && String(cc.cliente) !== '—' ? cc.cliente : (mb.cliente || '—')) + '',
      vendedor: vendedor,
      publico: cc.publico,
      licitacion: cc.licitacion,
      contratos: +cc.contratos || 0,
      fact_esp: fe,
      fact_real: fr,
      cump: cump,
      valor_fact: vf,
      margen_bruto: mbr,
      margen_pct: mPct,
      dif: dif,
      pot: pot
    };
  }

  // --- mediana; umbrales para cuadrantes ---
  function medianOf(arr) {
    var a = arr.filter(function (x) { return !isNaN(x); }).sort(function (x, y) { return x - y; });
    if (!a.length) return NaN;
    var mid = Math.floor(a.length / 2);
    if (a.length % 2) return a[mid];
    return (a[mid - 1] + a[mid]) / 2;
  }

  function assignQuadrant(row, tC, tM) {
    if (isNaN(row.cump) || isNaN(row.margen_pct) || isNaN(tC) || isNaN(tM)) return { cat: 'Sin dato', code: 'none' };
    var hiC = row.cump >= tC;
    var hiM = row.margen_pct >= tM;
    if (hiM && hiC) return { cat: 'Ideal', code: 'ideal' };
    if (hiM && !hiC) return { cat: 'Oportunidad', code: 'opp' };
    if (!hiM && hiC) return { cat: 'Revisar precios', code: 'prices' };
    return { cat: 'Acción urgente', code: 'urgent' };
  }

  function recomputeQuadrants(rows) {
    var cumps = rows.map(function (r) { return r.cump; });
    var margs = rows.map(function (r) { return r.margen_pct; });
    var tC; var tM;
    if (state.cumpCut === 'median') tC = medianOf(cumps);
    else tC = parseFloat(state.cumpCut);
    if (state.margCut === 'median') tM = medianOf(margs);
    else tM = parseFloat(state.margCut);
    if (isNaN(tC)) tC = 0.5;
    if (isNaN(tM)) tM = 0.1;
    state.cumpThresh = tC; state.margThresh = tM;
    for (var i = 0; i < rows.length; i++) {
      var q = assignQuadrant(rows[i], tC, tM);
      rows[i].categoria = q.cat;
      rows[i].categoriaCode = q.code;
    }
  }

  // --- validación mínima ---
  function validateCC(m) {
    var e = [];
    if (m.codigo < 0 && m.cliente < 0) e.push('Falta columna de cliente o código (Contract Compliance).');
    if (m.fact_esp < 0) e.push('No se detectó facturación esperada o meta (Contract Compliance).');
    if (m.fact_real < 0) e.push('No se detectó facturación real o neta (Contract Compliance).');
    return e;
  }
  function validateMB(m) {
    var e = [];
    if (m.codigo < 0 && m.cliente < 0) e.push('Falta columna de cliente o código (MB).');
    if (m.valor < 0) e.push('No se detectó valor facturado (MB).');
    if (m.margen < 0) e.push('No se detectó margen bruto (MB).');
    return e;
  }

  // --- registro de columnas en consola ---
  function logDetection(label, map, headers) {
    console.log('[' + label + '] columnas en archivo:', headers);
    console.log('[' + label + '] mapeo detectado:', map);
  }

  // --- KPIs y resumen global ---
  function kpiSums(rows) {
    var te = 0, tr = 0, tmb = 0, tf = 0, tco = 0, negM = 0, oppN = 0, urgN = 0, totalCl = rows.length;
    for (var j = 0; j < rows.length; j++) {
      var rr = rows[j];
      tco += rr.contratos || 0;
      if (!isNaN(rr.fact_esp)) te += rr.fact_esp;
      if (!isNaN(rr.fact_real)) tr += rr.fact_real;
      tmb += isNaN(rr.margen_bruto) ? 0 : rr.margen_bruto;
      tf += isNaN(rr.valor_fact) ? 0 : rr.valor_fact;
      if (rr.margen_pct < 0) negM += 1;
      if (rr.categoriaCode === 'opp') oppN += 1;
      if (rr.categoriaCode === 'urgent') urgN += 1;
    }
    return {
      totalCl: totalCl, tco: tco, te: te, tr: tr, tmb: tmb, tf: tf,
      cumpPond: te > 0 ? tr / te : NaN,
      margPond: tf > 0 ? tmb / tf : NaN,
      negM: negM, oppN: oppN, urgN: urgN
    };
  }

  // --- top concentración ---
  function topShare(rows, n, getVal) {
    var s = rows.map(function (r) { return { v: getVal(r), c: r.cliente }; }).filter(function (x) { return !isNaN(x.v); }).sort(function (a, b) { return b.v - a.v; });
    var tot = 0; for (var k = 0; k < s.length; k++) tot += s[k].v;
    if (tot <= 0) return 0;
    var top = 0; for (var t = 0; t < Math.min(n, s.length); t++) top += s[t].v;
    return top / tot;
  }

  // --- riesgos y oportunidades (listas) ---
  function listRisks(rows) {
    var a = rows.slice();
    a.sort(function (a, b) { return (b.fact_real * (1 - (isNaN(b.cump) ? 0 : b.cump))) - (a.fact_real * (1 - (isNaN(a.cump) ? 0 : a.cump))); });
    return a;
  }
  function listOpp(rows) {
    return rows.filter(function (r) { return r.categoriaCode === 'opp'; }).sort(function (a, b) { return (b.pot || 0) - (a.pot || 0); });
  }
  function listLowMarginHighFact(rows) {
    return rows.filter(function (r) { return !isNaN(r.margen_pct) && r.margen_pct < state.margThresh && r.valor_fact > 0; })
      .sort(function (a, b) { return b.valor_fact - a.valor_fact; });
  }

  // --- segmentos ---
  function groupSegment(rows, keyfn) {
    var g = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var k = keyfn(r);
      if (!g[k]) g[k] = { label: k, n: 0, te: 0, tr: 0, tmb: 0, tf: 0 };
      g[k].n += 1;
      g[k].te += isNaN(r.fact_esp) ? 0 : r.fact_esp;
      g[k].tr += isNaN(r.fact_real) ? 0 : r.fact_real;
      g[k].tmb += isNaN(r.margen_bruto) ? 0 : r.margen_bruto;
      g[k].tf += isNaN(r.valor_fact) ? 0 : r.valor_fact;
    }
    return Object.keys(g).map(function (x) { return g[x]; });
  }

  // --- UI: formateo ---
  function fmtPct(x) { if (isNaN(x)) return '—'; return (x * 100).toFixed(1) + '%'; }
  function fmtMoney(x) { if (isNaN(x)) return '—'; return new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 }).format(x); }

  // --- render ---
  function showSections(show) {
    var ids = ['section-executive', 'section-kpis', 'section-insights', 'section-chart', 'section-risks', 'section-ops', 'section-segments', 'section-table'];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (el) el.hidden = !show;
    }
  }

  function renderExec(rows) {
    var k = kpiSums(rows);
    var c5 = topShare(rows, 5, function (r) { return r.fact_real; });
    var c10 = topShare(rows, 10, function (r) { return r.fact_real; });
    var kopp = rows.filter(function (r) { return r.categoriaCode === 'opp'; }).length;
    var highValLowC = rows.filter(function (r) { return r.fact_real > k.tr * 0.02 && (isNaN(r.cump) ? true : r.cump < 0.85); }).length;
    var parts = [
      'El cumplimiento ponderado general (Σ real / Σ esperado) es de ' + (isNaN(k.cumpPond) ? 'N/D' : fmtPct(k.cumpPond) + '.'),
      k.negM + ' ' + (k.negM === 1 ? 'cliente tiene' : 'clientes tienen') + ' margen % negativo en su línea agregada de MB.',
      'El top 5 concentra ' + (isNaN(c5) || c5 === 0 ? '—' : (c5 * 100).toFixed(0) + '% de la facturación real; el top 10 concentra ' + (c10 * 100).toFixed(0) + '%.'),
      (kopp || '0') + ' ' + (kopp === 1 ? 'cliente' : 'clientes') + ' con alto margen (por umbral) y bajo cumplimiento: oportunidad de cerrar brecha hacia la meta de contrato.',
      (highValLowC ? 'Hay ' + highValLowC + ' clientes con carga material y cumplimiento por debajo del 85% — revisar prioridad y riesgo (tabla de riesgos).' : 'Revisar clientes con brecha de cumplimiento (mapa y tabla).')
    ];
    var h = document.getElementById('executive-text');
    h.innerHTML = parts.map(function (p) { return '<div class="exec-line">' + p + '</div>'; }).join('');
  }

  function readThresholdsFromUI() {
    var a = document.getElementById('threshold-cumplimiento');
    var b = document.getElementById('threshold-margen');
    if (a) state.cumpCut = a.value;
    if (b) state.margCut = b.value;
  }

  function renderKPIs(rows) {
    var k = kpiSums(rows);
    var items = [
      { label: 'Clientes', v: String(k.totalCl), c: '' },
      { label: 'Contratos (total)', v: String(Math.round(k.tco)), c: '' },
      { label: 'Fact. esperada', v: fmtMoney(k.te), c: '' },
      { label: 'Fact. real', v: fmtMoney(k.tr), c: '' },
      { label: 'Cumpl. ponderado', v: isNaN(k.cumpPond) ? '—' : fmtPct(k.cumpPond), c: isNaN(k.cumpPond) ? '' : (k.cumpPond < 0.9 ? 'warn' : 'ok') },
      { label: 'Margen bruto total', v: fmtMoney(k.tmb), c: '' },
      { label: 'Margen % ponderado', v: isNaN(k.margPond) ? '—' : fmtPct(k.margPond), c: '' },
      { label: 'Neg. margin (N)', v: String(k.negM), c: k.negM > 0 ? 'danger' : 'ok' },
      { label: 'En oportunidad', v: String(k.oppN), c: k.oppN > 0 ? 'warn' : '' },
      { label: 'En acción urgente', v: String(k.urgN), c: k.urgN > 0 ? 'danger' : 'ok' }
    ];
    var g = document.getElementById('kpi-grid');
    g.innerHTML = items.map(function (it) {
      return '<div class="kpi-card"><div class="kpi-label">' + it.label + '</div><div class="kpi-value' + (it.c ? ' ' + it.c : '') + '">' + it.v + '</div></div>';
    }).join('');
  }

  function renderInsights(rows) {
    var k = kpiSums(rows);
    var ul = document.getElementById('insight-list');
    var t = [];
    t.push('Brecha de contrato: potencial agregado (suma de max(0, fact. esperada − real)) pesa ' + (function () {
      var s = 0; for (var i = 0; i < rows.length; i++) s += isNaN(rows[i].pot) || rows[i].pot < 0 ? 0 : rows[i].pot; return fmtMoney(s);
    }()) + ' en términos de facturación no materializada aún (según cálculo de potencial de línea).');
    if (!isNaN(k.margPond) && k.margPond < 0.12) t.push('Margen % agregado por debajo del 12%: presión de rentabilidad a nivel de cartera.');
    if (k.oppN > 0) t.push(k.oppN + ' clientes combinan buena tasa de margen con bajo avance de cumplimiento: priorizar cierre comercial o condiciones de contrato.');
    ul.innerHTML = t.map(function (x) { return '<li>' + x + '</li>'; }).join('');
  }

  var COL = { ideal: '#22c55e', opp: '#f59e0b', prices: '#60a5fa', urgent: '#ef4444', none: '#6b7280' };

  function maxCumpPct(rows) {
    var m = 100;
    for (var i = 0; i < rows.length; i++) {
      if (!isNaN(rows[i].cump)) m = Math.max(m, rows[i].cump * 100);
    }
    return m * 1.05;
  }

  function drawChart(rows) {
    if (typeof Plotly === 'undefined') { console.warn('Plotly no disponible'); return; }
    var tC = state.cumpThresh; var tM = state.margThresh;
    if (isNaN(tC)) tC = 0.5; if (isNaN(tM)) tM = 0.1;
    var names = { ideal: 'Ideal', opp: 'Oportunidad', prices: 'Revisar precios', urgent: 'Acción urgente', none: 'Sin dato' };
    var data = ['ideal', 'opp', 'prices', 'urgent', 'none'].map(function (code) {
      var subset = rows.filter(function (r) { return (r.categoriaCode || 'none') === code && !isNaN(r.cump) && !isNaN(r.margen_pct); });
      return {
        x: subset.map(function (r) { return r.cump * 100; }),
        y: subset.map(function (r) { return r.margen_pct * 100; }),
        text: subset.map(function (r) { return r.cliente || ''; }),
        name: names[code] || code,
        customdata: subset.map(function (r) {
          return 'Cód: ' + (r.codigo || '—') + '<br>Vend: ' + (r.vendedor || '—') + '<br>Fact. esp.: ' + fmtMoney(r.fact_esp) + '<br>Fact. real: ' + fmtMoney(r.fact_real) +
            '<br>M. bruto: ' + fmtMoney(r.margen_bruto) + '<br>Potencial: ' + fmtMoney(r.pot) + '<br><b>' + (r.categoria || '') + '</b>';
        }),
        mode: 'markers', type: 'scatter',
        hovertemplate: '<b>' + (names[code] || code) + '</b><br>Cliente: %{text}<br>Cumpl. %: %{x:.1f}%<br>Marg. %: %{y:.1f}%<br>%{customdata}<extra></extra>',
        marker: { size: subset.map(function (r) { return Math.sqrt(Math.max(0, r.fact_real) / 1e3) * 1.2 + 6; }), sizemode: 'diameter', color: COL[code] || '#888', line: { width: 0.5, color: '#0f1419' } }
      };
    });
    var xMax = Math.max(120, maxCumpPct(rows));
    var layout = {
      paper_bgcolor: '#121920', plot_bgcolor: '#0f1419', font: { color: '#e8edf4' },
      margin: { t: 40, r: 20, b: 48, l: 56 },
      xaxis: { title: 'Cumplimiento % (real / esperado)', range: [0, xMax] },
      yaxis: { title: 'Margen % (m.b. / valor fact.)' },
      shapes: [
        { type: 'line', x0: tC * 100, x1: tC * 100, y0: 0, y1: 1, yref: 'paper', line: { color: 'rgba(59,130,246,0.45)', width: 1, dash: 'dot' } },
        { type: 'line', y0: tM * 100, y1: tM * 100, x0: 0, x1: 1, xref: 'paper', line: { color: 'rgba(245,158,11,0.4)', width: 1, dash: 'dot' } }
      ],
      legend: { orientation: 'h' },
      hoverlabel: { bgcolor: '#1a222d' }
    };
    Plotly.react('plotly-chart', data, layout, { displayModeBar: true, responsive: true });
  }

  function buildRiskList(rows) {
    var byKey = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.margen_pct < 0) {
        byKey[r.key] = { r: r, w: 1e6 + Math.abs(r.margen_bruto || 0) };
        continue;
      }
      if (!isNaN(r.cump) && r.cump < 0.8 && r.fact_real > 0) {
        var s = r.fact_real * (0.8 - r.cump);
        if (!byKey[r.key] || s > (byKey[r.key].s || 0)) byKey[r.key] = { r: r, s: s, w: s };
      }
    }
    for (var j = 0; j < rows.length; j++) {
      var r2 = rows[j];
      if (isNaN(r2.margen_pct) || r2.margen_pct >= (state.margThresh || 0.1) || r2.valor_fact <= 0) continue;
      if (!byKey[r2.key]) byKey[r2.key] = { r: r2, w: r2.valor_fact * 0.0001 };
    }
    var a = [];
    for (var k6 in byKey) { if (byKey.hasOwnProperty(k6)) a.push(byKey[k6].r); }
    a.sort(function (x, y) {
      return y.fact_real * (isNaN(y.cump) ? 0.5 : (1 - y.cump)) - x.fact_real * (isNaN(x.cump) ? 0.5 : (1 - x.cump));
    });
    return a;
  }

  function fillSimpleTable(tbody, thead, headers, list, rowFn) {
    thead.innerHTML = '<tr>' + headers.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr>';
    tbody.innerHTML = list.map(function (r) { return rowFn(r); }).join('');
  }

  function renderRisks(rows) {
    var th = document.querySelector('#table-risks thead');
    var tb = document.querySelector('#table-risks tbody');
    var L = buildRiskList(rows).slice(0, 40);
    th.innerHTML = '<tr><th>Cliente</th><th>Código</th><th>Fact. real</th><th>Cump %</th><th>Margen %</th><th>Valor fact. MB</th><th>Cat.</th></tr>';
    tb.innerHTML = L.map(function (r) { return '<tr' + (r.margen_pct < 0 ? ' class="row-danger"' : '') + '><td data-text>' + (r.cliente || '') + '</td><td data-text>' + (r.codigo || '') + '</td><td>' + fmtMoney(r.fact_real) + '</td><td>' + fmtPct(r.cump) + '</td><td>' + fmtPct(r.margen_pct) + '</td><td>' + fmtMoney(r.valor_fact) + '</td><td class="cat-' + (r.categoriaCode || 'unknown') + '">' + (r.categoria || '') + '</td></tr>'; }).join('');
  }

  function renderOpps(rows) {
    var th2 = document.querySelector('#table-ops thead');
    var tb2 = document.querySelector('#table-ops tbody');
    var L2 = listOpp(rows);
    th2.innerHTML = '<tr><th>Cliente</th><th>Código</th><th>Potencial (esp - real)</th><th>Fact. esperada</th><th>Fact. real</th><th>Cump %</th><th>Margen %</th></tr>';
    tb2.innerHTML = L2.map(function (r) { return '<tr class="row-warn"><td data-text>' + (r.cliente || '') + '</td><td data-text>' + (r.codigo || '') + '</td><td>' + fmtMoney(r.pot) + '</td><td>' + fmtMoney(r.fact_esp) + '</td><td>' + fmtMoney(r.fact_real) + '</td><td>' + fmtPct(r.cump) + '</td><td>' + fmtPct(r.margen_pct) + '</td></tr>'; }).join('');
  }

  function renderSegments(rows) {
    function one(tableId, list) {
      var th3 = document.querySelector('#' + tableId + ' thead');
      var tb3 = document.querySelector('#' + tableId + ' tbody');
      th3.innerHTML = '<tr><th>Segmento</th><th>Clientes</th><th>Fact. esp.</th><th>Fact. real</th><th>Cumpl. % *</th><th>M. bruto</th><th>Margen % *</th></tr>';
      tb3.innerHTML = list.map(function (s) {
        var cP = s.te > 0 ? s.tr / s.te : NaN, mP = s.tf > 0 ? s.tmb / s.tf : NaN;
        return '<tr><td data-text>' + s.label + '</td><td>' + s.n + '</td><td>' + fmtMoney(s.te) + '</td><td>' + fmtMoney(s.tr) + '</td><td>' + fmtPct(cP) + '</td><td>' + fmtMoney(s.tmb) + '</td><td>' + fmtPct(mP) + '</td></tr>';
      }).join('');
    }
    one('table-seg-vendedor', groupSegment(rows, function (r) { return r.vendedor || '—'; }));
    one('table-seg-publico', groupSegment(rows, function (r) { return r.publico || '—'; }));
    one('table-seg-licit', groupSegment(rows, function (r) { return r.licitacion || '—'; }));
  }

  function distinctVals(rows, fn) {
    var s = {}; for (var i = 0; i < rows.length; i++) { s[fn(rows[i])] = 1; }
    return Object.keys(s).sort();
  }

  function populateFilters(rows) {
    var opts = function (elId, allLabel, values) { var s = document.getElementById(elId); s.innerHTML = '<option value="">' + allLabel + '</option>'; values.forEach(function (v) { var o = document.createElement('option'); o.value = v; o.textContent = v; s.appendChild(o); }); };
    opts('filter-vendedor', 'Todos', distinctVals(rows, function (r) { return r.vendedor || '—'; }));
    opts('filter-publico', 'Todos', distinctVals(rows, function (r) { return r.publico || '—'; }));
    opts('filter-licit', 'Todos', distinctVals(rows, function (r) { return r.licitacion || '—'; }));
    opts('filter-cat', 'Todas', ['Ideal', 'Oportunidad', 'Revisar precios', 'Acción urgente', 'Sin dato'].filter(function (c) { return rows.some(function (r) { return (r.categoria || '') === c; }); })
    );
  }

  function filterMain(rows) {
    var q = (document.getElementById('table-search') && document.getElementById('table-search').value || '').toLowerCase();
    var fv = document.getElementById('filter-vendedor').value;
    var fp = document.getElementById('filter-publico').value;
    var fl = document.getElementById('filter-licit').value;
    var fc = document.getElementById('filter-cat').value;
    var flt = document.getElementById('filter-bajo-cump').value;
    var tC2 = state.cumpThres;
    return rows.filter(function (r) {
      if (q && (String(r.cliente) + ' ' + String(r.codigo)).toLowerCase().indexOf(q) < 0) return false;
      if (fv && (r.vendedor || '—') !== fv) return false;
      if (fp && (r.publico || '—') !== fp) return false;
      if (fl && (r.licitacion || '—') !== fl) return false;
      if (fc && (r.categoria || '') !== fc) return false;
      if (flt === 'lowCump' && !(r.cump < tC2)) return false;
      if (flt === 'lowMargin' && !(r.margen_pct < (state.margThresh || 0))) return false;
      return true;
    });
  }

  var MAIN_COLS = [
    { k: 'codigo', t: 'Código' }, { k: 'cliente', t: 'Cliente' }, { k: 'vendedor', t: 'Vendedor' },
    { k: 'publico', t: 'Publ./priv.' }, { k: 'licitacion', t: 'Licit.' }, { k: 'contratos', t: 'N° ctes.' },
    { k: 'fact_esp', t: 'Fact. esp.' }, { k: 'fact_real', t: 'Fact. real' },
    { k: 'cump', t: 'Cump. %' }, { k: 'valor_fact', t: 'Val. fact. MB' }, { k: 'margen_bruto', t: 'M. bruto' },
    { k: 'margen_pct', t: 'Marg. %' }, { k: 'pot', t: 'Potencial' }, { k: 'categoria', t: 'Categoría' }
  ];

  function valCell(k, r) {
    if (k === 'cump' || k === 'margen_pct') return fmtPct(r[k]);
    if (k === 'contratos') return (r[k] == null || isNaN(r[k])) ? '—' : String(Math.round(r[k]));
    if (k === 'fact_esp' || k === 'fact_real' || k === 'valor_fact' || k === 'margen_bruto' || k === 'pot') return fmtMoney(r[k]);
    if (k === 'categoria') return '<span class="cat-' + (r.categoriaCode || 'unknown') + '">' + (r.categoria || '') + '</span>';
    return (r[k] == null || r[k] === undefined) ? '—' : String(r[k]);
  }

  function isTextCol(c) {
    return c === 'cliente' || c === 'codigo' || c === 'vendedor' || c === 'publico' || c === 'licitacion' || c === 'categoria';
  }

  function renderMainTable() {
    var R = state.rows; if (!R) return;
    var F = filterMain(R);
    var c = state.sortMain.col, d = state.sortMain.dir;
    F.sort(function (a, b) {
      var x = a[c], y = b[c];
      if (isTextCol(c)) {
        x = (x + '').toLowerCase(); y = (y + '').toLowerCase();
        if (x < y) return -1 * d; if (x > y) return 1 * d; return 0;
      }
      if (isNaN(x) && isNaN(y)) return 0; if (isNaN(x)) return 1; if (isNaN(y)) return -1; return d * (x - y);
    });
    var th = document.querySelector('#table-main thead');
    var tb = document.querySelector('#table-main tbody');
    th.innerHTML = '<tr>' + MAIN_COLS.map(function (col) { return '<th class="sortable" data-k="' + col.k + '">' + col.t + (state.sortMain.col === col.k ? (d > 0 ? ' \u25B2' : ' \u25BC') : '') + '</th>'; }).join('') + '</tr>';
    tb.innerHTML = F.map(function (r) {
      var cl = [];
      if (r.margen_pct < 0) cl.push('row-danger');
      if (!isNaN(r.cump) && (state.cumpThresh != null) && r.cump < state.cumpThresh && r.fact_esp > 0) cl.push('row-warn');
      return '<tr' + (cl.length ? ' class="' + cl.join(' ') + '"' : '') + '>' + MAIN_COLS.map(function (h) { return '<td' + (isTextCol(h.k) ? ' data-text' : '') + '>' + valCell(h.k, r) + '</td>'; }).join('') + '</tr>';
    }).join('');
  }

  function refreshView() {
    if (!state.rows) return;
    readThresholdsFromUI();
    recomputeQuadrants(state.rows);
    showSections(true);
    renderExec(state.rows);
    renderKPIs(state.rows);
    renderInsights(state.rows);
    drawChart(state.rows);
    renderRisks(state.rows);
    renderOpps(state.rows);
    renderSegments(state.rows);
    renderMainTable();
  }

  function runFullRender() {
    refreshView();
    if (state.rows) populateFilters(state.rows);
  }

  function exportCSV() {
    if (!state.rows) return;
    var F = filterMain(state.rows);
    var h = MAIN_COLS.map(function (c) { return c.t; }).join(';') + '\n';
    var b = h + F.map(function (r) { return MAIN_COLS.map(function (c) { var v = c.k === 'cump' || c.k === 'margen_pct' ? (isNaN(r[c.k]) ? '' : r[c.k]) : r[c.k]; return (v == null ? '' : String(v).replace(/[;\n\r]/g, ' ')); }).join(';'); }).join('\n');
    var bb = new Blob([ '\ufeff' + b ], { type: 'text/csv;charset=utf-8;' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(bb); a.download = 'clientes_cruce.csv';
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
  }

  function setFileStatus(id, name, nRows, nCols, extra) {
    var el = document.getElementById(id);
    if (!el) return;
    if (!name) { el.setAttribute('data-empty', '1'); el.textContent = 'Archivo no cargado'; return; }
    el.removeAttribute('data-empty');
    el.textContent = 'Archivo: ' + name + '\nFilas: ' + nRows + ' · Columnas: ' + nCols + (extra || '');
  }

  function setUploadError(msg) {
    var a = document.getElementById('upload-alerts');
    if (a) { a.hidden = false; a.textContent = msg; }
  }

  function runPipeline() {
    if (!state.rawCC || !state.rawMB) return;
    var cco = matrixToObjects(state.rawCC.matrix);
    var mbo = matrixToObjects(state.rawMB.matrix);
    var ccm = mapColumnsCC(cco.headers);
    var mbm = mapColumnsMB(mbo.headers);
    logDetection('Contract Compliance', ccm.map, cco.headers);
    logDetection('MB Q1', mbm.map, mbo.headers);
    var err = validateCC(ccm.map).concat(validateMB(mbm.map));
    if (err.length) {
      setUploadError(err.join(' '));
      showSections(false);
      return;
    }
    var a = document.getElementById('upload-alerts');
    if (a) a.hidden = true;
    var hint = document.getElementById('hint-both');
    if (hint) hint.hidden = false;
    var ccAgg = aggregateCC(cco.rows, ccm.map);
    var mbAgg = aggregateMB(mbo.rows, mbm.map);
    state.rows = mergeData(ccAgg, mbAgg);
    state.ccMeta = cco;
    state.mbMeta = mbo;
    if (!state.rows.length) {
      setUploadError('No se obtuvieron clientes al cruzar. Revisa códigos/nombres o duplicados vacíos.');
      return;
    }
    runFullRender();
  }

  function handleFileInput(which) {
    return function (e) {
      var f = e.target.files && e.target.files[0];
      if (!f) return;
      if (typeof XLSX === 'undefined') { setUploadError('SheetJS no se cargó. Comprueba la conexión a internet.'); return; }
      readSheetToMatrix(f, function (err, res) {
        if (err) { setUploadError(err.message); return; }
        if (which === 'cc') {
          state.rawCC = { name: f.name, matrix: res.matrix, sheetName: res.sheetName };
          var t = matrixToObjects(res.matrix);
          setFileStatus('status-cc', f.name, t.rows.length, t.headers.length, ' · Hoja: ' + (res.sheetName || ''));
        } else {
          state.rawMB = { name: f.name, matrix: res.matrix, sheetName: res.sheetName };
          var t2 = matrixToObjects(res.matrix);
          setFileStatus('status-mb', f.name, t2.rows.length, t2.headers.length, ' · Hoja: ' + (res.sheetName || ''));
        }
        if (state.rawCC && state.rawMB) runPipeline();
      });
    };
  }

  function wireEvents() {
    var fc = document.getElementById('file-cc');
    var fm = document.getElementById('file-mb');
    if (fc) fc.addEventListener('change', handleFileInput('cc'));
    if (fm) fm.addEventListener('change', handleFileInput('mb'));
    var btnR = document.getElementById('btn-refresh-quadrants');
    if (btnR) btnR.addEventListener('click', function () { if (state.rows) refreshView(); });
    var tc = document.getElementById('threshold-cumplimiento');
    var tm = document.getElementById('threshold-margen');
    if (tc) tc.addEventListener('change', function () { if (state.rows) refreshView(); });
    if (tm) tm.addEventListener('change', function () { if (state.rows) refreshView(); });
    var exp = document.getElementById('btn-export-csv');
    if (exp) exp.addEventListener('click', exportCSV);
    var fSearch = document.getElementById('table-search');
    if (fSearch) fSearch.addEventListener('input', renderMainTable);
    ['filter-vendedor', 'filter-publico', 'filter-licit', 'filter-cat', 'filter-bajo-cump'].forEach(function (id) {
      var s = document.getElementById(id);
      if (s) s.addEventListener('change', renderMainTable);
    });
    var mainT = document.getElementById('table-main');
    if (mainT) {
      mainT.addEventListener('click', function (ev) {
        var t = ev.target;
        if (!t || !t.getAttribute) return;
        if (t.classList && t.classList.contains('sortable')) {
          var k = t.getAttribute('data-k');
          if (k) {
            if (state.sortMain.col === k) state.sortMain.dir = -state.sortMain.dir; else { state.sortMain.col = k; state.sortMain.dir = 1; }
            renderMainTable();
          }
        }
      });
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wireEvents);
  else wireEvents();
})();