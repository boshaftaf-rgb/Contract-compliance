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
    vendedor: ['vendedor', 'vend', 'asesor', 'comercial', 'seller', 'ejecutiv', 'representant'],
    publico: ['publico', 'privado', 'sector', 'público', 'tipo', 'gob', 'gobierno', 'público/priv', 'público/privado'],
    licitacion: ['licit', 'concurso', 'lic tacion', 'remate'],
    fact_esp: ['esperad', 'meta', 'objetiv', 'target', 'budget', 'proyectad', 'planead', 'presup', 'monto esper'],
    fact_real: ['real', 'neta', 'neto', 'facturado', 'cobr', 'vend', 'monto real', 'fact. real', 'fact neta', 'reali'],
    contratos: ['contrat', 'n contrat', 'num contrat', 'contracts', 'cte', 'no contrat', 'cant contrat', 'n° contrat']
  };

  var MB_KEYS = {
    codigo: ['codigo', 'código', 'id cliente', 'idcliente', 'cod cliente', 'code', 'nro codigo', 'cód', 'cód.'],
    cliente: ['cliente', 'razon', 'razón', 'customer', 'nombre', 'deudor', 'empresa', 'nombre cliente'],
    vendedor: ['vendedor', 'vend', 'asesor', 'comercial', 'seller', 'ejecutiv', 'representant'],
    /** fact. neto/venta neta: planillas frecuentes; si no, no matchea "valor facturado" en el encabezado. */
    valor: ['valor factur', 'fact neto', 'fact. neto', 'ventas', 'facturación', 'facturacion', 'monto', 'importe', 'ingreso', 'revenue', 'billing', 'vta', 'vta neta', 'venta neta', 'importe neta'],
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

  /** Filtros solo sección 5. vend/pub/lic: null = mostrar todos; Set = OR dentro del tipo */
  var mapState = { vend: null, pub: null, lic: null, search: '', minC: 0, minCump: 0, colorBy: 'quadrant', sizeBy: 'contratos', xLog: false };

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

  /**
   * Convierte cumpl. / margen a cifra 0-100+ para el eje (ratio 0.56 -> 56; "56%" -> 56).
   * null/undefined/NaN -> null
   */
  function toPercentValue(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number' && (isNaN(value) || value !== value)) return null;
    if (typeof value === 'string') {
      var t = value.trim();
      if (t === '' || t === '—' || t === '-') return null;
      t = t.replace(/%/g, ' ').replace(/,/g, '.').replace(/\s+/g, '');
      if (t === '') return null;
    }
    var n = (typeof value === 'string') ? parseNumber(value) : value;
    if (typeof n !== 'number' || isNaN(n) || n !== n) return null;
    if (n >= 0 && n <= 1) return n * 100;
    return n;
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
      // Campos opcionales: evita enganchar columnas incorrectas por parecido débil.
      if ((k === 'vendedor' || k === 'publico' || k === 'licitacion') && r.score < 3) m[k] = -1;
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
      if (k2 === 'vendedor' && r2.score < 3) m[k2] = -1;
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

  /** Código de cliente (SAP, etc.): puro numérico → misma clave aunque venga en col. "Cliente" en un archivo y "Código" en otro. */
  function looksLikeClientId(s) {
    if (s == null) return false;
    var t = String(s).trim();
    if (t.length < 5) return false;
    return /^\d+$/.test(t);
  }

  function clientKey(row, map, headers) {
    var ci = map.codigo;
    var cj = map.cliente;
    var code = (ci >= 0 && row[ci] != null) ? String(row[ci]).trim() : '';
    if (code) return 'C:' + normKey(code);
    var name = (cj >= 0 && row[cj] != null) ? String(row[cj]).trim() : '';
    if (looksLikeClientId(name)) return 'C:' + normKey(name);
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

  function cleanVendorName(s) {
    if (s == null) return '—';
    var t = String(s).replace(/\s+/g, ' ').trim();
    if (!t || t === '-' || t === '—') return '—';
    return t.toUpperCase();
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
      // "N° contrat" casi siempre es identificador (miles de millones al sumar); "cant" es cantidad pequeña
      var cadd = 1;
      if (map.contratos >= 0) {
        var cn = parseNumber(line[map.contratos]);
        if (isNaN(cn) || cn <= 0) cadd = 1;
        else if (cn > 1e4) cadd = 1;
        else cadd = Math.max(1, Math.round(cn));
      }
      g[k].contratos += cadd;
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
    normalizeVendorNames(final);
    return final;
  }

  function normalizeVendorNames(rows) {
    if (!rows || !rows.length) return;
    var byNorm = {};
    for (var i = 0; i < rows.length; i++) {
      var v = cleanVendorName(rows[i].vendedor);
      var nk = normKey(v);
      if (!nk) nk = '—';
      if (!byNorm[nk]) byNorm[nk] = { total: 0, forms: {} };
      byNorm[nk].total += 1;
      byNorm[nk].forms[v] = (byNorm[nk].forms[v] || 0) + 1;
    }
    var canon = {};
    Object.keys(byNorm).forEach(function (k) {
      if (k === '—') { canon[k] = '—'; return; }
      var forms = byNorm[k].forms;
      var best = null, bestN = -1;
      Object.keys(forms).forEach(function (f) {
        var n = forms[f];
        if (n > bestN) { bestN = n; best = f; return; }
        if (n === bestN && best && f.length > best.length) best = f;
      });
      canon[k] = best || '—';
    });
    for (var j = 0; j < rows.length; j++) {
      var v2 = cleanVendorName(rows[j].vendedor);
      var nk2 = normKey(v2);
      if (!nk2) nk2 = '—';
      rows[j].vendedor = canon[nk2] || '—';
    }
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
    if (cc.vendedor && cc.vendedor !== '—') vendedor = cleanVendorName(cc.vendedor);
    else if (mb.vendedor && mb.vendedor !== '—') vendedor = cleanVendorName(mb.vendedor);
    var dif = isNaN(fr) || isNaN(fe) ? NaN : (fr - fe);
    var pot = isNaN(fe) || isNaN(fr) ? NaN : (fe - fr);
    var ca = (cc.cliente && String(cc.cliente) !== '—') ? String(cc.cliente).trim() : '';
    var cmb = (mb.cliente && String(mb.cliente) !== '—') ? String(mb.cliente).trim() : '';
    var clienteOut = cmb;
    if (ca) {
      if (!cmb) clienteOut = ca;
      else if (looksLikeClientId(ca) && !looksLikeClientId(cmb) && cmb.length > 2) clienteOut = cmb;
      else clienteOut = ca;
    } else if (!cmb) clienteOut = '—';
    return {
      key: cc.key,
      codigo: (cc.codigo && String(cc.codigo).trim() && cc.codigo !== '—') ? String(cc.codigo).trim() : (mb.codigo && String(mb.codigo).trim() ? String(mb.codigo).trim() : '—'),
      cliente: String(clienteOut),
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
  function fmtInt(x) { if (isNaN(x)) return '—'; return new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 }).format(Math.round(x)); }

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
    state.cumpCut = 'median';
    state.margCut = 'median';
  }

  function renderKPIs(rows) {
    var k = kpiSums(rows);
    var items = [
      { label: 'Clientes', v: fmtInt(k.totalCl), c: '' },
      { label: 'Contratos (total)', v: fmtInt(k.tco), c: '' },
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
    var sumPot = 0, nBrecha = 0;
    for (var i0 = 0; i0 < rows.length; i0++) {
      var p0 = rows[i0].pot;
      if (!isNaN(p0) && p0 > 0) { sumPot += p0; nBrecha += 1; }
    }
    var te = k.te;
    if (sumPot > 0.5) {
      var shareMeta = (te > 0) ? (sumPot / te) * 100 : NaN;
      var a = 'Aún faltan unos ' + fmtMoney(sumPot) + ' para alinear lo facturado con la meta (diferencia sumada solo en clientes donde el real aún queda bajo el esperado).';
      if (!isNaN(shareMeta) && te > 0) a += ' Eso ronda el ' + shareMeta.toFixed(0) + '% de la facturación meta total, así que pesa con fuerza en el cierre hacia el objetivo.';
      a += ' Aplica a ' + nBrecha + (nBrecha === 1 ? ' cliente' : ' clientes') + '. Revisa oportunidades (tabla) y el mapa para bajar de ahí, no en promedio.';
      t.push(a);
    } else {
      t.push('A nivel de cartera, la facturación real y la meta casi se compensan: no se ve un hueco agregado. Igual mira al detalle: detrás de un promedio “bueno” a veces se esconde unos pocos clientes muy atrasados.');
    }
    if (!isNaN(k.cumpPond)) {
      if (k.cumpPond < 0.9) t.push('Cumplimiento comercial (ponderado): ' + fmtPct(k.cumpPond) + ', por debajo del 90% respecto a la meta agregada: conviene asegurar cierre o revisar riesgo en las cuentas atrasadas.');
      else t.push('Cumplimiento comercial (ponderado) en rango aceptable: ' + fmtPct(k.cumpPond) + ' frente a la meta, salvo excepciones que aún muestre el mapa o la tabla de riesgos.');
    }
    var conPot = rows.filter(function (r) { return !isNaN(r.pot) && r.pot > 0; });
    conPot.sort(function (a, b) { return b.pot - a.pot; });
    if (conPot.length >= 1) {
      var nShow3 = Math.min(3, conPot.length);
      var parts2 = [];
      for (var h2 = 0; h2 < nShow3; h2++) {
        var r = conPot[h2];
        var nm = String(r.cliente || r.codigo || '—');
        if (nm.length > 40) nm = nm.substring(0, 37) + '…';
        parts2.push(nm + ' (' + fmtMoney(r.pot) + ')');
      }
      var tail = conPot.length > 3 ? ' Además, ' + (conPot.length - 3) + ' con brecha en cifra menor, pero a seguir al radar.' : '';
      t.push('Cuentas con mayor atraso hacia la meta (por monto a recuperar): ' + parts2.join(' · ') + '.' + tail);
    }
    if (!isNaN(k.margPond) && k.margPond < 0.12) t.push('Margen % de la cartera por debajo del 12%: presión de rentabilidad en el agregado; revisa precio y mix con margen aceptable.');
    if (k.oppN > 0) t.push(k.oppN + ' cliente' + (k.oppN === 1 ? '' : 's') + ' con buen margen pero bajo avance hacia la meta: prioridad comercial o de condición de contrato (cuadrante Oportunidad en el mapa).');
    ul.innerHTML = t.map(function (x) { return '<li>' + x + '</li>'; }).join('');
  }

  var COL = { ideal: '#22c55e', opp: '#f59e0b', prices: '#3b82f6', urgent: '#ef4444', none: '#6b7280' };
  var VENDOR_COLS = ['#0ea5e9', '#8b5cf6', '#f97316', '#10b981', '#ec4899', '#6366f1', '#14b8a6', '#eab308', '#64748b', '#a855f7', '#d946ef', '#0d9488'];

  function strHash(s) {
    var h = 0; var u = String(s == null ? '' : s);
    for (var i = 0; i < u.length; i++) h = ((h << 5) - h) + u.charCodeAt(i) | 0;
    return Math.abs(h);
  }

  function getBubbleSize(r, sizeBy) {
    var v = 12;
    if (sizeBy === 'uniform') v = 20;
    else if (sizeBy === 'valor_fact') v = 6 + Math.min(32, Math.sqrt(Math.max(0, +r.valor_fact || 0) / 3e3) * 2.2);
    else if (sizeBy === 'contratos') v = 6 + Math.min(32, Math.sqrt(Math.max(0, +r.contratos || 0)) * 2.5);
    else if (sizeBy === 'fact_esp') v = 6 + Math.min(32, Math.sqrt(Math.max(0, +r.fact_esp || 0) / 3e3) * 2.2);
    else if (sizeBy === 'margen_bruto') v = 6 + Math.min(32, Math.sqrt(Math.max(0, Math.abs(+r.margen_bruto || 0)) / 1e3) * 2.2);
    if (isNaN(v) || v < 8) v = 8;
    if (v > 44) v = 44;
    return v;
  }

  function colorForMapRow(r, colorBy) {
    if (colorBy === 'quadrant') return COL[r.categoriaCode] || COL.none;
    if (colorBy === 'vendedor') return VENDOR_COLS[strHash(r.vendedor || '—') % VENDOR_COLS.length];
    if (colorBy === 'publico') {
      if (r.publico === 'Público') return '#0284c7';
      if (r.publico === 'Privado') return '#64748b';
      return '#94a3b8';
    }
    if (colorBy === 'licitacion') {
      if (r.licitacion === 'Sí') return '#16a34a';
      if (r.licitacion === 'No') return '#ef4444';
      return '#94a3b8';
    }
    return COL[r.categoriaCode] || '#64748b';
  }

  function passMapSetFilters(r) {
    if (mapState.vend) {
      var v = (r.vendedor || '—');
      if (!mapState.vend.has(v)) return { ok: false, reason: 'vendedor' };
    }
    if (mapState.pub) {
      var p = (r.publico || '—');
      if (!mapState.pub.has(p)) return { ok: false, reason: 'público/privado' };
    }
    if (mapState.lic) {
      var l = (r.licitacion || '—');
      if (!mapState.lic.has(l)) return { ok: false, reason: 'licitación' };
    }
    return { ok: true };
  }

  /**
   * Medianas de cumpl. y margen (0–1) con datos reales, para imputar solo eje faltante (evita L en ejes a 0).
   */
  function computeMedianAxisImputes(allRows) {
    var cR = [], mR = [];
    for (var u = 0; u < allRows.length; u++) {
      var w = allRows[u], fe = +w.fact_esp, fr = +w.fact_real, vf = +w.valor_fact, mbr = +w.margen_bruto;
      if (fe > 0 && !isNaN(fr) && isFinite(fr) && isFinite(fe)) cR.push(fr / fe);
      if (vf > 0 && isFinite(mbr) && isFinite(vf)) mR.push(mbr / vf);
    }
    var mc = medianOf(cR);
    var mm = medianOf(mR);
    if (isNaN(mc)) mc = 0.5;
    if (isNaN(mm)) mm = 0.1;
    return { imputeXPct: mc * 100, imputeYPct: mm * 100 };
  }

  /** Cumplimiento en % para eje: si no hay meta CC, se imputa mediana (no 0) para despegar del eje Y. */
  function chartAxisCumplPct(r, imputeXPct) {
    var v = toPercentValue(r.cump);
    if (v != null && isFinite(v)) return v;
    var fe = +r.fact_esp, fr = +r.fact_real;
    if (fe > 0 && isFinite(fe) && isFinite(fr)) return (fr / fe) * 100;
    if ((fe === 0 || isNaN(fe)) && fr > 0) return 100;
    return imputeXPct;
  }
  function chartAxisMargenPct(r, imputeYPct) {
    var v2 = toPercentValue(r.margen_pct);
    if (v2 != null && isFinite(v2)) return v2;
    var vf = +r.valor_fact, mb = +r.margen_bruto;
    if (vf > 0 && isFinite(vf) && isFinite(mb)) return (mb / vf) * 100;
    return imputeYPct;
  }

  function getChartData(allRows) {
    var nCross = allRows && allRows.length ? allRows.length : 0;
    var discarded = [];
    var out = [];
    var afterVend = 0, withXY = 0;
    var q = (document.getElementById('map-search') && document.getElementById('map-search').value || '').trim().toLowerCase();
    var minCRaw = document.getElementById('map-min-contratos') ? parseInt(document.getElementById('map-min-contratos').value, 10) : 0;
    var minC = (isNaN(minCRaw) ? 0 : minCRaw);
    minC = Math.max(0, minC);
    var minCumpIn = document.getElementById('map-min-cump') ? document.getElementById('map-min-cump').value : '0';
    var minCumpP = Math.max(0, Math.min(200, parseFloat(minCumpIn) || 0));
    var colorBy = (document.getElementById('map-color-by') && document.getElementById('map-color-by').value) || 'quadrant';
    var sizeBy = (document.getElementById('map-size-by') && document.getElementById('map-size-by').value) || 'contratos';
    var xLog = (document.getElementById('map-x-scale') && document.getElementById('map-x-scale').value) === 'log';
    var imputes = computeMedianAxisImputes(allRows);
    var tCu0 = state.cumpThresh, tMu0 = state.margThresh;
    if (isNaN(tCu0)) tCu0 = 0.5; if (isNaN(tMu0)) tMu0 = 0.1;
    for (var j = 0; j < nCross; j++) {
      var r2 = allRows[j];
      var prf = passMapSetFilters(r2);
      if (!prf.ok) {
        discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'filtro: ' + prf.reason });
        continue;
      }
      afterVend++;
      if (q) {
        var blob = (String(r2.cliente || '') + ' ' + String(r2.codigo || '')).toLowerCase();
        if (blob.indexOf(q) < 0) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'búsqueda' }); continue; }
      }
      if (minC > 0 && (+r2.contratos || 0) < minC) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'N° contratos < mín.' }); continue; }
      var xP = chartAxisCumplPct(r2, imputes.imputeXPct);
      var yP = chartAxisMargenPct(r2, imputes.imputeYPct);
      if (isNaN(xP) || isNaN(yP)) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'eje x/y no num.' }); continue; }
      withXY++;
      if (minCumpP > 0 && xP < minCumpP) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'cumpl. < mín. %' }); continue; }
      if (xLog && xP <= 0) xP = 0.01;
      var sz = getBubbleSize(r2, sizeBy);
      if (isNaN(sz) || sz < 8) sz = 8;
      var qMap = assignQuadrant(
        { cump: xP / 100, margen_pct: yP / 100 },
        tCu0, tMu0
      );
      var plotQ = (qMap && qMap.code) || 'none';
      var cMark;
      if (colorBy === 'quadrant') cMark = COL[plotQ] || COL.none;
      else cMark = colorForMapRow(r2, colorBy);
      out.push({
        row: r2, x: xP, y: yP, size: sz,
        plotQuadrantCode: plotQ, color: cMark
      });
    }
    return { nCross: nCross, afterVend: afterVend, withXY: withXY, out: out, discarded: discarded, colorBy: colorBy, sizeBy: sizeBy, xLog: xLog, minCumpP: minCumpP, minC: minC, q: q };
  }

  function debugChartData(g, allRows) {
    console.log('[Mapa 5] Registros cruzados (total filas):', g.nCross);
    console.log('[Mapa 5] Tras chips vendedor / público / licitación:', g.afterVend);
    console.log('[Mapa 5] Con cumpl. y margen % válidos (tras búsq./N° ctes., antes corte cumpl. mín.):', g.withXY);
    console.log('[Mapa 5] Puntos finales a Plotly (tras cumpl. mín. y eje log):', g.out.length);
    var sample = g.out.slice(0, 5).map(function (o) { return { cliente: o.row.cliente, x: o.x, y: o.y, cump: o.row.cump, marg: o.row.margen_pct }; });
    console.log('[Mapa 5] Primeros 5 puntos enviados a Plotly:', sample);
    var d = g.discarded;
    if (d && d.length) {
      var by = {};
      for (var k = 0; k < d.length; k++) { var rs = d[k].reason; by[rs] = (by[rs] || 0) + 1; }
      console.log('[Mapa 5] Registros descartados (conteo por motivo):', by);
    } else console.log('[Mapa 5] Registros descartados: ninguno');
  }

  function buildMapPointHover(o) {
    var r = o.row;
    return [
      (r.cliente || '—'), 'Cód: ' + (r.codigo || '—'), 'Vend: ' + (r.vendedor || '—'),
      'Publ./priv.: ' + (r.publico || '—'), 'Licit.: ' + (r.licitacion || '—'),
      'Cumpl. %: ' + o.x.toFixed(1), 'Margen %: ' + o.y.toFixed(1),
      'Fact. esp.: ' + fmtMoney(r.fact_esp), 'Fact. real: ' + fmtMoney(r.fact_real),
      'Val. fact. MB: ' + fmtMoney(r.valor_fact), 'M. bruto: ' + fmtMoney(r.margen_bruto),
      'Potencial: ' + fmtMoney(r.pot), 'Categoría: ' + (r.categoria || '—')
    ].join('<br>');
  }

  function buildPlotlyFromGetChartData(g) {
    var tC = state.cumpThresh, tM = state.margThresh;
    if (isNaN(tC)) tC = 0.5; if (isNaN(tM)) tM = 0.1;
    var tCx = tC * 100, tMy = tM * 100;
    var qNames = { ideal: 'Ideal', opp: 'Oportunidad', prices: 'Revisar precios', urgent: 'Acción urgente', none: 'Sin dato' };
    var out = g.out, colorBy = g.colorBy, xLog = g.xLog;
    var codes = ['ideal', 'opp', 'prices', 'urgent', 'none'];
    var data = [];
    if (colorBy === 'quadrant') {
      for (var c2 = 0; c2 < codes.length; c2++) {
        var code2 = codes[c2];
        var sel2 = out.filter(function (o) { return (o.plotQuadrantCode || o.row.categoriaCode || 'none') === code2; });
        if (!sel2.length) continue;
        var lab = qNames[code2] || code2;
        var prefix = '<b>' + lab + '</b><br>';
        data.push({
          x: sel2.map(function (o) { return o.x; }),
          y: sel2.map(function (o) { return o.y; }),
          text: sel2.map(function (o) { return o.row.cliente || ''; }),
          name: lab,
          mode: 'markers',
          type: 'scatter',
          hovertext: sel2.map(function (o) { return prefix + buildMapPointHover(o); }),
          hoverinfo: 'text',
          hovertemplate: '%{hovertext}<extra></extra>',
          marker: {
            size: sel2.map(function (o) { return o.size; }),
            sizemode: 'diameter',
            sizemin: 8,
            color: COL[code2] || '#888',
            line: { width: 0.8, color: 'rgba(255,255,255,0.2)' }
          }
        });
      }
    } else if (out.length) {
      data.push({
        x: out.map(function (o) { return o.x; }),
        y: out.map(function (o) { return o.y; }),
        text: out.map(function (o) { return o.row.cliente || ''; }),
        name: 'Clientes',
        mode: 'markers',
        type: 'scatter',
        showlegend: false,
        hovertext: out.map(buildMapPointHover),
        hoverinfo: 'text',
        hovertemplate: '%{hovertext}<extra></extra>',
        marker: {
          size: out.map(function (o) { return o.size; }),
          sizemode: 'diameter',
          sizemin: 8,
          color: out.map(function (o) { return o.color; }),
          line: { width: 0.6, color: 'rgba(255,255,255,0.2)' }
        }
      });
    }
    var lineMed = { color: 'rgba(148,163,184,0.6)', width: 1, dash: 'dash' };
    var cornerAnnot = [
      { xref: 'paper', yref: 'paper', x: 0.99, y: 0.98, xanchor: 'right', yanchor: 'top', text: 'Ideal', showarrow: false, font: { size: 11, color: '#22c55e' } },
      { xref: 'paper', yref: 'paper', x: 0.01, y: 0.98, xanchor: 'left', yanchor: 'top', text: 'Oportunidad', showarrow: false, font: { size: 11, color: '#f59e0b' } },
      { xref: 'paper', yref: 'paper', x: 0.99, y: 0.04, xanchor: 'right', yanchor: 'bottom', text: 'Revisar precios', showarrow: false, font: { size: 11, color: '#3b82f6' } },
      { xref: 'paper', yref: 'paper', x: 0.01, y: 0.04, xanchor: 'left', yanchor: 'bottom', text: 'Acción urgente', showarrow: false, font: { size: 11, color: '#ef4444' } }
    ];
    return { data: data, layout: {
      paper_bgcolor: '#1a222d', plot_bgcolor: '#121920',
      font: { color: '#e8edf4', size: 12 },
      margin: { t: 44, r: 20, b: 56, l: 64 },
      xaxis: {
        title: { text: 'Cumplimiento de contrato (%)' },
        type: xLog ? 'log' : 'linear',
        autorange: true,
        showgrid: true,
        gridcolor: 'rgba(45,58,74,0.85)',
        zeroline: true, zerolinecolor: 'rgba(100,116,139,0.4)'
      },
      yaxis: {
        title: { text: 'Margen / ventas %' },
        type: 'linear',
        autorange: true,
        showgrid: true,
        gridcolor: 'rgba(45,58,74,0.85)',
        zeroline: true, zerolinecolor: 'rgba(100,116,139,0.4)'
      },
      legend: { orientation: 'h', yanchor: 'bottom', y: 1.12, x: 0, font: { size: 11, color: '#e8edf4' } },
      shapes: [
        { type: 'line', x0: tCx, x1: tCx, y0: 0, y1: 1, yref: 'paper', line: lineMed },
        { type: 'line', y0: tMy, y1: tMy, x0: 0, x1: 1, xref: 'paper', line: lineMed }
      ],
      annotations: cornerAnnot,
      hoverlabel: { bgcolor: '#0f1419', bordercolor: '#2d3a4a', font: { color: '#e8edf4' } }
    } };
  }

  function updateMapMedianBar() {
    var el = document.getElementById('map-median-text');
    if (!el) return;
    var tC = state.cumpThresh, tM = state.margThresh;
    if (isNaN(tC)) tC = 0.5; if (isNaN(tM)) tM = 0.1;
    el.textContent = 'Líneas de mediana: Cumplimiento = ' + (tC * 100).toFixed(1) + '%, margen = ' + (tM * 100).toFixed(1) + '% (cartera completa).';
  }

  function updateStrategicMap() {
    var box = document.getElementById('strategic-map-empty');
    updateMapMedianBar();
    if (typeof Plotly === 'undefined') { console.warn('Plotly no disponible'); if (box) { box.hidden = false; box.textContent = 'No se pudo cargar Plotly.'; } return; }
    var R = state.rows; if (!R) return;
    var g = getChartData(R);
    debugChartData(g, R);
    if (g.out.length === 0) {
      if (box) { box.hidden = false; box.textContent = 'No hay clientes graficables con los filtros actuales.'; }
      try { Plotly.purge('plotly-chart'); } catch (e) {}
    } else {
      if (box) box.hidden = true;
      var b = buildPlotlyFromGetChartData(g);
      if (!b.data.length) { if (box) { box.hidden = false; } try { Plotly.purge('plotly-chart'); } catch (e2) {} return; }
      Plotly.react('plotly-chart', b.data, b.layout, { displayModeBar: true, responsive: true });
    }
  }

  function drawChart(rows) { updateStrategicMap(); }

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

  function topVendorVals(rows, limit) {
    var n = limit || 3;
    var cnt = {};
    for (var i = 0; i < rows.length; i++) {
      var v = String(rows[i].vendedor || '—').trim();
      if (!v) v = '—';
      cnt[v] = (cnt[v] || 0) + 1;
    }
    var names = Object.keys(cnt).filter(function (v2) { return v2 !== '—'; });
    names.sort(function (a, b) {
      var d = (cnt[b] || 0) - (cnt[a] || 0);
      if (d !== 0) return d;
      return a.localeCompare(b, 'es', { sensitivity: 'base' });
    });
    if (!names.length) return [];
    return names.slice(0, n);
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
    if (state.rows) {
      buildVendChipsFromRows(state.rows);
      resetMapFilters();
      configureMapSlidersFromRows(state.rows);
      updateMapChipCounts(state.rows);
    }
    refreshView();
    if (state.rows) populateFilters(state.rows);
  }

  function syncSliderLabels() {
    var c = document.getElementById('map-min-contratos');
    var cu = document.getElementById('map-min-cump');
    var lv = document.getElementById('map-min-contratos-val');
    var lcu = document.getElementById('map-min-cump-val');
    if (c && lv) lv.textContent = c.value;
    if (cu && lcu) lcu.textContent = cu.value;
  }

  function configureMapSlidersFromRows(rows) {
    if (!rows || !rows.length) return;
    var mx = 1;
    for (var i = 0; i < rows.length; i++) mx = Math.max(mx, Math.ceil(Math.max(0, +rows[i].contratos || 0)));
    mx = Math.min(200, Math.max(50, mx));
    var el = document.getElementById('map-min-contratos');
    if (el && el.type === 'range') {
      el.min = '0';
      el.max = String(mx);
      if (+el.value > mx) el.value = String(mx);
    }
    syncSliderLabels();
  }

  function resetMapFilters() {
    mapState = { vend: null, pub: null, lic: null };
    var ms = document.getElementById('map-search'); if (ms) ms.value = '';
    var mc = document.getElementById('map-min-contratos'); if (mc) { mc.value = '0'; if (mc.type === 'range' && +mc.value > +mc.max) mc.value = mc.max; }
    var mcu = document.getElementById('map-min-cump'); if (mcu) mcu.value = '0';
    var cby = document.getElementById('map-color-by'); if (cby) cby.value = 'quadrant';
    var sby = document.getElementById('map-size-by'); if (sby) sby.value = 'contratos';
    var xsc = document.getElementById('map-x-scale'); if (xsc) xsc.value = 'linear';
    syncSliderLabels();
    syncMapChipUI();
  }

  function countBy(rows, fn) {
    var o = {};
    for (var i = 0; i < rows.length; i++) {
      var k = fn(rows[i]);
      o[k] = (o[k] || 0) + 1;
    }
    return o;
  }

  function updateMapChipCounts(rows) {
    if (!rows) return;
    var cv = countBy(rows, function (r) { return r.vendedor || '—'; });
    var cpub = countBy(rows, function (r) { return r.publico || '—'; });
    var clic = countBy(rows, function (r) { return r.licitacion || '—'; });
    document.querySelectorAll('#map-chips-vend [data-cnt-v]').forEach(function (el) {
      var k = el.getAttribute('data-cnt-v');
      el.textContent = '(' + (cv[k] || 0) + ')';
    });
    var aPub = cpub['Público'] || 0, aPrv = cpub['Privado'] || 0;
    var elPub = document.querySelector('#map-chips-pub [data-cnt="pub:Publico"]');
    var elPrv = document.querySelector('#map-chips-pub [data-cnt="pub:Privado"]');
    if (elPub) elPub.textContent = '(' + aPub + ')';
    if (elPrv) elPrv.textContent = '(' + aPrv + ')';
    var aSi = clic['Sí'] || 0, aNo = clic['No'] || 0;
    var elS = document.querySelector('#map-chips-lic [data-cnt="lic:Si"]');
    var elN = document.querySelector('#map-chips-lic [data-cnt="lic:No"]');
    if (elS) elS.textContent = '(' + aSi + ')';
    if (elN) elN.textContent = '(' + aNo + ')';
  }

  function buildVendChipsFromRows(rows) {
    var wrap = document.getElementById('map-chips-vend');
    if (!wrap) return;
    var d = topVendorVals(rows, 3);
    wrap.textContent = '';
    if (!d.length) {
      var b0 = document.createElement('button');
      b0.type = 'button';
      b0.className = 'map-chip';
      b0.disabled = true;
      b0.textContent = 'Sin dato de vendedor';
      wrap.appendChild(b0);
      return;
    }
    for (var i = 0; i < d.length; i++) {
      var v = d[i], vs = String(v);
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'map-chip';
      b.setAttribute('data-vend', vs);
      b.setAttribute('aria-pressed', 'false');
      b.appendChild(document.createTextNode(vs + ' '));
      var sp = document.createElement('span');
      sp.className = 'chip-cnt';
      sp.setAttribute('data-cnt-v', vs);
      sp.textContent = '(0)';
      b.appendChild(sp);
      wrap.appendChild(b);
    }
  }

  function syncMapChipUI() {
    var tv = document.getElementById('map-vend-todos');
    if (tv) { tv.classList.toggle('is-active', !mapState.vend); }
    document.querySelectorAll('#map-chips-vend [data-vend]').forEach(function (b) {
      var v = b.getAttribute('data-vend');
      var on = mapState.vend && mapState.vend.has(v);
      b.classList.toggle('is-active', !!on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    var tPub = document.getElementById('map-pub-todos');
    if (tPub) tPub.classList.toggle('is-active', !mapState.pub);
    document.querySelectorAll('#map-chips-pub [data-pub]').forEach(function (b) {
      var p = b.getAttribute('data-pub');
      if (p === 'all') return;
      var on = mapState.pub && mapState.pub.has(p);
      b.classList.toggle('is-active', !!on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    var tLic = document.getElementById('map-lic-todos');
    if (tLic) tLic.classList.toggle('is-active', !mapState.lic);
    document.querySelectorAll('#map-chips-lic [data-lic]').forEach(function (b) {
      var p = b.getAttribute('data-lic');
      if (p === 'all') return;
      var on = mapState.lic && mapState.lic.has(p);
      b.classList.toggle('is-active', !!on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
  }

  function exportMapCSV() {
    if (!state.rows) return;
    var g = getChartData(state.rows);
    if (!g.out.length) return;
    var h = 'Cliente;Código;Cumpl %;Margen %;Vendedor;Público/priv.;Licitación;Fact. esp.;Fact. real;Val. fact. MB;M. bruto;Potencial;Categoría\n';
    var b = h + g.out.map(function (o) {
      var r = o.row;
      return [r.cliente, r.codigo, o.x.toFixed(1), o.y.toFixed(1), r.vendedor, r.publico, r.licitacion,
        isNaN(r.fact_esp) ? '' : r.fact_esp, isNaN(r.fact_real) ? '' : r.fact_real, isNaN(r.valor_fact) ? '' : r.valor_fact,
        isNaN(r.margen_bruto) ? '' : r.margen_bruto, isNaN(r.pot) ? '' : r.pot, r.categoria
      ].map(function (c) { return (c == null ? '' : String(c).replace(/[;\n\r]/g, ' ')); }).join(';');
    }).join('\n');
    var link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([ '\ufeff' + b ], { type: 'text/csv;charset=utf-8;' }));
    link.download = 'mapa_estrategico_filtro.csv';
    link.click();
    setTimeout(function () { URL.revokeObjectURL(link.href); }, 3000);
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

  function handleMapVendClick(dv) {
    if (dv === 'all') mapState.vend = null;
    else {
      if (!mapState.vend) mapState.vend = new Set();
      if (mapState.vend.has(dv)) mapState.vend.delete(dv); else mapState.vend.add(dv);
      if (mapState.vend.size === 0) mapState.vend = null;
    }
    syncMapChipUI();
    if (state.rows) updateMapChipCounts(state.rows);
    updateStrategicMap();
  }
  function handleMapPubClick(p) {
    if (p === 'all') mapState.pub = null;
    else {
      if (!mapState.pub) mapState.pub = new Set();
      if (mapState.pub.has(p)) mapState.pub.delete(p); else mapState.pub.add(p);
      if (mapState.pub.size === 0) mapState.pub = null;
    }
    syncMapChipUI();
    if (state.rows) updateMapChipCounts(state.rows);
    updateStrategicMap();
  }
  function handleMapLicClick(l) {
    if (l === 'all') mapState.lic = null;
    else {
      if (!mapState.lic) mapState.lic = new Set();
      if (mapState.lic.has(l)) mapState.lic.delete(l); else mapState.lic.add(l);
      if (mapState.lic.size === 0) mapState.lic = null;
    }
    syncMapChipUI();
    if (state.rows) updateMapChipCounts(state.rows);
    updateStrategicMap();
  }

  function wireMapSection() {
    var sec = document.getElementById('section-chart');
    if (sec) {
      sec.addEventListener('click', function (ev) {
        var t = ev.target;
        if (t && t.classList && t.classList.contains('map-todos-link')) {
          ev.preventDefault();
          var w = t.getAttribute('data-todos');
          if (w === 'vend') handleMapVendClick('all');
          else if (w === 'pub') handleMapPubClick('all');
          else if (w === 'lic') handleMapLicClick('all');
          return;
        }
        if (!t || !t.closest) return;
        var chip = t.closest('.map-chip');
        if (!chip || !sec.contains(chip)) return;
        if (chip.hasAttribute('data-vend')) {
          ev.preventDefault();
          handleMapVendClick(chip.getAttribute('data-vend'));
          return;
        }
        if (chip.hasAttribute('data-pub')) {
          ev.preventDefault();
          handleMapPubClick(chip.getAttribute('data-pub'));
          return;
        }
        if (chip.hasAttribute('data-lic')) {
          ev.preventDefault();
          handleMapLicClick(chip.getAttribute('data-lic'));
        }
      });
    }
    var ms = document.getElementById('map-search');
    if (ms) ms.addEventListener('input', function () { if (state && state.rows) updateStrategicMap(); });
    var mcr = document.getElementById('map-min-contratos');
    if (mcr) {
      mcr.addEventListener('input', function () { syncSliderLabels(); if (state && state.rows) updateStrategicMap(); });
    }
    var mcu2 = document.getElementById('map-min-cump');
    if (mcu2) {
      mcu2.addEventListener('input', function () { syncSliderLabels(); if (state && state.rows) updateStrategicMap(); });
    }
    ['map-color-by', 'map-size-by', 'map-x-scale'].forEach(function (id) {
      var s = document.getElementById(id);
      if (s) s.addEventListener('change', function () { if (state && state.rows) updateStrategicMap(); });
    });
    var br = document.getElementById('map-btn-reset');
    if (br) br.addEventListener('click', function () { if (state.rows) { resetMapFilters(); if (state.rows) configureMapSlidersFromRows(state.rows); updateMapChipCounts(state.rows); } updateStrategicMap(); });
    var be = document.getElementById('map-btn-export');
    if (be) be.addEventListener('click', exportMapCSV);
  }

  function wireEvents() {
    var fc = document.getElementById('file-cc');
    var fm = document.getElementById('file-mb');
    if (fc) fc.addEventListener('change', handleFileInput('cc'));
    if (fm) fm.addEventListener('change', handleFileInput('mb'));
    wireMapSection();
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