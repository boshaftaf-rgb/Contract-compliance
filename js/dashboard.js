/**
 * Dashboard inteligencia comercial
 * Capa visual construida sobre el motor comercial reutilizable.
 */
import { processCommercialExcels } from "./core/index.js";

(function () {
  'use strict';

  var state = {
    analysisData: null,
    files: {
      territorySalesFile: null,
      marginsFile: null,
      contractsFile: null
    },
    rows: null,
    sortMain: { col: 'cliente', dir: 1 },
    cumpCut: 'median',
    margCut: 'median',
    cumpThresh: null,
    margThresh: null,
    validationSummary: null
  };

  /**
   * Filtros sección 5 (mapa estratégico): territorio, público/privado, licitación.
   * null = todos; Set = OR dentro de cada dimensión.
   */
  var mapState = { territory: null, pub: null, lic: null, search: '', minC: 0, minCump: 0, colorBy: 'quadrant', sizeBy: 'contratos', xLog: false };

  /** Etiqueta unificada cuando no hay territorio en datos. */
  var SIN_TERRITORIO = 'Sin territorio';

  /** Coincide con `data-todos` del botón «Todos» de la tarjeta Territorio en index.html (histórico: vend). */
  var MAP_TERRITORY_TODOS_DOM = 'vend';

  /** Coincide con `value` de la opción «Territorio» en `#map-color-by` en index.html (histórico: vendedor). */
  var MAP_COLOR_BY_TERRITORY_SELECT_VALUE = 'vendedor';

  function isBlankTerritoryValue(v) {
    if (v == null) return true;
    var s = String(v).trim();
    if (!s) return true;
    if (s === '—' || s === '-' || s === '–') return true;
    return false;
  }

  function normalizeDashboardTerritoryLabel(v) {
    if (isBlankTerritoryValue(v)) return SIN_TERRITORIO;
    return String(v).trim();
  }

  /**
   * Solo capa dashboard: índice customer_code → territorio desde filas ya normalizadas del motor.
   * Última fila gana si hay duplicados de código.
   */
  function buildTerritoryByCustomerCodeFromSales(territorySalesRows) {
    var map = {};
    var arr = Array.isArray(territorySalesRows) ? territorySalesRows : [];
    for (var i = 0; i < arr.length; i++) {
      var row = arr[i];
      var code = String(row.customer_code != null ? row.customer_code : '');
      if (!code) continue;
      if (isBlankTerritoryValue(row.territory)) continue;
      map[code] = String(row.territory).trim();
    }
    return map;
  }

  function resolveDashboardTerritoryForCustomer(customer, territoryByCode) {
    var code = String(customer.customer_code != null ? customer.customer_code : '');
    if (!isBlankTerritoryValue(customer.territory)) return normalizeDashboardTerritoryLabel(customer.territory);
    var fromSales = code ? territoryByCode[code] : undefined;
    if (!isBlankTerritoryValue(fromSales)) return normalizeDashboardTerritoryLabel(fromSales);
    return SIN_TERRITORIO;
  }

  function logTerritoryDashboardDebug(territorySalesRows, territoryByCode, rows) {
    var ts = Array.isArray(territorySalesRows) ? territorySalesRows : [];
    var uniq = {};
    for (var u = 0; u < ts.length; u++) {
      var tv = ts[u].territory;
      if (!isBlankTerritoryValue(tv)) uniq[String(tv).trim()] = 1;
    }
    var uniqKeys = Object.keys(uniq).sort(function (a, b) { return a.localeCompare(b, 'es', { sensitivity: 'base' }); });
    var pairs = [];
    var seenSig = {};
    for (var p = 0; p < ts.length && pairs.length < 10; p++) {
      var tr = ts[p];
      var cc = String(tr.customer_code != null ? tr.customer_code : '');
      if (!cc || isBlankTerritoryValue(tr.territory)) continue;
      var sig = cc + '\t' + String(tr.territory).trim();
      if (seenSig[sig]) continue;
      seenSig[sig] = 1;
      pairs.push(cc + ' → ' + String(tr.territory).trim());
    }
    var withT = 0, sinT = 0;
    for (var r = 0; r < rows.length; r++) {
      if ((rows[r].territory || SIN_TERRITORIO) === SIN_TERRITORIO) sinT += 1;
      else withT += 1;
    }
    console.log('[Mapa 5][Territorio] registros en analysisData.territorySales:', ts.length);
    console.log('[Mapa 5][Territorio] códigos cliente con territorio en lookup (desde ventas):', Object.keys(territoryByCode || {}).length);
    console.log('[Mapa 5][Territorio] territorios únicos (valores en ventas):', uniqKeys.length, uniqKeys);
    console.log('[Mapa 5][Territorio] primeros 10 pares customer_code → territory (desde territorySales):', pairs);
    console.log('[Mapa 5][Territorio] clientes filas dashboard con territorio distinto de «' + SIN_TERRITORIO + '»:', withT);
    console.log('[Mapa 5][Territorio] clientes en «' + SIN_TERRITORIO + '»:', sinT);
  }

  function finiteNumber(value) {
    if (value === null || value === undefined || value === '') return NaN;
    var n = Number(value);
    return Number.isFinite(n) ? n : NaN;
  }

  function metricOrZero(value) {
    var n = finiteNumber(value);
    return isNaN(n) ? 0 : n;
  }

  /** Convierte porcentajes del motor a ratio 0-1 para las rutinas visuales existentes. */
  function toRatio(value) {
    var n = finiteNumber(value);
    if (isNaN(n)) return NaN;
    return Math.abs(n) > 1.5 ? n / 100 : n;
  }

  /** Convierte cumpl. / margen a cifra 0-100+ para el eje. */
  function toPercentValue(value) {
    var n = finiteNumber(value);
    if (isNaN(n)) return null;
    if (n >= -1.5 && n <= 1.5) return n * 100;
    return n;
  }

  async function loadCommercialData(ref) {
    var territorySalesFile = ref.territorySalesFile;
    var marginsFile = ref.marginsFile;
    var contractsFile = ref.contractsFile;
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS no se cargo. Comprueba la conexion a internet antes de procesar.');
    }
    return processCommercialExcels({
      territorySalesFile: territorySalesFile,
      marginsFile: marginsFile,
      contractsFile: contractsFile
    });
  }

  function buildValidationSummary(analysisData) {
    var customers = analysisData && analysisData.customers ? analysisData.customers : [];
    return {
      customers: customers.length,
      contracts: analysisData && analysisData.contracts ? analysisData.contracts.length : 0,
      customersWithSales: customers.filter(function (c) { return !!c.aparece_en_ventas; }).length,
      customersWithMargins: customers.filter(function (c) { return !!c.aparece_en_margenes; }).length,
      customersWithContracts: customers.filter(function (c) { return !!c.aparece_en_contratos; }).length
    };
  }

  function mapAnalysisDataToDashboardModel(analysisData) {
    var customers = analysisData && Array.isArray(analysisData.customers) ? analysisData.customers : [];
    var territorySalesSrc = analysisData && Array.isArray(analysisData.territorySales) ? analysisData.territorySales : [];
    var territoryByCustomerCode = buildTerritoryByCustomerCodeFromSales(territorySalesSrc);
    var rows = customers.map(function (c) {
      var factEsp = finiteNumber(c.facturacion_esperada_total);
      var factReal = finiteNumber(c.facturado_neto_total);
      var valorFact = finiteNumber(c.valor_facturado_total != null ? c.valor_facturado_total : c.valor_facturado_periodo_activo);
      var margenBruto = finiteNumber(c.margen_bruto_total != null ? c.margen_bruto_total : c.margen_bruto_periodo_activo);
      var cump = toRatio(c.cumplimiento_facturacion_promedio);
      if (isNaN(cump) && !isNaN(factEsp) && factEsp !== 0 && !isNaN(factReal)) cump = factReal / factEsp;
      var margenPct = toRatio(c.margen_sobre_ventas != null ? c.margen_sobre_ventas : c.margen_sobre_ventas_periodo_activo);
      if (isNaN(margenPct) && !isNaN(valorFact) && valorFact !== 0 && !isNaN(margenBruto)) margenPct = margenBruto / valorFact;
      var territory = resolveDashboardTerritoryForCustomer(c, territoryByCustomerCode);
      return {
        key: String(c.customer_code || ''),
        codigo: String(c.customer_code || '-'),
        cliente: c.customer_display_name || c.customer_name || c.names_raw_joined || String(c.customer_code || '-'),
        territory: territory,
        publico: '-',
        licitacion: '-',
        contratos: metricOrZero(c.contratos_total != null ? c.contratos_total : c.contratos_total_historicos),
        fact_esp: isNaN(factEsp) ? NaN : factEsp,
        fact_real: isNaN(factReal) ? NaN : factReal,
        cump: cump,
        valor_fact: isNaN(valorFact) ? NaN : valorFact,
        margen_bruto: isNaN(margenBruto) ? NaN : margenBruto,
        margen_pct: margenPct,
        dif: (!isNaN(factReal) && !isNaN(factEsp)) ? factReal - factEsp : NaN,
        pot: (!isNaN(factReal) && !isNaN(factEsp)) ? factEsp - factReal : NaN,
        aparece_en_ventas: !!c.aparece_en_ventas,
        aparece_en_margenes: !!c.aparece_en_margenes,
        aparece_en_contratos: !!c.aparece_en_contratos
      };
    });
    logTerritoryDashboardDebug(territorySalesSrc, territoryByCustomerCode, rows);
    return {
      rows: rows,
      validationSummary: buildValidationSummary(analysisData),
      analysisData: analysisData
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

  // --- KPIs y resumen global ---
  function kpiSums(rows) {
    // Agregados globales usados por KPI cards, resumen ejecutivo e insights.
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
    a.sort(function (a, b) {
      var sb = (!isNaN(b.fact_real) && !isNaN(b.cump)) ? (b.fact_real * (1 - b.cump)) : -Infinity;
      var sa = (!isNaN(a.fact_real) && !isNaN(a.cump)) ? (a.fact_real * (1 - a.cump)) : -Infinity;
      return sb - sa;
    });
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
    // Resumen ejecutivo (punto 2): frases gerenciales calculadas con indicadores agregados.
    var k = kpiSums(rows);
    var c5 = topShare(rows, 5, function (r) { return r.fact_real; });
    var c10 = topShare(rows, 10, function (r) { return r.fact_real; });
    var kopp = rows.filter(function (r) { return r.categoriaCode === 'opp'; }).length;
    var highValLowC = rows.filter(function (r) { return !isNaN(r.cump) && r.fact_real > k.tr * 0.02 && r.cump < 0.85; }).length;
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
    // Insights automáticos (punto 4): detecta brechas, cuentas críticas y señales de margen.
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
  /** Paleta para coloración discreta por territorio (select «Territorio» en mapa). */
  var TERRITORY_DISCRETE_COLORS = ['#0ea5e9', '#8b5cf6', '#f97316', '#10b981', '#ec4899', '#6366f1', '#14b8a6', '#eab308', '#64748b', '#a855f7', '#d946ef', '#0d9488'];

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
    if (colorBy === MAP_COLOR_BY_TERRITORY_SELECT_VALUE) {
      return TERRITORY_DISCRETE_COLORS[strHash(r.territory || SIN_TERRITORIO) % TERRITORY_DISCRETE_COLORS.length];
    }
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

  /** Filtros por chips del mapa estratégico (territorio, público/privado, licitación). */
  function passStrategicMapChipFilters(r) {
    if (mapState.territory) {
      var terr = (r.territory != null && String(r.territory) !== '') ? String(r.territory) : SIN_TERRITORIO;
      if (!mapState.territory.has(terr)) return { ok: false, reason: 'territorio' };
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

  /** Cumplimiento en % para eje: acepta valor calculado con base real del mismo cliente. */
  function resolveCumplAxis(r) {
    var v = toPercentValue(r.cump);
    if (v != null && isFinite(v)) return { value: v, source: 'calculado (fact. real / fact. esperada)' };
    var fe = +r.fact_esp, fr = +r.fact_real;
    if (fe > 0 && isFinite(fe) && isFinite(fr)) return { value: (fr / fe) * 100, source: 'calculado (fact. real / fact. esperada)' };
    return { value: null, source: 'faltante', missing: 'cumplimiento (fact. esp./fact. real)' };
  }
  /** Margen en % para eje: acepta valor calculado con base real del mismo cliente. */
  function resolveMargenAxis(r) {
    var v2 = toPercentValue(r.margen_pct);
    if (v2 != null && isFinite(v2)) return { value: v2, source: 'calculado (margen bruto / valor facturado)' };
    var vf = +r.valor_fact, mb = +r.margen_bruto;
    if (vf > 0 && isFinite(vf) && isFinite(mb)) return { value: (mb / vf) * 100, source: 'calculado (margen bruto / valor facturado)' };
    return { value: null, source: 'faltante', missing: 'margen (valor fact./margen bruto)' };
  }

  function getChartData(allRows) {
    // Construye dataset final para Plotly aplicando:
    // 1) chips/filtros, 2) búsqueda, 3) umbrales, 4) ejes y tamaño/color de burbuja.
    var nCross = allRows && allRows.length ? allRows.length : 0;
    var discarded = [];
    var quality = { xMissing: 0, yMissing: 0, bothMissing: 0, plotted: 0, evaluated: 0 };
    var out = [];
    var afterChipFilters = 0, withXY = 0;
    var q = (document.getElementById('map-search') && document.getElementById('map-search').value || '').trim().toLowerCase();
    var minCRaw = document.getElementById('map-min-contratos') ? parseInt(document.getElementById('map-min-contratos').value, 10) : 0;
    var minC = (isNaN(minCRaw) ? 0 : minCRaw);
    minC = Math.max(0, minC);
    var minCumpIn = document.getElementById('map-min-cump') ? document.getElementById('map-min-cump').value : '0';
    var minCumpP = Math.max(0, Math.min(200, parseFloat(minCumpIn) || 0));
    var colorBy = (document.getElementById('map-color-by') && document.getElementById('map-color-by').value) || 'quadrant';
    var sizeBy = (document.getElementById('map-size-by') && document.getElementById('map-size-by').value) || 'contratos';
    var xLog = (document.getElementById('map-x-scale') && document.getElementById('map-x-scale').value) === 'log';
    var tCu0 = state.cumpThresh, tMu0 = state.margThresh;
    if (isNaN(tCu0)) tCu0 = 0.5; if (isNaN(tMu0)) tMu0 = 0.1;
    for (var j = 0; j < nCross; j++) {
      var r2 = allRows[j];
      quality.evaluated += 1;
      var prf = passStrategicMapChipFilters(r2);
      if (!prf.ok) {
        discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'filtro: ' + prf.reason });
        continue;
      }
      afterChipFilters++;
      if (q) {
        var blob = (String(r2.cliente || '') + ' ' + String(r2.codigo || '')).toLowerCase();
        if (blob.indexOf(q) < 0) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'búsqueda' }); continue; }
      }
      if (minC > 0 && (+r2.contratos || 0) < minC) { discarded.push({ key: r2.key, cliente: r2.cliente, reason: 'N° contratos < mín.' }); continue; }
      var xMeta = resolveCumplAxis(r2);
      var yMeta = resolveMargenAxis(r2);
      var xP = xMeta.value;
      var yP = yMeta.value;
      if (!isFinite(xP) || !isFinite(yP)) {
        if (!isFinite(xP) && !isFinite(yP)) quality.bothMissing += 1;
        else if (!isFinite(xP)) quality.xMissing += 1;
        else quality.yMissing += 1;
        var missReason = (!isFinite(xP) && !isFinite(yP))
          ? (xMeta.missing + ' + ' + yMeta.missing)
          : (!isFinite(xP) ? xMeta.missing : yMeta.missing);
        discarded.push({ key: r2.key, cliente: r2.cliente, reason: missReason });
        continue;
      }
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
        plotQuadrantCode: plotQ, color: cMark,
        xSource: xMeta.source, ySource: yMeta.source
      });
      quality.plotted += 1;
    }
    return { nCross: nCross, afterChipFilters: afterChipFilters, withXY: withXY, out: out, discarded: discarded, colorBy: colorBy, sizeBy: sizeBy, xLog: xLog, minCumpP: minCumpP, minC: minC, q: q, quality: quality };
  }

  function debugChartData(g, allRows) {
    console.log('[Mapa 5] Registros cruzados (total filas):', g.nCross);
    console.log('[Mapa 5] Tras chips territorio / público / licitación:', g.afterChipFilters);
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
      (r.cliente || '—'), 'Cód: ' + (r.codigo || '—'), 'Territorio: ' + (r.territory || '—'),
      'Publ./priv.: ' + (r.publico || '—'), 'Licit.: ' + (r.licitacion || '—'),
      'Cumpl. %: ' + o.x.toFixed(1), 'Margen %: ' + o.y.toFixed(1),
      'Origen eje X: ' + (o.xSource || '—'), 'Origen eje Y: ' + (o.ySource || '—'),
      'Fact. esp.: ' + fmtMoney(r.fact_esp), 'Fact. real: ' + fmtMoney(r.fact_real),
      'Val. fact. MB: ' + fmtMoney(r.valor_fact), 'M. bruto: ' + fmtMoney(r.margen_bruto),
      'Potencial: ' + fmtMoney(r.pot), 'Categoría: ' + (r.categoria || '—')
    ].join('<br>');
  }

  function buildPlotlyFromGetChartData(g) {
    // Arma trazas/layout del mapa estratégico (punto 5) con líneas de mediana por cartera completa.
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

  function updateMapMedianBar(g) {
    var el = document.getElementById('map-median-text');
    if (!el) return;
    var tC = state.cumpThresh, tM = state.margThresh;
    if (isNaN(tC)) tC = 0.5; if (isNaN(tM)) tM = 0.1;
    var base = 'Líneas de mediana: Cumplimiento = ' + (tC * 100).toFixed(1) + '%, margen = ' + (tM * 100).toFixed(1) + '% (cartera completa).';
    if (!g || !g.quality) { el.textContent = base; return; }
    var q = g.quality;
    var miss = q.xMissing + q.yMissing + q.bothMissing;
    var qualityTxt = ' Evaluados: ' + q.evaluated + '. Graficados: ' + q.plotted + '. Excluidos por faltantes: ' + miss +
      ' (sin X: ' + q.xMissing + ', sin Y: ' + q.yMissing + ', sin ambos: ' + q.bothMissing + ').';
    el.textContent = base + qualityTxt;
  }

  function updateStrategicMap() {
    var box = document.getElementById('strategic-map-empty');
    if (typeof Plotly === 'undefined') { console.warn('Plotly no disponible'); if (box) { box.hidden = false; box.textContent = 'No se pudo cargar Plotly.'; } return; }
    var R = state.rows; if (!R) return;
    var g = getChartData(R);
    updateMapMedianBar(g);
    debugChartData(g, R);
    if (g.out.length === 0) {
      if (box) {
        var q = g.quality || { xMissing: 0, yMissing: 0, bothMissing: 0 };
        box.hidden = false;
        box.textContent = 'No hay clientes graficables con los filtros actuales. Faltantes: sin X=' + q.xMissing + ', sin Y=' + q.yMissing + ', sin ambos=' + q.bothMissing + '.';
      }
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
      var sy = (!isNaN(y.fact_real) && !isNaN(y.cump)) ? (y.fact_real * (1 - y.cump)) : -Infinity;
      var sx = (!isNaN(x.fact_real) && !isNaN(x.cump)) ? (x.fact_real * (1 - x.cump)) : -Infinity;
      return sy - sx;
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
    one('table-seg-vendedor', groupSegment(rows, function (r) { return r.territory || SIN_TERRITORIO; }));
    one('table-seg-publico', groupSegment(rows, function (r) { return r.publico || '—'; }));
    one('table-seg-licit', groupSegment(rows, function (r) { return r.licitacion || '—'; }));
  }

  function distinctVals(rows, fn) {
    var s = {}; for (var i = 0; i < rows.length; i++) { s[fn(rows[i])] = 1; }
    return Object.keys(s).sort();
  }

  function populateFilters(rows) {
    var opts = function (elId, allLabel, values) { var s = document.getElementById(elId); s.innerHTML = '<option value="">' + allLabel + '</option>'; values.forEach(function (v) { var o = document.createElement('option'); o.value = v; o.textContent = v; s.appendChild(o); }); };
    opts('filter-vendedor', 'Todos', distinctVals(rows, function (r) { return r.territory || SIN_TERRITORIO; }));
    opts('filter-publico', 'Todos', distinctVals(rows, function (r) { return r.publico || '—'; }));
    opts('filter-licit', 'Todos', distinctVals(rows, function (r) { return r.licitacion || '—'; }));
    opts('filter-cat', 'Todas', ['Ideal', 'Oportunidad', 'Revisar precios', 'Acción urgente', 'Sin dato'].filter(function (c) { return rows.some(function (r) { return (r.categoria || '') === c; }); })
    );
  }

  function filterMain(rows) {
    var q = (document.getElementById('table-search') && document.getElementById('table-search').value || '').toLowerCase();
    var filterTerritory = document.getElementById('filter-vendedor').value;
    var fp = document.getElementById('filter-publico').value;
    var fl = document.getElementById('filter-licit').value;
    var fc = document.getElementById('filter-cat').value;
    var flt = document.getElementById('filter-bajo-cump').value;
    var tC2 = state.cumpThresh;
    return rows.filter(function (r) {
      if (q && (String(r.cliente) + ' ' + String(r.codigo)).toLowerCase().indexOf(q) < 0) return false;
      if (filterTerritory && (r.territory || SIN_TERRITORIO) !== filterTerritory) return false;
      if (fp && (r.publico || '—') !== fp) return false;
      if (fl && (r.licitacion || '—') !== fl) return false;
      if (fc && (r.categoria || '') !== fc) return false;
      if (flt === 'lowCump' && !(r.cump < tC2)) return false;
      if (flt === 'lowMargin' && !(r.margen_pct < (state.margThresh || 0))) return false;
      return true;
    });
  }

  var MAIN_COLS = [
    { k: 'codigo', t: 'Código' }, { k: 'cliente', t: 'Cliente' }, { k: 'territory', t: 'Territorio' },
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
    return c === 'cliente' || c === 'codigo' || c === 'territory' || c === 'publico' || c === 'licitacion' || c === 'categoria';
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
    renderRisks(state.rows);
    renderOpps(state.rows);
    renderSegments(state.rows);
    renderMainTable();
    try {
      drawChart(state.rows);
    } catch (e) {
      console.error('No se pudo renderizar el mapa estratégico:', e);
    }
  }

  function runFullRender() {
    if (state.rows) {
      buildMapTerritoryChipsFromRows(state.rows);
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
    mapState.territory = null;
    mapState.pub = null;
    mapState.lic = null;
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
    var cv = countBy(rows, function (r) { return r.territory || SIN_TERRITORIO; });
    var cpub = countBy(rows, function (r) { return r.publico || '—'; });
    var clic = countBy(rows, function (r) { return r.licitacion || '—'; });
    document.querySelectorAll('#map-chips-vend [data-cnt-territory]').forEach(function (el) {
      var k = el.getAttribute('data-cnt-territory');
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

  function buildMapTerritoryChipsFromRows(rows) {
    var wrap = document.getElementById('map-chips-vend');
    if (!wrap) return;
    wrap.textContent = '';
    if (!rows || !rows.length) {
      var bEmpty = document.createElement('button');
      bEmpty.type = 'button';
      bEmpty.className = 'map-chip';
      bEmpty.disabled = true;
      bEmpty.textContent = 'Sin clientes';
      wrap.appendChild(bEmpty);
      return;
    }
    var cv = countBy(rows, function (r) { return r.territory || SIN_TERRITORIO; });
    var realTerr = Object.keys(cv).filter(function (k) { return k !== SIN_TERRITORIO; });
    realTerr.sort(function (a, b) {
      var d = (cv[b] || 0) - (cv[a] || 0);
      if (d !== 0) return d;
      return a.localeCompare(b, 'es', { sensitivity: 'base' });
    });
    function appendChip(label) {
      var vs = String(label);
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'map-chip';
      b.setAttribute('data-territory', vs);
      b.setAttribute('aria-pressed', 'false');
      b.appendChild(document.createTextNode(vs + ' '));
      var sp = document.createElement('span');
      sp.className = 'chip-cnt';
      sp.setAttribute('data-cnt-territory', vs);
      sp.textContent = '(0)';
      b.appendChild(sp);
      wrap.appendChild(b);
    }
    for (var i = 0; i < realTerr.length; i++) appendChip(realTerr[i]);
    if ((cv[SIN_TERRITORIO] || 0) > 0) appendChip(SIN_TERRITORIO);
    if (!realTerr.length && !(cv[SIN_TERRITORIO] > 0)) {
      var b0 = document.createElement('button');
      b0.type = 'button';
      b0.className = 'map-chip';
      b0.disabled = true;
      b0.textContent = 'Sin dato de territorio';
      wrap.appendChild(b0);
    }
  }

  function syncMapChipUI() {
    var tv = document.getElementById('map-vend-todos');
    if (tv) { tv.classList.toggle('is-active', !mapState.territory); }
    document.querySelectorAll('#map-chips-vend [data-territory]').forEach(function (b) {
      var v = b.getAttribute('data-territory');
      var on = mapState.territory && mapState.territory.has(v);
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
    var h = 'Cliente;Código;Cumpl %;Margen %;Territorio;Público/priv.;Licitación;Fact. esp.;Fact. real;Val. fact. MB;M. bruto;Potencial;Categoría\n';
    var b = h + g.out.map(function (o) {
      var r = o.row;
      return [r.cliente, r.codigo, o.x.toFixed(1), o.y.toFixed(1), r.territory, r.publico, r.licitacion,
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

  async function runPipeline() {
    var files = state.files;
    if (!files.territorySalesFile || !files.marginsFile || !files.contractsFile) return;
    setUploadError('Procesando archivos con el motor comercial...');
    try {
      var analysisData = await loadCommercialData({
        territorySalesFile: files.territorySalesFile,
        marginsFile: files.marginsFile,
        contractsFile: files.contractsFile
      });
      var dashboardModel = mapAnalysisDataToDashboardModel(analysisData);
      state.analysisData = analysisData;
      state.rows = dashboardModel.rows;
      state.validationSummary = dashboardModel.validationSummary;
      var a = document.getElementById('upload-alerts');
      if (a) {
        a.hidden = false;
        a.textContent = 'Procesado: ' + state.validationSummary.customers + ' clientes - ' + state.validationSummary.contracts + ' contratos - ventas: ' + state.validationSummary.customersWithSales + ' - margenes: ' + state.validationSummary.customersWithMargins + ' - contratos: ' + state.validationSummary.customersWithContracts + '.';
      }
      var hint = document.getElementById('hint-both');
      if (hint) hint.hidden = false;
      console.info('[Motor comercial] analysisData:', analysisData);
      console.info('[Motor comercial] validacion:', state.validationSummary);
      if (!state.rows.length) {
        setUploadError('El motor proceso los archivos, pero no devolvio clientes consolidados.');
        showSections(false);
        return;
      }
      runFullRender();
    } catch (err) {
      console.error('No se pudo procesar con el motor comercial:', err);
      setUploadError(err && err.message ? err.message : 'No se pudo procesar con el motor comercial.');
      showSections(false);
    }
  }

  function handleFileInput(which) {
    return function (e) {
      var f = e.target.files && e.target.files[0];
      if (!f) return;
      state.files[which] = f;
      if (which === 'territorySalesFile') setFileStatus('status-territory', f.name, '-', '-', '');
      else if (which === 'marginsFile') setFileStatus('status-mb', f.name, '-', '-', '');
      else if (which === 'contractsFile') setFileStatus('status-cc', f.name, '-', '-', '');
      if (state.files.territorySalesFile && state.files.marginsFile && state.files.contractsFile) runPipeline();
    };
  }

  function handleMapTerritoryChipClick(territoryKey) {
    if (territoryKey === 'all') mapState.territory = null;
    else {
      if (!mapState.territory) mapState.territory = new Set();
      if (mapState.territory.has(territoryKey)) mapState.territory.delete(territoryKey); else mapState.territory.add(territoryKey);
      if (mapState.territory.size === 0) mapState.territory = null;
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
          if (w === MAP_TERRITORY_TODOS_DOM) handleMapTerritoryChipClick('all');
          else if (w === 'pub') handleMapPubClick('all');
          else if (w === 'lic') handleMapLicClick('all');
          return;
        }
        if (!t || !t.closest) return;
        var chip = t.closest('.map-chip');
        if (!chip || !sec.contains(chip)) return;
        if (chip.hasAttribute('data-territory')) {
          ev.preventDefault();
          handleMapTerritoryChipClick(chip.getAttribute('data-territory'));
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
    var ft = document.getElementById('file-territory');
    if (fc) fc.addEventListener('change', handleFileInput('contractsFile'));
    if (fm) fm.addEventListener('change', handleFileInput('marginsFile'));
    if (ft) ft.addEventListener('change', handleFileInput('territorySalesFile'));
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