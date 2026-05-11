/**
 * Unión por `customer_code` y métricas consolidadas.
 *
 * Solo se usa `customer_code` como llave: es el identificador estable entre fuentes.
 * `customer_name` se usa solo para mostrar (prioridad en tabla consolidada), nunca para unir.
 *
 * No lee archivos ni toca el DOM.
 */
import {
  trimStr,
  parseLocaleNumber,
  joinUnique,
  parseAnalysisDateInput,
  parseContractDate,
  getContractStatus,
  dateToIsoDateLocal,
  startOfDay,
  normalizeCustomerCodeString,
  isValidNumericCustomerCode,
} from "./utils.js";

/**
 * Une filas normalizadas de ventas, márgenes y contratos en un Map por código.
 * @returns {Map<string, object>}
 */
export function mergeCustomersByCode(salesRows, marginsRows, contractsRows) {
  const map = new Map();

  const ensure = (code) => {
    const key = String(code);
    if (!map.has(key)) {
      map.set(key, {
        customer_code: key,
        names: new Set(),
        sales_rows: [],
        margin_rows: [],
        contract_rows: [],
      });
    }
    return map.get(key);
  };

  (salesRows || []).forEach((r) => {
    if (!isValidNumericCustomerCode(r.customer_code)) return;
    const code = normalizeCustomerCodeString(r.customer_code);
    const n = ensure(code);
    if (r.customer_name) n.names.add(r.customer_name);
    n.sales_rows.push(r);
  });
  (marginsRows || []).forEach((r) => {
    if (!isValidNumericCustomerCode(r.customer_code)) return;
    const code = normalizeCustomerCodeString(r.customer_code);
    const n = ensure(code);
    if (r.customer_name) n.names.add(r.customer_name);
    n.margin_rows.push(r);
  });
  (contractsRows || []).forEach((r) => {
    if (!isValidNumericCustomerCode(r.customer_code)) return;
    const code = normalizeCustomerCodeString(r.customer_code);
    const n = ensure(code);
    if (r.customer_name) n.names.add(r.customer_name);
    n.contract_rows.push(r);
  });

  return map;
}

function cleanNameFragment(s) {
  let t = trimStr(s);
  if (!t) return "";
  t = t.replace(/\s+/g, " ");
  t = t.replace(/,+\s*$/, "").trim();
  return t;
}

function normalizeForNameCompare(s) {
  return cleanNameFragment(s)
    .toLowerCase()
    .replace(/\s*,\s*/g, ",")
    .replace(/\s*·\s*/g, "·");
}

function uniqueNonEmptyNamesInOrder(rows) {
  const seen = new Set();
  const out = [];
  (rows || []).forEach((r) => {
    const n = cleanNameFragment(r.customer_name || "");
    if (!n) return;
    const k = normalizeForNameCompare(n);
    if (seen.has(k)) return;
    seen.add(k);
    out.push(n);
  });
  return out;
}

/** Texto de auditoría por fuente (puede incluir varias variantes separadas por ·). */
function auditStringForSourceRows(rows) {
  const u = uniqueNonEmptyNamesInOrder(rows);
  return u.length ? u.join(" · ") : "";
}

function pickLongestAmongEquivalent(names) {
  const byKey = new Map();
  names.forEach((p) => {
    const k = normalizeForNameCompare(p);
    const prev = byKey.get(k);
    if (!prev || p.length > prev.length) byKey.set(k, p);
  });
  return Array.from(byKey.values());
}

function bestSingleNameFromAuditString(auditStr) {
  if (!auditStr) return "";
  const parts = auditStr.split(/\s*·\s*/).map(cleanNameFragment).filter(Boolean);
  if (!parts.length) return "";
  const uniq = pickLongestAmongEquivalent(parts);
  if (uniq.length === 1) return uniq[0];
  let best = uniq[0];
  for (let i = 1; i < uniq.length; i += 1) {
    const c = uniq[i];
    if (isStrictlyMoreCompleteName(c, best)) best = c;
    else if (!isStrictlyMoreCompleteName(best, c) && c.length > best.length) best = c;
  }
  return best;
}

function isStrictlyMoreCompleteName(a, b) {
  if (!b) return !!a;
  if (!a) return false;
  const na = normalizeForNameCompare(a);
  const nb = normalizeForNameCompare(b);
  if (na === nb) return a.length > b.length;
  if (nb.includes(na) && b.length > a.length) return true;
  return false;
}

function buildCustomerNameSources(rec) {
  return {
    margins: auditStringForSourceRows(rec.margin_rows),
    contracts: auditStringForSourceRows(rec.contract_rows),
    territory_sales: auditStringForSourceRows(rec.sales_rows),
  };
}

function pickCustomerDisplayName(sources) {
  const m = bestSingleNameFromAuditString(sources.margins);
  const c = bestSingleNameFromAuditString(sources.contracts);
  const t = bestSingleNameFromAuditString(sources.territory_sales);
  const ordered = [
    { source: 0, n: m },
    { source: 1, n: c },
    { source: 2, n: t },
  ].filter((x) => x.n);
  if (!ordered.length) return "";
  let best = ordered[0];
  for (let i = 1; i < ordered.length; i += 1) {
    const cur = ordered[i];
    if (isStrictlyMoreCompleteName(cur.n, best.n)) best = cur;
    else if (isStrictlyMoreCompleteName(best.n, cur.n)) continue;
    else if (cur.source < best.source) best = cur;
  }
  return best.n;
}

/**
 * Nombre para UI y objeto de auditoría por fuente (no es llave de unión).
 * @param {object} rec - Entrada del mapa `mergeCustomersByCode`
 */
export function getCustomerNamePresentation(rec) {
  const customer_name_sources = buildCustomerNameSources(rec);
  const customer_display_name = pickCustomerDisplayName(customer_name_sources);
  return {
    customer_display_name,
    customer_name_sources,
    names_raw_joined: joinUnique(rec.names),
  };
}

function countDistinctContractIds(rec) {
  const ids = new Set();
  (rec.contract_rows || []).forEach((r) => {
    if (r.contract_id) ids.add(String(r.contract_id).trim());
  });
  return ids.size;
}

/** Número finito solo si hay filas de origen; si no hay fuente, null (no confundir con cero “real”). */
function metricFromSource(hasRows, value) {
  if (!hasRows) return null;
  return Number.isFinite(value) ? value : null;
}

function countDuplicateContractRows(contract_rows) {
  const byId = new Map();
  (contract_rows || []).forEach((r) => {
    const id = String(r.contract_id || "").trim();
    if (!id) return;
    byId.set(id, (byId.get(id) || 0) + 1);
  });
  let extra = 0;
  byId.forEach((n) => {
    if (n > 1) extra += n - 1;
  });
  return extra;
}

/** Agrega métricas de las tres fuentes para un único código (objeto interno del Map). */
export function aggregateCustomer(rec, opts = {}) {
  const analysisDay =
    opts.contractsAnalysisDate != null
      ? parseAnalysisDateInput(opts.contractsAnalysisDate) || startOfDay(new Date())
      : startOfDay(new Date());
  const analysisIso = dateToIsoDateLocal(analysisDay) || "";

  const namePres = getCustomerNamePresentation(rec);
  const customer_name = namePres.customer_display_name || namePres.names_raw_joined || "";
  const territory = joinUnique(new Set(rec.sales_rows.map((r) => r.territory).filter(Boolean)));

  let sales_total = 0;
  rec.sales_rows.forEach((r) => {
    sales_total += Number(r.sales_amount || 0) || 0;
  });

  let facturacion = 0;
  let margen_bruto = 0;
  let weightedMarginPct = 0;
  let marginWeight = 0;
  rec.margin_rows.forEach((r) => {
    const f = parseLocaleNumber(r.valor_facturado_loc);
    const mb = parseLocaleNumber(r.margen_bruto_loc);
    const mp = parseLocaleNumber(r.margen_sobre_ventas_loc);
    facturacion += f;
    margen_bruto += mb;
    if (Number.isFinite(mp) && f !== 0) {
      weightedMarginPct += mp * f;
      marginWeight += f;
    }
  });
  const margen_sobre_ventas_pct = marginWeight ? weightedMarginPct / marginWeight : null;

  const contractIds = new Set();
  let valor_contrato = 0;
  let valor_neto_contrato_activo = 0;
  let facturacion_esperada = 0;
  let facturado_neto = 0;
  let cumplWeighted = 0;
  let cumplWeight = 0;
  let contracts_active = 0;
  let contracts_expired = 0;
  let contracts_sin_fecha_fin = 0;

  rec.contract_rows.forEach((r) => {
    if (r.contract_id) contractIds.add(String(r.contract_id).trim());
    const vNet = parseLocaleNumber(r.valor_neto_contrato);
    valor_contrato += vNet;
    facturacion_esperada += parseLocaleNumber(r.facturacion_esperada);
    facturado_neto += parseLocaleNumber(r.facturado_neto);
    const c = parseLocaleNumber(r.cumplimiento_facturacion_pct);
    const w = Math.abs(parseLocaleNumber(r.facturacion_esperada)) || 1;
    if (Number.isFinite(c)) {
      cumplWeighted += c * w;
      cumplWeight += w;
    }
    let endDate = null;
    if (r.contract_end_parsed) {
      endDate = parseAnalysisDateInput(r.contract_end_parsed);
    }
    if (!endDate) {
      endDate = parseContractDate(r.fin).date;
    }
    const status = getContractStatus(endDate, analysisDay);
    r.analysis_date_used = analysisIso;
    r.contract_status = status;
    if (status === "sin_fecha_fin") {
      contracts_sin_fecha_fin += 1;
    } else if (status === "activo") {
      contracts_active += 1;
      valor_neto_contrato_activo += vNet;
    } else {
      contracts_expired += 1;
    }
  });

  const cumplimiento_pct = cumplWeight ? cumplWeighted / cumplWeight : null;

  const posibles_contratos_duplicados = countDuplicateContractRows(rec.contract_rows);

  return {
    customer_code: String(rec.customer_code),
    customer_name,
    customer_display_name: namePres.customer_display_name,
    customer_name_sources: namePres.customer_name_sources,
    names_raw_joined: namePres.names_raw_joined,
    territory,
    sales_total,
    facturacion,
    margen_bruto,
    margen_sobre_ventas_pct,
    contract_ids: Array.from(contractIds).join(", "),
    valor_contrato,
    valor_neto_contrato_activo,
    valor_neto_contrato_historico: valor_contrato,
    facturacion_esperada,
    facturado_neto,
    cumplimiento_pct,
    contracts_active,
    contracts_expired,
    contracts_sin_fecha_fin,
    contratos_total_historicos: contractIds.size,
    posibles_contratos_duplicados,
  };
}

/**
 * @param {Map<string, object>} customerMap - Resultado de mergeCustomersByCode
 * @param {object} [opts]
 * @param {string|Date|null} [opts.contractsAnalysisDate] - Fecha de corte para vigencia de contratos
 * @param {number|null} [opts.territoryActiveYear] - Año de referencia declarado para ventas en consolidado
 */
export function buildConsolidatedCustomerRows(customerMap, opts = {}) {
  const customers = [];
  customerMap.forEach((rec) => {
    const agg = aggregateCustomer(rec, opts);
    const hasSales = rec.sales_rows.length > 0;
    const hasMargins = rec.margin_rows.length > 0;
    const hasContracts = rec.contract_rows.length > 0;

    const ventasTotal = metricFromSource(hasSales, agg.sales_total);
    const valorFacturado = metricFromSource(hasMargins, agg.facturacion);
    const margenBruto = metricFromSource(hasMargins, agg.margen_bruto);
    const margenSobreVentas = metricFromSource(hasMargins, agg.margen_sobre_ventas_pct);
    const valorNetoContrato = metricFromSource(hasContracts, agg.valor_contrato);
    const valorNetoContratoActivo = metricFromSource(hasContracts, agg.valor_neto_contrato_activo);
    const facturacionEsperada = metricFromSource(hasContracts, agg.facturacion_esperada);
    const facturadoNeto = metricFromSource(hasContracts, agg.facturado_neto);
    const cumplimiento = metricFromSource(hasContracts, agg.cumplimiento_pct);

    customers.push({
      customer_code: String(normalizeCustomerCodeString(rec.customer_code)),
      customer_name: agg.customer_display_name || null,
      customer_display_name: agg.customer_display_name || null,
      customer_name_sources: agg.customer_name_sources,
      names_raw_joined: agg.names_raw_joined || null,
      aparece_en_ventas: hasSales,
      aparece_en_margenes: hasMargins,
      aparece_en_contratos: hasContracts,
      ventas_periodo_activo: ventasTotal,
      ventas_anio_activo: opts.territoryActiveYear ?? null,
      ventas_total: ventasTotal,
      valor_facturado_periodo_activo: valorFacturado,
      margen_bruto_periodo_activo: margenBruto,
      margen_sobre_ventas_periodo_activo: margenSobreVentas,
      valor_facturado_total: valorFacturado,
      margen_bruto_total: margenBruto,
      margen_sobre_ventas: margenSobreVentas,
      contratos_total: countDistinctContractIds(rec),
      contratos_total_historicos: agg.contratos_total_historicos,
      valor_neto_contrato_activo: valorNetoContratoActivo,
      valor_neto_contrato_historico: valorNetoContrato,
      valor_neto_contrato_total: valorNetoContrato,
      facturacion_esperada_total: facturacionEsperada,
      facturado_neto_total: facturadoNeto,
      cumplimiento_facturacion_promedio: cumplimiento,
      contratos_activos: agg.contracts_active,
      contratos_vencidos: agg.contracts_expired,
      contratos_sin_fecha_fin: agg.contracts_sin_fecha_fin,
      posibles_contratos_duplicados: agg.posibles_contratos_duplicados,
      x_valor_facturado_total: valorFacturado,
      y_margen_sobre_ventas: margenSobreVentas,
      x_ventas_total: ventasTotal,
      y_margen_bruto_total: margenBruto,
      x_valor_neto_contrato_total: valorNetoContratoActivo,
      y_cumplimiento_facturacion_promedio: cumplimiento,
    });
  });
  customers.sort((a, b) => String(a.customer_code).localeCompare(String(b.customer_code), "es"));
  return customers;
}
