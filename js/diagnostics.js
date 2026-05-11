/**
 * Diagnóstico de filas: validación por celda/fila y utilidades para el panel avanzado (exclusiones).
 *
 * Las filas inválidas no se “pierden”: van a `excludedRows` con motivo, para revisión en UI
 * sin mezclar esa lista con el flujo principal de filas normalizadas.
 */
import {
  trimStr,
  valueToString,
  normalizeCustomerCodeString,
  isSummaryToken,
  hasAnyUsefulCell,
  escapeHtml,
  stringifyCell,
  parseContractDate,
  parseAnalysisDateInput,
  getContractStatus,
  dateToIsoDateLocal,
  startOfDay,
} from "./utils.js";

function isRepeatedHeaderText(lower) {
  return (
    (lower.includes("representante") && lower.includes("cliente")) ||
    lower === "cliente" ||
    lower === "representante" ||
    lower.includes("territorio / region") ||
    lower.includes("territorio / región")
  );
}

function isRepresentativeOrTerritoryText(lower) {
  return (
    (lower.includes("representante") && !/^\d/.test(lower)) ||
    (lower.includes("territorio") && !/^\d/.test(lower)) ||
    (lower.includes("region") && !/^\d/.test(lower))
  );
}

export function diagnoseTerritoryClienteCell(cell) {
  const text = trimStr(valueToString(cell));
  if (!text) return { ok: false, reason: "campo cliente vacío" };
  if (isSummaryToken(text)) return { ok: false, reason: "valor de resumen o total" };

  const lower = text.toLowerCase();
  if (
    lower.includes("subtotal") ||
    lower.includes("total general") ||
    lower.startsWith("total ")
  ) {
    return { ok: false, reason: "valor de resumen o total" };
  }

  if (isRepeatedHeaderText(lower)) {
    return { ok: false, reason: "encabezado repetido" };
  }

  if (isRepresentativeOrTerritoryText(lower)) {
    return { ok: false, reason: "fila de representante o territorio" };
  }

  const startsNumeric = /^\s*\d+/.test(text);
  if (!startsNumeric) {
    return { ok: false, reason: "no inicia con código numérico" };
  }

  const onlyDigits = /^\s*(\d{6,12})\s*$/.test(text);
  if (onlyDigits) {
    return { ok: false, reason: "inicia con código pero no tiene nombre después" };
  }

  const m = text.match(/^\s*(\d{6,12})\s+(.+)$/);
  if (!m) return { ok: false, reason: "formato no reconocido" };

  const code = normalizeCustomerCodeString(m[1]);
  if (!/^\d{6,12}$/.test(code)) return { ok: false, reason: "formato no reconocido" };
  const name = trimStr(m[2].replace(/^[-–—|/:]+\s*/, ""));
  if (!name) return { ok: false, reason: "inicia con código pero no tiene nombre después" };

  return { ok: true, customer_code: code, customer_name: name };
}

export function diagnoseMarginsKey(codeRaw, row) {
  const code = trimStr(valueToString(codeRaw));
  if (!code) return { ok: false, reason: "Cód. Pagador vacío" };
  if (isSummaryToken(code)) return { ok: false, reason: "fila de total o resumen" };
  if (!hasAnyUsefulCell(row)) return { ok: false, reason: "fila sin datos útiles" };
  const normalized = normalizeCustomerCodeString(code);
  if (!/^\d+$/.test(normalized)) return { ok: false, reason: "Cód. Pagador no numérico" };
  if (!normalized) return { ok: false, reason: "formato no reconocido" };
  return { ok: true };
}

export function diagnoseContractsRow(codeRaw, contractRaw, row) {
  const code = trimStr(valueToString(codeRaw));
  if (!code) return { ok: false, reason: "Solicitante Cliente vacío" };
  if (isSummaryToken(code)) return { ok: false, reason: "fila de total o resumen" };
  if (!hasAnyUsefulCell(row)) return { ok: false, reason: "fila sin datos útiles" };
  const normalized = normalizeCustomerCodeString(code);
  if (!/^\d+$/.test(normalized)) return { ok: false, reason: "Solicitante Cliente no numérico" };
  if (!trimStr(valueToString(contractRaw))) return { ok: false, reason: "Contrato vacío" };
  if (!normalized) return { ok: false, reason: "formato no reconocido" };
  return { ok: true };
}

/** Fragmento HTML de tabla para filas excluidas (diagnóstico avanzado). */
export function previewExcludedRows(rows, context) {
  if (!rows.length) return "";
  const thead = `
        <tr>
          <th>Perfil</th>
          <th>Hoja</th>
          <th>Columna evaluada</th>
          <th>Fila Excel</th>
          <th>Valor columna cliente</th>
          <th>Motivo de exclusión</th>
        </tr>
      `;
  const tbody = rows
    .map(
      (r) => `
          <tr>
            <td>${escapeHtml(stringifyCell(context.profile))}</td>
            <td>${escapeHtml(stringifyCell(r.source_sheet || context.sheetName))}</td>
            <td>${escapeHtml(stringifyCell(context.keyColumn))}</td>
            <td>${escapeHtml(stringifyCell(r.excel_row))}</td>
            <td>${escapeHtml(stringifyCell(r.customer_value))}</td>
            <td>${escapeHtml(stringifyCell(r.reason))}</td>
          </tr>
        `
    )
    .join("");
  return `<table><thead>${thead}</thead><tbody>${tbody}</tbody></table>`;
}

/**
 * Consolida metadatos de exclusión por archivo para el estado global `diagnostics`
 * y para trazabilidad en consola (ver validationLog).
 * @param {object[]} parseResults - Salidas de parseWorkbookForSlot en orden de slots
 * @returns {object[]}
 */
export function buildAdvancedDiagnostics(parseResults) {
  return (parseResults || [])
    .filter((p) => p && p.diagnosticMeta)
    .map((p) => ({
      slot: p.slot,
      fileName: p.fileName,
      profile: p.diagnosticMeta.profile,
      sheetName: p.diagnosticMeta.sheetName,
      keyColumn: p.diagnosticMeta.keyColumn,
      rawRowCount: p.diagnosticMeta.rawRowCount,
      normalizedCount: p.diagnosticMeta.normalizedCount,
      excludedCount: p.diagnosticMeta.excludedCount,
      excludedRows: [...(p.diagnosticMeta.excludedRows || [])],
      consolidationUsesSheets: p.diagnosticMeta.consolidationUsesSheets || [],
      sheetsExcludedFromConsolidationSum: p.diagnosticMeta.sheetsExcludedFromConsolidationSum || [],
      aggregateHeadersDetected: p.diagnosticMeta.aggregateHeadersDetected || [],
      territoryDetailVsTotal: p.diagnosticMeta.territoryDetailVsTotal || null,
    }));
}

/**
 * Metadatos de política de consolidación (hoja activa, columnas agregadas, riesgo de traslape).
 * @param {object[]} parseResults
 * @returns {object[]}
 */
export function buildConsolidationDiagnosticsList(parseResults) {
  return (parseResults || [])
    .filter((p) => p && p.consolidationMeta)
    .map((p) => ({
      kind: "consolidation_policy",
      slot: p.slot,
      fileName: p.fileName,
      activeSheetNames: p.consolidationMeta.activeSheetNames || [],
      inactiveSheetNames: p.consolidationMeta.inactiveSheetNames || [],
      multiSheetOverlapRisk: !!p.consolidationMeta.multiSheetOverlapRisk,
      sheetsRelation: p.consolidationMeta.sheetsRelation || "",
      territoryColumnAnalysis: p.consolidationMeta.territoryColumnAnalysis || null,
      aggregateHeadersDetected: p.consolidationMeta.aggregateHeadersDetected || [],
      fullNormalizedRowCount: p.consolidationMeta.fullNormalizedRowCount ?? 0,
      consolidationRowCount: p.consolidationMeta.consolidationRowCount ?? 0,
    }));
}

/**
 * Fragmento HTML solo para «diagnóstico avanzado» (archivo contratos): parseo DD/MM, ISO, serial Excel;
 * conteos activo/vencido/sin fin; muestras de raw → parseado. No modifica filas ni la vista principal.
 */
export function buildContractDateDiagnosticsHtml(normalizedRows, analysisDateRaw) {
  const rows = normalizedRows || [];
  const analysisDay = parseAnalysisDateInput(analysisDateRaw) || startOfDay(new Date());
  const analysisLabel = dateToIsoDateLocal(analysisDay) || "—";
  let activo = 0;
  let vencido = 0;
  let sinFin = 0;
  let noParseable = 0;
  let ambiguousCount = 0;
  const samples = [];
  const badSamples = [];

  rows.forEach((r) => {
    const pr = parseContractDate(r.fin);
    const st = getContractStatus(pr.date, analysisDay);
    if (pr.ambiguous) ambiguousCount += 1;
    if (st === "activo") activo += 1;
    else if (st === "vencido") vencido += 1;
    else sinFin += 1;
    const raw = trimStr(pr.raw || valueToString(r.fin));
    if (raw && !pr.date) noParseable += 1;
    const id = trimStr(String(r.contract_id || ""));
    if (samples.length < 10 && id) {
      samples.push({
        id,
        raw,
        iso: pr.iso || "—",
        status: st,
        warn: pr.warning || "",
      });
    }
    if (badSamples.length < 8 && raw && !pr.date) {
      badSamples.push({ id: id || "—", raw, warn: pr.warning || "—" });
    }
  });

  const sampleLines = samples
    .map(
      (s) =>
        `<li><code>${escapeHtml(s.id)}</code>: raw <code>${escapeHtml(s.raw || "—")}</code> → parseado <code>${escapeHtml(s.iso)}</code> · <strong>${escapeHtml(s.status)}</strong>${s.warn ? ` · ${escapeHtml(s.warn)}` : ""}</li>`,
    )
    .join("");
  const badLines = badSamples
    .map(
      (s) =>
        `<li><code>${escapeHtml(s.id)}</code>: <code>${escapeHtml(s.raw)}</code> — ${escapeHtml(s.warn)}</li>`,
    )
    .join("");

  return `<p class="muted sample-caption"><strong>Contratos — fechas de fin (DD/MM/AAAA México) vs análisis ${escapeHtml(analysisLabel)}</strong></p>
    <ul class="diag-list">
      <li><strong>Activos:</strong> ${activo} · <strong>Vencidos:</strong> ${vencido} · <strong>Sin fecha fin / no clasificable:</strong> ${sinFin} · <strong>Texto de fin no parseable:</strong> ${noParseable}</li>
      <li><strong>Interpretación ambigua (Date nativa):</strong> ${ambiguousCount}</li>
    </ul>
    <p class="muted sample-caption"><strong>Ejemplos (contract_id · raw · parseado · estado)</strong></p>
    <ul class="diag-list">${sampleLines || "<li>—</li>"}</ul>
    ${
      badLines
        ? `<p class="muted sample-caption"><strong>Fechas de fin no parseables (muestra)</strong></p><ul class="diag-list">${badLines}</ul>`
        : ""
    }`;
}
