/**
 * Vista activa de consolidación: evita sumar hojas/periodos traslapados en `customers`.
 * `normalizedData` / `rawData` conservan todas las hojas; el merge usa solo filas de consolidación.
 */
import { collectHeaders, normalizeHeader, isLikelyAggregateTotalHeader } from "./utils.js";
import { FILE_SLOT, createDefaultActiveDataView } from "./state.js";

export { createDefaultActiveDataView };

/** Etiqueta heurística de relación entre hojas del mismo libro y perfil. */
export function classifySheetsRelation(qualified) {
  if (!qualified || qualified.length <= 1) return "single_sheet";
  const years = new Set(qualified.map((h) => h.year).filter((y) => y != null));
  const currencies = new Set(qualified.map((h) => h.currency).filter(Boolean));
  if (years.size > 1) return "possible_distinct_periods";
  if (currencies.size > 1) return "possible_distinct_currencies";
  return "unknown_overlap_risk";
}

export function detectAggregateHeadersInRawRows(rowsRaw, maxRows = 120) {
  const labels = collectHeaders((rowsRaw || []).slice(0, maxRows));
  return labels.filter((h) => isLikelyAggregateTotalHeader(normalizeHeader(h)));
}

/**
 * @param {object} row
 * @param {object} ctx
 */
export function enrichNormalizedRow(row, ctx) {
  if (!row || typeof row !== "object") return row;
  return {
    ...row,
    source_file: ctx.source_file ?? null,
    source_sheet: ctx.source_sheet ?? row._source_sheet ?? null,
    profile: ctx.profile ?? null,
    detected_year: ctx.detected_year ?? null,
    detected_month: ctx.detected_month ?? null,
    detected_period: ctx.detected_period ?? null,
    detected_currency: ctx.detected_currency ?? null,
    detected_view: ctx.detected_view ?? null,
    is_aggregate_column: ctx.is_aggregate_column ?? false,
    is_monthly_column: ctx.is_monthly_column ?? false,
    is_total_column: ctx.is_total_column ?? false,
  };
}

export function enrichRows(rows, ctx) {
  return (rows || []).map((r) => enrichNormalizedRow(r, ctx));
}

export function tagRawRowsWithSource(rowsRaw, source_sheet, source_file) {
  return (rowsRaw || []).map((row) => ({
    ...row,
    __source_sheet: source_sheet,
    __source_file: source_file,
  }));
}

/**
 * @param {object[]} files - parseResults ordenados
 * @param {ReturnType<typeof createDefaultActiveDataView>} view - mutado
 */
export function applyActiveDataViewFromFiles(files, view) {
  const ts = files.find((f) => f.slot === FILE_SLOT.TERRITORY_SALES);
  const mg = files.find((f) => f.slot === FILE_SLOT.MARGINS);
  const ct = files.find((f) => f.slot === FILE_SLOT.CONTRACTS);

  if (ts?.consolidationMeta) {
    const m = ts.consolidationMeta;
    view.territorySales.activeSheets = m.activeSheetNames ? m.activeSheetNames.slice() : [];
    view.territorySales.activeYear = m.detectedYearPrimary ?? null;
  }
  if (mg?.consolidationMeta) {
    view.margins.activeSheet = mg.consolidationMeta.activeSheetNames?.[0] ?? null;
    view.margins.activePeriod = mg.consolidationMeta.detectedPeriodPrimary ?? null;
  }
  if (ct?.consolidationMeta) {
    view.contracts.activeSheets = ct.consolidationMeta.activeSheetNames
      ? ct.consolidationMeta.activeSheetNames.slice()
      : [];
    if (!view.contracts.analysisDate) {
      view.contracts.analysisDate = new Date().toISOString().slice(0, 10);
    }
  }
}

/**
 * Filas para merge de clientes (vista activa, una hoja por defecto si hay riesgo de traslape).
 */
export function consolidationNormalizedRowsForSlot(files, slot) {
  const f = (files || []).find((x) => x.slot === slot);
  if (!f) return [];
  return f.normalizedRowsConsolidation ?? f.normalizedRows ?? [];
}
