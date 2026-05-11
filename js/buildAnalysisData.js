/**
 * Capa de preparación: empaqueta datos ya normalizados y consolidados para futuros módulos
 * (dashboards, filtros, KPIs, scatter 2x2, exportación). No renderiza UI.
 *
 * `customer_code` en `customers` debe ser string; las métricas numéricas ausentes van como null
 * donde no hay fuente (no se inventan ceros “semánticos” en tasas sin datos).
 */
import { listMissingExpectedColumns } from "./expectedColumns.js";
import {
  discoverAvailableFieldsBySlot,
  discoverRelatedColumnHintsBySlot,
  mergePossibleBusinessFields,
} from "./fieldDiscovery.js";
import { buildDataQualityReport } from "./dataQualityReport.js";
import { getCustomerNamePresentation } from "./mergeCustomers.js";

/** Quita el libro SheetJS (no serializable) de cada resultado de parseo. */
function filesSnapshotWithoutWorkbook(files) {
  return (files || []).map((f) => {
    if (!f || typeof f !== "object") return f;
    const { workbook: _wb, ...rest } = f;
    return rest;
  });
}

/** Mapa código → cruces por fuente, en forma de objeto plano (Sets → arrays). */
function customersByCodeToPlain(map) {
  const out = {};
  if (!map || typeof map.forEach !== "function") return out;
  map.forEach((rec, key) => {
    const np = getCustomerNamePresentation(rec);
    out[String(key)] = {
      customer_code: rec.customer_code,
      names: Array.from(rec.names || []),
      customer_display_name: np.customer_display_name,
      customer_name_sources: np.customer_name_sources,
      names_raw_joined: np.names_raw_joined,
      sales_rows: rec.sales_rows || [],
      margin_rows: rec.margin_rows || [],
      contract_rows: rec.contract_rows || [],
    };
  });
  return out;
}

/** Evita duplicar arrays grandes de filas en metadata (solo resumen por hoja). */
function summarizeHojasDetectadas(hd) {
  if (!hd) return { territorySales: [], margins: [], contracts: [] };
  const mapEntry = (e) => {
    if (!e) return e;
    const {
      normalizedRows: _nr,
      rowsRaw: _rr,
      excludedRowsSample: _ex,
      ...rest
    } = e;
    return rest;
  };
  return {
    territorySales: (hd.territorySales || []).map(mapEntry),
    margins: (hd.margins || []).map(mapEntry),
    contracts: (hd.contracts || []).map(mapEntry),
  };
}

/** @param {object} appState - Estado tras sync + merge (rawData, normalizedData, files, customers, diagnostics). */
export function buildAnalysisData(appState) {
  const rawData = appState.rawData;
  const normalizedData = appState.normalizedData;
  const files = appState.files || [];
  const customers = appState.customers || [];
  const diagnostics = appState.diagnostics || [];

  const availableFields = discoverAvailableFieldsBySlot(rawData);
  const relatedBySlot = discoverRelatedColumnHintsBySlot(rawData);
  const possibleBusinessFields = mergePossibleBusinessFields(relatedBySlot);

  const slots = files.map((f) => ({
    slot: f.slot,
    fileName: f.fileName,
    profile: f.profile,
    sheetName: f.sheetName || "",
    source_sheet: f.source_sheet || f.sheetName || "",
    year: f.year ?? null,
    period: f.period ?? null,
    currency: f.currency ?? null,
    detectionMethod: f.detection?.method ?? "",
    detectionReason: f.detection?.reason ?? "",
    keyColumn: f.diagnosticMeta?.keyColumn ?? f.territoryMeta?.clientColumn ?? null,
    rawRowCount: f.rowsRaw?.length ?? 0,
    normalizedRowCount: f.normalizedRows?.length ?? 0,
    normalizedRowCountConsolidation: f.normalizedRowsConsolidation?.length ?? 0,
    excludedCount: f.diagnosticMeta?.excludedCount ?? 0,
    missingExpectedColumns: listMissingExpectedColumns(f.profile, f.rowsRaw),
    relatedColumnHints: relatedBySlot[f.slot] || {},
    hojasDetectadas: summarizeHojasDetectadas(f.hojasDetectadas),
    sheetSelection: f.sheetSelection || null,
    sheetsUsed: f.sheetSelection?.sheetsUsed || (f.sheetName ? [f.sheetName] : []),
    consolidationMeta: f.consolidationMeta
      ? {
          activeSheetNames: f.consolidationMeta.activeSheetNames,
          inactiveSheetNames: f.consolidationMeta.inactiveSheetNames,
          multiSheetOverlapRisk: f.consolidationMeta.multiSheetOverlapRisk,
          sheetsRelation: f.consolidationMeta.sheetsRelation,
          territoryColumnAnalysis: f.consolidationMeta.territoryColumnAnalysis,
          aggregateHeadersDetected: f.consolidationMeta.aggregateHeadersDetected,
        }
      : null,
  }));

  const metadata = {
    generatedAt: new Date().toISOString(),
    slots,
    totals: {
      customersConsolidated: customers.length,
    },
    consolidatedCustomersNote:
      "La tabla `customers` y los campos `*_periodo_activo` usan solo la hoja/vista activa por archivo cuando hay varias hojas válidas; `normalizedData` conserva todas las filas etiquetadas.",
    activeDataView: appState.activeDataView
      ? JSON.parse(JSON.stringify(appState.activeDataView))
      : null,
  };

  const quality = buildDataQualityReport({
    files,
    normalizedData,
    customers,
    possibleBusinessFields,
    diagnostics,
  });

  return {
    activeDataView: appState.activeDataView
      ? JSON.parse(JSON.stringify(appState.activeDataView))
      : null,
    customers,
    margins: normalizedData.margins,
    contracts: normalizedData.contracts,
    territorySales: normalizedData.territorySales,
    normalizedData: {
      territorySales: normalizedData.territorySales,
      margins: normalizedData.margins,
      contracts: normalizedData.contracts,
    },
    rawData: {
      territorySales: rawData.territorySales,
      margins: rawData.margins,
      contracts: rawData.contracts,
    },
    customersByCode: customersByCodeToPlain(appState.customersByCode),
    files: filesSnapshotWithoutWorkbook(files),
    diagnostics,
    metadata,
    availableFields,
    possibleBusinessFields,
    quality,
  };
}

/** Indica si `analysisData` tiene el paquete mínimo generado en esta corrida (objeto interno tipo JSON). */
export function isAnalysisDataReady(analysisData) {
  return !!(analysisData && analysisData.metadata && typeof analysisData.metadata.generatedAt === "string");
}
