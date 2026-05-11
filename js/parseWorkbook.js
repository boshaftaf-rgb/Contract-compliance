/**
 * Construye el resultado de parseo para un slot (perfil esperado + libro XLSX).
 * Usa análisis por contenido de todas las hojas; el nombre de la hoja solo apoya metadatos.
 *
 * Regla de consolidación: `normalizedRows` = todas las hojas calificadas (para raw/normalized completos).
 * `normalizedRowsConsolidation` = solo la hoja activa elegida (no sumar hojas con riesgo de traslape).
 */
import { analyzeWorkbookSheets } from "./sheetAnalyzer.js";
import { detectProfileFromFileName } from "./profiles.js";
import {
  classifySheetsRelation,
  detectAggregateHeadersInRawRows,
  enrichRows,
  tagRawRowsWithSource,
} from "./consolidationSupport.js";

function emptyParse({ slot, expectedProfile, fileName, reason, alerts = [], workbook = null, extras = {} }) {
  return {
    slot,
    role: expectedProfile,
    fileName,
    workbook,
    sheetName: "",
    source_sheet: "",
    year: null,
    period: null,
    currency: null,
    detection: {
      profile: "unknown",
      method: "unknown",
      reason,
      sheetName: "",
      scores: { margins: 0, contracts: 0, territory_sales: 0 },
    },
    profile: "unknown",
    rowsRaw: [],
    normalizedRows: [],
    normalizedRowsConsolidation: [],
    consolidationMeta: null,
    territoryMeta: null,
    diagnosticMeta: null,
    alerts,
    hojasDetectadas: { territorySales: [], margins: [], contracts: [] },
    sheetSelection: null,
    ...extras,
  };
}

function maxScoresFromHojas(hojasDetectadas) {
  const maxOf = (arr) => (arr && arr.length ? Math.max(...arr.map((h) => h.confidenceScore || 0)) : 0);
  return {
    margins: maxOf(hojasDetectadas.margins),
    contracts: maxOf(hojasDetectadas.contracts),
    territory_sales: maxOf(hojasDetectadas.territorySales),
  };
}

function slotListForExpected(hojasDetectadas, expectedProfile) {
  if (expectedProfile === "territory_sales") return hojasDetectadas.territorySales;
  if (expectedProfile === "margins") return hojasDetectadas.margins;
  return hojasDetectadas.contracts;
}

function getQualifiedSheetsSorted(hojasDetectadas, expectedProfile) {
  const list = slotListForExpected(hojasDetectadas, expectedProfile);
  const qualified = list.filter((h) => h.confidenceScore >= 68);
  qualified.sort(
    (a, b) =>
      b.normalizedRowCount - a.normalizedRowCount ||
      b.confidenceScore - a.confidenceScore ||
      String(a.sheetName).localeCompare(String(b.sheetName)),
  );
  return qualified;
}

function mergeYears(qualified) {
  const ys = qualified.map((h) => h.year).filter((y) => y != null);
  if (!ys.length) return null;
  const first = ys[0];
  return ys.every((y) => y === first) ? first : null;
}

function buildParseFromQualifiedSheets({
  slot,
  expectedProfile,
  fileName,
  workbook,
  qualified,
  hojasDetectadas,
  fileLevelReason,
}) {
  const active = qualified[0];
  const inactive = qualified.slice(1);
  const sheetLabels = qualified.map((h) => h.source_sheet || h.sheetName);
  const activeLabel = active.source_sheet || active.sheetName;
  const sheetRelation = classifySheetsRelation(qualified);
  const multi = qualified.length > 1;

  const alerts = [...new Set(qualified.flatMap((h) => h.alerts || []))];
  if (multi) {
    alerts.push(
      `Consolidado principal: solo la hoja «${activeLabel}». Otras ${inactive.length} hoja(s) con el mismo perfil se conservan en el dataset normalizado completo y no se suman automáticamente (relación: ${sheetRelation}).`,
    );
  }

  const rowsRaw = qualified.flatMap((h) =>
    tagRawRowsWithSource(h.rowsRaw || [], h.sheetName, fileName),
  );

  const normalizedRowsAll = qualified.flatMap((h) => {
    const ctx = {
      source_file: fileName,
      source_sheet: h.sheetName,
      profile: expectedProfile,
      detected_year: h.year ?? null,
      detected_month: null,
      detected_period: h.period ?? null,
      detected_currency: h.currency ?? null,
      detected_view: null,
      is_aggregate_column: false,
      is_monthly_column: false,
      is_total_column: false,
    };
    return enrichRows(h.normalizedRows || [], ctx);
  });

  const consolidationCtx = {
    source_file: fileName,
    source_sheet: active.sheetName,
    profile: expectedProfile,
    detected_year: active.year ?? null,
    detected_month: null,
    detected_period: active.period ?? null,
    detected_currency: active.currency ?? null,
    detected_view: "consolidation_active_sheet",
    is_aggregate_column: false,
    is_monthly_column: false,
    is_total_column: false,
  };
  const normalizedRowsConsolidation = enrichRows(active.normalizedRows || [], consolidationCtx);

  const rawRowTotal = qualified.reduce((s, h) => s + (h.rawRowCount || 0), 0);
  const excludedTotal = qualified.reduce((s, h) => s + (h.excludedCount || 0), 0);
  const excludedRowsCombined = qualified
    .flatMap((h) =>
      (h.diagnosticMeta?.excludedRows || h.excludedRowsSample || []).map((er) => ({
        ...er,
        source_sheet: h.sheetName,
      })),
    )
    .slice(0, 25);

  const scores = maxScoresFromHojas(hojasDetectadas);
  const joinedAll = sheetLabels.join(" · ");
  const detection = {
    profile: expectedProfile,
    method: "sheet_content",
    reason:
      fileLevelReason ||
      (multi
        ? `Dataset completo: ${qualified.length} hojas (${joinedAll}). Consolidado: solo «${activeLabel}».`
        : `Hoja «${activeLabel}» (confianza ${active.confidenceScore}, ${active.normalizedRowCount} filas normalizadas).`),
    sheetName: multi ? `${activeLabel} (activa) · otras: ${inactive.map((i) => i.sheetName).join(", ") || "—"}` : activeLabel,
    scores,
  };

  const aggregateHeadersActive =
    expectedProfile === "margins" || expectedProfile === "contracts"
      ? detectAggregateHeadersInRawRows(active.rowsRaw).slice(0, 20)
      : [];

  const consolidationMeta = {
    activeSheetNames: [active.sheetName],
    inactiveSheetNames: inactive.map((h) => h.sheetName),
    multiSheetOverlapRisk: multi,
    sheetsRelation: sheetRelation,
    detectedYearPrimary: active.year ?? null,
    detectedPeriodPrimary: active.period ?? null,
    detectedCurrencyPrimary: active.currency ?? null,
    territoryColumnAnalysis: expectedProfile === "territory_sales" ? active.territoryMeta?.columnAnalysis ?? null : null,
    aggregateHeadersDetected: aggregateHeadersActive,
    consolidationRowCount: normalizedRowsConsolidation.length,
    fullNormalizedRowCount: normalizedRowsAll.length,
  };

  let territoryMeta = null;
  if (expectedProfile === "territory_sales" && active.territoryMeta) {
    const tm = active.territoryMeta;
    territoryMeta = {
      ...tm,
      activeSheetForConsolidation: activeLabel,
      otherSheetsNotSummed: inactive.map((h) => h.sheetName),
      rawRowCount: rawRowTotal,
      normalizedCountFull: normalizedRowsAll.length,
      normalizedCountConsolidation: normalizedRowsConsolidation.length,
      excludedCount: excludedTotal,
      excludedRows: excludedRowsCombined.slice(0, 20),
      sampleClients: normalizedRowsConsolidation.slice(0, 5),
    };
  }

  const diagnosticMeta = {
    profile: expectedProfile,
    sheetName: detection.sheetName,
    keyColumn: active.diagnosticMeta?.keyColumn ?? active.territoryMeta?.clientColumn ?? "",
    rawRowCount: rawRowTotal,
    normalizedCount: normalizedRowsAll.length,
    excludedCount: excludedTotal,
    excludedRows: excludedRowsCombined,
    consolidationUsesSheets: consolidationMeta.activeSheetNames,
    sheetsExcludedFromConsolidationSum: consolidationMeta.inactiveSheetNames,
    aggregateHeadersDetected: aggregateHeadersActive,
    territoryDetailVsTotal: consolidationMeta.territoryColumnAnalysis || null,
  };

  return {
    slot,
    role: expectedProfile,
    fileName,
    workbook,
    sheetName: activeLabel,
    source_sheet: multi ? `${activeLabel} (consolidado)` : activeLabel,
    year: mergeYears(qualified),
    period: [...new Set(qualified.map((h) => h.period).filter(Boolean))].join(" · ") || null,
    currency: qualified.map((h) => h.currency).find((c) => c != null) ?? null,
    detection,
    profile: expectedProfile,
    rowsRaw,
    normalizedRows: normalizedRowsAll,
    normalizedRowsConsolidation,
    consolidationMeta,
    territoryMeta,
    diagnosticMeta,
    alerts,
    hojasDetectadas,
    sheetSelection: {
      expectedProfile,
      sheetsUsed: sheetLabels,
      sheetsUsedInConsolidation: [active.sheetName],
      sheetsNotSummedInConsolidation: inactive.map((h) => h.sheetName),
      primarySheetName: active.sheetName,
      alternateSheetNames: inactive.map((h) => h.sheetName),
      selectableSheetNames: [],
      alternatesDetail: qualified.map((h) => ({
        sheetName: h.sheetName,
        confidenceScore: h.confidenceScore,
        normalizedRowCount: h.normalizedRowCount,
        rawRowCount: h.rawRowCount,
        excludedCount: h.excludedCount,
      })),
    },
  };
}

export function parseWorkbookForSlot({
  slot,
  expectedProfile,
  fileName,
  workbook,
  stubReason,
  stubAlerts,
}) {
  if (!workbook) {
    return emptyParse({
      slot,
      expectedProfile,
      fileName: fileName || "(sin archivo)",
      reason: stubReason || "No se seleccionó archivo para este origen.",
      alerts:
        stubAlerts && stubAlerts.length
          ? stubAlerts
          : [
              "Sin archivo: la unión y los totales no incluirán datos de este origen hasta que cargue el .xlsx correspondiente.",
            ],
    });
  }

  const sheetNames = workbook.SheetNames || [];
  if (!sheetNames.length) {
    return emptyParse({
      slot,
      expectedProfile,
      fileName,
      reason: "No hay hoja legible en el libro (archivo vacío o sin pestañas).",
      workbook,
      alerts: [
        "El libro no contiene hojas utilizables. Revise que el archivo no esté corrupto o vacío.",
      ],
    });
  }

  const analysis = analyzeWorkbookSheets(workbook, expectedProfile, fileName);
  const { hojasDetectadas, fileLevelReason } = analysis;
  const qualified = getQualifiedSheetsSorted(hojasDetectadas, expectedProfile);

  if (!qualified.length) {
    const fn = detectProfileFromFileName(fileName);
    const scores = maxScoresFromHojas(hojasDetectadas);
    const reason =
      fileLevelReason ||
      (fn.profile === expectedProfile
        ? "Ninguna hoja alcanzó el umbral de confianza; el nombre del archivo sugiere el tipo correcto."
        : "No se encontró una hoja con columnas reconocibles para este tipo de archivo.");
    const alerts = [reason];
    if (fn.profile === expectedProfile && fn.reason) alerts.push(fn.reason);
    return emptyParse({
      slot,
      expectedProfile,
      fileName,
      reason,
      workbook,
      alerts,
      extras: {
        hojasDetectadas,
        sheetSelection: {
          expectedProfile,
          primarySheetName: null,
          alternateSheetNames: slotListForExpected(hojasDetectadas, expectedProfile).map((h) => h.sheetName),
          selectableSheetNames: [],
          alternatesDetail: [],
        },
        detection: {
          profile: "unknown",
          method: "sheet_content",
          reason,
          sheetName: "",
          scores,
        },
      },
    });
  }

  return buildParseFromQualifiedSheets({
    slot,
    expectedProfile,
    fileName,
    workbook,
    qualified,
    hojasDetectadas,
    fileLevelReason,
  });
}
