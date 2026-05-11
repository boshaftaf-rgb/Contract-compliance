/**
 * Indicadores de calidad y consistencia para auditoría (sin UI).
 * Alimenta el resumen compacto de preparación de datos.
 */
import { parseLocaleNumber, parseContractDate, isValidNumericCustomerCode } from "./utils.js";

function countRowsMissingValidCode(normalizedRows) {
  let n = 0;
  (normalizedRows || []).forEach((row) => {
    if (!isValidNumericCustomerCode(row.customer_code)) n += 1;
  });
  return n;
}

/**
 * @param {object} params
 * @param {object[]} params.files
 * @param {object} params.normalizedData
 * @param {object[]} params.customers - filas consolidadas
 * @param {object} params.possibleBusinessFields
 * @param {object[]} params.diagnostics
 */
export function buildDataQualityReport({
  files,
  normalizedData,
  customers,
  possibleBusinessFields,
  diagnostics,
}) {
  const filesUnrecognized = (files || []).filter((f) => f.workbook && f.profile === "unknown").length;
  const emptyRawSheets = (files || []).filter((f) => f.workbook && (!f.rowsRaw || f.rowsRaw.length === 0)).length;

  let excludedRowsTotal = 0;
  (diagnostics || []).forEach((d) => {
    excludedRowsTotal += d.excludedCount || 0;
  });

  let contractsWithoutEndDate = 0;
  (normalizedData?.contracts || []).forEach((row) => {
    if (!parseContractDate(row.fin).date) contractsWithoutEndDate += 1;
  });

  let negativeMarginRows = 0;
  let zeroBilledRows = 0;
  (normalizedData?.margins || []).forEach((row) => {
    const mb = parseLocaleNumber(row.margen_bruto_loc);
    const vf = parseLocaleNumber(row.valor_facturado_loc);
    if (Number.isFinite(mb) && mb < 0) negativeMarginRows += 1;
    if (Number.isFinite(vf) && vf === 0) zeroBilledRows += 1;
  });

  const rowsMissingCode =
    countRowsMissingValidCode(normalizedData?.territorySales) +
    countRowsMissingValidCode(normalizedData?.margins) +
    countRowsMissingValidCode(normalizedData?.contracts);

  let clientsSingleSource = 0;
  let clientsMultiSource = 0;
  (customers || []).forEach((c) => {
    const n = [c.aparece_en_ventas, c.aparece_en_margenes, c.aparece_en_contratos].filter(Boolean).length;
    if (n === 1) clientsSingleSource += 1;
    if (n >= 2) clientsMultiSource += 1;
  });

  const possibleBucketsWithHits = Object.entries(possibleBusinessFields || {}).filter(([, list]) => list.length);

  return {
    filesUnrecognized,
    emptyRawSheets,
    excludedRowsTotal,
    contractsWithoutEndDate,
    negativeMarginRows,
    zeroBilledRows,
    normalizedRowsMissingValidCustomerCode: rowsMissingCode,
    clientsSingleSource,
    clientsMultiSource,
    possibleBusinessBucketsHit: possibleBucketsWithHits.length,
    diagnosticsEntries: (diagnostics || []).length,
  };
}
