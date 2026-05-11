/**
 * Normalización por perfil: produce filas con `customer_code` y lista de exclusiones.
 * Funciones puras (sin DOM). Los motivos de exclusión alimentan el diagnóstico avanzado.
 *
 * Reglas de exclusión (resumen):
 * - Ventas territorio: celdas vacías, totales, encabezados repetidos, filas que no parecen cliente+código.
 * - Márgenes: código pagador vacío/no numérico, filas de resumen, filas sin datos útiles.
 * - Contratos: solicitante vacío/no numérico, contrato vacío, mismas reglas de resumen.
 */
import {
  findColumnKey,
  collectHeaders,
  normalizeHeader,
  normalizeCustomerCodeString,
  trimStr,
  valueToString,
  parseLocaleNumber,
  parseContractDate,
  isLikelyIdColumn,
  isTerritoryRegionHeader,
  monthlyColumnHint,
  isLikelyAggregateTotalHeader,
} from "./utils.js";
import {
  diagnoseMarginsKey,
  diagnoseContractsRow,
  diagnoseTerritoryClienteCell,
} from "./diagnostics.js";
import {
  MARGINS_COLUMNS,
  CONTRACTS_COLUMNS,
  TERRITORY_SALES_CLIENT_COLUMNS,
  TERRITORY_SALES_REGION_COLUMNS,
} from "./columnLabels.js";

/**
 * @typedef {object} NormalizeResult
 * @property {object[]} rows
 * @property {string} keyColumn
 * @property {number} excludedCount
 * @property {object[]} excludedRows
 * @property {string} [userMessage] - Mensaje para UI si falta columna clave
 */

/**
 * Normaliza filas del archivo de márgenes.
 * @param {object[]} rows - Filas crudas con __excel_row
 * @returns {NormalizeResult & { rows: object[] }}
 */
export function normalizeMarginsFile(rows) {
  const colPayer = findColumnKey(rows, MARGINS_COLUMNS.PAYER_CODE);
  const colName = findColumnKey(rows, MARGINS_COLUMNS.PAYER_NAME);
  const colFact = findColumnKey(rows, MARGINS_COLUMNS.BILLED_VALUE);
  const colMb = findColumnKey(rows, MARGINS_COLUMNS.GROSS_MARGIN);
  const colMp = findColumnKey(rows, MARGINS_COLUMNS.MARGIN_OVER_SALES);
  if (!colPayer) {
    return {
      rows: [],
      keyColumn: "Cód. Pagador (no detectada)",
      userMessage:
        "No se encontró la columna «Cód. Pagador». Sin ella no se puede obtener `customer_code` en márgenes.",
      excludedCount: rows.length,
      excludedRows: rows.slice(0, 20).map((row) => ({
        excel_row: row.__excel_row || "",
        customer_value: "",
        reason: "formato no reconocido",
      })),
    };
  }

  const out = [];
  const excludedRows = [];
  rows.forEach((row) => {
    const codeRaw = row[colPayer];
    const diag = diagnoseMarginsKey(codeRaw, row);
    if (!diag.ok) {
      if (excludedRows.length < 20) {
        excludedRows.push({
          excel_row: row.__excel_row || "",
          customer_value: trimStr(valueToString(codeRaw)),
          reason: diag.reason,
        });
      }
      return;
    }
    out.push({
      customer_code: normalizeCustomerCodeString(codeRaw),
      customer_name: colName ? trimStr(valueToString(row[colName])) : "",
      valor_facturado_loc: colFact ? row[colFact] : "",
      margen_bruto_loc: colMb ? row[colMb] : "",
      margen_sobre_ventas_loc: colMp ? row[colMp] : "",
      _source: row,
    });
  });
  return {
    rows: out,
    keyColumn: colPayer,
    excludedCount: Math.max(0, rows.length - out.length),
    excludedRows,
  };
}

/**
 * Normaliza filas del archivo de contratos.
 * @param {object[]} rows
 */
export function normalizeContractsFile(rows) {
  const colCode = findColumnKey(rows, CONTRACTS_COLUMNS.CLIENT_CODE);
  const colName = findColumnKey(rows, CONTRACTS_COLUMNS.CLIENT_NAME);
  const colContract = findColumnKey(rows, CONTRACTS_COLUMNS.CONTRACT_ID);
  const colInicio = findColumnKey(rows, CONTRACTS_COLUMNS.START);
  const colFin = findColumnKey(rows, CONTRACTS_COLUMNS.END);
  const colValor = findColumnKey(rows, CONTRACTS_COLUMNS.NET_CONTRACT_VALUE);
  const colEsp = findColumnKey(rows, CONTRACTS_COLUMNS.EXPECTED_BILLING);
  const colNet = findColumnKey(rows, CONTRACTS_COLUMNS.NET_BILLED);
  const colCumpl = findColumnKey(rows, CONTRACTS_COLUMNS.BILLING_COMPLIANCE_PCT);
  if (!colCode) {
    return {
      rows: [],
      keyColumn: "Solicitante Cliente (no detectada)",
      userMessage:
        "No se encontró la columna «Solicitante Cliente». Sin ella no se puede obtener `customer_code` en contratos.",
      excludedCount: rows.length,
      excludedRows: rows.slice(0, 20).map((row) => ({
        excel_row: row.__excel_row || "",
        customer_value: "",
        reason: "formato no reconocido",
      })),
    };
  }

  const out = [];
  const excludedRows = [];
  rows.forEach((row) => {
    const codeRaw = row[colCode];
    const diag = diagnoseContractsRow(codeRaw, colContract ? row[colContract] : "", row);
    if (!diag.ok) {
      if (excludedRows.length < 20) {
        excludedRows.push({
          excel_row: row.__excel_row || "",
          customer_value: trimStr(valueToString(codeRaw)),
          reason: diag.reason,
        });
      }
      return;
    }
    const finCell = colFin ? row[colFin] : null;
    const finParsed = parseContractDate(finCell);
    out.push({
      customer_code: normalizeCustomerCodeString(codeRaw),
      customer_name: colName ? trimStr(valueToString(row[colName])) : "",
      contract_id: colContract ? trimStr(valueToString(row[colContract])) : "",
      inicio: colInicio ? row[colInicio] : "",
      fin: colFin ? row[colFin] : "",
      contract_end_raw: trimStr(valueToString(finCell ?? "")),
      contract_end_parsed: finParsed.iso,
      date_parse_warning: finParsed.warning,
      analysis_date_used: null,
      contract_status: null,
      valor_neto_contrato: colValor ? row[colValor] : "",
      facturacion_esperada: colEsp ? row[colEsp] : "",
      facturado_neto: colNet ? row[colNet] : "",
      cumplimiento_facturacion_pct: colCumpl ? row[colCumpl] : "",
      _source: row,
    });
  });
  return {
    rows: out,
    keyColumn: colCode,
    excludedCount: Math.max(0, rows.length - out.length),
    excludedRows,
  };
}

/**
 * Clasifica columnas de ventas: mensuales, agregadas (Total, etc.) e ignoradas (auxiliares).
 * Solo monthlyColumns entran en ventas_periodo_activo (salvo que no haya ninguna; entonces fallback a agregados).
 */
function partitionTerritorySalesColumns(headers, colRep, colTerr) {
  const monthly = [];
  const aggregate = [];
  const ignored = [];
  (headers || []).forEach((h) => {
    if (h === colRep || h === colTerr) {
      ignored.push(h);
      return;
    }
    const n = normalizeHeader(h);
    if (n.includes("representante") && n.includes("cliente")) {
      ignored.push(h);
      return;
    }
    if (isTerritoryRegionHeader(n)) {
      ignored.push(h);
      return;
    }
    if (isLikelyIdColumn(n)) {
      ignored.push(h);
      return;
    }
    if (isLikelyAggregateTotalHeader(n)) {
      aggregate.push(h);
      return;
    }
    if (monthlyColumnHint(h)) {
      monthly.push(h);
      return;
    }
    ignored.push(h);
  });
  return { monthly, aggregate, ignored };
}

/** Columna de total preferida para validar suma de meses (misma fila). */
function pickAggregateColumnForRowValidation(aggregateHeaders) {
  if (!aggregateHeaders?.length) return null;
  let best = aggregateHeaders[0];
  let bestScore = -1;
  aggregateHeaders.forEach((h) => {
    const n = normalizeHeader(h);
    let s = 0;
    if (/total\s*general|grand\s*total/.test(n)) s += 4;
    else if (n === "total" || /^total\s/.test(n)) s += 2;
    if (/acumulado|ytd|summary/.test(n)) s += 1;
    if (s > bestScore) {
      bestScore = s;
      best = h;
    }
  });
  return best;
}

function findTerritoryColumn(rows) {
  if (!rows.length) return "";
  const keys = Object.keys(rows[0] || {});
  const map = new Map(keys.map((k) => [normalizeHeader(k), k]));
  for (const c of TERRITORY_SALES_REGION_COLUMNS) {
    const n = normalizeHeader(c);
    if (map.has(n)) return map.get(n);
  }
  let best = "";
  let bestScore = 0;
  keys.forEach((k) => {
    const n = normalizeHeader(k);
    let s = 0;
    if (n.includes("territorio")) s += 2;
    if (n.includes("region")) s += 2;
    if (s > bestScore) {
      bestScore = s;
      best = k;
    }
  });
  return bestScore >= 2 ? best : "";
}

function findTerritoryClientColumn(rows) {
  const byName = findColumnKey(rows, TERRITORY_SALES_CLIENT_COLUMNS);
  if (byName) return byName;
  return inferClientColumnByNumericPrefix(rows, 500);
}

function parseTerritoryClienteCell(cell) {
  const diagnosed = diagnoseTerritoryClienteCell(cell);
  if (!diagnosed.ok) return null;
  return { customer_code: diagnosed.customer_code, customer_name: diagnosed.customer_name };
}

function inferClientColumnByNumericPrefix(rows, maxRows) {
  if (!rows.length) return "";
  const keys = Object.keys(rows[0] || {});
  const counts = new Map();
  const limit = Math.min(rows.length, maxRows);

  for (let i = 0; i < limit; i += 1) {
    const row = rows[i];
    keys.forEach((k) => {
      if (parseTerritoryClienteCell(row[k])) {
        counts.set(k, (counts.get(k) || 0) + 1);
      }
    });
  }

  let bestKey = "";
  let best = 0;
  counts.forEach((v, k) => {
    if (v > best) {
      best = v;
      bestKey = k;
    }
  });
  return best > 0 ? bestKey : "";
}

/**
 * Normaliza filas del archivo de ventas por territorio.
 * @param {object[]} rows
 */
export function normalizeTerritorySalesFile(rows) {
  const colTerr = findTerritoryColumn(rows);
  const colClient = findTerritoryClientColumn(rows);
  if (!colClient) {
    return {
      rows: [],
      userMessage:
        "No se detectó una columna de cliente tipo «Representante / Cliente» (o equivalente). Sin ella no se puede obtener `customer_code` en ventas por territorio.",
      excludedCount: rows.length,
      excludedRows: rows.slice(0, 20).map((row) => ({
        excel_row: row.__excel_row || "",
        customer_value: "",
        reason: "formato no reconocido",
      })),
      clientColumn: "— (no detectada)",
      territoryColumn: colTerr || "",
      columnAnalysis: null,
    };
  }
  const headers = collectHeaders(rows);
  const { monthly, aggregate, ignored } = partitionTerritorySalesColumns(headers, colClient, colTerr);
  const monthlyOrdered = monthly.slice().sort((a, b) => headers.indexOf(a) - headers.indexOf(b));
  const grandCol = pickAggregateColumnForRowValidation(aggregate);
  const salesFromMonthlyOnly = monthlyOrdered.length > 0;
  const TOLERANCE_VS_TOTAL = 1.0;

  const out = [];
  const excludedRows = [];
  let rowTotalMismatchCount = 0;

  rows.forEach((row) => {
    const cell = row[colClient];
    const diagnosis = diagnoseTerritoryClienteCell(cell);
    if (!diagnosis.ok) {
      if (excludedRows.length < 20) {
        excludedRows.push({
          excel_row: row.__excel_row || "",
          customer_value: trimStr(valueToString(cell)),
          reason: diagnosis.reason,
        });
      }
      return;
    }

    let monthly_sum = 0;
    monthlyOrdered.forEach((h) => {
      monthly_sum += parseLocaleNumber(row[h]);
    });

    let sales_amount = monthly_sum;
    if (!salesFromMonthlyOnly && aggregate.length > 0) {
      sales_amount = 0;
      aggregate.forEach((h) => {
        sales_amount += parseLocaleNumber(row[h]);
      });
    }

    let aggregate_total = null;
    if (grandCol != null && aggregate.length > 0) {
      aggregate_total = parseLocaleNumber(row[grandCol]);
    }

    const monthly_validation_warnings = [];
    if (
      salesFromMonthlyOnly &&
      grandCol != null &&
      aggregate_total != null &&
      Number.isFinite(aggregate_total) &&
      aggregate.length > 0
    ) {
      if (Math.abs(monthly_sum - aggregate_total) > TOLERANCE_VS_TOTAL) {
        monthly_validation_warnings.push(
          "La suma de columnas mensuales no coincide con Total general para esta fila.",
        );
        rowTotalMismatchCount += 1;
      }
    }

    const sales_calculation = {
      monthly_columns_used: monthlyOrdered.slice(),
      aggregate_columns_excluded: aggregate.slice(),
      ignored_columns: ignored.slice(),
      monthly_sum,
      aggregate_total,
      difference_vs_aggregate:
        salesFromMonthlyOnly && aggregate_total != null && Number.isFinite(aggregate_total)
          ? monthly_sum - aggregate_total
          : null,
      aggregate_column_used_for_validation: grandCol || null,
      monthly_validation_warnings,
      computed_from_monthly_only: salesFromMonthlyOnly,
      aggregate_fallback_used: !salesFromMonthlyOnly && aggregate.length > 0,
    };

    out.push({
      customer_code: diagnosis.customer_code,
      customer_name: diagnosis.customer_name,
      territory: colTerr ? trimStr(valueToString(row[colTerr])) : "",
      sales_amount,
      sales_computed_from_aggregate_columns_only: !salesFromMonthlyOnly && aggregate.length > 0,
      sales_calculation,
      _source: row,
    });
  });

  const validationNotes = [];
  if (rowTotalMismatchCount > 0) {
    validationNotes.push(
      `En ${rowTotalMismatchCount} fila(s) la suma de columnas mensuales difiere del total de referencia (tolerancia ${TOLERANCE_VS_TOTAL}).`,
    );
  }

  const columnAnalysis = {
    monthlyColumns: monthlyOrdered.slice(),
    aggregateColumns: aggregate.slice(),
    ignoredColumns: ignored.slice(),
    firstMonthlyColumn: monthlyOrdered[0] || null,
    monthlyCount: monthlyOrdered.length,
    detailColumnsUsed: monthlyOrdered.slice(),
    aggregateColumnsExcludedFromSum: salesFromMonthlyOnly ? aggregate.slice() : [],
    aggregateColumnsUsedAsPrimary: !salesFromMonthlyOnly ? aggregate.slice() : [],
    salesComputedFromDetailOnly: salesFromMonthlyOnly,
    aggregateExcludedFromSum: salesFromMonthlyOnly && aggregate.length > 0,
    aggregateColumnForValidation: grandCol || null,
    rowTotalMismatchCount,
    validationNotes,
  };

  return {
    rows: out,
    excludedCount: Math.max(0, rows.length - out.length),
    excludedRows,
    clientColumn: colClient,
    territoryColumn: colTerr || "",
    columnAnalysis,
  };
}
