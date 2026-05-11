/**
 * Analiza todas las hojas de un libro por contenido (columnas / patrones).
 * Una misma hoja puede figurar en más de un perfil si las señales superan el umbral.
 * La hoja principal del slot esperado se elige por filas normalizadas (sin sumar traslapes).
 */
import {
  sheetToAoa,
  sheetToJsonRows,
  rowsFromAoa,
  findTerritoryHeaderRow,
  buildHeaderLabelsFromRow,
} from "./excelReader.js";
import { collectHeaders, normalizeHeader, yearMonthHint, salesAmountHint } from "./utils.js";
import {
  scoreMarginsHeadersWeighted,
  scoreContractsHeadersWeighted,
  scoreTerritorySalesFromSheet,
  MARGINS_REQUIRED_HEADERS,
  CONTRACTS_REQUIRED_HEADERS,
  detectProfileFromFileName,
} from "./profiles.js";
import {
  normalizeTerritorySalesFile,
  normalizeMarginsFile,
  normalizeContractsFile,
} from "./normalizers.js";
import {
  inferYearFromContext,
  inferCurrencyHint,
  inferPeriodHint,
  headerSetFromPhysicalLabels,
} from "./sheetContext.js";

const MIN_LIST_PROFILE = 60;
const MIN_PRIMARY = 68;

function territoryPartialScore(aoa, headerSetFallback) {
  const pick = findTerritoryHeaderRow(aoa, 20);
  const normHeaders = pick.headers.map((h) => normalizeHeader(h));
  const headerBlob = normHeaders.join(" | ");
  const hasClient =
    headerBlob.includes("representante") ||
    headerBlob.includes("cliente") ||
    normHeaders.some((h) => h === "cliente" || h === "representante");
  const hasTerr =
    normHeaders.some((h) => {
      const n = normalizeHeader(h);
      return n.includes("territorio") || n.includes("region") || n.includes("zona");
    }) ||
    headerBlob.includes("territorio") ||
    headerBlob.includes("region");
  const hasMonth =
    normHeaders.some((h) => yearMonthHint(normalizeHeader(h)) || salesAmountHint(normalizeHeader(h))) ||
    headerBlob.includes("total");
  return Math.round(hasClient * 34 + hasTerr * 33 + hasMonth * 33);
}

function physicalHeadersFromJsonRows(rows) {
  return collectHeaders(rows.slice(0, 80));
}

function matchedPhysicalColumns(headerSetNormalized, requiredNormList, physicalLabels) {
  const out = [];
  (physicalLabels || []).forEach((phys) => {
    const n = normalizeHeader(phys);
    if (requiredNormList.includes(n)) out.push(phys);
  });
  return out;
}

function buildRowsForProfile(profile, aoa) {
  if (profile === "territory_sales") {
    const headerPick = findTerritoryHeaderRow(aoa, 20);
    const rowsRaw = rowsFromAoa(aoa, headerPick.index, headerPick.headers);
    return { rowsRaw, headerPick };
  }
  const headers = buildHeaderLabelsFromRow(aoa, 0);
  const rowsRaw = rowsFromAoa(aoa, 0, headers);
  return { rowsRaw, headerPick: { index: 0, score: 0, headers } };
}

function normalizeForProfile(profile, rowsRaw, sheetName, headerPick) {
  if (profile === "territory_sales") {
    const norm = normalizeTerritorySalesFile(rowsRaw);
    return {
      normalizedRows: norm.rows,
      territoryMeta: {
        headerRow1Based: headerPick.index + 1,
        headerScore: headerPick.score,
        clientColumn: norm.clientColumn,
        territoryColumn: norm.territoryColumn || "—",
        rawRowCount: rowsRaw.length,
        normalizedCount: norm.rows.length,
        excludedCount: norm.excludedCount,
        excludedRows: norm.excludedRows.slice(0, 20),
        sampleClients: norm.rows.slice(0, 5),
        columnAnalysis: norm.columnAnalysis || null,
      },
      diagnosticMeta: {
        profile: "territory_sales",
        sheetName,
        keyColumn: norm.clientColumn,
        rawRowCount: rowsRaw.length,
        normalizedCount: norm.rows.length,
        excludedCount: norm.excludedCount,
        excludedRows: norm.excludedRows.slice(0, 20),
      },
      userMessages: [
        ...(norm.userMessage ? [norm.userMessage] : []),
        ...((norm.columnAnalysis && norm.columnAnalysis.validationNotes) || []),
      ],
    };
  }
  if (profile === "margins") {
    const norm = normalizeMarginsFile(rowsRaw);
    return {
      normalizedRows: norm.rows,
      territoryMeta: null,
      diagnosticMeta: {
        profile: "margins",
        sheetName,
        keyColumn: norm.keyColumn,
        rawRowCount: rowsRaw.length,
        normalizedCount: norm.rows.length,
        excludedCount: norm.excludedCount,
        excludedRows: norm.excludedRows.slice(0, 20),
      },
      userMessages: [
        ...(norm.userMessage ? [norm.userMessage] : []),
        ...((norm.columnAnalysis && norm.columnAnalysis.validationNotes) || []),
      ],
    };
  }
  if (profile === "contracts") {
    const norm = normalizeContractsFile(rowsRaw);
    return {
      normalizedRows: norm.rows,
      territoryMeta: null,
      diagnosticMeta: {
        profile: "contracts",
        sheetName,
        keyColumn: norm.keyColumn,
        rawRowCount: rowsRaw.length,
        normalizedCount: norm.rows.length,
        excludedCount: norm.excludedCount,
        excludedRows: norm.excludedRows.slice(0, 20),
      },
      userMessages: [
        ...(norm.userMessage ? [norm.userMessage] : []),
        ...((norm.columnAnalysis && norm.columnAnalysis.validationNotes) || []),
      ],
    };
  }
  return {
    normalizedRows: [],
    territoryMeta: null,
    diagnosticMeta: null,
    userMessages: [],
  };
}

function makeEntry(
  profile,
  sheetName,
  confidenceScore,
  reason,
  physHeaders,
  headerSetNorm,
  pack,
  marginW,
  contractW,
  rowsRaw,
) {
  const requiredNorm =
    profile === "margins"
      ? MARGINS_REQUIRED_HEADERS
      : profile === "contracts"
        ? CONTRACTS_REQUIRED_HEADERS
        : [];

  const detectedColumns =
    profile === "territory_sales" && pack.territoryMeta
      ? [pack.territoryMeta.clientColumn, pack.territoryMeta.territoryColumn].filter(Boolean)
      : matchedPhysicalColumns(headerSetNorm, requiredNorm, physHeaders);

  const year = inferYearFromContext(sheetName, physHeaders);
  const currency = profile === "contracts" ? inferCurrencyHint(sheetName, headerSetNorm) : null;
  const period = inferPeriodHint(sheetName);

  const dm = pack.diagnosticMeta;
  return {
    source_sheet: sheetName,
    sheetName,
    profile,
    confidenceScore,
    detectedColumns,
    normalizedRows: pack.normalizedRows,
    normalizedRowCount: pack.normalizedRows.length,
    rowsRaw: rowsRaw || [],
    rawRowCount: dm?.rawRowCount ?? (rowsRaw || []).length,
    territoryMeta: pack.territoryMeta,
    diagnosticMeta: pack.diagnosticMeta,
    excludedCount: dm?.excludedCount ?? 0,
    excludedRowsSample: (dm?.excludedRows || []).slice(0, 20),
    reason,
    year,
    period,
    currency,
    alerts: pack.userMessages || [],
    marginHeaderScore: marginW.score,
    contractHeaderScore: contractW.score,
  };
}

/**
 * @param {object} workbook
 * @param {string} expectedProfile - territory_sales | margins | contracts
 * @param {string} fileName
 */
export function analyzeWorkbookSheets(workbook, expectedProfile, fileName) {
  const sheetNames = workbook?.SheetNames || [];
  const territorySales = [];
  const margins = [];
  const contracts = [];

  sheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const jsonRows = sheetToJsonRows(sheet);
    const physHeaders = physicalHeadersFromJsonRows(jsonRows);
    const headerSetNorm = headerSetFromPhysicalLabels(physHeaders);

    const marginW = scoreMarginsHeadersWeighted(headerSetNorm);
    const contractW = scoreContractsHeadersWeighted(headerSetNorm);
    const aoa = sheetToAoa(sheet);
    const territoryBlock = scoreTerritorySalesFromSheet(aoa, headerSetNorm);
    let territoryScore = territoryBlock.score;
    if (territoryScore < 100) {
      territoryScore = Math.max(territoryScore, territoryPartialScore(aoa, headerSetNorm));
    }

    const territoryReason =
      territoryScore >= 100 ? territoryBlock.detail : `territorio (parcial): ${territoryScore}`;

    if (territoryScore >= MIN_LIST_PROFILE) {
      const { rowsRaw, headerPick } = buildRowsForProfile("territory_sales", aoa);
      const pack = normalizeForProfile("territory_sales", rowsRaw, sheetName, headerPick);
      territorySales.push(
        makeEntry(
          "territory_sales",
          sheetName,
          territoryScore,
          territoryReason,
          physHeaders,
          headerSetNorm,
          pack,
          marginW,
          contractW,
          rowsRaw,
        ),
      );
    }

    if (marginW.score >= MIN_LIST_PROFILE) {
      const { rowsRaw, headerPick } = buildRowsForProfile("margins", aoa);
      const pack = normalizeForProfile("margins", rowsRaw, sheetName, headerPick);
      margins.push(
        makeEntry(
          "margins",
          sheetName,
          marginW.score,
          marginW.detail,
          physHeaders,
          headerSetNorm,
          pack,
          marginW,
          contractW,
          rowsRaw,
        ),
      );
    }

    if (contractW.score >= MIN_LIST_PROFILE) {
      const { rowsRaw, headerPick } = buildRowsForProfile("contracts", aoa);
      const pack = normalizeForProfile("contracts", rowsRaw, sheetName, headerPick);
      contracts.push(
        makeEntry(
          "contracts",
          sheetName,
          contractW.score,
          contractW.detail,
          physHeaders,
          headerSetNorm,
          pack,
          marginW,
          contractW,
          rowsRaw,
        ),
      );
    }
  });

  const slotList =
    expectedProfile === "territory_sales"
      ? territorySales
      : expectedProfile === "margins"
        ? margins
        : contracts;
  const hasPrimary = slotList.some((h) => h.confidenceScore >= MIN_PRIMARY);

  const fn = detectProfileFromFileName(fileName);
  let fileLevelReason = "";
  if (!hasPrimary && fn.profile === expectedProfile) {
    fileLevelReason =
      "Ninguna hoja alcanzó el umbral de confianza; el nombre del archivo sugiere el tipo correcto. Revise encabezados.";
  } else if (!hasPrimary) {
    fileLevelReason = "No se encontró una hoja con columnas reconocibles para este tipo de archivo.";
  }

  return {
    hojasDetectadas: { territorySales, margins, contracts },
    fileLevelReason,
  };
}
