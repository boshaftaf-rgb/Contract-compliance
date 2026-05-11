/**
 * Lectura de libros y hojas XLSX (SheetJS en el navegador).
 * Solo estructura tabular (AOA / filas); no interpreta perfiles ni `customer_code`.
 */
import {
  normalizeHeader,
  trimStr,
  valueToString,
  yearMonthHint,
  salesAmountHint,
  isTerritoryRegionHeader,
} from "./utils.js";

/**
 * Límites y validación básica de seguridad al leer XLSX en el navegador.
 * Mitigan DoS por archivos enormes y rechazan archivos que no son ZIP/XLSX
 * (la extensión `.xlsx` y el `accept` del input no son confiables).
 */
const MAX_XLSX_BYTES = 50 * 1024 * 1024;
const XLSX_MAGIC = [0x50, 0x4b, 0x03, 0x04];

export function readWorkbook(file) {
  const XLSX = globalThis.XLSX;
  if (!file || typeof file.arrayBuffer !== "function") {
    return Promise.reject(new Error("Archivo no válido."));
  }
  if (typeof file.size === "number" && file.size > MAX_XLSX_BYTES) {
    return Promise.reject(
      new Error("El archivo es demasiado grande. El límite permitido es de 50 MB.")
    );
  }
  return file.arrayBuffer().then((buf) => {
    const head = new Uint8Array(buf, 0, Math.min(XLSX_MAGIC.length, buf.byteLength));
    const validMagic =
      head.length === XLSX_MAGIC.length && XLSX_MAGIC.every((b, i) => head[i] === b);
    if (!validMagic) {
      throw new Error("El archivo no parece ser un .xlsx válido.");
    }
    return XLSX.read(buf, { type: "array", cellDates: true });
  });
}

/**
 * Lee un .xlsx sin propagar excepción (evita romper toda la app por un solo archivo dañado).
 * @returns {Promise<{ ok: true, workbook: object } | { ok: false, workbook: null, error: string }>}
 */
export async function readWorkbookFromFile(file) {
  try {
    const workbook = await readWorkbook(file);
    return { ok: true, workbook };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, workbook: null, error: message };
  }
}

export function sheetToAoa(sheet) {
  const XLSX = globalThis.XLSX;
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: true,
  });
}

export function sheetToJsonRows(sheet) {
  const XLSX = globalThis.XLSX;
  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
    blankrows: false,
  });
}

const MONTH_NAMES_ES = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];

function scoreTerritoryHeaderCandidateRow(row) {
  let score = 0;
  const cells = row.map((c) => normalizeHeader(valueToString(c)));
  cells.forEach((t) => {
    if (!t) return;
    if (t.includes("representante")) score += 3;
    if (t.includes("cliente")) score += 2;
    if (isTerritoryRegionHeader(t)) score += 3;
    if (t.includes("territorio")) score += 2;
    if (t.includes("region")) score += 1;
    if (t.includes("total")) score += 2;
    if (yearMonthHint(t) || salesAmountHint(t)) score += 2;
    if (MONTH_NAMES_ES.some((m) => t.includes(m))) score += 2;
  });
  return score;
}

export function buildHeaderLabelsFromRow(aoa, headerRowIndex) {
  const row = Array.isArray(aoa[headerRowIndex]) ? aoa[headerRowIndex] : [];
  const used = new Map();
  return row.map((cell, idx) => {
    const label = trimStr(valueToString(cell));
    const base = label || `__col_${idx + 1}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    return count ? `${base}__${count + 1}` : base;
  });
}

export function findTerritoryHeaderRow(aoa, maxScan) {
  const limit = Math.min(aoa.length, maxScan);
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    const s = scoreTerritoryHeaderCandidateRow(row);
    if (s > bestScore) {
      bestScore = s;
      bestIdx = i;
    }
  }

  const headers = buildHeaderLabelsFromRow(aoa, bestIdx);
  return { index: bestIdx, score: bestScore, headers };
}

export function rowsFromAoa(aoa, headerRowIndex, headers) {
  const rows = [];
  for (let r = headerRowIndex + 1; r < aoa.length; r += 1) {
    const line = Array.isArray(aoa[r]) ? aoa[r] : [];
    const obj = {};
    let has = false;
    headers.forEach((h, j) => {
      const v = line[j] !== undefined && line[j] !== null ? line[j] : "";
      obj[h] = v;
      if (trimStr(valueToString(v)) !== "") has = true;
    });
    if (has) {
      obj.__excel_row = r + 1;
      rows.push(obj);
    }
  }
  return rows;
}
