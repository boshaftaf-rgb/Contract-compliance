/**
 * Columnas esperadas por perfil (encabezados normalizados, como en detección de perfil).
 * Sirve para auditar qué falta en el Excel sin inferir datos.
 */
import {
  normalizeHeader,
  yearMonthHint,
  salesAmountHint,
  isTerritoryRegionHeader,
} from "./utils.js";

const MARGINS_REQUIRED_NORMALIZED = [
  "cod. pagador",
  "nombre pagad.",
  "valor facturado loc",
  "margen bruto loc",
  "margen / ventas loc",
];

const CONTRACTS_REQUIRED_NORMALIZED = [
  "contrato",
  "solicitante cliente",
  "nombre del solicitante",
  "inicio",
  "fin",
  "valor neto contrato",
  "facturación esperada",
  "facturado neto",
  "cumplimiento facturación%",
];

function normalizedKeySetFromRawRows(rowsRaw, maxRows = 80) {
  const set = new Set();
  (rowsRaw || []).slice(0, maxRows).forEach((row) => {
    Object.keys(row || {}).forEach((key) => {
      if (key === "__excel_row" || key === "_source") return;
      set.add(normalizeHeader(key));
    });
  });
  return set;
}

/**
 * Lista encabezados requeridos que no aparecen en las filas crudas (muestra).
 * @param {string} profile - margins | contracts | territory_sales | unknown
 * @param {object[]} rowsRaw
 * @returns {string[]} Etiquetas legibles (normalizadas) faltantes
 */
export function listMissingExpectedColumns(profile, rowsRaw) {
  if (!rowsRaw || !rowsRaw.length) {
    if (profile === "margins") return [...MARGINS_REQUIRED_NORMALIZED];
    if (profile === "contracts") return [...CONTRACTS_REQUIRED_NORMALIZED];
    if (profile === "territory_sales") {
      return ["representante/cliente", "territorio o región", "columnas de periodo o importe"];
    }
    return [];
  }

  const keys = normalizedKeySetFromRawRows(rowsRaw);

  if (profile === "margins") {
    return MARGINS_REQUIRED_NORMALIZED.filter((h) => !keys.has(h));
  }
  if (profile === "contracts") {
    return CONTRACTS_REQUIRED_NORMALIZED.filter((h) => !keys.has(h));
  }
  if (profile === "territory_sales") {
    const missing = [];
    const arr = Array.from(keys);
    const hasRepClient = arr.some(
      (h) => (h.includes("representante") && h.includes("cliente")) || h === "cliente" || h === "representante"
    );
    const hasTerr = arr.some((h) => isTerritoryRegionHeader(h));
    const hasMonthOrAmount = arr.some((h) => yearMonthHint(h) || salesAmountHint(h));
    if (!hasRepClient) missing.push("representante/cliente (columna cliente)");
    if (!hasTerr) missing.push("territorio / región");
    if (!hasMonthOrAmount) missing.push("periodo o importe (mes/año/ventas)");
    return missing;
  }
  return [];
}
