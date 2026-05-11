/**
 * Metadatos de contexto por hoja (año, moneda, periodo) a partir de contenido y nombre.
 * El nombre de la hoja solo apoya; no es criterio de inclusión/exclusión.
 */
import { normalizeHeader } from "./utils.js";

/**
 * @param {string} sheetName
 * @param {string[]} physicalHeaderLabels - nombres de columna como en Excel
 */
export function inferYearFromContext(sheetName, physicalHeaderLabels) {
  const re = /\b(20\d{2})\b/;
  const fromName = String(sheetName || "").match(re);
  if (fromName) return Number(fromName[1]);
  for (const label of physicalHeaderLabels || []) {
    const m = String(label).match(re);
    if (m) return Number(m[1]);
  }
  return null;
}

/**
 * @param {Set<string>} normalizedHeaders
 */
export function inferCurrencyHint(sheetName, normalizedHeaders) {
  const blob =
    String(sheetName || "") +
    " " +
    Array.from(normalizedHeaders || [])
      .join(" ")
      .toLowerCase();
  if (/\busd\b|dolares|dólares|dollar|us\$/.test(blob)) return "USD";
  if (/\bmxn\b|peso mexicano|pesos mx/.test(blob)) return "MXN";
  if (/\bml\b|moneda local|loc\.?$/i.test(blob)) return "ML";
  return null;
}

/** Texto libre para periodo (p. ej. nombre de hoja) si no hay modelo formal. */
export function inferPeriodHint(sheetName) {
  const t = String(sheetName || "").trim();
  return t || null;
}

export function headerSetFromPhysicalLabels(labels) {
  return new Set((labels || []).map((h) => normalizeHeader(h)));
}
