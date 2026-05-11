/**
 * Descubre columnas reales del Excel (raw) para metadata y futuros filtros.
 * No infiere valores de negocio: solo clasifica nombres de columna por palabras clave.
 */
import { normalizeHeader } from "./utils.js";
import { FILE_SLOT } from "./state.js";

/**
 * Palabras clave por categoría (encabezado ya normalizado).
 * `vendedor` solo con columnas explícitas de rol comercial; «Representante / Cliente» va a `cliente_mixto`.
 */
export const POSSIBLE_BUSINESS_BUCKETS = {
  vendedor: [
    "vendedor",
    "sales rep",
    "sales representative",
    "ejecutivo",
    "representante comercial",
    "agente",
    "ejecutivo de cuenta",
  ],
  cliente_mixto: [],
  sector: ["sector", "tipo cliente", "tipo de cliente"],
  licitacion: ["licitacion", "licitación", "concurso", "adjudicacion", "adjudicación", "tender", "bid", "rfp"],
  publicoPrivado: ["publico", "público", "privado", "gobierno"],
  territorio: ["territorio", "region", "región"],
};

function isClienteMixtoColumn(nh) {
  const hasRep = nh.includes("representante");
  const hasCli = nh.includes("cliente");
  if (hasRep && hasCli) return true;
  if (nh.includes("representante / cliente") || nh.includes("rep / cliente")) return true;
  return false;
}

function bucketsForNormalizedHeader(nh) {
  const hits = [];
  if (isClienteMixtoColumn(nh)) {
    hits.push("cliente_mixto");
    return hits;
  }
  Object.entries(POSSIBLE_BUSINESS_BUCKETS).forEach(([bucket, keywords]) => {
    if (bucket === "cliente_mixto") return;
    if (keywords.some((kw) => nh.includes(kw))) hits.push(bucket);
  });
  return hits;
}

/**
 * @param {object[]} rowsRaw
 * @param {number} maxRows
 * @returns {string[]} Nombres físicos de columnas (como en el Excel)
 */
export function collectRawColumnLabels(rowsRaw, maxRows = 40) {
  const set = new Set();
  (rowsRaw || []).slice(0, maxRows).forEach((row) => {
    Object.keys(row || {}).forEach((key) => {
      if (key === "__excel_row" || key === "_source") return;
      set.add(key);
    });
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "es"));
}

/**
 * Campos de negocio candidatos detectados en columnas reales (por slot).
 * @returns {Record<string, Record<string, string[]>>} slot → bucket → [nombres columna]
 */
export function discoverRelatedColumnHintsBySlot(rawData) {
  const slots = [FILE_SLOT.TERRITORY_SALES, FILE_SLOT.MARGINS, FILE_SLOT.CONTRACTS];
  const out = {};
  slots.forEach((slot) => {
    const rows = rawData[slot] || [];
    const labels = collectRawColumnLabels(rows, 120);
    const byBucket = {
      vendedor: [],
      cliente_mixto: [],
      sector: [],
      licitacion: [],
      publicoPrivado: [],
      territorio: [],
    };
    const seen = {
      vendedor: new Set(),
      cliente_mixto: new Set(),
      sector: new Set(),
      licitacion: new Set(),
      publicoPrivado: new Set(),
      territorio: new Set(),
    };
    labels.forEach((label) => {
      const nh = normalizeHeader(label);
      bucketsForNormalizedHeader(nh).forEach((bucket) => {
        if (!seen[bucket].has(label)) {
          seen[bucket].add(label);
          byBucket[bucket].push(label);
        }
      });
    });
    out[slot] = byBucket;
  });
  return out;
}

/**
 * Union global de columnas candidatas (mismo formato que pide analysisData.possibleBusinessFields).
 */
export function mergePossibleBusinessFields(relatedBySlot) {
  const merged = {
    vendedor: [],
    cliente_mixto: [],
    sector: [],
    licitacion: [],
    publicoPrivado: [],
    territorio: [],
  };
  const seen = {
    vendedor: new Set(),
    cliente_mixto: new Set(),
    sector: new Set(),
    licitacion: new Set(),
    publicoPrivado: new Set(),
    territorio: new Set(),
  };
  Object.values(relatedBySlot || {}).forEach((bySlot) => {
    Object.entries(bySlot).forEach(([bucket, labels]) => {
      (labels || []).forEach((label) => {
        if (!seen[bucket].has(label)) {
          seen[bucket].add(label);
          merged[bucket].push(label);
        }
      });
    });
  });
  return merged;
}

/**
 * @param {object} rawData - state.rawData
 */
export function discoverAvailableFieldsBySlot(rawData) {
  return {
    territorySales: collectRawColumnLabels(rawData.territorySales, 200),
    margins: collectRawColumnLabels(rawData.margins, 200),
    contracts: collectRawColumnLabels(rawData.contracts, 200),
  };
}
