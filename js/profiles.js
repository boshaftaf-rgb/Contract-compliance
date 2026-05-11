/**
 * Detección de perfil de archivo: `margins`, `contracts`, `territory_sales` o `unknown`.
 *
 * Orden de decisión:
 * 1) Puntuación por conjunto de columnas esperadas en cada hoja (máxima confianza).
 * 2) Si no hay ganador, pistas desde el nombre del archivo (respaldo).
 * 3) Si hay empate entre perfiles, el nombre del archivo desempata cuando coincide.
 */
import { sheetToAoa, findTerritoryHeaderRow, sheetToJsonRows } from "./excelReader.js";
import {
  collectHeaders,
  normalizeHeader,
  yearMonthHint,
  salesAmountHint,
  isTerritoryRegionHeader,
  countKeywordHits,
} from "./utils.js";

function maxBy(list, fn) {
  let best = null;
  let bestScore = -1;
  list.forEach((item) => {
    const v = fn(item);
    const s = v && typeof v.score === "number" ? v.score : 0;
    if (s > bestScore) {
      bestScore = s;
      best = item;
    }
  });
  return { item: best || list[0], score: Math.max(bestScore, 0) };
}

function pickWinner(contenders) {
  if (!contenders.length) return null;
  return contenders.slice().sort((a, b) => b.score - a.score)[0];
}

function buildReason(profile, candidates, sheetName) {
  const row = candidates.find((c) => c.sheetName === sheetName);
  if (!row) return "";
  if (profile === "margins") return row.margins.detail;
  if (profile === "contracts") return row.contracts.detail;
  if (profile === "territory_sales") return row.territory.detail;
  return "";
}

function fallbackSheet(candidates, profile) {
  if (!candidates.length || profile === "unknown") return "";
  const scored = candidates
    .map((c) => {
      let s = 0;
      if (profile === "margins") s = c.margins.score;
      if (profile === "contracts") s = c.contracts.score;
      if (profile === "territory_sales") s = c.territory.score;
      return { sheetName: c.sheetName, s };
    })
    .sort((a, b) => b.s - a.s);
  return scored[0] ? scored[0].sheetName : "";
}

export const MARGINS_REQUIRED_HEADERS = [
  "cod. pagador",
  "nombre pagad.",
  "valor facturado loc",
  "margen bruto loc",
  "margen / ventas loc",
];

/** Señales opcionales (no bastan solas, suben confianza al comparar hojas). */
export const MARGINS_OPTIONAL_HEADERS = ["costos", "costo", "coste"];

function scoreMarginsHeaders(headerSet) {
  const required = MARGINS_REQUIRED_HEADERS;
  const matched = required.filter((h) => headerSet.has(h));
  const score = matched.length === required.length ? 100 : 0;
  return {
    score,
    detail: `margins: ${matched.length}/${required.length} columnas clave`,
  };
}

/** Puntuación 0–100 por cobertura de columnas (contenido, no nombre de hoja). */
export function scoreMarginsHeadersWeighted(headerSet) {
  const required = MARGINS_REQUIRED_HEADERS;
  const matched = required.filter((h) => headerSet.has(h));
  let score = Math.round((100 * matched.length) / required.length);
  const optionalHits = MARGINS_OPTIONAL_HEADERS.filter((h) => headerSet.has(h)).length;
  if (optionalHits) score = Math.min(100, score + 5);
  return {
    score,
    matchedKeys: matched,
    detail: `margins: ${matched.length}/${required.length} columnas clave${optionalHits ? ` · costos: sí` : ""}`,
  };
}

export const CONTRACTS_REQUIRED_HEADERS = [
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

function scoreContractsHeaders(headerSet) {
  const required = CONTRACTS_REQUIRED_HEADERS;
  const matched = required.filter((h) => headerSet.has(h));
  const score = matched.length === required.length ? 100 : 0;
  return {
    score,
    detail: `contracts: ${matched.length}/${required.length} columnas clave`,
  };
}

export function scoreContractsHeadersWeighted(headerSet) {
  const required = CONTRACTS_REQUIRED_HEADERS;
  const matched = required.filter((h) => headerSet.has(h));
  const score = Math.round((100 * matched.length) / required.length);
  return {
    score,
    matchedKeys: matched,
    detail: `contracts: ${matched.length}/${required.length} columnas clave`,
  };
}

function scoreTerritorySalesHeaders(headerSet) {
  const headers = Array.from(headerSet);
  const hasRep = headers.some((h) => h.includes("representante") && h.includes("cliente"));
  const hasTerritory = headers.some((h) => isTerritoryRegionHeader(h));
  const hasTimeOrAmount = headers.some((h) => yearMonthHint(h) || salesAmountHint(h));
  const score = hasRep && hasTerritory && hasTimeOrAmount ? 100 : 0;
  return {
    score,
    detail: `territory_sales: rep/cliente ${hasRep ? "sí" : "no"} · terr/región ${
      hasTerritory ? "sí" : "no"
    } · mes/año/importe ${hasTimeOrAmount ? "sí" : "no"}`,
  };
}

export function scoreTerritorySalesFromSheet(aoa, headerSetFallback) {
  const pick = findTerritoryHeaderRow(aoa, 20);
  const normHeaders = pick.headers.map((h) => normalizeHeader(h));
  const headerBlob = normHeaders.join(" | ");

  const hasClientHint =
    headerBlob.includes("representante") ||
    headerBlob.includes("cliente") ||
    normHeaders.some((h) => h === "cliente" || h === "representante");
  const hasTerritoryHint =
    normHeaders.some((h) => isTerritoryRegionHeader(h)) ||
    headerBlob.includes("territorio") ||
    headerBlob.includes("region");
  const hasMonthOrTotal =
    normHeaders.some((h) => yearMonthHint(h) || salesAmountHint(h)) || headerBlob.includes("total");

  const strict = hasClientHint && hasTerritoryHint && hasMonthOrTotal;
  if (strict) {
    return {
      score: 100,
      detail: `territory_sales (fila encabezado ~${pick.index + 1}, score ${pick.score}): cliente ${hasClientHint ? "sí" : "no"} · terr ${hasTerritoryHint ? "sí" : "no"} · mes/total/importe ${hasMonthOrTotal ? "sí" : "no"}`,
    };
  }

  const loose = scoreTerritorySalesHeaders(headerSetFallback);
  if (loose.score) return loose;
  return {
    score: 0,
    detail: `territory_sales: sin encabezado claro en primeras 20 filas (mejor score ${pick.score}).`,
  };
}

export function detectProfileFromFileName(fileName) {
  const base = String(fileName || "")
    .replace(/\.[^.]+$/, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();

  const scores = {
    margins: countKeywordHits(base, ["mb", "margen", "margin", "facturacion", "billing"]),
    contracts: countKeywordHits(base, ["contract", "contrato", "compliance", "cumplimiento"]),
    territory_sales: countKeywordHits(base, ["ventas", "sales", "territorio", "territory"]),
  };

  const sorted = Object.keys(scores).sort((a, b) => scores[b] - scores[a]);
  const best = sorted[0];
  if (!scores[best]) return { profile: "unknown", score: 0, reason: "" };
  return {
    profile: best,
    score: scores[best],
    reason: `Nombre sugiere ${best} (${scores[best]} coincidencias).`,
  };
}

/**
 * Detecta el perfil del libro (margins, contracts, territory_sales) o "unknown".
 * Primero se puntúan columnas por hoja; si no hay ganador, se usa el nombre del archivo.
 * @param {object} workbook - Libro SheetJS
 * @param {string} fileName - Nombre del archivo (solo respaldo / desempate)
 */
export function detectFileProfile(workbook, fileName) {
  if (!workbook || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
    return {
      profile: "unknown",
      method: "unknown",
      reason: "El libro no contiene hojas o no se pudo leer.",
      sheetName: "",
      scores: { margins: 0, contracts: 0, territory_sales: 0 },
    };
  }

  const sheetNames = workbook.SheetNames || [];
  const candidates = [];

  sheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetToJsonRows(sheet);
    const headers = collectHeaders(rows);
    const headerSet = new Set(headers.map((h) => normalizeHeader(h)));

    const margins = scoreMarginsHeaders(headerSet);
    const contracts = scoreContractsHeaders(headerSet);
    const aoa = sheetToAoa(sheet);
    const territory = scoreTerritorySalesFromSheet(aoa, headerSet);

    candidates.push({ sheetName, margins, contracts, territory });
  });

  const bestMargins = maxBy(candidates, (c) => c.margins.score);
  const bestContracts = maxBy(candidates, (c) => c.contracts.score);
  const bestTerritory = maxBy(candidates, (c) => c.territory.score);

  const contenders = [
    { profile: "margins", sheet: bestMargins.item.sheetName, score: bestMargins.score },
    { profile: "contracts", sheet: bestContracts.item.sheetName, score: bestContracts.score },
    {
      profile: "territory_sales",
      sheet: bestTerritory.item.sheetName,
      score: bestTerritory.score,
    },
  ].filter((c) => c.score > 0);

  let winner = pickWinner(contenders);
  let method = "columns";
  let reason = winner ? buildReason(winner.profile, candidates, winner.sheet) : "";

  if (!winner) {
    const fn = detectProfileFromFileName(fileName);
    if (fn.profile !== "unknown") {
      winner = { profile: fn.profile, sheet: "", score: fn.score };
      method = "filename";
      reason = fn.reason;
    }
  } else {
    const ties = contenders.filter((c) => c.score === winner.score);
    if (ties.length > 1) {
      const fn = detectProfileFromFileName(fileName);
      const pick = ties.find((t) => t.profile === fn.profile);
      if (pick && fn.profile !== "unknown") {
        winner = pick;
        method = "columns+filename";
        reason = `Empate por columnas; desempate por nombre: ${fn.reason}`;
      }
    }
  }

  const profile = winner ? winner.profile : "unknown";
  return {
    profile,
    method: profile === "unknown" ? "unknown" : method,
    reason,
    sheetName:
      winner && winner.sheet
        ? winner.sheet
        : fallbackSheet(candidates, profile),
    scores: {
      margins: bestMargins.score,
      contracts: bestContracts.score,
      territory_sales: bestTerritory.score,
    },
  };
}
