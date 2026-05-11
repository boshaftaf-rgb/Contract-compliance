/**
 * Utilidades compartidas (texto, números, fechas).
 * `escapeHtml` + `stringifyCell` en la capa de render evitan inyección HTML desde celdas Excel;
 * los datos no se ejecutan como código (no eval, no new Function).
 */
export function normalizeHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

export function trimStr(v) {
  return String(v ?? "").trim();
}

export function valueToString(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

export function normalizeCustomerCodeString(value) {
  let t = trimStr(valueToString(value));
  t = t.replace(/\s+/g, "");
  if (/^\d+\.0+$/.test(t)) t = t.replace(/\.0+$/, "");
  return t;
}

export function isSummaryToken(text) {
  const t = trimStr(text).toLowerCase();
  return (
    !t ||
    t === "-" ||
    t === "totals" ||
    t === "total" ||
    t === "subtotal" ||
    t === "grand total"
  );
}

export function isValidNumericCustomerCode(value) {
  if (isSummaryToken(value)) return false;
  const code = normalizeCustomerCodeString(value);
  return /^\d+$/.test(code);
}

export function hasAnyUsefulCell(row) {
  return Object.keys(row || {}).some((k) => {
    if (k === "__excel_row") return false;
    return trimStr(valueToString(row[k])) !== "";
  });
}

/** Coincidencias de palabras clave (p. ej. nombre de archivo) */
export function countKeywordHits(text, keywords) {
  return keywords.reduce((acc, kw) => (text.includes(kw) ? acc + 1 : acc), 0);
}

/** Columnas y filas */
export function collectHeaders(rows) {
  const set = new Set();
  rows.slice(0, 300).forEach((row) => {
    Object.keys(row || {}).forEach((k) => set.add(k));
  });
  return Array.from(set);
}

export function findColumnKey(rows, candidates) {
  if (!rows.length) return "";
  const map = new Map();
  Object.keys(rows[0] || {}).forEach((h) => map.set(normalizeHeader(h), h));
  for (const c of candidates) {
    const k = normalizeHeader(c);
    if (map.has(k)) return map.get(k);
  }
  return "";
}

/** Columna de territorio / región (encabezado ya normalizado). */
export function isTerritoryRegionHeader(n) {
  return n.includes("territorio") || n.includes("region") || n.includes("zona");
}

export function isLikelyIdColumn(n) {
  return (
    n.includes("cod") ||
    n.includes("cód") ||
    n.includes("id") ||
    n.includes("contrato") ||
    n.includes("factura") ||
    n.includes("pedido")
  );
}

/**
 * Encabezado de columna que representa mes / año (inglés y español, 2 o 4 cifras de año).
 * Incluye Jan-26, Ene-26, Enero 2026, etc. No debe usarse solo para sumar: combinar con exclusión de agregados.
 */
export function monthlyColumnHint(rawHeader) {
  const n = normalizeHeader(rawHeader);
  if (!n) return false;
  if (/\b(19|20)\d{2}\b/.test(n)) return true;
  if (/\bmes\b/.test(n) || /\banio\b|\baño\b|\byear\b/.test(n)) return true;
  if (
    /\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b/.test(
      n,
    )
  )
    return true;
  if (
    /\b(january|february|march|april|june|july|august|september|october|november|december)\b/.test(n)
  )
    return true;
  const abbrEn = "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec";
  const abbrEs = "ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic";
  if (new RegExp(`\\b(?:${abbrEn}|${abbrEs})\\b[\\s\\-_.\\/]+\\d{2,4}\\b`, "i").test(n)) return true;
  if (new RegExp(`\\b(?:${abbrEn}|${abbrEs})\\d{2,4}\\b`, "i").test(n)) return true;
  if (new RegExp(`\\b(?:${abbrEn}|${abbrEs})\\b(?=\\s*$)`, "i").test(n)) return true;
  const legacy = [
    "ene",
    "feb",
    "mar",
    "abr",
    "may",
    "jun",
    "jul",
    "ago",
    "sep",
    "oct",
    "nov",
    "dic",
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
    "jan",
    "feb",
    "apr",
    "aug",
    "sep",
    "oct",
    "nov",
    "dec",
  ];
  return legacy.some((m) => {
    const re = new RegExp(`\\b${m}\\b`, "i");
    return re.test(n);
  });
}

/** @deprecated Preferir monthlyColumnHint; se mantiene para compatibilidad en otros módulos. */
export function yearMonthHint(h) {
  return monthlyColumnHint(h);
}

export function salesAmountHint(h) {
  return (
    h.includes("venta") ||
    h.includes("ventas") ||
    h.includes("importe") ||
    h.includes("monto") ||
    h.includes("factur") ||
    h.includes("billing") ||
    h.includes("amount")
  );
}

/** Encabezado normalizado que parece columna de total / agregado (no sumar junto con meses o detalle). */
export function isLikelyAggregateTotalHeader(n) {
  const t = String(n || "").trim();
  if (!t) return false;
  if (/\bgrand\s*total\b|\btotal\s*general\b|\btotal\s+acumulado\b|\btotal\s+acumulada\b/.test(t)) return true;
  if (/\bacumulado\b|\bacumulada\b|\bsummary\b|\bsuma\s+total\b/.test(t)) return true;
  if (t === "total" || t.startsWith("total ") || t.endsWith(" total")) return true;
  if (/\btotal\b/.test(t) && (t.includes("general") || t.includes("global") || t.includes("ytd") || t.includes("acum"))) return true;
  return false;
}

/** Números y porcentajes */
export function parseLocaleNumber(value) {
  const s = trimStr(valueToString(value));
  if (!s || isSummaryToken(s)) return 0;
  let t = s.replace(/[^\d,.\-+%]/g, "");
  const isPct = t.includes("%");
  if (isPct) t = t.replace(/%/g, "");
  if (!t) return 0;

  const hasComma = t.includes(",");
  const hasDot = t.includes(".");
  if (hasComma && hasDot) {
    if (t.lastIndexOf(",") > t.lastIndexOf(".")) {
      t = t.replace(/\./g, "").replace(",", ".");
    } else {
      t = t.replace(/,/g, "");
    }
  } else if (hasComma) {
    const parts = t.split(",");
    if (parts.length === 2 && parts[1].length <= 2) t = parts[0] + "." + parts[1];
    else t = t.replace(/,/g, "");
  }

  const n = Number(t);
  if (!Number.isFinite(n)) return 0;
  return isPct ? n : n;
}

/** Número genérico (pocas fracciones; solo UI). */
export function formatNum(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 4 }).format(n);
}

/** Importes / montos: solo presentación (no alterar valores en datos). */
export function formatAmount(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
}

/** Conteos enteros en UI. */
export function formatIntegerDisplay(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(n));
}

/** Solo presentación: hasta 2 decimales, sin forzar enteros. */
export function formatPct(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return `${new Intl.NumberFormat("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(n)} %`;
}

export function stringifyCell(v) {
  if (v instanceof Date) return v.toISOString();
  return trimStr(valueToString(v)) || "—";
}

export function joinUnique(setLike) {
  const arr = Array.from(setLike || []).filter(Boolean);
  return arr.length ? arr.join(" · ") : "";
}

/**
 * Fechas genéricas (p. ej. celdas no ligadas a contratos). No usar para columna «Fin» de contratos:
 * ver `parseContractDate` (DD/MM/YYYY México).
 */
export function parseExcelDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const s = trimStr(valueToString(value));
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]) - 1;
    const year = Number(m[3].length === 2 ? "20" + m[3] : m[3]);
    const dt = new Date(year, month, day);
    if (!isNaN(dt.getTime())) return dt;
  }
  return null;
}

export function startOfDay(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** Año YY → YYYY (00–49 → 2000–2049; 50–99 → 1950–1999). */
export function expandTwoDigitContractYear(yy) {
  const n = Number(yy);
  if (!Number.isFinite(n) || n < 0 || n > 99) return null;
  if (n <= 49) return 2000 + n;
  return 1900 + n;
}

/** Fecha calendario local como YYYY-MM-DD (sin hora; útil para trazabilidad). */
export function dateToIsoDateLocal(d) {
  if (!d || isNaN(d.getTime())) return null;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Serial de fecha Excel (días) → medianoche local del día civil correspondiente.
 */
export function excelSerialToLocalDate(serial) {
  const n = Math.floor(Number(serial));
  if (!Number.isFinite(n) || n < 1 || n > 2958465) return null;
  const baseUtcMs = Date.UTC(1899, 11, 31);
  const ms = baseUtcMs + n * 86400000;
  const u = new Date(ms);
  if (isNaN(u.getTime())) return null;
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
}

/**
 * Fecha de análisis (p. ej. activeDataView.contracts.analysisDate).
 * ISO YYYY-MM-DD se interpreta como día civil en hora local (no UTC de `Date.parse`).
 */
export function parseAnalysisDateInput(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return startOfDay(value);
  const s = trimStr(valueToString(value));
  if (!s) return null;
  const isoDay = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (isoDay) {
    const y = Number(isoDay[1]);
    const mo = Number(isoDay[2]) - 1;
    const d = Number(isoDay[3]);
    const dt = new Date(y, mo, d);
    if (!isNaN(dt.getTime())) return startOfDay(dt);
  }
  const p = parseContractDate(s);
  return p.date ? startOfDay(p.date) : null;
}

/**
 * Fecha de fin de contrato: DD/MM/YYYY o DD-MM-YYYY por defecto (México), ISO, serial Excel, Date.
 * @returns {{ date: Date|null, iso: string|null, warning: string|null, ambiguous: boolean, raw: string }}
 */
export function parseContractDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const d = startOfDay(value);
    const iso = dateToIsoDateLocal(d);
    return { date: d, iso, warning: null, ambiguous: false, raw: iso || "" };
  }
  if (value === undefined || value === null) {
    return { date: null, iso: null, warning: null, ambiguous: false, raw: "" };
  }
  if (typeof value === "string" && !trimStr(value)) {
    return { date: null, iso: null, warning: null, ambiguous: false, raw: "" };
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const d = excelSerialToLocalDate(value);
    if (d) {
      const sd = startOfDay(d);
      return { date: sd, iso: dateToIsoDateLocal(sd), warning: null, ambiguous: false, raw: String(value) };
    }
    return {
      date: null,
      iso: null,
      warning: "número no reconocido como fecha Excel",
      ambiguous: false,
      raw: String(value),
    };
  }

  const s = trimStr(valueToString(value));
  if (!s) return { date: null, iso: null, warning: null, ambiguous: false, raw: "" };

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) {
      const y = Number(m[1]);
      const mo = Number(m[2]) - 1;
      const d = Number(m[3]);
      const dt = new Date(y, mo, d);
      if (!isNaN(dt.getTime()) && dt.getFullYear() === y && dt.getMonth() === mo && dt.getDate() === d) {
        const sd = startOfDay(dt);
        return { date: sd, iso: dateToIsoDateLocal(sd), warning: null, ambiguous: false, raw: s };
      }
    }
  }

  const ymd = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (ymd) {
    const y = Number(ymd[1]);
    const mo = Number(ymd[2]) - 1;
    const d = Number(ymd[3]);
    const dt = new Date(y, mo, d);
    if (!isNaN(dt.getTime()) && dt.getFullYear() === y && dt.getMonth() === mo && dt.getDate() === d) {
      const sd = startOfDay(dt);
      return { date: sd, iso: dateToIsoDateLocal(sd), warning: null, ambiguous: false, raw: s };
    }
  }

  const dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]) - 1;
    let year = Number(dmy[3]);
    if (String(dmy[3]).length === 2) {
      const fy = expandTwoDigitContractYear(year);
      if (fy == null) {
        return { date: null, iso: null, warning: "año de 2 dígitos inválido", ambiguous: false, raw: s };
      }
      year = fy;
    }
    const dt = new Date(year, month, day);
    if (
      !isNaN(dt.getTime()) &&
      dt.getFullYear() === year &&
      dt.getMonth() === month &&
      dt.getDate() === day
    ) {
      const sd = startOfDay(dt);
      return { date: sd, iso: dateToIsoDateLocal(sd), warning: null, ambiguous: false, raw: s };
    }
    return { date: null, iso: null, warning: "fecha DD/MM/AAAA inválida", ambiguous: false, raw: s };
  }

  const asNum = Number(String(s).replace(",", ".").trim());
  if (Number.isFinite(asNum) && asNum > 20000 && asNum < 600000) {
    const d = excelSerialToLocalDate(asNum);
    if (d) {
      const sd = startOfDay(d);
      return { date: sd, iso: dateToIsoDateLocal(sd), warning: null, ambiguous: false, raw: s };
    }
  }

  const tryJs = new Date(s);
  if (!isNaN(tryJs.getTime())) {
    const sd = startOfDay(tryJs);
    return {
      date: sd,
      iso: dateToIsoDateLocal(sd),
      warning: "interpretación por Date nativa; verificar formato",
      ambiguous: true,
      raw: s,
    };
  }

  return { date: null, iso: null, warning: "no se pudo interpretar la fecha de fin", ambiguous: false, raw: s };
}

/**
 * Vigencia respecto a la fecha de análisis (días completos, hora local).
 * @param {Date|null} contractEndDate - Fin de contrato o null
 * @param {Date} analysisDate - Fecha de análisis (normalmente startOfDay)
 * @returns {"activo"|"vencido"|"sin_fecha_fin"}
 */
export function getContractStatus(contractEndDate, analysisDate) {
  if (!contractEndDate || isNaN(contractEndDate.getTime())) return "sin_fecha_fin";
  if (!analysisDate || isNaN(analysisDate.getTime())) return "sin_fecha_fin";
  const end = startOfDay(contractEndDate);
  const ana = startOfDay(analysisDate);
  if (end.getTime() >= ana.getTime()) return "activo";
  return "vencido";
}

export function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
