/**
 * Registro en consola para validar internamente un procesamiento (sin enviar datos a servidor).
 * Útil al depurar perfiles, conteos y la tabla consolidada.
 *
 * Seguridad: el resumen incluye nombres reales de archivo y métricas por slot,
 * que pueden ser sensibles en entornos compartidos. En producción solo se imprime
 * si se activa explícitamente `globalThis.__LECTURA_DATOS_DEBUG__ = true`.
 */

function isLogSnapshotEnabled() {
  if (globalThis.__LECTURA_DATOS_DEBUG__ === true) return true;
  const host = globalThis.location?.hostname;
  return host === "localhost" || host === "127.0.0.1";
}

function sumExcluded(diagnostics) {
  return diagnostics.reduce((acc, d) => acc + (d.excludedCount || 0), 0);
}

/**
 * @param {object} params
 * @param {object[]} params.files - Resultados de parseo por slot
 * @param {object} params.normalizedData - { territorySales, margins, contracts }
 * @param {object[]} params.diagnostics
 * @param {Map} params.customersByCode
 * @param {object[]} params.customers - filas consolidadas
 * @param {object | null} [params.analysisData] - Paquete analysisData (calidad + campos)
 */
export function logProcessingSnapshot({
  files,
  normalizedData,
  diagnostics,
  customersByCode,
  customers,
  analysisData,
}) {
  if (!isLogSnapshotEnabled()) return;
  const lines = [
    "[Lectura datos] Resumen de procesamiento",
    "— Archivos / perfil detectado —",
  ];

  (files || []).forEach((f) => {
    lines.push(
      `  · ${f.slot}: ${f.fileName} → perfil=${f.profile} (${f.detection?.method || "?"}) hoja=${f.sheetName || "—"}`
    );
  });

  lines.push(
    "— Filas normalizadas —",
    `  · territorySales: ${(normalizedData?.territorySales || []).length}`,
    `  · margins: ${(normalizedData?.margins || []).length}`,
    `  · contracts: ${(normalizedData?.contracts || []).length}`,
    "— Exclusiones (metadatos agregados) —",
    `  · entradas de diagnóstico: ${(diagnostics || []).length}, filas excluidas (total reportado): ${sumExcluded(diagnostics || [])}`,
    "— Clientes —",
    `  · únicos en índice (Map): ${customersByCode?.size ?? 0}`,
    `  · filas tabla consolidada: ${(customers || []).length}`,
  );

  if (analysisData?.metadata?.generatedAt) {
    lines.push(
      "— analysisData —",
      `  · generado: ${analysisData.metadata.generatedAt}`,
      `  · calidad: excl. total ${analysisData.quality?.excludedRowsTotal ?? "?"}, ` +
        `clientes un solo origen ${analysisData.quality?.clientsSingleSource ?? "?"}, ` +
        `multi-origen ${analysisData.quality?.clientsMultiSource ?? "?"}`
    );
  }

  console.info(lines.join("\n"));
}
