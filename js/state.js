/**
 * Estado central de la aplicación (única fuente de verdad en memoria).
 *
 * Regla de negocio: la unión entre archivos usa solo `customer_code` (identificador
 * estable del ERP). `customer_name` puede variar por archivo o fila y no debe usarse
 * como llave para evitar duplicados falsos o pérdida de filas al unir.
 */
export const FILE_SLOT = {
  TERRITORY_SALES: "territorySales",
  MARGINS: "margins",
  CONTRACTS: "contracts",
};

export function createAppState() {
  return {
    /** Resultado de parseo por cada slot (siempre 3 entradas tras un proceso completo; puede incluir stubs si falta archivo). */
    files: [],
    /** Filas crudas extraídas del Excel (sin reglas de cliente). */
    rawData: {
      territorySales: [],
      margins: [],
      contracts: [],
    },
    /** Filas listas para análisis (con customer_code cuando aplica). */
    normalizedData: {
      territorySales: [],
      margins: [],
      contracts: [],
    },
    /** Resumen de exclusiones y contexto para el panel de diagnóstico avanzado. */
    diagnostics: [],
    /** Mapa código → agregación por fuente (ventas / márgenes / contratos). */
    customersByCode: new Map(),
    /** Tabla consolidada `customers` (una fila por customer_code). */
    customers: [],
    /**
     * Base interna del análisis (objeto plano tipo JSON): normalizados, customers, cruces por código,
     * resumen de archivos, diagnósticos, campos detectados y calidad. Sin libro Excel (`workbook`).
     * Se regenera tras cada procesamiento exitoso.
     */
    analysisData: null,
    /** Vista activa por perfil para el consolidado principal (no sustituye el dataset completo en normalizedData). */
    activeDataView: createDefaultActiveDataView(),
  };
}

export function createDefaultActiveDataView() {
  return {
    territorySales: {
      activeYear: null,
      activeMonths: [],
      activeSheets: [],
      allowMultiYearAggregation: false,
    },
    margins: {
      activeSheet: null,
      activePeriod: null,
      allowMultiSheetAggregation: false,
    },
    contracts: {
      mode: "active_or_historical",
      analysisDate: null,
      activeSheets: [],
      allowMultiSheetAggregation: false,
    },
  };
}

/**
 * Restaura el estado y deja listo para un nuevo ciclo.
 * @param {ReturnType<typeof createAppState>} state
 */
export function resetAppState(state) {
  state.files = [];
  state.rawData.territorySales = [];
  state.rawData.margins = [];
  state.rawData.contracts = [];
  state.normalizedData.territorySales = [];
  state.normalizedData.margins = [];
  state.normalizedData.contracts = [];
  state.diagnostics = [];
  state.customersByCode = new Map();
  state.customers = [];
  state.analysisData = null;
  state.activeDataView = createDefaultActiveDataView();
}

/**
 * Sincroniza arrays planos y el mapa a partir de los resultados de parseo en `files`.
 * @param {ReturnType<typeof createAppState>} state
 */
export function syncNormalizedAndRawFromFiles(state) {
  const bySlot = Object.fromEntries(state.files.map((f) => [f.slot, f]));
  state.rawData.territorySales = bySlot[FILE_SLOT.TERRITORY_SALES]?.rowsRaw ?? [];
  state.rawData.margins = bySlot[FILE_SLOT.MARGINS]?.rowsRaw ?? [];
  state.rawData.contracts = bySlot[FILE_SLOT.CONTRACTS]?.rowsRaw ?? [];
  state.normalizedData.territorySales = bySlot[FILE_SLOT.TERRITORY_SALES]?.normalizedRows ?? [];
  state.normalizedData.margins = bySlot[FILE_SLOT.MARGINS]?.normalizedRows ?? [];
  state.normalizedData.contracts = bySlot[FILE_SLOT.CONTRACTS]?.normalizedRows ?? [];
}
