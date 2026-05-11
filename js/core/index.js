/**
 * Entrada reutilizable del motor de datos comerciales.
 *
 * Esta función encapsula lectura, normalización, consolidación por `customer_code`
 * y construcción de `analysisData` para que un proyecto externo de dashboard pueda
 * importar el procesamiento sin depender de la UI actual ni de renderizado visual.
 */
import { readWorkbookFromFile } from "../excelReader.js";
import { parseWorkbookForSlot } from "../parseWorkbook.js";
import { mergeCustomersByCode, buildConsolidatedCustomerRows } from "../mergeCustomers.js";
import { buildAdvancedDiagnostics, buildConsolidationDiagnosticsList } from "../diagnostics.js";
import {
  createAppState,
  resetAppState,
  syncNormalizedAndRawFromFiles,
  FILE_SLOT,
} from "../state.js";
import { applyActiveDataViewFromFiles, consolidationNormalizedRowsForSlot } from "../consolidationSupport.js";
import { buildAnalysisData } from "../buildAnalysisData.js";

const SLOT_ORDER = [FILE_SLOT.TERRITORY_SALES, FILE_SLOT.MARGINS, FILE_SLOT.CONTRACTS];

const EXPECTED_PROFILE_BY_SLOT = {
  [FILE_SLOT.TERRITORY_SALES]: "territory_sales",
  [FILE_SLOT.MARGINS]: "margins",
  [FILE_SLOT.CONTRACTS]: "contracts",
};

function expectedProfileForSlot(slot) {
  return EXPECTED_PROFILE_BY_SLOT[slot];
}

function buildInputPlan({ territorySalesFile, marginsFile, contractsFile }) {
  return [
    {
      slot: FILE_SLOT.TERRITORY_SALES,
      expected: expectedProfileForSlot(FILE_SLOT.TERRITORY_SALES),
      file: territorySalesFile,
    },
    {
      slot: FILE_SLOT.MARGINS,
      expected: expectedProfileForSlot(FILE_SLOT.MARGINS),
      file: marginsFile,
    },
    {
      slot: FILE_SLOT.CONTRACTS,
      expected: expectedProfileForSlot(FILE_SLOT.CONTRACTS),
      file: contractsFile,
    },
  ];
}

function buildProfileValidationError(parsed) {
  const expected = expectedProfileForSlot(parsed.slot);
  if (parsed.profile === "unknown") {
    const sheets = (parsed.workbook?.SheetNames || []).join(", ");
    return new Error(
      `El archivo «${parsed.fileName}» (${parsed.slot}) no tiene ninguna hoja reconocible como «${expected}» por contenido de columnas. Hojas en el libro: ${sheets || "—"}. ${parsed.detection?.reason || ""}`,
    );
  }

  return new Error(
    `El archivo «${parsed.fileName}» (${parsed.slot}) no coincide con el tipo esperado. Se detectó perfil «${parsed.profile}» y se requiere «${expected}».`,
  );
}

/**
 * Procesa los tres Excel comerciales y devuelve `analysisData`.
 *
 * Uso previsto desde otro proyecto:
 * const analysisData = await processCommercialExcels({
 *   territorySalesFile,
 *   marginsFile,
 *   contractsFile,
 * });
 *
 * @param {{ territorySalesFile?: File, marginsFile?: File, contractsFile?: File }} files
 * @param {{ state?: ReturnType<typeof createAppState> }} [options] - Uso interno de la UI actual para conservar su `appState`.
 * @returns {Promise<object>} analysisData listo para reutilización por dashboards futuros.
 */
export async function processCommercialExcels(
  { territorySalesFile, marginsFile, contractsFile },
  options = {},
) {
  const state = options.state || createAppState();
  resetAppState(state);

  const parseResults = [];
  for (const plan of buildInputPlan({ territorySalesFile, marginsFile, contractsFile })) {
    if (!plan.file) {
      parseResults.push(
        parseWorkbookForSlot({
          slot: plan.slot,
          expectedProfile: plan.expected,
          fileName: "(no cargado)",
          workbook: null,
        }),
      );
      continue;
    }

    const readResult = await readWorkbookFromFile(plan.file);
    if (!readResult.ok) {
      parseResults.push(
        parseWorkbookForSlot({
          slot: plan.slot,
          expectedProfile: plan.expected,
          fileName: plan.file.name,
          workbook: null,
          stubReason: `No se pudo leer el archivo: ${readResult.error}`,
          stubAlerts: [
            "Compruebe que el .xlsx no esté corrupto ni abierto en exclusión en otra aplicación.",
          ],
        }),
      );
      continue;
    }

    parseResults.push(
      parseWorkbookForSlot({
        slot: plan.slot,
        expectedProfile: plan.expected,
        fileName: plan.file.name,
        workbook: readResult.workbook,
      }),
    );
  }

  for (const parsed of parseResults) {
    if (!parsed.workbook) continue;
    const expected = expectedProfileForSlot(parsed.slot);
    if (parsed.profile === "unknown" || parsed.profile !== expected) {
      const err = buildProfileValidationError(parsed);
      err.internalAnalysisNotice = "No se pudo generar la base interna de análisis.";
      err.statusMessage = err.message;
      throw err;
    }
  }

  const ordered = SLOT_ORDER.map((slot) => parseResults.find((p) => p.slot === slot)).filter(Boolean);
  state.files = ordered;
  syncNormalizedAndRawFromFiles(state);
  applyActiveDataViewFromFiles(ordered, state.activeDataView);
  state.diagnostics = [
    ...buildAdvancedDiagnostics(ordered),
    ...buildConsolidationDiagnosticsList(ordered),
  ];

  state.customersByCode = mergeCustomersByCode(
    consolidationNormalizedRowsForSlot(ordered, FILE_SLOT.TERRITORY_SALES),
    consolidationNormalizedRowsForSlot(ordered, FILE_SLOT.MARGINS),
    consolidationNormalizedRowsForSlot(ordered, FILE_SLOT.CONTRACTS),
  );
  state.customers = buildConsolidatedCustomerRows(state.customersByCode, {
    contractsAnalysisDate: state.activeDataView.contracts.analysisDate,
    territoryActiveYear: state.activeDataView.territorySales.activeYear,
  });
  state.analysisData = buildAnalysisData(state);

  return state.analysisData;
}
