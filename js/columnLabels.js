/**
 * Etiquetas de columnas esperadas en los Excel (candidatos para búsqueda flexible).
 * La coincidencia real usa normalizeHeader + findColumnKey en utils.
 */

/** Archivo de márgenes / facturación */
export const MARGINS_COLUMNS = {
  PAYER_CODE: ["Cód. Pagador"],
  PAYER_NAME: ["Nombre pagad."],
  BILLED_VALUE: ["Valor facturado Loc"],
  GROSS_MARGIN: ["Margen bruto Loc"],
  MARGIN_OVER_SALES: ["Margen / ventas Loc"],
};

/** Archivo de contratos / compliance */
export const CONTRACTS_COLUMNS = {
  CLIENT_CODE: ["Solicitante Cliente"],
  CLIENT_NAME: ["Nombre del Solicitante"],
  CONTRACT_ID: ["Contrato"],
  START: ["Inicio"],
  END: ["Fin"],
  NET_CONTRACT_VALUE: ["Valor Neto Contrato"],
  EXPECTED_BILLING: ["Facturación Esperada"],
  NET_BILLED: ["Facturado Neto"],
  BILLING_COMPLIANCE_PCT: ["Cumplimiento Facturación%"],
};

/** Ventas por territorio: columna que combina código + nombre de cliente */
export const TERRITORY_SALES_CLIENT_COLUMNS = [
  "Representante / Cliente",
  "Representante/Cliente",
  "Representante - Cliente",
  "Representante",
  "Cliente",
];

/** Territorio / región */
export const TERRITORY_SALES_REGION_COLUMNS = [
  "Territorio / Región",
  "Territorio / Region",
  "Territorio",
  "Región",
  "Region",
];
