# Dependencias vendorizadas

Esta carpeta contiene **copias locales** de librerías de terceros usadas por el dashboard estático. No se cargan desde CDN en tiempo de ejecución.

## Contenido

| Archivo | Librería | Versión exacta | URL de origen (descarga) |
|-----------|-----------|------------------|---------------------------|
| `xlsx.full.min.js` | SheetJS (xlsx) | **0.18.5** | https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js |
| `plotly-2.30.0.min.js` | Plotly.js | **2.30.0** | https://cdn.plot.ly/plotly-2.30.0.min.js |

## Fecha de incorporación

**2026-05-11** (copias alineadas con las URLs anteriores).

## Política de actualización

- **No actualizar** estas versiones sin **prueba funcional** explícita.
- Tras cualquier cambio de versión, validar como mínimo:
  - **Procesamiento de Excel** (carga de los tres archivos comerciales y motor `processCommercialExcels`).
  - **Renderizado Plotly** (mapa estratégico, hover, filtros, exportación si aplica).
- Si se sustituyen los binarios, actualizar este README (versión, URL, fecha) y `index.html` si cambian nombres de archivo.

## Licencias

Respetar las licencias de cada proyecto (SheetJS / Plotly). Consultar el encabezado de cada `.min.js` o la documentación oficial del proveedor.
