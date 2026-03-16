# Contexto Actualizado - RS Acciones (Apps Script)

## Objetivo
WebApp de Google Apps Script para gestionar:
- Acciones (alta, edición, tablero, exportación).
- Entidades (privadas, áreas GCBA, personas contacto).
- Calendario mensual de acciones.
- Mediciones/KPIs.

Incluye lógica de catálogos, validaciones, edición por modal y exportaciones.

## Stack y archivos principales
- Backend Apps Script: `src/Code.js`
- Estructura UI: `src/Index.html`
- Lógica frontend (JS): `src/App.html`
- Estilos: `src/Style.html`
- Manifest: `src/appsscript.json`

## Hojas de datos (fuente principal)
- `ACCIONES`
- `VALIDACIONES`
- `AREAS_GCBA`
- `ENTIDADES_PRIVADAS`
- `PERSONAS`
- (y hojas auxiliares usadas por funciones de bootstrap/admin en `Code.js`)

## Flujo técnico general
1. `doGet()` renderiza `Index` e incluye `Style` + `App`.
2. `getUIBootstrap()` devuelve:
   - columnas, labels, catálogos, validaciones, estados.
3. `listAcciones()` carga datos para tablero, calendario y mediciones.
4. Frontend mantiene estado global y renderiza vistas/tablas/modales.

## Vistas y módulos
- Tablero Acciones
  - Filtros: búsqueda, año, mes, fecha, estado, tipo, eje, línea temática.
  - Botón `Invertir orden` (más reciente/antiguo).
  - Edición por modal.
  - Exportación Excel.
- Carga Acciones
  - Campos por estado con lógica de habilitación/bloqueo.
  - Participantes y combos dinámicos.
- Carga Entidades
  - Módulos: Entidades Privadas, Áreas GCBA, Personas Contacto.
  - Alta/vinculación/desvinculación de contactos.
  - Modal de edición en Tablero Entidades.
- Tablero Entidades
  - Filtros por tipo y atributos.
  - Exportación Excel.
- Mediciones
  - KPIs (Total, Realizadas, En desarrollo, Idea, Canceladas).
  - Tabla por línea temática.
  - Exportación (xlsx si backend disponible, fallback csv).
- Calendario
  - Vista desktop completa.
  - Versión mobile adaptada con interacción por día y acciones.

## Cambios recientes importantes (estado actual)
- Tablero Acciones:
  - `fecha_carga` oculta en UI.
  - `fecha_carga` se mantiene en exportación Excel.
  - Columna `nombre_accion` compactada y multilinea controlada.
  - Resto de columnas compactadas.
  - Columna `comentarios` ensanchada para mejor lectura (sin truncado agresivo).
- Exportaciones:
  - CSV con BOM UTF-8 (`\uFEFF`) para evitar corrupción de tildes/ñ en Excel.
  - Aplica a descarga CSV genérica y fallback de mediciones.
- Responsive:
  - Navegación mobile con menú hamburguesa.
  - Ajustes de layouts, filtros, tablas y calendario para celular.
- Entidades:
  - Modal de edición funcional.
  - Gestión visual de contactos vinculados y creación/vinculación desde modal.

## Convenciones de implementación
- Fechas: ISO `YYYY-MM-DD` para persistencia.
- IDs normalizados por utilidades backend.
- Headers con normalización robusta en backend para tolerar variantes.
- Campos bloqueados en acciones: `fecha_carga`, `nombre_accion`, `timestamp`, `usuario`, `id_accion`.

## Estado frontend relevante (App.html)
- Variables principales:
  - `ALL`, `FILTERED`, `META`
  - `DISPLAY_COLUMNS` (UI)
  - `EXPORT_COLUMNS` (export; puede incluir columnas ocultas en UI)
  - `VALIDATIONS`, `CATALOGS`, `ESTADOS`
- Render principal:
  - `renderHeader()`, `renderTable()`, `applyFilters()`
- Export:
  - `buildExportMatrixFromFiltered_()`
  - `exportFilteredAccionesXLSX()`
  - `exportFilteredEntidadesXLSX()`
  - `medExport_()` / `medExportCSV_()`

## Deploy y sincronización
- Descargar cambios remotos: `clasp pull`
- Subir cambios locales: `clasp push`
- Deploy WebApp: `clasp deploy`

## Nota operativa
Si cambios no aparecen en la web:
1. Hard refresh del navegador (`Ctrl+F5`).
2. Crear nueva versión en Implementaciones de Apps Script si estás usando deployment versionado.
