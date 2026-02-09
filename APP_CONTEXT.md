# Contexto de la Aplicaci?n (RS ? Acciones)

## Objetivo
Aplicaci?n web en Google Apps Script para registrar, administrar y visualizar **acciones** y **entidades** relacionadas. Incluye carga de acciones, tablero de entidades, mediciones, calendario y m?dulos de administraci?n de cat?logos, con edici?n in-app y contactos m?ltiples.

## Stack y estructura
- **Apps Script (backend):** `src/Code.js`
- **UI (HTML):** `src/Index.html`
- **L?gica front-end (JS embebido):** `src/App.html`
- **Estilos:** `src/Style.html`
- **Configuraci?n Apps Script:** `src/appsscript.json`

## Hojas de c?lculo usadas
Las hojas se encuentran en el Spreadsheet activo asociado al Script. Estructura base (seg?n archivo RS_DATOS.xlsx):

- **ACCIONES**
  - `timestamp`, `usuario`, `id_accion`, `fecha_carga`, `fecha_accion`, `estado_accion`, `nombre_accion`, `tipo_accion`, `eje`, `lineaTematica`, `produccion`,
    `area_gcba_org_1_id`, `area_gcba_org_2_id`, `area_gcba_org_3_id`, `ente_org_1_id`, `ente_org_2_id`, `ente_org_3_id`, `lugar`, `territorio`,
    `contratacion_personal`, `estado_invitado`, `participantes_entidad_json`, `participantes_persona_json`, `participantes_extras_json`, `meta_tipo`,
    `impacto_numerico`, `comentarios`, `updated_at`, `update_user`, `activo`
- **VALIDACIONES**: `campo`, `valor`, `activo`, `updated_at`, `User`
- **AREAS_GCBA**: `id_area`, `area_nombre`, `activo`, `contacto`, `fecha_contacto`, `updated_at`, `User`
- **ENTIDADES_PRIVADAS**: `id_ente`, `ente_nombre`, `sector`, `rubro`, `web`, `zona_cobertura`, `zona_comuna`, `direccion`, `red_discapacidad`, `observaciones`, `activo`, `updated_at`, `fecha_contacto`, `contacto`
- **PERSONAS**: `id_persona`, `apellido`, `nombre`, `dni`, `mail`, `telefono`, `area`, `rol`, `activo`, `updated_at`, `area_id`, `cargo`, `observaciones`, `red_discapacidad`, `gcba`, `ente_nombre_ref`, `update_user`

## Flujo general de la app
1. **doGet()** devuelve `Index.html` con parciales `Style.html` y `App.html`.
2. **Bootstrap inicial** desde `getUIBootstrap()`:
   - Columnas a mostrar
   - Cat?logos (?reas, entes, personas)
   - Validaciones
   - Estados y etiquetas
3. **Carga de datos** desde `listAcciones()`.
4. **Render de tabla y filtros** en el front.

## Vistas principales
- **Tablero Acciones**: listado con filtros + edici?n.
- **Tablero Entidades**: listado din?mico de entidades/?reas/personas con filtros por Sector/Rubro/Comuna y edici?n.
- **Carga Acciones**: alta de acciones.
- **Carga Entidades (Admin)**: alta de entidades privadas y ?reas GCBA con uno o m?s contactos.
- **Mediciones**: KPIs y tabla por l?nea tem?tica, exportaci?n CSV.
- **Calendario**: vista mensual con acciones por fecha, filtro por estado y buscador.

## Backend (Code.js) ? funciones clave
- **WebApp**: `doGet()`, `include()`
- **Bootstrap UI**: `getUIBootstrap()`
- **Acciones**:
  - `listAcciones()`
  - `getAccionByRow()`
  - `getAccionByIdAccion()`
  - `createAccion()`
  - `updateAccion()`
  - `listAccionesCalendar()`
- **Admin**: alta y validaciones de entidades/?reas/personas
  - **Entidades con m?ltiples contactos**: `createEntidadPrivadaWithContacto()` acepta `contactos[]`
  - **?reas con m?ltiples contactos**: `createAreaWithContacto()` acepta `contactos[]`
- **Cat?logos**: `readCatalog_()`, `readPersonasCatalog_()`
- **Utilidades**: normalizaci?n de headers, fechas, JSON, IDs, etc.

## Front-end (App.html) ? funciones clave
- **Estado global**: `ALL`, `FILTERED`, `DISPLAY_COLUMNS`, `VALIDATIONS`, `CATALOGS`.
- **Filtros del tablero**:
  - Estado, A?o, L?nea tem?tica, Eje, Tipo de acci?n.
  - Las opciones de L?nea/Eje/Tipo se recalculan seg?n el A?o seleccionado.
- **Tabla**: renderiza filas y permite abrir modal de edici?n.
- **Carga**: formulario completo con participantes, combos y estados.
- **Admin**: combos y flujos espec?ficos por tipo + carga de contactos m?ltiples.
- **Mediciones**: KPIs + tabla por l?nea, exportaci?n CSV.
- **Calendario**: grilla mensual con acciones, filtro por estado y b?squeda.

## C?mo se utiliza (usuario final)
1. Abrir la webapp publicada.
2. Navegar por pesta?as:
   - **Tablero Acciones**: buscar, filtrar, editar.
   - **Carga Acciones**: registrar nuevas acciones.
   - **Carga Entidades**: registrar entidades, ?reas y personas (con m?ltiples contactos).
   - **Mediciones**: ver KPIs y exportar.
   - **Calendario**: vista de acciones por fecha.

## Convenciones importantes
- **IDs**: `id_accion` se genera autom?ticamente en `createAccion()`.
- **IDs ?reas**: formato de 4 d?gitos (`AR-0001`), coherente con Acciones.
- **Contactos**: columnas `contacto` en ENTIDADES/?REAS guardan IDs de persona separados por ` | `.
- **Fechas**: formato ISO `YYYY-MM-DD`.
- **Headers**: normalizaci?n robusta (case-insensitive, sin acentos, sin espacios).
- **Columnas bloqueadas**: `fecha_carga`, `nombre_accion`, `timestamp`, `usuario`, `id_accion`.

## Notas de mantenimiento
- Si se agregan columnas nuevas en `ACCIONES`, revisar:
  - `CFG_MAIN.DISPLAY_COLUMNS`
  - Render de tabla en `App.html`
- Si se agregan nuevas validaciones, revisar `VALIDACIONES`.
- Si se modifica encoding, revisar caracteres especiales (acentos).
- Tablero Entidades:
  - El campo `contacto` se muestra como nombres (mapeo desde cat?logo de personas).
  - La edici?n de entidades/?reas/personas est? habilitada desde el tablero.

## Deploy y sincronizaci?n
- **pull**: `clasp pull`
- **push**: `clasp push`
- **publish**: `clasp deploy`

---
Este documento est? pensado como gu?a base para futuros asistentes y desarrolladores.
