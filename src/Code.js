/******************************************************
 * RS — WebApp ACCIONES (Carga + Editar + Admin)
 * build: 2026-01-13 (AR)
 *
 * ✅ Code.gs unificado y compatible con Index.html (3 views)
 * - doGet() + include()
 * - Bootstrap UI: getDefaults(), getBuildInfo(), getUIBootstrap()
 * - Acciones: listAcciones(), getAccionByRow(), createAccion(), updateAccion()
 * - Admin: VALIDACIONES + AREAS + ENTES + PERSONAS (create + preview + addValidation)
 *
 * ✅ Adaptaciones para Admin nuevo (Index+App nuevos):
 * - adminListCamposValidacion() para combo “Campo”
 * - getAdminBootstrap() opcional (campos + areas + previews)
 * - aliases de funciones para no romper App viejo
 ******************************************************/

/* =========================================================
 * CONFIG
 * ========================================================= */

const CFG_MAIN = {
  SHEET_ACCIONES: "ACCIONES",
  SHEET_AREAS: "AREAS_GCBA",
  SHEET_ENTES: "ENTIDADES_PRIVADAS",
  SHEET_PERSONAS: "PERSONAS",
  SHEET_VALIDACIONES: "VALIDACIONES",
  SHEET_VALIDACIONES_FILTROS: "VALIDACIONES_FILTROS",

  // Columnas que el UI usa (tabla + editor) — NO incluye *_extras_json
  DISPLAY_COLUMNS: [
    "fecha_carga",
    "fecha_accion",
    "estado_accion",
    "nombre_accion",
    "tipo_accion",
    "eje",
    "lineaTematica",
    "produccion",
    "area_gcba_org_1_id",
    "area_gcba_org_2_id",
    "area_gcba_org_3_id",
    "ente_org_1_id",
    "ente_org_2_id",
    "ente_org_3_id",
    "lugar",
    "territorio",
    "contratacion_personal",
    "estado_invitado",
    "participantes_entidad_json",
    "participantes_persona_json",
    "meta_tipo",
    "impacto_numerico",
    "comentarios",
  ],

  // Campos NO editables desde UI
  LOCKED_FIELDS: new Set(["fecha_carga", "nombre_accion", "timestamp", "usuario", "id_accion"]),

  // Campos tipo fecha (solo yyyy-mm-dd)
  DATE_ONLY: new Set(["fecha_carga", "fecha_accion"]),

  // Defaults / fallbacks
  FALLBACK_ESTADOS_ACCION: ["IDEA", "DESARROLLO", "REALIZADO", "CANCELADO"],
  FALLBACK_SI_NO: ["SI", "NO"],
};

const CFG_ADMIN = {
  validacionesSheetName: "VALIDACIONES",
  areasSheetName: "AREAS_GCBA",
  entesSheetName: "ENTIDADES_PRIVADAS",
  personasSheetName: "PERSONAS",
};

const RS_BUILD = "RS | build 2026-01-13 AR";

const ENTIDADES_MAP = {
  entidad_privada: {
    sheet: "ENTIDADES_PRIVADAS",
    label: "Entidades Privadas",
    columns: [
      "id_ente","ente_nombre","sector","rubro","web","zona_cobertura","zona_comuna","direccion",
      "red_discapacidad","observaciones","activo","updated_at","fecha_contacto","contacto"
    ],
    lock: new Set(["id_ente","updated_at","update_user"])
  },
  area_gcba: {
    sheet: "AREAS_GCBA",
    label: "Áreas GCBA",
    columns: ["id_area","area_nombre","activo","contacto","fecha_contacto","updated_at","user"],
    lock: new Set(["id_area","updated_at","update_user","user"])
  },
  persona: {
    sheet: "PERSONAS",
    label: "Personas",
    columns: [
      "id_persona","apellido","nombre","dni","mail","telefono","area","rol","activo","updated_at",
      "area_id","cargo","observaciones","red_discapacidad","gcba","ente_nombre_ref"
    ],
    lock: new Set(["id_persona","updated_at","update_user"])
  }
};

/* =========================================================
 * WEBAPP ENTRY
 * ========================================================= */
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("RS • Acciones")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/** include HTML parciales */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Defaults server-side (compatible con tu UI) */
function getDefaults() {
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const fechaISO = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  const userEmail = safeUserEmail_();

  return {
    fechaISO,
    fechaCargaISO: fechaISO,
    userEmail,
  };
}

function getBuildInfo() {
  const ss = SpreadsheetApp.getActive();
  return {
    build: RS_BUILD,
    scriptId: ScriptApp.getScriptId(),
    ssId: ss.getId(),
    ssName: ss.getName(),
    ts: new Date().toISOString(),
  };
}

/* =========================================================
 * ===================== CARGA + EDITAR =====================
 * ========================================================= */

/**
 * Bootstrap UI.
 * Devuelve catálogos + validaciones + labels + lockedFields.
 */
function getUIBootstrap() {
  const areas = readCatalog_(CFG_MAIN.SHEET_AREAS, "id_area", "area_nombre");
  const entes = readCatalog_(CFG_MAIN.SHEET_ENTES, "id_ente", "ente_nombre");
  const personas = readPersonasCatalog_();

  const validations = readValidaciones_();

  const estadosInv = pickValidation_(validations, [
    "estadosinvitacion",
    "estadoinvitado",
    "estado_invitado",
  ]);

  const participantStates = uniq_(["A invitar", ...estadosInv, "Cancelado"]);

  // ====== ✅ CALENDARIO: filtros + colores ======
  const estadosAccion = uniq_(
    pickValidation_(validations, ["estado_accion", "estadoaccion"]).length
      ? pickValidation_(validations, ["estado_accion", "estadoaccion"])
      : CFG_MAIN.FALLBACK_ESTADOS_ACCION
  ).map(normalizeEstado_);

  const lineasTematicas = uniq_(
    pickValidation_(validations, ["lineaTematica", "lineatematica", "linea_tematica"])
  );

  const calendar = {
    estadosAccion,
    lineasTematicas,
    colorsByLinea: buildColorsByLinea_(lineasTematicas),
  };

  // Labels
  const labels = {
    timestamp: "Timestamp",
    usuario: "Usuario",
    fecha_carga: "Fecha de carga",
    fecha_accion: "Fecha de acción",
    estado_accion: "Estado",
    nombre_accion: "Nombre de la acción",
    tipo_accion: "Tipo de acción",
    eje: "Eje Gobierno",
    lineaTematica: "Línea temática",
    produccion: "Producción",
    area_gcba_org_1_id: "Área GCBA Organizadora 1",
    area_gcba_org_2_id: "Área GCBA Organizadora 2",
    area_gcba_org_3_id: "Área GCBA Organizadora 3",
    ente_org_1_id: "Ente Organizador 1 (Privado)",
    ente_org_2_id: "Ente Organizador 2 (Privado)",
    ente_org_3_id: "Ente Organizador 3 (Privado)",
    lugar: "Lugar",
    territorio: "Territorio",
    contratacion_personal: "Contratación de personal",
    estado_invitado: "Estado invitación",
    participantes_entidad_json: "Participantes (Entidad)",
    participantes_persona_json: "Participantes (Persona)",
    meta_tipo: "Tipo de indicador",
    impacto_numerico: "Impacto (numérico)",
    comentarios: "Comentarios",
    updated_at: "Actualizado",
    update_user: "Actualizado por",
    activo: "Activo",
  };

  return {
    ok: true,
    displayColumns: CFG_MAIN.DISPLAY_COLUMNS,
    lockedFields: Array.from(CFG_MAIN.LOCKED_FIELDS),
    catalogs: { areas, entes, personas },
    validations,
    participantStates,
    labels,

    // ✅ nuevo: soporte calendario
    calendar,
  };
}



/**
 * (Opcional) Bootstrap para Admin nuevo (si tu App.html lo prefiere)
 */
function getAdminBootstrap() {
  const areas = readCatalog_(CFG_ADMIN.areasSheetName, "id_area", "area_nombre");
  const campos = adminListCamposValidacion();
  return {
    ok: true,
    camposValidacion: campos,
    areasCatalog: areas,
  };
}

/**
 * Lista acciones + enrich nombres (areas / entes).
 */
function listAcciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const areasMap = toMap_(readCatalog_(CFG_MAIN.SHEET_AREAS, "id_area", "area_nombre"));
  const entesMap = toMap_(readCatalog_(CFG_MAIN.SHEET_ENTES, "id_ente", "ente_nombre"));

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    return {
      ok: true,
      count: 0,
      rows: [],
      meta: { sheet: CFG_MAIN.SHEET_ACCIONES, lastRow, lastCol, spreadsheetId: ss.getId() },
    };
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headersNorm = values[0].map((h) => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h, i) => {
    if (h) idx[h] = i;
  });

  const mustHave = ["id_accion", ...CFG_MAIN.DISPLAY_COLUMNS.filter((c) => c !== "updated_at" && c !== "update_user")];
  const missing = mustHave.filter((c) => idx[c] === undefined);
  const missingHard = missing.filter((c) => !["updated_at", "update_user", "activo"].includes(c));
  if (missingHard.length) throw new Error(`Faltan columnas en ACCIONES: ${missingHard.join(", ")}`);

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rowNumber = r + 1;

    const obj = {
      rowNumber,
      id_accion: String(row[idx["id_accion"]] || "").trim(),
    };

    for (const col of CFG_MAIN.DISPLAY_COLUMNS) {
      if (idx[col] === undefined) {
        obj[col] = "";
        continue;
      }
      obj[col] = formatByType_(col, row[idx[col]]);
    }

    obj.estado_accion = normalizeEstado_(obj.estado_accion);

    for (const c of ["area_gcba_org_1_id", "area_gcba_org_2_id", "area_gcba_org_3_id"]) {
      const id = String(obj[c] || "").trim();
      obj[c + "_name"] = id ? (areasMap[id] || "") : "";
    }
    for (const c of ["ente_org_1_id", "ente_org_2_id", "ente_org_3_id"]) {
      const id = String(obj[c] || "").trim();
      obj[c + "_name"] = id ? (entesMap[id] || "") : "";
    }

    if (!obj.id_accion && !obj.nombre_accion) continue;
    rows.push(obj);
  }

  return {
    ok: true,
    count: rows.length,
    rows,
    meta: { sheet: CFG_MAIN.SHEET_ACCIONES, lastRow, lastCol, spreadsheetId: ss.getId() },
  };
}

/**
 * Para edición: trae el objeto de una fila puntual.
 */
function getAccionByRow(rowNumber) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error("rowNumber inválido");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lastCol = sh.getLastColumn();
  const headersNorm = sh.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h, i) => (idx[h] = i));

  const row = sh.getRange(rn, 1, 1, lastCol).getValues()[0];

  const obj = { rowNumber: rn };
  obj.id_accion = idx["id_accion"] !== undefined ? String(row[idx["id_accion"]] || "").trim() : "";

  for (const col of CFG_MAIN.DISPLAY_COLUMNS) {
    if (idx[col] === undefined) obj[col] = "";
    else obj[col] = formatByType_(col, row[idx[col]]);
  }

  obj.estado_accion = normalizeEstado_(obj.estado_accion);

  return { ok: true, row: obj };
}
/* =========================================================
 * ======================= CALENDARIO =======================
 * ========================================================= */

/**
 * Devuelve acciones para un mes visible del calendario.
 * params:
 *  - year: number (ej 2026)
 *  - month: number 1..12
 *  - estado: "TODOS" | "IDEA" | "DESARROLLO" | "REALIZADO" | "CANCELADO"
 *  - lineas: array<string> (lineaTematica permitidas)
 *
 * output:
 *  - events: [{ id, rowNumber, date, title, estado, lineaTematica, color, meta }]
 */
function listAccionesCalendar(params){
  params = params || {};
  const year  = Number(params.year);
  const month = Number(params.month);
  if (!year || !month || month < 1 || month > 12) throw new Error("Parámetros inválidos: year/month");

  const estadoIn = String(params.estado || "TODOS").trim().toUpperCase();
  const estadoFilter = (estadoIn && estadoIn !== "TODOS") ? normalizeEstado_(estadoIn) : "";

  const lineasArr = Array.isArray(params.lineas) ? params.lineas : [];
  const lineasSet = new Set(lineasArr.map(x => String(x||"").trim()).filter(Boolean));

  // ventana: mes visible + margen para grilla (semana anterior y posterior)
  const start = new Date(year, month - 1, 1);
  start.setDate(start.getDate() - 7);
  const end = new Date(year, month, 1);
  end.setDate(end.getDate() + 7);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, year, month, count:0, events:[] };

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const headersNorm = values[0].map(h => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h,i)=>{ if (h) idx[h] = i; });

  const iId     = idx["id_accion"];
  const iFecha  = idx["fecha_accion"];
  const iEstado = idx["estado_accion"];
  const iNombre = idx["nombre_accion"];
  const iTipo   = idx["tipo_accion"];
  const iLinea  = idx["lineaTematica"];
  const iAct    = idx["activo"];

  const missing = [];
  if (iFecha === undefined) missing.push("fecha_accion");
  if (iEstado === undefined) missing.push("estado_accion");
  if (iNombre === undefined) missing.push("nombre_accion");
  if (missing.length) throw new Error(`Faltan columnas en ACCIONES (Calendario): ${missing.join(", ")}`);

  // colores por línea (bootstrap)
  const validations = readValidaciones_();
  const lineasBase = uniq_(pickValidation_(validations, ["lineaTematica","lineatematica","linea_tematica"]));
  const colorsByLinea = buildColorsByLinea_(lineasBase);

  const events = [];
  for (let r=1; r<values.length; r++){
    const row = values[r];
    const rowNumber = r + 1;

    // activo (si existe)
    if (iAct !== undefined){
      const act = String(row[iAct] || "SI").trim().toUpperCase();
      if (act === "NO") continue;
    }

    // fecha
    const rawFecha = row[iFecha];
    const iso = coerceToISODate_(rawFecha);
    if (!iso) continue;

    const d = isoToDate_(iso);
    if (!d) continue;
    if (d < start || d >= end) continue;

    // estado
    const est = normalizeEstado_(row[iEstado]);
    if (estadoFilter && est !== estadoFilter) continue;

    // línea temática
    const linea = (iLinea === undefined) ? "" : String(row[iLinea] || "").trim();
    if (lineasSet.size && !lineasSet.has(linea)) continue;

    // título (como en tu screenshot: podés tunearlo después en front)
    const nombre = String(row[iNombre] || "").trim();
    const tipo = (iTipo === undefined) ? "" : String(row[iTipo] || "").trim();

    const title = tipo ? `[${tipo}] ${nombre}` : nombre;

    const idAcc = (iId === undefined) ? "" : String(row[iId] || "").trim();
    const color = colorsByLinea[linea] || colorsByLinea["__default"] || "#DDE6EE";

    events.push({
      id: idAcc || `ROW_${rowNumber}`,
      rowNumber,
      date: iso,           // YYYY-MM-DD
      title,
      estado: est,
      lineaTematica: linea,
      color,               // hex
      meta: {
        tipo_accion: tipo,
      }
    });
  }

  // orden por fecha y luego título
  events.sort((a,b)=>{
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return String(a.title||"").localeCompare(String(b.title||""), "es");
  });

  return { ok:true, year, month, count: events.length, events };
}

/**
 * Detalle por click: trae la acción por id_accion (si existe),
 * y si no existe id_accion, podés usar getAccionByRow(rowNumber).
 */
function getAccionByIdAccion(id_accion){
  const id = String(id_accion || "").trim();
  if (!id) throw new Error("id_accion inválido");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, found:false };

  const values = sh.getRange(1,1,lr,lc).getValues();
  const headersNorm = values[0].map(h => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h,i)=>{ if (h) idx[h] = i; });

  const iId = idx["id_accion"];
  if (iId === undefined) throw new Error(`ACCIONES no tiene la columna id_accion`);

  for (let r=1; r<values.length; r++){
    const row = values[r];
    const rid = String(row[iId] || "").trim();
    if (rid !== id) continue;

    // reutilizamos tu formato de edición
    return getAccionByRow(r + 1);
  }

  return { ok:true, found:false };
}

/**
 * ✅ CREATE: crea nueva acción y genera id_accion
 */
function createAccion(payload) {
  payload = payload || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lastCol = sh.getLastColumn();
  const headersRaw = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  const headers = headersRaw.map((h) => normalizeHeader_(h));
  const idx0 = {};
  headers.forEach((h, i) => {
    if (h) idx0[h] = i;
  });

  const nombre = String(getPayload_(payload, "nombre_accion") || "").trim();
  const fechaISO = coerceToISODate_(getPayload_(payload, "fecha_accion"));
  const estado = normalizeEstado_(getPayload_(payload, "estado_accion"));
  const tipo = String(getPayload_(payload, "tipo_accion") || "").trim();

  if (!nombre) throw new Error("Nombre de la acción es obligatorio.");
  if (!fechaISO) throw new Error("Fecha de acción es obligatoria.");
  if (!estado) throw new Error("Estado de acción es obligatorio.");
  if (!tipo) throw new Error("Tipo de acción es obligatorio.");

  const idBase = buildAccionId_(fechaISO, tipo, nombre);

  let idFinal = idBase;
  if (idx0["id_accion"] !== undefined) {
    const lr = sh.getLastRow();
    if (lr >= 2) {
      const col = idx0["id_accion"] + 1;
      const existing = sh
        .getRange(2, col, lr - 1, 1)
        .getValues()
        .flat()
        .map((v) => String(v || "").trim())
        .filter(Boolean);
      idFinal = ensureUniqueId_(idBase, new Set(existing));
    }
  }

  const now = new Date();
  const user = safeUserEmail_();

  setPayload_(payload, "id_accion", idFinal);
  setPayload_(payload, "fecha_accion", fechaISO);
  setPayload_(payload, "estado_accion", estado);

  if (idx0["fecha_carga"] !== undefined && !getPayload_(payload, "fecha_carga")) setPayload_(payload, "fecha_carga", toISODate_(now));
  if (idx0["usuario"] !== undefined && !getPayload_(payload, "usuario")) setPayload_(payload, "usuario", user);
  if (idx0["timestamp"] !== undefined && !getPayload_(payload, "timestamp")) setPayload_(payload, "timestamp", formatTimestamp_(now));
  if (idx0["activo"] !== undefined && !getPayload_(payload, "activo")) setPayload_(payload, "activo", "SI");

  setPayload_(payload, "estado_accion", normalizeEstado_(getPayload_(payload, "estado_accion")));
  if (idx0["impacto_numerico"] !== undefined) {
    const imp = getPayload_(payload, "impacto_numerico");
    setPayload_(payload, "impacto_numerico", imp === null || imp === undefined ? "" : String(imp).trim());
  }

  const row = headers.map((h) => {
    if (!h) return "";

    let v = getPayload_(payload, h);
    if (v === undefined) v = payload[h];

    if (/_json$/.test(h)) v = coerceToJsonString_(v);
    if (CFG_MAIN.DATE_ONLY.has(h)) v = coerceToISODate_(v);

    if (v === null || v === undefined) return "";
    return v;
  });

  sh.appendRow(row);

  return { ok: true, id_accion: idFinal };
}

/**
 * Actualiza una fila SOLO columnas permitidas.
 */
function updateAccion(rowNumber, updates) {
  const rn = Number(rowNumber);
  if (!rn || rn < 2) throw new Error("rowNumber inválido");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lastCol = sh.getLastColumn();
  const headersNorm = sh.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h, i) => {
    if (h) idx[h] = i + 1; // 1-based
  });

  const allowed = new Set(CFG_MAIN.DISPLAY_COLUMNS);

  for (const key in updates || {}) {
    const k = normalizeHeader_(key);

    if (!allowed.has(k)) continue;
    if (CFG_MAIN.LOCKED_FIELDS.has(k)) continue;

    const colNum = idx[k];
    if (!colNum) continue;

    let v = updates[key];

    if (/_json$/.test(k)) v = coerceToJsonString_(v);

    if (CFG_MAIN.DATE_ONLY.has(k)) {
      const iso = coerceToISODate_(v);
      if (!iso) sh.getRange(rn, colNum).clearContent();
      else sh.getRange(rn, colNum).setValue(iso);
      continue;
    }

    if (k === "estado_accion") v = normalizeEstado_(v);

    sh.getRange(rn, colNum).setValue(v === null || v === undefined ? "" : v);
  }

  const now = new Date();
  const user = safeUserEmail_();

  if (idx["updated_at"]) sh.getRange(rn, idx["updated_at"]).setValue(now);
  if (idx["update_user"]) sh.getRange(rn, idx["update_user"]).setValue(user);

  return { ok: true };
}

/* =========================================================
 * ========================= ADMIN ==========================
 * ========================================================= */

/** Normaliza strings para comparar duplicados (admin) */
function norm_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

/** Asegura hoja + headers (si existe, NO pisa nada; si no hay headers, los crea) */
function ensureSheetHeaders_(sheetName, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  const r1 = sh.getRange(1, 1, 1, Math.max(headers.length, 1)).getValues()[0] || [];
  const hasAnyHeader = r1.some((v) => String(v || "").trim() !== "");
  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Devuelve map campo -> [valores activos] */
function getValidacionesMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.validacionesSheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_ADMIN.validacionesSheetName}"`);

  const v = sh.getDataRange().getValues();
  if (!v || v.length < 2) return {};

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const iCampo = headers.indexOf("campo");
  const iValor = headers.indexOf("valor");
  const iAct = headers.indexOf("activo");

  if (iCampo === -1 || iValor === -1) return {};

  const out = {};
  rows.forEach((r) => {
    const campo = String(r[iCampo] || "").trim();
    const valor = String(r[iValor] || "").trim();
    if (!campo || !valor) return;

    if (iAct !== -1 && String(r[iAct] || "").trim().toUpperCase() === "NO") return;

    const key = normalizeValidationKey_(campo);
    if (!out[key]) out[key] = [];
    out[key].push(valor);
  });

  Object.keys(out).forEach((k) => {
    out[k] = [...new Set(out[k].map((x) => String(x).trim()))]
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, "es"));
  });

  return out;
}
// ✅ Endpoint público para el Front (Tablero Acciones)
// Devuelve un mapa "bonito" + raw normalizado (compatibilidad)
function getValidacionesMap(){
  const raw = getValidacionesMap_(); // ya existe, devuelve claves normalizadas

  // helper: traer por varias posibles keys (normalizadas)
  const pick = (candidates) => {
    for (const c of candidates){
      const k = normalizeValidationKey_(c);
      if (Array.isArray(raw[k]) && raw[k].length) return raw[k];
    }
    return [];
  };

  return {
    ok: true,

    // ✅ mapa normalizado (por si tu front viejo lo usa)
    raw,

    // ✅ mapa "bonito" (para chips/selects del Tablero)
    EstadoAccion: pick(["EstadoAccion","estado_accion","estadoaccion","estado"]),
    EstadosInvitacion: pick(["EstadosInvitacion","estado_invitado","estadoinvitado","estado invitado"]),
    Produccion: pick(["Produccion","produccion"]),
    ContratacionPersonal: pick(["ContratacionPersonal","contratacion_personal","contratacionpersonal"]),
    Lugar: pick(["Lugar","lugar"]),
    TiposAccion: pick(["TiposAccion","tipo_accion","tiposaccion","tipoaccion"]),
    Ejes: pick(["Ejes","eje","ejes"]),
    lineaTematica: pick(["lineaTematica","lineatematica","linea_tematica","linea temática"]),
    MetaTipos: pick(["MetaTipos","meta_tipo","metatipos"]),
    Territorios: pick(["Territorios","territorio","territorios"])
  };
}

/**
 * ✅ NUEVO: lista de campos (para combo editable “Campo” en Admin)
 */
function adminListCamposValidacion() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.validacionesSheetName);
  if (!sh) return [];

  const v = sh.getDataRange().getValues();
  if (!v || v.length < 2) return [];

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);
  const iCampo = headers.indexOf("campo");
  const iAct = headers.indexOf("activo");

  if (iCampo === -1) return [];

  const set = new Map(); // norm -> original
  rows.forEach((r) => {
    const campo = String(r[iCampo] || "").trim();
    if (!campo) return;

    if (iAct !== -1 && String(r[iAct] || "").trim().toUpperCase() === "NO") return;

    const key = norm_(campo);
    if (!set.has(key)) set.set(key, campo);
  });

  const list = Array.from(set.values()).sort((a, b) => a.localeCompare(b, "es"));

  const pinned = [
    "estado_accion",
    "tipo_accion",
    "eje",
    "lineaTematica",
    "produccion",
    "lugar",
    "territorio",
    "contratacion_personal",
    "meta_tipo",
    "estado_invitado",
  ];

  const pinnedPretty = pinned
    .map((x) => {
      const found = list.find((y) => norm_(y) === norm_(x));
      return found || x;
    })
    .filter((x, i, arr) => arr.indexOf(x) === i);

  const merged = [...pinnedPretty, ...list.filter((x) => !pinnedPretty.some((p) => norm_(p) === norm_(x)))];
  return merged;
}

/** Preview simple de validaciones por campo */
function adminPreviewValidaciones(campo, limit) {
  const lim = Math.min(Math.max(Number(limit || 30), 1), 200);
  const m = getValidacionesMap_();
  const key = normalizeValidationKey_(campo);
  const list = m[key] || [];
  return list.slice(0, lim);
}

/**
 * Agrega un valor a VALIDACIONES (evita duplicados activos).
 */
function addValidationValue(campo, valor) {
  const c = String(campo || "").trim();
  const v = String(valor || "").trim();
  if (!c) throw new Error("Campo inválido.");
  if (!v) throw new Error("El valor está vacío.");

  const sh = ensureSheetHeaders_(CFG_ADMIN.validacionesSheetName, [
    "campo",
    "valor",
    "activo",
    "updated_at",
    "update_user",
  ]);

  const data = sh.getDataRange().getValues();
  const headers = (data[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const iCampo = headers.indexOf("campo");
  const iValor = headers.indexOf("valor");
  const iAct = headers.indexOf("activo");

  const rows = data.slice(1);

  const exists = rows.some((r) => {
    const sameCampo = norm_(r[iCampo]) === norm_(c);
    const sameValor = norm_(r[iValor]) === norm_(v);
    const activeOk = iAct === -1 ? true : String(r[iAct] || "").trim().toUpperCase() !== "NO";
    return sameCampo && sameValor && activeOk;
  });

  if (exists) return { ok: true, created: false, message: "Ya existía (no se duplicó).", campo: c, valor: v };

  const now = new Date();
  const userEmail = safeUserEmail_();

  sh.appendRow([c, v, "SI", now, userEmail]);
  return { ok: true, created: true, message: "Agregado.", campo: c, valor: v };
}

/** Alias por compatibilidad */
function addValidation(campo, valor) {
  return addValidationValue(campo, valor);
}

/** Genera siguiente ID con prefijo + padding */
function nextId_(sheetName, idHeaderCandidates, prefix, pad) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`No existe la hoja "${sheetName}"`);

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return `${prefix}-${String(1).padStart(pad, "0")}`;

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const findIdx = (cands) => {
    for (const c of cands) {
      const idx = headers.indexOf(String(c).toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const idIdx = findIdx(idHeaderCandidates);
  const colIdx = idIdx !== -1 ? idIdx : 0;

  let maxN = 0;
  rows.forEach((r) => {
    const raw = String(r[colIdx] || "").trim();
    if (!raw) return;
    const m = raw.match(new RegExp(`^${prefix}-(\\d+)$`));
    if (!m) return;
    const n = Number(m[1]);
    if (!isNaN(n)) maxN = Math.max(maxN, n);
  });

  const next = maxN + 1;
  return `${prefix}-${String(next).padStart(pad, "0")}`;
}

/** Genera N IDs consecutivos (usa nextId_ como base) */
function nextIds_(sheetName, idHeaderCandidates, prefix, pad, count) {
  const n = Math.max(1, Number(count || 1));
  const first = nextId_(sheetName, idHeaderCandidates, prefix, pad);
  if (n === 1) return [first];

  const m = String(first).match(/^([A-Z]+)-(\d+)$/i);
  if (!m) return [first];

  const pref = m[1];
  const start = Number(m[2]);
  const width = m[2].length;

  const out = [];
  for (let i = 0; i < n; i++) {
    out.push(`${pref}-${String(start + i).padStart(width, "0")}`);
  }
  return out;
}

/**
 * CREATE Ente Privado — ENTIDADES_PRIVADAS
 */
function ensureEntidadesPrivadasHeaders_() {
  const headersFinal = [
    "id_ente",
    "ente_nombre",
    "sector",
    "rubro",
    "web",
    "zona_cobertura",
    "zona_comuna",
    "direccion",
    "red_discapacidad",
    "observaciones",
    "activo",
    "updated_at",
    "fecha_contacto",
    "contacto",
  ];

  const sh = ensureSheetHeaders_(CFG_ADMIN.entesSheetName, headersFinal);

  const lastCol = sh.getLastColumn();
  const r1 = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const existing = r1.map(h => normalizeHeader_(h)).filter(Boolean);

  const missing = headersFinal.filter(h => !existing.includes(h));
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }

  return sh;
}

function createEntePrivado(ente) {
  const e = ente || {};

  const ente_nombre = String(e.ente_nombre || "").trim();
  if (ente_nombre.length < 3) throw new Error("Nombre de ente muy corto (mín 3).");

  const zona_cobertura = String(e.zona_cobertura || "").trim();
  const zona_comuna = String(e.zona_comuna || "").trim();
  const fecha_contacto = coerceToISODate_(e.fecha_contacto || "");

  const sector = String(e.sector || "").trim();
  const rubro = String(e.rubro || "").trim();
  const web = String(e.web || "").trim();
  const direccion = String(e.direccion || "").trim();
  const red_discapacidad = String(e.red_discapacidad || "").trim();
  const observaciones = String(e.observaciones || "").trim();
  const activo = String(e.activo || "SI").trim().toUpperCase() === "NO" ? "NO" : "SI";

  // referencia a persona contacto (PE-XXXX)
  const contacto = String(e.contacto || "").trim();

  const sh = ensureEntidadesPrivadasHeaders_();

  const lastCol = sh.getLastColumn();
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader_(h));
  const idx = {};
  hdr.forEach((h, i) => { if (h) idx[h] = i; });

  if (idx["ente_nombre"] === undefined) throw new Error(`La hoja "${CFG_ADMIN.entesSheetName}" no tiene la columna "ente_nombre".`);

  // evitar duplicado por nombre (activo != NO)
  const data = sh.getDataRange().getValues();
  const rows = data.slice(1);
  const iNombre = idx["ente_nombre"];
  const iAct = idx["activo"];

  const exists = rows.some((r) => {
    const same = norm_(r[iNombre]) === norm_(ente_nombre);
    const activeOk = iAct === undefined ? true : String(r[iAct] || "").trim().toUpperCase() !== "NO";
    return same && activeOk;
  });

  if (exists) return { ok: true, created: false, message: "Ya existía (no se duplicó).", ente_nombre };

  const id = nextId_(CFG_ADMIN.entesSheetName, ["id_ente", "id", "codigo"], "EN", 3);
  const now = new Date();

  const row = new Array(lastCol).fill("");
  row[idx["id_ente"]] = id;
  row[idx["ente_nombre"]] = ente_nombre;
  if (idx["sector"] !== undefined) row[idx["sector"]] = sector;
  if (idx["rubro"] !== undefined) row[idx["rubro"]] = rubro;
  if (idx["web"] !== undefined) row[idx["web"]] = web;

  if (idx["zona_cobertura"] !== undefined) row[idx["zona_cobertura"]] = zona_cobertura;
  if (idx["zona_comuna"] !== undefined) row[idx["zona_comuna"]] = zona_comuna;

  if (idx["direccion"] !== undefined) row[idx["direccion"]] = direccion;
  if (idx["red_discapacidad"] !== undefined) row[idx["red_discapacidad"]] = red_discapacidad;
  if (idx["observaciones"] !== undefined) row[idx["observaciones"]] = observaciones;
  if (idx["activo"] !== undefined) row[idx["activo"]] = activo;
  if (idx["updated_at"] !== undefined) row[idx["updated_at"]] = now;

  if (idx["fecha_contacto"] !== undefined) row[idx["fecha_contacto"]] = fecha_contacto || "";
  if (idx["contacto"] !== undefined) row[idx["contacto"]] = contacto || "";

  sh.appendRow(row);
  const rowNumber = sh.getLastRow();

  return { ok: true, created: true, id, ente_nombre, rowNumber };
}

/** Alias compat */
function createEnte(ente) { return createEntePrivado(ente); }

/** Preview Entes (top N) */
function adminPreviewEntes(limit) {
  const lim = Math.min(Math.max(Number(limit || 30), 1), 200);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.entesSheetName);
  if (!sh) return [];

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const idx = (k) => headers.indexOf(k);
  const iId = idx("id_ente");
  const iNombre = idx("ente_nombre");
  const iAct = idx("activo");

  return rows
    .map((r) => ({
      id: iId === -1 ? "" : String(r[iId] || "").trim(),
      name: iNombre === -1 ? "" : String(r[iNombre] || "").trim(),
      activo: iAct === -1 ? "SI" : String(r[iAct] || "SI").trim().toUpperCase(),
    }))
    .filter((x) => x.id && x.name && x.activo !== "NO")
    .sort((a, b) => a.name.localeCompare(b.name, "es"))
    .slice(0, lim);
}

/** CREATE Área GCBA — AREAS_GCBA */
function createAreaGCBA(area_nombre) {
  const name = String(area_nombre || "").trim();
  if (name.length < 3) throw new Error("Nombre de área muy corto (mín 3).");

  const sh = ensureSheetHeaders_(CFG_ADMIN.areasSheetName, [
    "id_area",
    "area_nombre",
    "sigla",
    "mail",
    "telefono",
    "activo",
    "updated_at",
    "update_user",
  ]);

  const v = sh.getDataRange().getValues();
  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const nameIdx = headers.indexOf("area_nombre");
  const actIdx = headers.indexOf("activo");

  const exists = rows.some((r) => {
    const activeOk = actIdx === -1 ? true : String(r[actIdx] || "").trim().toUpperCase() !== "NO";
    return activeOk && norm_(r[nameIdx]) === norm_(name);
  });

  if (exists) return { ok: true, created: false, message: "Ya existía (no se duplicó).", name };

  const id = nextId_(CFG_ADMIN.areasSheetName, ["id_area", "id", "codigo"], "AR", 3);

  const now = new Date();
  const userEmail = safeUserEmail_();

  sh.appendRow([id, name, "", "", "", "SI", now, userEmail]);
  const rowNumber = sh.getLastRow();

  return { ok: true, created: true, id, name, rowNumber };
}

/** Alias compat */
function createArea(name) { return createAreaGCBA(name); }

/** Preview Áreas (top N) */
function adminPreviewAreas(limit) {
  const lim = Math.min(Math.max(Number(limit || 30), 1), 200);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.areasSheetName);
  if (!sh) return [];

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const idx = (k) => headers.indexOf(k);
  const iId = idx("id_area");
  const iNombre = idx("area_nombre");
  const iAct = idx("activo");

  return rows
    .map((r) => ({
      id: iId === -1 ? "" : String(r[iId] || "").trim(),
      name: iNombre === -1 ? "" : String(r[iNombre] || "").trim(),
      activo: iAct === -1 ? "SI" : String(r[iAct] || "SI").trim().toUpperCase(),
    }))
    .filter((x) => x.id && x.name && x.activo !== "NO")
    .sort((a, b) => a.name.localeCompare(b.name, "es"))
    .slice(0, lim);
}

/** CREATE Persona — PERSONAS (headers extendidos) */
function ensurePersonasHeadersForContacto_() {
  const headersFinal = [
    "id_persona",
    "apellido",
    "nombre",
    "dni",
    "mail",
    "telefono",

    // compat
    "area_id",
    "rol",

    // nuevos para contacto
    "cargo",
    "observaciones",
    "red_discapacidad",
    "gcba",
    "ente_nombre_ref",

    "activo",
    "updated_at",
    "update_user",
  ];

  const sh = ensureSheetHeaders_(CFG_ADMIN.personasSheetName, headersFinal);

  const lastCol = sh.getLastColumn();
  const r1 = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const existing = r1.map(h => normalizeHeader_(h)).filter(Boolean);

  const missing = headersFinal.filter(h => !existing.includes(h));
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }

  return sh;
}

/**
 * ✅ ENDPOINT PRINCIPAL NUEVO (ENTE + CONTACTO ATÓMICO)
 * payload = { entidad: {...}, contacto: {...} }
 */
function createEntidadPrivadaWithContacto(payload){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let enteRow = null;
  let personaRow = null;
  let personaRows = [];

  try {
    payload = payload || {};
    const entidad  = payload.entidad  || payload.ente || payload.entidad_privada || {};
    const contactos = Array.isArray(payload.contactos)
      ? payload.contactos
      : (payload.contacto ? [payload.contacto] : []);

    // =========================
    // 1) VALIDACIONES COMPLETAS
    // =========================
    const reqEnte = ["ente_nombre", "sector", "rubro", "zona_cobertura", "zona_comuna"];
    const reqPersona = ["nombre", "apellido"]; // si DNI obligatorio -> agregá "dni"

    reqEnte.forEach(k => {
      if (!String(entidad[k] || "").trim()) throw new Error(`Entidad: falta "${k}".`);
    });
    if (!contactos.length) throw new Error("Debe cargar al menos un contacto.");
    contactos.forEach((c, i) => {
      reqPersona.forEach(k => {
        if (!String(c?.[k] || "").trim()) throw new Error(`Contacto ${i+1}: falta "${k}".`);
      });
    });

    const enteNombre = String(entidad.ente_nombre || "").trim();
    const dnis  = contactos.map(c => String(c?.dni || "").trim()).filter(Boolean);
    const mails = contactos.map(c => String(c?.mail || "").trim()).filter(Boolean);

    // =========================
    // 2) PRECHECKS (ANTES DE ESCRIBIR)
    // =========================
    if (enteNombreExistsActive_(enteNombre)) {
      throw new Error("Ya existía una entidad con ese nombre (activa).");
    }
    dnis.forEach(dni => {
      if (personaDniExistsActive_(dni)) throw new Error(`Ya existía una persona con el DNI ${dni}.`);
    });
    mails.forEach(mail => {
      if (personaMailExistsActive_(mail)) throw new Error(`Ya existía una persona con el mail ${mail}.`);
    });

    // =========================
    // 3) RESERVA ID PERSONA
    // =========================
    const peIds = nextIds_(CFG_ADMIN.personasSheetName, ["id_persona","id","codigo"], "PE", 4, contactos.length);
    const contactoStr = peIds.join(" | ");

    // =========================
    // 4) ESCRIBIR ENTE + PERSONA
    // =========================

    // 4.1 ENTE con contacto PE-XXXX
    const enteRes = createEntePrivado({
      ...entidad,
      contacto: contactoStr,
    });

    if (!enteRes?.ok || !enteRes?.created) {
      throw new Error(enteRes?.message || "No se pudo crear la entidad privada.");
    }
    enteRow = enteRes.rowNumber || null;

    // 4.2 PERSONAS (varios contactos)
    ensurePersonasHeadersForContacto_();
    personaRows = [];
    for (let i = 0; i < contactos.length; i++){
      const c = contactos[i] || {};
      const personaRes = createPersona({
        id_persona: peIds[i],
        nombre: String(c.nombre || "").trim(),
        apellido: String(c.apellido || "").trim(),
        dni: String(c.dni || "").trim(),
        mail: String(c.mail || "").trim(),
        telefono: String(c.telefono || "").trim(),
        rol: String(c.rol || "").trim(),
        cargo: String(c.cargo || "").trim(),
        observaciones: String(c.observaciones || "").trim(),
        red_discapacidad: String(c.red_discapacidad || "").trim(),
        gcba: String(c.gcba || "").trim().toUpperCase() === "SI" ? "SI" : "NO",
        area_id: "",
        ente_nombre_ref: enteNombre,
      });

      if (!personaRes?.ok || !personaRes?.created) {
        throw new Error(personaRes?.message || "No se pudo crear la persona contacto.");
      }
      personaRows.push(personaRes.rowNumber || null);
    }
    personaRow = personaRows.filter(Boolean).slice(-1)[0] || null;

    return {
      ok: true,
      id_ente: enteRes.id,
      persona_ids: peIds,
      message: "Entidad privada + contactos creados.",
    };

  } catch (err) {
    // =========================
    // 5) ROLLBACK
    // =========================
    try { personaRows.forEach(rn => rollbackRow_(CFG_ADMIN.personasSheetName, rn)); } catch(_){}
    try { if (enteRow) rollbackRow_(CFG_ADMIN.entesSheetName, enteRow); } catch(_){}
    throw err;

  } finally {
    lock.releaseLock();
  }
}

/**
 * ✅ ENDPOINT: crea Área + Persona contacto (ATÓMICO)
 * payload = { area: { area_nombre }, contacto: {...} }
 */
function createAreaWithContacto(payload){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let areaRow = null;
  let personaRow = null;
  let personaRows = [];

  try {
    payload = payload || {};
    const area = payload.area || payload.entidad || payload.area_gcba || {};
    const contactos = Array.isArray(payload.contactos)
      ? payload.contactos
      : (payload.contacto ? [payload.contacto] : []);

    // =========================
    // 1) VALIDACIONES
    // =========================
    const area_nombre = String(area.area_nombre || "").trim();
    if (!area_nombre) throw new Error('Área: falta "area_nombre".');

    if (!contactos.length) throw new Error("Debe cargar al menos un contacto.");
    contactos.forEach((c, i) => {
      const nombre = String(c?.nombre || "").trim();
      const apellido = String(c?.apellido || "").trim();
      if (!nombre || !apellido) throw new Error(`Contacto ${i+1}: nombre y apellido son obligatorios.`);
    });

    const dnis  = contactos.map(c => String(c?.dni || "").trim()).filter(Boolean);
    const mails = contactos.map(c => String(c?.mail || "").trim()).filter(Boolean);

    // =========================
    // 2) PRECHECKS (ANTES DE ESCRIBIR)
    // =========================
    if (areaNombreExistsActive_(area_nombre)) throw new Error("Ya existía un área con ese nombre (activa).");
    dnis.forEach(dni => {
      if (personaDniExistsActive_(dni)) throw new Error(`Ya existía una persona con el DNI ${dni}.`);
    });
    mails.forEach(mail => {
      if (personaMailExistsActive_(mail)) throw new Error(`Ya existía una persona con el mail ${mail}.`);
    });

    // =========================
    // 3) RESERVA ID PERSONA
    // =========================
    const peIds = nextIds_(CFG_ADMIN.personasSheetName, ["id_persona","id","codigo"], "PE", 4, contactos.length);

    // =========================
    // 4) CREAR ÁREA (solo nombre)
    // =========================
    const areaRes = createAreaGCBA(area_nombre);
    if (!areaRes?.ok || !areaRes?.created) throw new Error(areaRes?.message || "No se pudo crear el área.");
    areaRow = areaRes.rowNumber || null;

    // =========================
    // 5) CREAR PERSONA (con area_id = AR-XXX)
    // =========================
    ensurePersonasHeadersForContacto_();

    personaRows = [];
    for (let i = 0; i < contactos.length; i++){
      const c = contactos[i] || {};
      const personaRes = createPersona({
        id_persona: peIds[i],
        nombre: String(c.nombre || "").trim(),
        apellido: String(c.apellido || "").trim(),
        dni: String(c.dni || "").trim(),
        mail: String(c.mail || "").trim(),
        telefono: String(c.telefono || "").trim(),
        rol: String(c.rol || "").trim(),
        cargo: String(c.cargo || "").trim(),
        observaciones: String(c.observaciones || "").trim(),
        red_discapacidad: String(c.red_discapacidad || "").trim(),
        gcba: "SI",
        area_id: areaRes.id,
        ente_nombre_ref: "",
      });

      if (!personaRes?.ok || !personaRes?.created) throw new Error(personaRes?.message || "No se pudo crear la persona contacto del área.");
      personaRows.push(personaRes.rowNumber || null);
    }
    personaRow = personaRows.filter(Boolean).slice(-1)[0] || null;

    // set contactos en área si existe columna
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(CFG_ADMIN.areasSheetName);
      const lastCol = sh.getLastColumn();
      const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader_(h));
      const idx = {};
      hdr.forEach((h, i)=>{ if (h) idx[h] = i + 1; });
      if (idx["contacto"]) sh.getRange(areaRow, idx["contacto"]).setValue(peIds.join(" | "));
      if (idx["fecha_contacto"]) sh.getRange(areaRow, idx["fecha_contacto"]).setValue(new Date());
    } catch(_){}

    return {
      ok: true,
      id_area: areaRes.id,
      persona_ids: peIds,
      message: "Área + contactos creados.",
    };

  } catch (err) {
    try { personaRows.forEach(rn => rollbackRow_(CFG_ADMIN.personasSheetName, rn)); } catch(_){}
    try { if (areaRow) rollbackRow_(CFG_ADMIN.areasSheetName, areaRow); } catch(_){}
    throw err;

  } finally {
    lock.releaseLock();
  }
}

function createPersona(person) {
  const p = person || {};
  const apellido = String(p.apellido || "").trim();
  const nombre = String(p.nombre || "").trim();
  if (apellido.length < 2 || nombre.length < 2) throw new Error("Apellido y nombre son obligatorios (mín 2).");

  const sh = ensurePersonasHeadersForContacto_();

  const v = sh.getDataRange().getValues();
  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const idx = (k) => headers.indexOf(k);
  const iDni = idx("dni");
  const iMail = idx("mail");
  const actIdx = idx("activo");

  const dni = String(p.dni || "").trim();
  const mail = String(p.mail || "").trim();

  if (dni && iDni !== -1) {
    const dupDni = rows.some((r) => {
      const activeOk = actIdx === -1 ? true : String(r[actIdx] || "").trim().toUpperCase() !== "NO";
      return activeOk && String(r[iDni] || "").trim() === dni;
    });
    if (dupDni) return { ok: true, created: false, message: "Ya existía una persona con ese DNI.", dni };
  }

  if (mail && iMail !== -1) {
    const dupMail = rows.some((r) => {
      const activeOk = actIdx === -1 ? true : String(r[actIdx] || "").trim().toUpperCase() !== "NO";
      return activeOk && norm_(r[iMail]) === norm_(mail);
    });
    if (dupMail) return { ok: true, created: false, message: "Ya existía una persona con ese mail.", mail };
  }

  const forcedId = String(p.id_persona || "").trim();
  const id = forcedId || nextId_(CFG_ADMIN.personasSheetName, ["id_persona", "id", "codigo"], "PE", 4);

  const now = new Date();
  const userEmail = safeUserEmail_();

  const lastCol = sh.getLastColumn();
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader_(h));
  const idxH = {};
  hdr.forEach((h, i) => { if (h) idxH[h] = i; });

  const row = new Array(lastCol).fill("");

  row[idxH["id_persona"]] = id;
  if (idxH["apellido"] !== undefined) row[idxH["apellido"]] = apellido;
  if (idxH["nombre"] !== undefined) row[idxH["nombre"]] = nombre;
  if (idxH["dni"] !== undefined) row[idxH["dni"]] = dni || "";
  if (idxH["mail"] !== undefined) row[idxH["mail"]] = mail || "";
  if (idxH["telefono"] !== undefined) row[idxH["telefono"]] = String(p.telefono || "").trim() || "";

  if (idxH["area_id"] !== undefined) row[idxH["area_id"]] = String(p.area_id || "").trim() || "";
  if (idxH["rol"] !== undefined) row[idxH["rol"]] = String(p.rol || "").trim() || "";

  if (idxH["cargo"] !== undefined) row[idxH["cargo"]] = String(p.cargo || "").trim() || "";
  if (idxH["observaciones"] !== undefined) row[idxH["observaciones"]] = String(p.observaciones || "").trim() || "";
  if (idxH["red_discapacidad"] !== undefined) row[idxH["red_discapacidad"]] = String(p.red_discapacidad || "").trim() || "";
  if (idxH["gcba"] !== undefined) row[idxH["gcba"]] = String(p.gcba || "").trim().toUpperCase() === "SI" ? "SI" : "NO";
  if (idxH["ente_nombre_ref"] !== undefined) row[idxH["ente_nombre_ref"]] = String(p.ente_nombre_ref || "").trim() || "";

  if (idxH["activo"] !== undefined) row[idxH["activo"]] = "SI";
  if (idxH["updated_at"] !== undefined) row[idxH["updated_at"]] = now;
  if (idxH["update_user"] !== undefined) row[idxH["update_user"]] = userEmail;

  sh.appendRow(row);
  const rowNumber = sh.getLastRow();

  return { ok: true, created: true, id, apellido, nombre, rowNumber };
}

/** Alias compat */
function createPersonaAdmin(p) { return createPersona(p); }

/** Preview Personas (top N) */
function adminPreviewPersonas(limit) {
  const lim = Math.min(Math.max(Number(limit || 30), 1), 200);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.personasSheetName);
  if (!sh) return [];

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const headers = (v[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const rows = v.slice(1);

  const idx = (k) => headers.indexOf(k);
  const iId = idx("id_persona");
  const iApe = idx("apellido");
  const iNom = idx("nombre");
  const iDni = idx("dni");
  const iMail = idx("mail");
  const iAct = idx("activo");

  return rows
    .map((r) => ({
      id: iId === -1 ? "" : String(r[iId] || "").trim(),
      name: `${String(iApe === -1 ? "" : r[iApe] || "").trim()}, ${String(iNom === -1 ? "" : r[iNom] || "").trim()}`.trim(),
      dni: iDni === -1 ? "" : String(r[iDni] || "").trim(),
      mail: iMail === -1 ? "" : String(r[iMail] || "").trim(),
      activo: iAct === -1 ? "SI" : String(r[iAct] || "SI").trim().toUpperCase(),
    }))
    .filter((x) => x.id && x.name && x.name.replace(",", "").trim() !== "" && x.activo !== "NO")
    .sort((a, b) => a.name.localeCompare(b.name, "es"))
    .slice(0, lim);
}

/* =========================================================
 * ===================== VALIDACIONES (shared) ==============
 * ========================================================= */
function readValidaciones_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_VALIDACIONES);
  if (!sh) return {};

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 2) return {};

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const headers = values[0].map((h) => normalizeHeader_(h));

  const hasCampo = headers.includes("campo");
  const hasValor = headers.includes("valor");

  // ✅ formato vertical: campo/valor(/activo)
  if (hasCampo && hasValor) {
    const iCampo = headers.indexOf("campo");
    const iValor = headers.indexOf("valor");
    const iActivo = headers.indexOf("activo");
    const out = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const campoRaw = String(row[iCampo] || "").trim();
      const valorRaw = String(row[iValor] || "").trim();
      if (!campoRaw || !valorRaw) continue;

      const activo = iActivo >= 0 ? String(row[iActivo] || "").trim().toUpperCase() : "SI";
      if (activo && activo !== "SI") continue;

      const key = normalizeValidationKey_(campoRaw);
      if (!out[key]) out[key] = [];
      out[key].push(valorRaw);
    }

    Object.keys(out).forEach((k) => (out[k] = uniq_(out[k])));

    if (!out["estadoaccion"] || !out["estadoaccion"].length) out["estadoaccion"] = CFG_MAIN.FALLBACK_ESTADOS_ACCION.slice();
    if (!out["contratacionpersonal"] || !out["contratacionpersonal"].length) out["contratacionpersonal"] = CFG_MAIN.FALLBACK_SI_NO.slice();

    return out;
  }

  // ✅ formato horizontal: columnas
  const idx = {};
  headers.forEach((h, i) => {
    if (h) idx[h] = i;
  });

  const wanted = [
    "tipo_accion",
    "eje",
    "lineaTematica",
    "produccion",
    "territorio",
    "contratacion_personal",
    "meta_tipo",
    "estado_invitado",
    "estado_accion",
  ];

  const out = {};
  for (const k of wanted) {
    if (idx[k] === undefined) continue;
    const col = idx[k];
    const arr = [];
    for (let r = 1; r < values.length; r++) {
      const v = String(values[r][col] || "").trim();
      if (v) arr.push(v);
    }
    out[k] = uniq_(arr);
  }

  if (!out["estado_accion"] || !out["estado_accion"].length) out["estado_accion"] = CFG_MAIN.FALLBACK_ESTADOS_ACCION.slice();
  if (!out["contratacion_personal"] || !out["contratacion_personal"].length) out["contratacion_personal"] = CFG_MAIN.FALLBACK_SI_NO.slice();

  return out;
}

/* =========================================================
 * ===================== CATÁLOGOS (shared) =================
 * ========================================================= */
function readCatalog_(sheetName, idColName, nameColName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return [];

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const headers = values[0].map((h) => normalizeHeader_(h));
  const idx = {};
  headers.forEach((h, i) => {
    if (h) idx[h] = i;
  });

  const idKey = normalizeHeader_(idColName);
  const nameKey = normalizeHeader_(nameColName);
  const actKey = idx["activo"] !== undefined ? "activo" : null;

  if (idx[idKey] === undefined || idx[nameKey] === undefined) return [];

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idx[idKey]] || "").trim();
    const name = String(row[idx[nameKey]] || "").trim();
    if (!id) continue;

    if (actKey) {
      const act = String(row[idx[actKey]] || "SI").trim().toUpperCase();
      if (act === "NO") continue;
    }

    out.push({ id, name });
  }

  out.sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "es"));
  return out;
}

function readPersonasCatalog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_PERSONAS);
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return [];

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const headers = values[0].map((h) => normalizeHeader_(h));
  const idx = {};
  headers.forEach((h, i) => {
    if (h) idx[h] = i;
  });

  const idI = idx["id_persona"];
  const nomI = idx["nombre"] ?? idx["nombres"] ?? idx["name"];
  const apeI = idx["apellido"] ?? idx["apellidos"] ?? idx["last_name"];
  const actI = idx["activo"];

  if (idI === undefined) return [];

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[idI] || "").trim();
    if (!id) continue;

    if (actI !== undefined) {
      const act = String(row[actI] || "SI").trim().toUpperCase();
      if (act === "NO") continue;
    }

    const nombre = nomI !== undefined ? String(row[nomI] || "").trim() : "";
    const apellido = apeI !== undefined ? String(row[apeI] || "").trim() : "";
    const full = [apellido, nombre].filter(Boolean).join(", ") || nombre || apellido || id;

    out.push({ id, name: full });
  }

  out.sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "es"));
  return out;
}

/* =========================================================
 * ===================== HELPERS (shared) ===================
 * ========================================================= */
function safeUserEmail_() {
  try {
    const a = Session.getActiveUser && Session.getActiveUser().getEmail();
    const e = Session.getEffectiveUser && Session.getEffectiveUser().getEmail();
    return (a || e || "").trim();
  } catch (e) {
    return "";
  }
}

/**
 * ✅ Normalización fuerte de headers:
 */
function normalizeHeader_(h) {
  let s = String(h || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_");

  const ALIASES = {
    "linea_tematica": "lineaTematica",
    "lineatematica": "lineaTematica",
    "subeje": "lineaTematica",
  };

  if (ALIASES[s]) return ALIASES[s];
  return s;
}

function normalizeValidationKey_(campoRaw) {
  return String(campoRaw || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .replace(/_/g, "");
}

function uniq_(arr) {
  const seen = new Set();
  const out = [];
  for (const x of arr || []) {
    const v = String(x ?? "").trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  return out;
}

function pickValidation_(validations, keys) {
  if (!validations) return [];
  for (const k of keys) {
    const nk = normalizeValidationKey_(k);
    if (Array.isArray(validations[nk]) && validations[nk].length) return validations[nk];

    const lk = normalizeHeader_(k);
    if (Array.isArray(validations[lk]) && validations[lk].length) return validations[lk];
  }
  return [];
}

function formatByType_(col, v) {
  if (v === null || v === undefined) return "";
  if (CFG_MAIN.DATE_ONLY.has(col)) return toISODate_(v);
  if (typeof v === "number") return String(v);
  return String(v).trim();
}

function toISODate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    const yyyy = v.getFullYear();
    const mm = String(v.getMonth() + 1).padStart(2, "0");
    const dd = String(v.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function coerceToISODate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return toISODate_(v);
  }
  const s = String(v || "").trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const mm = String(m[2]).padStart(2, "0");
    const yyyy = m[3];
    return `${yyyy}-${mm}-${dd}`;
  }

  const t = Date.parse(s);
  if (!isNaN(t)) return toISODate_(new Date(t));

  return "";
}

function normalizeEstado_(s) {
  const up = String(s || "").trim().toUpperCase();
  if (!up) return "";
  if (up === "REALIZADA") return "REALIZADO";
  if (up === "CANCELADA") return "CANCELADO";
  if (up === "IDEA" || up === "DESARROLLO" || up === "REALIZADO" || up === "CANCELADO") return up;
  return up;
}

function toMap_(arr) {
  const m = {};
  (arr || []).forEach((x) => {
    if (x?.id) m[x.id] = x.name || "";
  });
  return m;
}
function isoToDate_(iso){
  const s = String(iso||"").trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(y, mo-1, d);
  if (isNaN(dt.getTime())) return null;
  return dt;
}

/**
 * Colores estables por linea temática (hash -> HSL -> HEX).
 * Devuelve un mapa linea->#RRGGBB y un "__default".
 */
function buildColorsByLinea_(lineas){
  const out = {};
  const arr = Array.isArray(lineas) ? lineas : [];
  arr.forEach((name)=>{
    const key = String(name||"").trim();
    if (!key) return;
    out[key] = hashColor_(key);
  });
  out["__default"] = "#DDE6EE";
  return out;
}

function hashColor_(s){
  // hash simple estable
  const str = String(s||"");
  let h = 0;
  for (let i=0; i<str.length; i++){
    h = ((h << 5) - h) + str.charCodeAt(i);
    h |= 0;
  }
  // hue 0..360 (estable), sat/light fijos para look suave
  const hue = Math.abs(h) % 360;
  return hslToHex_(hue, 55, 85); // pastel claro para fondo
}

function hslToHex_(h, s, l){
  s /= 100; l /= 100;
  const c = (1 - Math.abs(2*l - 1)) * s;
  const x = c * (1 - Math.abs(((h/60) % 2) - 1));
  const m = l - c/2;
  let r=0,g=0,b=0;

  if (0 <= h && h < 60) { r=c; g=x; b=0; }
  else if (60 <= h && h < 120) { r=x; g=c; b=0; }
  else if (120 <= h && h < 180) { r=0; g=c; b=x; }
  else if (180 <= h && h < 240) { r=0; g=x; b=c; }
  else if (240 <= h && h < 300) { r=x; g=0; b=c; }
  else { r=c; g=0; b=x; }

  const toHex = (v) => {
    const n = Math.round((v + m) * 255);
    return n.toString(16).padStart(2,"0");
  };
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
}

/** ===== Payload helpers (normaliza keys) ===== */
function getPayload_(payload, key) {
  if (!payload) return undefined;
  const k = normalizeHeader_(key);
  if (payload[k] !== undefined) return payload[k];
  if (payload[key] !== undefined) return payload[key];
  for (const kk in payload) {
    if (normalizeHeader_(kk) === k) return payload[kk];
  }
  return undefined;
}

function setPayload_(payload, key, value) {
  if (!payload) return;
  const k = normalizeHeader_(key);
  payload[k] = value;
}

/** JSON stringify robusto */
function coerceToJsonString_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return String(v).trim();
  try {
    return JSON.stringify(v);
  } catch (e) {
    return "";
  }
}

/** ===== ID generation helpers ===== **/
function buildAccionId_(fechaISO, tipoAccion, nombreAccion) {
  const m = String(fechaISO).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const datePart = m ? `${m[1]}_${m[2]}_${m[3]}` : slug_(fechaISO);
  const tipoPart = slug_(tipoAccion);
  const nombrePart = slug_(nombreAccion);
  return [datePart, tipoPart, nombrePart].filter(Boolean).join("_");
}

function ensureUniqueId_(base, setExisting) {
  if (!setExisting.has(base)) return base;
  let i = 2;
  while (setExisting.has(`${base}_${i}`)) i++;
  return `${base}_${i}`;
}

function slug_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_");
}

function formatTimestamp_(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

/* =========================================================
 * ===== helpers para rollback / checks (NO romper nada) =====
 * ========================================================= */
function rollbackRow_(sheetName, rowNumber){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const rn = Number(rowNumber);
  if (rn && rn >= 2 && rn <= sh.getLastRow()) sh.deleteRow(rn);
}

function areaNombreExistsActive_(area_nombre){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.areasSheetName);
  if (!sh) return false;

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return false;

  const headers = (v[0]||[]).map(h=>String(h||"").trim().toLowerCase());
  const rows = v.slice(1);

  const iNombre = headers.indexOf("area_nombre");
  const iAct = headers.indexOf("activo");
  if (iNombre === -1) return false;

  const target = norm_(area_nombre);
  return rows.some(r=>{
    const activeOk = iAct === -1 ? true : String(r[iAct]||"").trim().toUpperCase() !== "NO";
    return activeOk && norm_(r[iNombre]) === target;
  });
}

function enteNombreExistsActive_(enteNombre){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.entesSheetName);
  if (!sh) return false;

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return false;

  const headers = (v[0]||[]).map(h=>String(h||"").trim().toLowerCase());
  const rows = v.slice(1);

  const iNombre = headers.indexOf("ente_nombre");
  const iAct = headers.indexOf("activo");
  if (iNombre === -1) return false;

  const target = norm_(enteNombre);
  return rows.some(r=>{
    const activeOk = iAct === -1 ? true : String(r[iAct]||"").trim().toUpperCase() !== "NO";
    return activeOk && norm_(r[iNombre]) === target;
  });
}

function personaDniExistsActive_(dni){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.personasSheetName);
  if (!sh) return false;

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return false;

  const headers = (v[0]||[]).map(h=>String(h||"").trim().toLowerCase());
  const rows = v.slice(1);

  const iDni = headers.indexOf("dni");
  const iAct = headers.indexOf("activo");
  if (iDni === -1) return false;

  const target = String(dni).trim();
  return rows.some(r=>{
    const activeOk = iAct === -1 ? true : String(r[iAct]||"").trim().toUpperCase() !== "NO";
    return activeOk && String(r[iDni]||"").trim() === target;
  });
}

function personaMailExistsActive_(mail){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_ADMIN.personasSheetName);
  if (!sh) return false;

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return false;

  const headers = (v[0]||[]).map(h=>String(h||"").trim().toLowerCase());
  const rows = v.slice(1);

  const iMail = headers.indexOf("mail");
  const iAct = headers.indexOf("activo");
  if (iMail === -1) return false;

  const target = norm_(mail);
  return rows.some(r=>{
    const activeOk = iAct === -1 ? true : String(r[iAct]||"").trim().toUpperCase() !== "NO";
    return activeOk && norm_(r[iMail]) === target;
  });
}

/**
 * Catálogo simple de ENTIDADES_PRIVADAS para UI si lo necesitás
 */
function listEntidades(payload){
  const kind = String(payload?.kind || "entidad_privada").trim();

  const cfg = ENTIDADES_MAP[kind];
  if (!cfg) return { ok:false, message:"kind inválido" };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.sheet);
  if (!sh) return { ok:false, message:`No existe la hoja ${cfg.sheet}` };

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { ok:true, kind, label: cfg.label, columns: cfg.columns, rows: [], meta:{count:0} };
  }

  // ✅ Index por header NORMALIZADO (robusto)
  const headerRaw = values[0].map(h => String(h || "").trim());
  const headerNorm = headerRaw.map(h => normalizeHeader_(h));

  const idx = {};
  headerNorm.forEach((h,i)=>{ if (h) idx[h] = i; });

  // columnas pedidas también normalizadas
  const colsNorm = cfg.columns.map(c => normalizeHeader_(c));

  const rows = [];
  for (let r=1; r<values.length; r++){
    const row = values[r];
    const empty = row.every(v => String(v ?? "").trim() === "");
    if (empty) continue;

    const obj = {};
    for (let k=0; k<cfg.columns.length; k++){
      const colOut = cfg.columns[k];       // nombre “bonito” para el front
      const colKey = colsNorm[k];          // clave normalizada para buscar en idx
      const j = idx[colKey];

      let v = (j === undefined) ? "" : row[j];

      if (v instanceof Date){
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[colOut] = (v === null || v === undefined) ? "" : String(v).trim();
    }

    // opcional: rowNumber para debug/acciones futuras
    obj.__rowNumber = r + 1;

    rows.push(obj);
  }

  return {
    ok:true,
    kind,
    label: cfg.label,
    columns: cfg.columns,
    rows,
    meta: { count: rows.length, sheet: cfg.sheet }
  };
}

/**
 * Actualiza una entidad / área / persona por rowNumber.
 * payload: { kind, rowNumber, updates }
 */
function updateEntidad(payload){
  payload = payload || {};
  const kind = String(payload.kind || "").trim();
  const rowNumber = Number(payload.rowNumber);
  const updates = payload.updates || {};

  if (!kind || !ENTIDADES_MAP[kind]) throw new Error("kind inválido");
  if (!rowNumber || rowNumber < 2) throw new Error("rowNumber inválido");

  const cfg = ENTIDADES_MAP[kind];
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(cfg.sheet);
  if (!sh) throw new Error(`No existe la hoja ${cfg.sheet}`);

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error("Hoja inválida");

  const headersRaw = sh.getRange(1,1,1,lastCol).getValues()[0];
  const headersNorm = headersRaw.map(h => normalizeHeader_(h));
  const idx = {};
  headersNorm.forEach((h,i)=>{ if (h) idx[h] = i + 1; });

  const allowed = new Set((cfg.columns || []).map(c => normalizeHeader_(c)));
  const locked = cfg.lock || new Set();

  let wrote = 0;
  for (const key in updates){
    const k = normalizeHeader_(key);
    if (locked.has(k)) continue;
    if (!allowed.has(k)) continue;

    const colNum = idx[k];
    if (!colNum) continue;

    let v = updates[key];
    if (v === null || v === undefined) v = "";
    v = String(v).trim();

    sh.getRange(rowNumber, colNum).setValue(v);
    wrote++;
  }

  const now = new Date();
  const user = safeUserEmail_();

  if (idx["updated_at"]) sh.getRange(rowNumber, idx["updated_at"]).setValue(now);
  if (idx["update_user"]) sh.getRange(rowNumber, idx["update_user"]).setValue(user);
  if (idx["user"]) sh.getRange(rowNumber, idx["user"]).setValue(user);

  return { ok:true, wrote, rowNumber, sheet: cfg.sheet };
}
function getMedicionesSummary(filters){
  filters = filters || {};
  const year = String(filters.year || "").trim(); // "2026" o ""
  const q = String(filters.q || "").trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_ACCIONES);
  if (!sh) throw new Error(`No existe la hoja "${CFG_MAIN.SHEET_ACCIONES}"`);

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, kpis:{ total:0, realizadas:0, desarrollo:0, idea:0, pendientes:0 }, years:[] };

  const values = sh.getRange(1,1,lr,lc).getValues();
  const headers = values[0].map(h => normalizeHeader_(h));
  const idx = {};
  headers.forEach((h,i)=>{ if(h) idx[h]=i; });

  const iFecha = idx["fecha_accion"];
  const iEstado = idx["estado_accion"];
  const iNombre = idx["nombre_accion"];
  const iId = idx["id_accion"];
  const iLinea = idx["lineaTematica"];
  const iAct = idx["activo"];

  if (iEstado === undefined) throw new Error("Falta columna estado_accion en ACCIONES");

  let total=0, realizadas=0, desarrollo=0, idea=0;
  const yearsSet = new Set();

  for (let r=1; r<values.length; r++){
    const row = values[r];

    // activo
    if (iAct !== undefined){
      const act = String(row[iAct] || "SI").trim().toUpperCase();
      if (act === "NO") continue;
    }

    const id = iId !== undefined ? String(row[iId]||"").trim() : "";
    const nom = iNombre !== undefined ? String(row[iNombre]||"").trim() : "";
    if (!id && !nom) continue;

    // año
    let rowYear = "";
    if (iFecha !== undefined){
      const iso = coerceToISODate_(row[iFecha]);
      if (iso && iso.length >= 4) rowYear = iso.slice(0,4);
      if (rowYear) yearsSet.add(rowYear);
      if (year && rowYear !== year) continue;
    } else {
      // si no hay fecha_accion, solo filtra si year viene vacío
      if (year) continue;
    }

    // búsqueda simple (nombre + linea + estado)
    if (q){
      const linea = iLinea !== undefined ? String(row[iLinea]||"").toLowerCase() : "";
      const estRaw = String(row[iEstado]||"");
      const hay = (nom+" "+linea+" "+estRaw).toLowerCase();
      if (!hay.includes(q)) continue;
    }

    total++;
    const est = normalizeEstado_(row[iEstado]);
    if (est === "REALIZADO") realizadas++;
    else if (est === "DESARROLLO") desarrollo++;
    else if (est === "IDEA") idea++;
  }

  const pendientes = total - realizadas;
  const yearsArr = Array.from(yearsSet).sort(); // "2024","2025","2026"

  return {
    ok:true,
    years: yearsArr,
    kpis: { total, realizadas, desarrollo, idea, pendientes }
  };
}
function exportAccionesXlsx(payload){
  payload = payload || {};
  const header = Array.isArray(payload.header) ? payload.header : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const sheetName = String(payload.sheetName || "Export");
  const meta = payload.meta || {};

  if (!header.length) throw new Error("Export: header vacío.");

  // 1) Crear spreadsheet temporal
  const ss = SpreadsheetApp.create("TMP_EXPORT_RS_ACCIONES");
  const sh = ss.getSheets()[0];
  sh.setName(sheetName);

  // 2) Escribir header + rows
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length){
    sh.getRange(2, 1, rows.length, header.length).setValues(
      rows.map(r => {
        const arr = Array.isArray(r) ? r : [];
        // asegurar largo exacto
        const out = new Array(header.length);
        for (let i=0;i<header.length;i++) out[i] = (arr[i] === undefined || arr[i] === null) ? "" : String(arr[i]);
        return out;
      })
    );
  }

  // (Opcional) Freeze + autosize
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, header.length); } catch(e){}

  // 3) Exportar como XLSX (Drive export)
  const fileId = ss.getId();
  const xlsxBlob = exportSpreadsheetAsXlsx_(fileId);

  // 4) Borrar el spreadsheet temporal (opcional, recomendado)
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e){}

  // 5) Nombre final del archivo
  const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const estado = String(meta.estado || "").trim();
  const q = String(meta.q || "").trim();
  const suffix = [estado ? "estado_" + estado : "", q ? "q_" + q.replace(/[^\w]+/g, "_").slice(0,30) : ""]
    .filter(Boolean).join("__");
  const filename = `RS_TableroAcciones_${ymd}${suffix ? "__" + suffix : ""}.xlsx`;

  return {
    filename,
    base64: Utilities.base64Encode(xlsxBlob.getBytes())
  };
}

function exportSpreadsheetAsXlsx_(spreadsheetId, filename) {
  filename = filename || "Tablero_Acciones.xlsx";

  const url =
    "https://www.googleapis.com/drive/v3/files/" +
    encodeURIComponent(spreadsheetId) +
    "/export?mimeType=" +
    encodeURIComponent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") +
    "&alt=media";

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error("Export XLSX falló. HTTP " + resp.getResponseCode() + ": " + resp.getContentText());
  }

  return resp.getBlob().setName(filename);
}
function getValidacionesFiltros(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG_MAIN.SHEET_VALIDACIONES_FILTROS || "VALIDACIONES_FILTROS");
  if (!sh) throw new Error('No existe la hoja VALIDACIONES_FILTROS');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return { ok:true, data:{ years:[], byYear:{} } };

  const values = sh.getRange(1,1,lr,lc).getValues();
  const headers = values[0].map(h => normalizeHeader_(h));

  const idx = (name) => headers.indexOf(normalizeHeader_(name));

  const cTipoYear = idx("Año_tipo_accion");
  const cTipoVal  = idx("tipo_accion_filtro");

  const cEjeYear  = idx("Año_eje");
  const cEjeVal   = idx("eje_filtro");

  const cLinYear  = idx("Año_lineaTematica");
  const cLinVal   = idx("lineaTematica_filtro");

  // fallback por posición como en tu screenshot (A..F)
  const P = {
    tipoYear: (cTipoYear !== -1 ? cTipoYear : 0),
    tipoVal:  (cTipoVal  !== -1 ? cTipoVal  : 1),
    ejeYear:  (cEjeYear  !== -1 ? cEjeYear  : 2),
    ejeVal:   (cEjeVal   !== -1 ? cEjeVal   : 3),
    linYear:  (cLinYear  !== -1 ? cLinYear  : 4),
    linVal:   (cLinVal   !== -1 ? cLinVal   : 5),
  };

  const byYear = {};
  const yearsSet = new Set();

  function push_(y, key, v){
    y = String(y||"").trim();
    v = String(v||"").trim();
    if (!y || !v) return;

    yearsSet.add(y);
    byYear[y] = byYear[y] || { tipo_accion:[], eje:[], lineaTematica:[] };
    byYear[y][key].push(v);
  }

  for (let r=1; r<values.length; r++){
    const row = values[r];
    push_(row[P.tipoYear], "tipo_accion", row[P.tipoVal]);
    push_(row[P.ejeYear],  "eje",         row[P.ejeVal]);
    push_(row[P.linYear],  "lineaTematica",row[P.linVal]);
  }

  const uniqKeepOrder = (arr) => {
    const s = new Set();
    const out = [];
    (arr || []).forEach(x => { x = String(x||"").trim(); if(x && !s.has(x)){ s.add(x); out.push(x); }});
    return out;
  };

  const byYearClean = {};
  Object.keys(byYear).forEach(y=>{
    byYearClean[y] = {
      tipo_accion:   uniqKeepOrder(byYear[y].tipo_accion),
      eje:          uniqKeepOrder(byYear[y].eje),
      lineaTematica:uniqKeepOrder(byYear[y].lineaTematica),
    };
  });

  const years = Array.from(yearsSet).sort((a,b)=>Number(a)-Number(b));
  return { ok:true, data:{ years, byYear: byYearClean } };
}
function getEjesFromValidacionesFiltros(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('VALIDACIONES_FILTROS');
  if (!sh) throw new Error('No existe la hoja VALIDACIONES_FILTROS');

  const lr = sh.getLastRow();
  if (lr < 2) return { ok:true, ejes:[] };

  // Col D = eje_filtro según tu screenshot
  const values = sh.getRange(2, 4, lr - 1, 1).getValues().flat();

  const seen = new Set();
  const ejes = [];
  values.forEach(v => {
    const s = String(v || '').trim();
    if (!s || seen.has(s)) return;
    seen.add(s);
    ejes.push(s);
  });

  return { ok:true, ejes };
}
