/**
 * Sur Consultores — Recepción de solicitudes de reunión y reservas Petrinovic
 *
 * Este código vive dentro de la cuenta de Google del usuario, asociado a una
 * planilla de Google Sheets. Recibe los envíos del formulario de la landing
 * page y de la Agenda Petrinovic, los anota como una fila nueva y avisa por
 * correo.
 *
 * Instalación: planilla -> Extensiones -> Apps Script -> pegar esto ->
 * Implementar -> Nueva implementación -> Aplicación web.
 */

// Correo que recibe el aviso de cada solicitud nueva.
const CORREO_AVISO = "fllancafil@surconsultores.org";

// Nombre de la pestaña donde se guardan las solicitudes.
const NOMBRE_HOJA = "Solicitudes";

const COLUMNAS = [
  "Fecha de la solicitud",
  "Nombre",
  "Empresa",
  "Correo",
  "Teléfono",
  "Industria",
  "Dotación",
  "Tema principal",
  "Fecha preferida",
  "Horario",
  "Mensaje",
  "Estado",
  "Origen",
];

// Agenda Petrinovic: cupos por día, sumando ambos servicios.
const CUPOS_POR_DIA = 50;
const HOJA_RESERVAS = "Reservas Petrinovic";

// Una reserva con este estado deja de ocupar cupos.
const ESTADO_ANULADA = "Anulada";

const COLUMNAS_RESERVAS = [
  "Fecha de la reserva",
  "Día reservado",
  "Servicio",
  "Cupos",
  "Empresa",
  "RUT",
  "Contacto",
  "Teléfono",
  "Correo",
  "Origen",
  "Estado",
];

/**
 * Recibe el formulario. La landing envía el contenido como texto plano para
 * evitar la verificación previa del navegador, que Apps Script no responde.
 */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);

    // Campo trampa: los robots lo rellenan, las personas no lo ven.
    if (datos.sitio_web) {
      return responder({ ok: true, ignorado: true });
    }

    if (datos.tipo === "reserva_petrinovic") {
      return responder(registrarReserva(datos));
    }

    const hoja = obtenerHoja();
    hoja.appendRow([
      new Date(),
      comoTexto(datos.nombre),
      comoTexto(datos.empresa),
      comoTexto(datos.email),
      comoTexto(datos.telefono),
      comoTexto(datos.industria),
      comoTexto(datos.dotacion),
      comoTexto(datos.tema),
      comoTexto(datos.fecha),
      comoTexto(datos.horario),
      comoTexto(datos.mensaje),
      "Nuevo",
      comoTexto(datos.origen || "Página web"),
    ]);

    enviarAviso(datos);
    return responder({ ok: true });
  } catch (error) {
    console.error(error);
    return responder({ ok: false, error: String(error) });
  }
}

/**
 * Con ?accion=disponibilidad entrega los cupos tomados por día para la Agenda
 * Petrinovic. Sin parámetros permite comprobar que la publicación quedó activa.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.accion === "disponibilidad") {
    return responder({ ok: true, capacidad: CUPOS_POR_DIA, reservados: cuposTomados() });
  }
  return responder({ ok: true, servicio: "Sur Consultores — solicitudes" });
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS);
    hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight("bold");
    hoja.setFrozenRows(1);
    hoja.setColumnWidth(1, 160);
    hoja.setColumnWidth(11, 320);
  }
  // Planillas creadas antes de agregar una columna: completa los títulos que falten.
  const titulos = hoja.getRange(1, 1, 1, COLUMNAS.length);
  const actuales = titulos.getValues()[0];
  if (actuales.some((valor, i) => valor !== COLUMNAS[i])) {
    titulos.setValues([COLUMNAS]).setFontWeight("bold");
  }
  return hoja;
}

/* ---------- Agenda Petrinovic ---------- */

/**
 * Anota una reserva si quedan cupos ese día. El candado evita que dos
 * empresas que reservan al mismo tiempo sobrepasen el máximo.
 */
function registrarReserva(datos) {
  const fecha = String(datos.fecha || "");
  const cupos = Number(datos.cupos);
  const zona = Session.getScriptTimeZone();
  const hoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  const dia = new Date(fecha + "T12:00:00").getDay();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || fecha <= hoy || dia === 0 || dia === 6) {
    return { ok: false, error: "Día no disponible" };
  }
  if (!Number.isInteger(cupos) || cupos < 1 || cupos > CUPOS_POR_DIA) {
    return { ok: false, error: "Cantidad de cupos no válida" };
  }

  const candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    const disponibles = CUPOS_POR_DIA - (cuposTomados()[fecha] || 0);
    if (cupos > disponibles) {
      return { ok: false, motivo: "sin_cupos", disponibles: Math.max(0, disponibles) };
    }

    obtenerHojaReservas().appendRow([
      new Date(),
      fecha,
      comoTexto(datos.servicio),
      cupos,
      comoTexto(datos.empresa),
      comoTexto(datos.rut),
      comoTexto(datos.nombre),
      comoTexto(datos.telefono),
      comoTexto(datos.email),
      comoTexto(datos.origen || "Página web"),
      "Pendiente de confirmación",
    ]);
    SpreadsheetApp.flush();
  } finally {
    candado.releaseLock();
  }

  enviarAvisoReserva(datos, cupos);
  return { ok: true, reserva: true };
}

/** Suma los cupos reservados por día, desde hoy en adelante. */
function cuposTomados() {
  const hoja = obtenerHojaReservas();
  const filas = hoja.getLastRow() - 1;
  if (filas < 1) return {};

  const zona = Session.getScriptTimeZone();
  const hoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  const valores = hoja.getRange(2, 1, filas, COLUMNAS_RESERVAS.length).getValues();
  const iDia = COLUMNAS_RESERVAS.indexOf("Día reservado");
  const iCupos = COLUMNAS_RESERVAS.indexOf("Cupos");
  const iEstado = COLUMNAS_RESERVAS.indexOf("Estado");
  const tomados = {};

  valores.forEach((fila) => {
    if (String(fila[iEstado]).trim() === ESTADO_ANULADA) return;
    // Sheets convierte "2026-10-05" en fecha; se vuelve a llevar a texto.
    const dia = fila[iDia] instanceof Date
      ? Utilities.formatDate(fila[iDia], zona, "yyyy-MM-dd")
      : String(fila[iDia]).trim();
    if (dia < hoy) return;
    tomados[dia] = (tomados[dia] || 0) + (Number(fila[iCupos]) || 0);
  });
  return tomados;
}

function obtenerHojaReservas() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA_RESERVAS);

  if (!hoja) {
    hoja = libro.insertSheet(HOJA_RESERVAS);
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS_RESERVAS);
    hoja.getRange(1, 1, 1, COLUMNAS_RESERVAS.length).setFontWeight("bold");
    hoja.setFrozenRows(1);
    hoja.setColumnWidth(1, 160);
    hoja.setColumnWidth(3, 220);
    hoja.setColumnWidth(5, 200);
  }
  return hoja;
}

function enviarAvisoReserva(datos, cupos) {
  const empresa = datos.empresa || "Sin empresa";
  const zona = Session.getScriptTimeZone();
  const dia = new Date(datos.fecha + "T12:00:00");
  const diaTexto = Utilities.formatDate(dia, zona, "dd-MM-yyyy");
  const cuposTexto = cupos + (cupos === 1 ? " cupo" : " cupos");

  const cuerpo = [
    (datos.nombre || "Una persona") + ", de " + empresa + ", reservó " + cuposTexto +
      " de " + String(datos.servicio || "-").toLowerCase() + " para el " + diaTexto + ".",
    "",
    "Servicio:   " + (datos.servicio || "-"),
    "Día:        " + diaTexto,
    "Cupos:      " + cupos,
    "",
    "Empresa:    " + empresa,
    "RUT:        " + (datos.rut || "-"),
    "Contacto:   " + (datos.nombre || "-"),
    "Teléfono:   " + (datos.telefono || "-"),
    "Correo:     " + (datos.email || "-"),
    "Origen:     " + (datos.origen || "Página web"),
    "",
    "Quedó anotada en la pestaña \"" + HOJA_RESERVAS + "\" de la planilla, como pendiente de confirmación.",
    "Si se anula, cambia su estado a \"" + ESTADO_ANULADA + "\" y los cupos vuelven a quedar libres.",
    "",
    "Puedes responder este correo para escribirle directamente.",
  ].join("\n");

  const opciones = { name: "Sur Consultores" };
  if (datos.email) {
    opciones.replyTo = datos.email;
  }

  MailApp.sendEmail(
    CORREO_AVISO,
    "Reserva Petrinovic — " + cuposTexto + " el " + diaTexto + " — " + empresa,
    cuerpo,
    opciones
  );
}

/* ---------- Solicitudes de reunión ---------- */

function enviarAviso(datos) {
  const empresa = datos.empresa || "Sin empresa";
  const cuerpo = [
    "Llegó una solicitud de reunión desde surconsultores.org",
    "",
    "Nombre:           " + (datos.nombre || "-"),
    "Empresa:          " + empresa,
    "Correo:           " + (datos.email || "-"),
    "Teléfono:         " + (datos.telefono || "-"),
    "Industria:        " + (datos.industria || "-"),
    "Dotación:         " + (datos.dotacion || "-"),
    "Tema principal:   " + (datos.tema || "-"),
    "Fecha preferida:  " + (datos.fecha || "-") + " (" + (datos.horario || "-") + ")",
    "Origen:           " + (datos.origen || "Página web"),
    "",
    "Mensaje:",
    datos.mensaje || "(sin mensaje)",
    "",
    "Puedes responder este correo para escribirle directamente.",
  ].join("\n");

  const opciones = { name: "Sur Consultores" };
  if (datos.email) {
    opciones.replyTo = datos.email;
  }

  MailApp.sendEmail(CORREO_AVISO, "Nueva solicitud de reunión — " + empresa, cuerpo, opciones);
}

/**
 * Evita que Sheets interprete un dato como fórmula. Un teléfono como
 * "+56 9 1234 5678" se convertiría en #ERROR! sin esta protección.
 */
function comoTexto(valor) {
  const texto = String(valor == null ? "" : valor);
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}

function responder(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto)).setMimeType(
    ContentService.MimeType.JSON
  );
}
