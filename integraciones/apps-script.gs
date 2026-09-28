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

  // La reserva ya quedó anotada: un correo de cliente inválido no debe anularla.
  try {
    enviarConfirmacionCliente(datos, cupos);
  } catch (error) {
    console.error("No se pudo enviar la confirmación al cliente: " + error);
  }
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

  GmailApp.sendEmail(
    CORREO_AVISO,
    "Reserva Petrinovic — " + cuposTexto + " el " + diaTexto + " — " + empresa,
    cuerpo,
    opciones
  );
}

const DIAS_SEMANA = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
  "septiembre", "octubre", "noviembre", "diciembre"];

/** Acuse de recibo para la empresa que reservó, al correo que dejó en el formulario. */
function enviarConfirmacionCliente(datos, cupos) {
  const correo = String(datos.email || "").trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return;

  const dia = new Date(datos.fecha + "T12:00:00");
  const diaLargo = DIAS_SEMANA[dia.getDay()] + " " + dia.getDate() + " de " + MESES[dia.getMonth()];
  const diaCorto = Utilities.formatDate(dia, Session.getScriptTimeZone(), "dd-MM-yyyy");
  const cuposTexto = cupos + (cupos === 1 ? " cupo" : " cupos");
  const nombre = String(datos.nombre || "").trim().split(/\s+/)[0];

  const cuerpo = [
    "Hola" + (nombre ? " " + nombre : "") + ",",
    "",
    "Tu solicitud de reserva quedó ingresada en nuestro sistema:",
    "",
    "Servicio:   " + (datos.servicio || "-"),
    "Día:        " + diaLargo,
    "Cupos:      " + cupos,
    "Empresa:    " + (datos.empresa || "-"),
    "RUT:        " + (datos.rut || "-"),
    "",
    "La reserva está PENDIENTE DE CONFIRMACIÓN. Te contactaremos para confirmarla.",
    "",
    "Ante cualquier duda, responde este correo o escríbenos por WhatsApp al +56 9 4818 4418.",
    "",
    "Sur Consultores",
    "surconsultores.org",
  ].join("\n");

  GmailApp.sendEmail(correo, "Recibimos tu reserva — " + cuposTexto + " el " + diaCorto, cuerpo, {
    name: "Sur Consultores",
    replyTo: CORREO_AVISO,
  });
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

  GmailApp.sendEmail(CORREO_AVISO, "Nueva solicitud de reunión — " + empresa, cuerpo, opciones);
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

/* ---------- Resumen diario por correo ---------- */

// Destinatario y horas (de 0 a 23, en la zona horaria del proyecto) del resumen.
// Se envía de lunes a viernes; el del lunes incluye lo llegado el fin de semana.
const CORREO_RESUMEN = "fdollancafil@outlook.com";
const HORAS_RESUMEN = [8, 20];

/**
 * Ejecutar una sola vez desde el editor para programar el resumen. Si se
 * vuelve a ejecutar, reemplaza la programación anterior en vez de duplicarla.
 */
function instalarResumenDiario() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "enviarResumen")
    .forEach((t) => ScriptApp.deleteTrigger(t));

  HORAS_RESUMEN.forEach((hora) => {
    ScriptApp.newTrigger("enviarResumen").timeBased().everyDays(1).atHour(hora).nearMinute(0).create();
  });
}

/**
 * Envía las reservas Petrinovic desde hoy en adelante, agrupadas por día, y
 * las solicitudes de reunión llegadas desde el resumen anterior.
 */
function enviarResumen() {
  const zona = Session.getScriptTimeZone();
  const ahora = new Date();

  // Sábado (6) y domingo (7): no se envía ni se mueve la marca del último resumen.
  if (Number(Utilities.formatDate(ahora, zona, "u")) >= 6) return;

  const propiedades = PropertiesService.getScriptProperties();
  const anterior = new Date(Number(propiedades.getProperty("ultimoResumen")) || ahora.getTime() - 12 * 3600 * 1000);
  const hoy = Utilities.formatDate(ahora, zona, "yyyy-MM-dd");

  // Reservas vigentes, agrupadas por día
  const hojaReservas = obtenerHojaReservas();
  const filasReservas = hojaReservas.getLastRow() - 1;
  const reservas = filasReservas < 1 ? [] : hojaReservas
    .getRange(2, 1, filasReservas, COLUMNAS_RESERVAS.length)
    .getValues()
    .map((fila) => ({
      recibida: fila[0],
      dia: fila[1] instanceof Date ? Utilities.formatDate(fila[1], zona, "yyyy-MM-dd") : String(fila[1]).trim(),
      servicio: fila[2],
      cupos: Number(fila[3]) || 0,
      empresa: fila[4],
      rut: fila[5],
      contacto: fila[6],
      telefono: fila[7],
      correo: fila[8],
      estado: String(fila[10]).trim(),
    }))
    .filter((r) => r.dia >= hoy && r.estado !== ESTADO_ANULADA)
    .sort((a, b) => (a.dia < b.dia ? -1 : a.dia > b.dia ? 1 : 0));

  const nuevas = reservas.filter((r) => r.recibida instanceof Date && r.recibida > anterior).length;

  // Solicitudes de reunión nuevas
  const hojaSolicitudes = obtenerHoja();
  const filasSolicitudes = hojaSolicitudes.getLastRow() - 1;
  const solicitudes = filasSolicitudes < 1 ? [] : hojaSolicitudes
    .getRange(2, 1, filasSolicitudes, COLUMNAS.length)
    .getValues()
    .filter((fila) => fila[0] instanceof Date && fila[0] > anterior);

  const celda = "padding:6px 10px;border-bottom:1px solid #E8E5DF;text-align:left;vertical-align:top;white-space:nowrap;";
  const titulo = "font-family:Arial,sans-serif;color:#2A333B;margin:24px 0 8px;";
  const html = [];

  html.push('<div style="font-family:Arial,sans-serif;font-size:14px;color:#2A333B;">');
  html.push("<p>Resumen de la planilla de Sur Consultores al " +
    Utilities.formatDate(ahora, zona, "dd-MM-yyyy 'a las' HH:mm") + ".</p>");

  html.push('<h3 style="' + titulo + '">Reservas Petrinovic desde hoy (' + reservas.length +
    (nuevas ? ", " + nuevas + (nuevas === 1 ? " nueva" : " nuevas") : "") + ")</h3>");
  if (reservas.length === 0) {
    html.push("<p>No hay reservas vigentes.</p>");
  } else {
    html.push('<table style="border-collapse:collapse;font-size:13px;">');
    html.push("<tr>" + ["Día", "Servicio", "Cupos", "Empresa", "RUT", "Contacto", "Teléfono", "Correo", "Estado"]
      .map((t) => '<th style="' + celda + 'background:#EDEAE3;">' + t + "</th>").join("") + "</tr>");

    let diaActual = null;
    reservas.forEach((r) => {
      if (r.dia !== diaActual) {
        diaActual = r.dia;
        const tomados = reservas.filter((x) => x.dia === r.dia).reduce((s, x) => s + x.cupos, 0);
        const fecha = Utilities.formatDate(new Date(r.dia + "T12:00:00"), zona, "dd-MM-yyyy");
        html.push('<tr><td colspan="9" style="' + celda + 'background:#EAF0ED;font-weight:bold;">' + fecha +
          " — " + tomados + " de " + CUPOS_POR_DIA + " cupos tomados</td></tr>");
      }
      const esNueva = r.recibida instanceof Date && r.recibida > anterior;
      html.push("<tr>" + [
        Utilities.formatDate(new Date(r.dia + "T12:00:00"), zona, "dd-MM") + (esNueva ? ' <b style="color:#567468;">Nueva</b>' : ""),
        escaparHtml(r.servicio), r.cupos, escaparHtml(r.empresa), escaparHtml(r.rut), escaparHtml(r.contacto),
        escaparHtml(r.telefono), escaparHtml(r.correo), escaparHtml(r.estado),
      ].map((v) => '<td style="' + celda + '">' + v + "</td>").join("") + "</tr>");
    });
    html.push("</table>");
  }

  html.push('<h3 style="' + titulo + '">Solicitudes de reunión nuevas (' + solicitudes.length + ")</h3>");
  if (solicitudes.length === 0) {
    html.push("<p>No llegaron solicitudes nuevas desde el resumen anterior.</p>");
  } else {
    html.push('<table style="border-collapse:collapse;font-size:13px;">');
    html.push("<tr>" + ["Recibida", "Nombre", "Empresa", "Correo", "Teléfono", "Tema", "Fecha preferida"]
      .map((t) => '<th style="' + celda + 'background:#EDEAE3;">' + t + "</th>").join("") + "</tr>");
    solicitudes.forEach((f) => {
      html.push("<tr>" + [
        Utilities.formatDate(f[0], zona, "dd-MM HH:mm"), escaparHtml(f[1]), escaparHtml(f[2]), escaparHtml(f[3]),
        escaparHtml(f[4]), escaparHtml(f[7]), (f[8] instanceof Date ? Utilities.formatDate(f[8], zona, "dd-MM-yyyy") : escaparHtml(f[8])) + " (" + escaparHtml(f[9]) + ")",
      ].map((v) => '<td style="' + celda + '">' + v + "</td>").join("") + "</tr>");
    });
    html.push("</table>");
  }

  html.push('<p style="margin-top:24px;"><a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() +
    '" style="color:#567468;">Abrir la planilla</a></p></div>');

  const asunto = "Resumen Sur Consultores — " + Utilities.formatDate(ahora, zona, "dd-MM HH:mm") +
    " — " + reservas.length + (reservas.length === 1 ? " reserva, " : " reservas, ") +
    solicitudes.length + (solicitudes.length === 1 ? " solicitud nueva" : " solicitudes nuevas");

  GmailApp.sendEmail(CORREO_RESUMEN, asunto, "Abre este correo en un lector que muestre HTML.", {
    htmlBody: html.join(""),
    name: "Sur Consultores",
  });

  propiedades.setProperty("ultimoResumen", String(ahora.getTime()));
}

function escaparHtml(valor) {
  return String(valor == null ? "" : valor)
    .replace(/^'/, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
