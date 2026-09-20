/**
 * Sur Consultores — Recepción de solicitudes de reunión
 *
 * Este código vive dentro de la cuenta de Google del usuario, asociado a una
 * planilla de Google Sheets. Recibe los envíos del formulario de la landing
 * page, los anota como una fila nueva y avisa por correo.
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
    ]);

    enviarAviso(datos);
    return responder({ ok: true });
  } catch (error) {
    console.error(error);
    return responder({ ok: false, error: String(error) });
  }
}

/** Permite comprobar en el navegador que la publicación quedó activa. */
function doGet() {
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
  return hoja;
}

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
