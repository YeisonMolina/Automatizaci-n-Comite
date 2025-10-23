// ---------------- CONFIG ----------------
const BASE_FOLDER_NAME = "Correos_Adjuntos"; // carpeta base
const DOMAIN_INTERNAL = "@tudominio.edu.co"; // cambia por tu dominio
// ----------------------------------------

/**
 * Procesa correos no leídos, guarda adjuntos y registra la información en la hoja.
 */
function procesarCorreos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();

  const columnas = [
    "ID", "Código_Interno", "Remitente", "Asunto", "Fecha_Recepción", "Fecha_Procesamiento",
    "Clasificación", "Tipo_Adjunto", "Interno/Externo",
    "Ruta_Correo", "Ruta_Carpeta", "Adjuntos", "Prioridad"
  ];

  // Crear encabezados si hoja vacía
  if (hoja.getLastRow() === 0) hoja.appendRow(columnas);

  // Asegurar que existan todas las columnas
  let encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(String);
  columnas.forEach(col => {
    if (!encabezados.includes(col)) {
      hoja.getRange(1, hoja.getLastColumn() + 1).setValue(col);
      encabezados.push(col);
    }
  });

  const colId = encabezados.indexOf("ID") + 1;
  const lastRow = hoja.getLastRow();
  const existingIds = lastRow >= 2 ? hoja.getRange(2, colId, lastRow - 1, 1).getValues().flat().map(String) : [];

  const baseFolder = getOrCreateFolderByName(BASE_FOLDER_NAME);
  const now = new Date();
  const threads = GmailApp.search("is:unread in:inbox");
  const maxPerRun = 400;
  let processed = 0;

  for (let thread of threads) {
    const mensajes = thread.getMessages();
    for (let mensaje of mensajes) {
      if (processed >= maxPerRun) break;

      const id = String(mensaje.getId());
      if (existingIds.includes(id)) continue;

      const remitente = mensaje.getFrom() || "";
      const asunto = mensaje.getSubject() || "";
      const fechaRecep = mensaje.getDate() || "";
      const fechaProc = now;
      const interno = remitente.toLowerCase().includes(DOMAIN_INTERNAL.toLowerCase()) ? "Interno" : "Externo";
      const rutaCorreo = `https://mail.google.com/mail/u/0/#inbox/${id}`;

      const codigoInterno = generarCodigoInterno(id);
      const prioridad = asignarPrioridad(codigoInterno, interno);

      const attachments = mensaje.getAttachments() || [];
      const validAttachments = attachments.filter(a => {
        const ct = (a.getContentType() || "").toLowerCase();
        const ext = (a.getName().split(".").pop() || "").toLowerCase();
        const isImage = ct.startsWith("image/") || ["jpg", "jpeg", "png", "gif", "bmp", "tiff", "svg", "webp"].includes(ext);
        return !isImage;
      });

      let urlCarpeta = "Sin adjuntos";
      let nombresAdjuntos = [];

      if (validAttachments.length > 0) {
        const clasificacionInicial = "Pendiente por clasificar";
        const carpetaClasif = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/${clasificacionInicial}`);
        const nombreSolic = limpiarNombreSolicitante(remitente);
        const carpetaSolicitud = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/${clasificacionInicial}/${codigoInterno}_${nombreSolic}`);
        urlCarpeta = carpetaSolicitud.getUrl();

        for (let att of validAttachments) {
          const safeName = `${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${att.getName()}`;
          const file = carpetaSolicitud.createFile(att);
          file.setName(safeName);
          nombresAdjuntos.push(safeName);
        }
      }

      const tiposAdj = validAttachments.length > 0
        ? validAttachments.map(a => (a.getName().split(".").pop() || "").toUpperCase()).join(", ")
        : "";

      const etiqueta = GmailApp.getUserLabelByName("Pendiente por clasificar") || GmailApp.createLabel("Pendiente por clasificar");
      thread.addLabel(etiqueta);

      hoja.appendRow([
        id, codigoInterno, remitente, asunto, fechaRecep, fechaProc,
        "Pendiente por clasificar", tiposAdj, interno,
        rutaCorreo, urlCarpeta, nombresAdjuntos.join(", "), prioridad
      ]);

      processed++;
    }
  }

  SpreadsheetApp.getUi().alert(`Procesados ${processed} correos nuevos.`);
}

/**
 * onEdit instalable: actualiza etiqueta y carpeta según nueva clasificación
 */
function onEditInstallable(e) {
  if (!e) return;
  const hoja = e.source.getActiveSheet();
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(String);
  const colClasif = encabezados.indexOf("Clasificación") + 1;
  const colId = encabezados.indexOf("ID") + 1;
  const colRemitente = encabezados.indexOf("Remitente") + 1;
  const colCodigo = encabezados.indexOf("Código_Interno") + 1;
  const colRutaCarpeta = encabezados.indexOf("Ruta_Carpeta") + 1;

  if (e.range.getColumn() !== colClasif || e.range.getRow() === 1) return;

  const fila = e.range.getRow();
  const nuevaClasif = String(e.value || "").trim();
  if (!nuevaClasif) return;

  const id = String(hoja.getRange(fila, colId).getValue());
  const remitente = hoja.getRange(fila, colRemitente).getValue();
  const codigo = hoja.getRange(fila, colCodigo).getValue();
  const rutaActual = hoja.getRange(fila, colRutaCarpeta).getValue();

  try {
    const mensaje = GmailApp.getMessageById(id);
    const hilo = mensaje.getThread();
    hilo.getLabels().forEach(l => hilo.removeLabel(l));
    const newLabel = GmailApp.getUserLabelByName(nuevaClasif) || GmailApp.createLabel(nuevaClasif);
    hilo.addLabel(newLabel);
  } catch (err) {
    Logger.log("Error actualizando etiqueta Gmail: " + err);
  }

  try {
    const nombreSolic = limpiarNombreSolicitante(remitente);
    const carpetaClasif = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/${nuevaClasif}`);
    const carpetaSolicitud = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/${nuevaClasif}/${codigo}_${nombreSolic}`);

    if (rutaActual && rutaActual !== "Sin adjuntos") {
      const match = rutaActual.match(/[-\w]{25,}/);
      if (match) {
        const idFolderAnt = match[0];
        const carpetaAnt = DriveApp.getFolderById(idFolderAnt);
        const files = carpetaAnt.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          carpetaSolicitud.addFile(f);
          carpetaAnt.removeFile(f);
        }
      }
    }

    hoja.getRange(fila, colRutaCarpeta).setValue(carpetaSolicitud.getUrl());
  } catch (err) {
    Logger.log("Error moviendo archivos en Drive: " + err);
  }
}

/** Genera código interno basado en el ID del correo */
function generarCodigoInterno(id) {
  const hash = id.split("").reduce((a, c) => (a + c.charCodeAt(0)) % 99999, 0);
  return "REQ-" + Utilities.formatString("%05d", hash);
}

/** Asigna prioridad: interna > externa, y por orden de llegada */
function asignarPrioridad(codigo, tipo) {
  const base = parseInt(codigo.replace(/\D/g, ""));
  return tipo === "Interno" ? base : base + 50000;
}

/** Crea o devuelve una carpeta por ruta */
function getOrCreateFolderByName(path) {
  const parts = path.split("/");
  let folder = DriveApp.getRootFolder();
  for (let name of parts) {
    let iter = folder.getFoldersByName(name.trim());
    folder = iter.hasNext() ? iter.next() : folder.createFolder(name.trim());
  }
  return folder;
}

/** Limpia nombre del remitente */
function limpiarNombreSolicitante(remitente) {
  if (!remitente) return "Solicitante";
  const match = remitente.match(/^(.+?)\s*<.+?>$/);
  let nombre = match ? match[1] : remitente;
  nombre = nombre.replace(/[^\w\s\-\._@]/g, "").replace(/\s+/g, "_");
  return nombre || "Solicitante";
}
