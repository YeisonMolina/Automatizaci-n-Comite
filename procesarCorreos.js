// ---------------- CONFIG ----------------
const BASE_FOLDER_NAME = "SolicitudesComite";
const DOMAIN_INTERNAL = "@uniboyaca.edu.co";
const MAX_THREADS_PER_RUN = 80; // máximo hilos por ejecución
const MAX_THREADS_WARNING = 70; // umbral para alerta por correo
const EMAIL_ADMIN = "yeisond.molina@gmail.com"; // destinatario de notificación
// ----------------------------------------

function procesarCorreos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    logEvento("WARN", "Otra ejecución en curso, abortando...");
    return;
  }

  let processed = 0;
  let errores = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getActiveSheet();
    const columnas = [
      "ID", "ID_CORREO", "Fecha_recepción", "Remitente", "Asunto", "Cuerpo",
      "Periodo", "Mes", "Clasificación", "Fecha_Procesamiento", "Tipo_Adjunto",
      "Interno/Externo", "Link_Correo", "Ruta_Adjuntos", "Nombre_Adjuntos",
      "Estado", "id_interno", "Fecha_respuesta"
    ];

    // Asegurar encabezados
    const currentHeaders = hoja.getRange(1, 1, 1, hoja.getLastColumn() || 1).getValues()[0];
    columnas.forEach((col, i) => {
      if (!currentHeaders[i] || currentHeaders[i] !== col) {
        hoja.getRange(1, i + 1).setValue(col);
      }
    });

    const idx = {};
    columnas.forEach((c, i) => idx[c] = i + 1);

    const lastRow = hoja.getLastRow();
    const colIdCorreo = idx["ID_CORREO"];
    const existingIds = lastRow >= 2
      ? hoja.getRange(2, colIdCorreo, lastRow - 1, 1).getValues().flat().filter(String)
      : [];
    const idRowMap = {};
    existingIds.forEach((id, i) => idRowMap[id] = i + 2);

    const props = PropertiesService.getScriptProperties();
    let lastIncrement = parseInt(props.getProperty("lastIncrement") || "0", 10) || 0;
    let nextIncrement = lastIncrement;

    const now = new Date();
    const threads = GmailApp.search("is:unread in:inbox", 0, MAX_THREADS_PER_RUN);
    threads.reverse();
    const labelProcesando = GmailApp.getUserLabelByName("Procesando") || GmailApp.createLabel("Procesando");
    const labelPendiente = GmailApp.getUserLabelByName("Pendiente Revisión") || GmailApp.createLabel("Pendiente Revisión");
    const meseNames = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

    for (let thread of threads) {
      try {
        const mensajes = thread.getMessages();
        if (!mensajes.some(m => m.isUnread && m.isUnread())) continue;

        const threadId = String(thread.getId());
        thread.addLabel(labelProcesando);

        const latestMsg = mensajes[mensajes.length - 1];
        const firstMsg = mensajes[0];
        const fechaRecep = firstMsg.getDate() || now;
        const remitente = latestMsg.getFrom() || "";
        const asunto = latestMsg.getSubject() || "";
        const cuerpo = (() => {
          let texto = latestMsg.getPlainBody ? latestMsg.getPlainBody() : latestMsg.getBody();
          if (texto && texto.length > 49000) {
            texto = texto.substring(0, 49000) + "\n\n... (texto truncado por límite de celda)";
          }
          return texto;
        })();
        const interno = remitente.toLowerCase().includes(DOMAIN_INTERNAL.toLowerCase()) ? "Interno" : "Externo";
        const year = fechaRecep.getFullYear();
        const monthIndex = fechaRecep.getMonth();
        const periodo = `${year}_${monthIndex <= 5 ? "10" : "20"}`;
        const mes = meseNames[monthIndex];
        const rutaCorreo = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

        // Adjuntos (no imágenes)
        const attachmentsToSave = mensajes.flatMap(m => {
          try {
            if (!m.isUnread()) return [];
            return (m.getAttachments() || []).filter(a => {
              const ext = (a.getName().split(".").pop() || "").toLowerCase();
              const ct = (a.getContentType() || "").toLowerCase();
              const isImage = ct.startsWith("image/") || ["jpg","jpeg","png","gif","bmp","tiff","svg","webp"].includes(ext);
              return !isImage;
            });
          } catch {
            return [];
          }
        });

        if (idRowMap[threadId]) {
          // 📩 Hilo existente → verificar si tiene más de una respuesta
          const fila = idRowMap[threadId];
          const numMensajes = mensajes.length;
          const estado = numMensajes > 1 ? "Actualización" : "Pendiente Revisión";

          hoja.getRange(fila, idx["Cuerpo"], 1, 5).setValues([[cuerpo, periodo, mes, estado, now]]);

          let carpetaSolicitud = null;
          const rutaActual = hoja.getRange(fila, idx["Ruta_Adjuntos"]).getValue();
          if (rutaActual && rutaActual !== "Sin adjuntos") {
            const match = String(rutaActual).match(/[-\w]{25,}/);
            if (match) {
              try { carpetaSolicitud = DriveApp.getFolderById(match[0]); } catch {}
            }
          }
          const codigo = hoja.getRange(fila, idx["ID"]).getValue();
          if (!carpetaSolicitud && attachmentsToSave.length > 0) {
            const nombreSolic = limpiarNombreSolicitante(remitente);
            carpetaSolicitud = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/Pendiente Revisión/${codigo}_${nombreSolic}`);
            hoja.getRange(fila, idx["Ruta_Adjuntos"]).setValue(carpetaSolicitud.getUrl());
          }

          if (carpetaSolicitud && attachmentsToSave.length > 0) {
            const existingNames = hoja.getRange(fila, idx["Nombre_Adjuntos"]).getValue();
            const nameList = existingNames ? existingNames.split(",").map(s => s.trim()).filter(Boolean) : [];
            attachmentsToSave.forEach((att, i) => {
              const safeName = `${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${codigo}_${i+1}_${sanitizeFilename(att.getName())}`;
              carpetaSolicitud.createFile(att).setName(safeName);
              nameList.push(safeName);
            });
            hoja.getRange(fila, idx["Nombre_Adjuntos"]).setValue(nameList.join(", "));
          }

          thread.addLabel(labelPendiente);
        } else {
          // 📥 Nueva solicitud
          nextIncrement++;
          const incrementalId = nextIncrement;
          let urlCarpeta = "Sin adjuntos", nombresAdjuntos = [];

          if (attachmentsToSave.length > 0) {
            const clasificacionInicial = "Pendiente Revisión";
            const nombreSolic = limpiarNombreSolicitante(remitente);
            const carpetaSolicitud = getOrCreateFolderByName(`${BASE_FOLDER_NAME}/${clasificacionInicial}/${incrementalId}_${nombreSolic}`);
            urlCarpeta = carpetaSolicitud.getUrl();

            attachmentsToSave.forEach((att, i) => {
              const safeName = `${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}_${incrementalId}_${i+1}_${sanitizeFilename(att.getName())}`;
              carpetaSolicitud.createFile(att).setName(safeName);
              nombresAdjuntos.push(safeName);
            });
          }

          thread.addLabel(labelPendiente);
          hoja.appendRow([
            incrementalId, threadId, fechaRecep, remitente, asunto, cuerpo,
            periodo, mes, "Pendiente Revisión", now,
            attachmentsToSave.map(a => a.getName().split(".").pop().toUpperCase()).join(", "),
            interno, rutaCorreo, urlCarpeta, nombresAdjuntos.join(", "),
            "Nuevo", "", ""
          ]);
        }

        thread.removeLabel(labelProcesando);
        processed++;

      } catch (e) {
        errores.push("Error procesando hilo: " + e);
        logEvento("ERROR", "Error procesando hilo: " + e);
      }
    }

    if (nextIncrement > lastIncrement)
      props.setProperty("lastIncrement", String(nextIncrement));

    hoja.getRange(hoja.getLastRow() + 1, 1).setNote(`Última ejecución: ${new Date().toLocaleString()} - ${processed} hilos procesados`);
    logEvento("INFO", `Procesados ${processed} hilos.`);

  } catch (err) {
    errores.push("Error global: " + err);
    logEvento("CRITICAL", "Error global: " + err);
  } finally {
    // 📬 Enviar notificación si hay errores o demasiados hilos
    if (errores.length > 0 || processed > MAX_THREADS_WARNING) {
      const subject = errores.length > 0
        ? "⚠️ Error en procesamiento de correos"
        : "⚠️ Alerta: alto volumen de correos procesados";
      const body = `
Se ejecutó el script de procesamiento de correos.

📊 Procesados: ${processed}
❗ Errores: ${errores.length}

Detalles:
${errores.join("\n\n") || "Sin errores."}

Fecha: ${new Date().toLocaleString()}
`;
      MailApp.sendEmail(EMAIL_ADMIN, subject, body);
    }
    lock.releaseLock();
  }
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
  return nombre.replace(/[^\w\s\-\._@]/g, "").replace(/\s+/g, "_") || "Solicitante";
}

/** Sanitiza nombres de archivo */
function sanitizeFilename(name) {
  let s = (name || "file").replace(/[\u0000-\u001f\u007f<>:"\/\\|?*\x00-\x1F]/g, "").replace(/\s+/g, "_");
  return s.substring(0, 200);
}

/** Log de eventos o errores en hoja "LOGS" */
function logEvento(tipo, mensaje) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLog = ss.getSheetByName("LOGS") || ss.insertSheet("LOGS");
  hojaLog.appendRow([new Date(), tipo, mensaje]);
}


function resetearId() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("lastIncrement", "0");
  Logger.log("ID reiniciado a 0");
}