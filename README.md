# AutomatizaciónProcesos

Resumen
- Script en Google Apps Script para procesar hilos de Gmail (no leídos) y gestionar adjuntos en Google Drive, manteniendo seguimiento por hilo en una Google Sheet.
- Registra/actualiza filas por hilo (ID_CORREO = threadId) con datos: ID incremental, fecha, remitente, asunto, cuerpo, periodo, mes, clasificación, estado, adjuntos, rutas, etc.
- Soporta actualización de hilos (estado "Nuevo" / "Actualización"), evita reprocesos concurrentes y persiste el último incremental con PropertiesService.

Archivos principales
- procesarCorreos.js — Lógica principal de ingestión y persistencia.

Requisitos y permisos
- Cuenta Google con permisos para:
  - Gmail (GmailApp): leer mensajes, etiquetas.
  - Drive (DriveApp): crear carpetas, crear/transferir archivos.
  - Sheets (SpreadsheetApp): leer/escribir hoja.
  - PropertiesService y LockService.
- Scopes típicos: https://www.googleapis.com/auth/gmail.modify, https://www.googleapis.com/auth/drive, https://www.googleapis.com/auth/spreadsheets, https://www.googleapis.com/auth/script.properties

Configuración (constantes)
- BASE_FOLDER_NAME — carpeta raíz en Drive para adjuntos (p. ej. "SolicitudesComite").
- DOMAIN_INTERNAL — dominio para diferenciar Interno/Externo.
- MAX_THREADS_PER_RUN — (opcional) tope de hilos por ejecución.
- EMAIL_ADMIN — correo para notificaciones (si se usa).

Estructura de columnas (actual)
1. ID (incremental)
2. ID_CORREO (threadId)
3. Fecha_recepción
4. Remitente
5. Asunto
6. Cuerpo
7. Periodo (YYYY_10 / YYYY_20)
8. Mes (Enero...Diciembre)
9. Clasificación
10. Fecha_Procesamiento
11. Tipo_Adjunto
12. Interno/Externo
13. Link_Correo
14. Ruta_Adjuntos
15. Nombre_Adjuntos
16. Estado
17. id_interno
18. Fecha_respuesta

Proceso de ejecución (manual)
1. Abrir el proyecto en el editor de Google Apps Script (o desde la Sheet: Extensiones > Apps Script).
2. Autorizar los permisos la primera vez que se ejecute una función que toca Gmail/Drive/Sheet.
3. Ejecutar manualmente `procesarCorreos()` (botón Run) para procesar hilos no leídos.
4. Verificar la hoja: nuevas filas o actualizaciones por hilo.
5. Opcional: ejecutar `resetearId()` para reiniciar contador (usa con precaución).

Proceso de ejecución (automatizado)
- Crear un trigger de tiempo (ej. cada 5-10 minutos) para `procesarCorreos`:
  - En Apps Script: Triggers (Reloj) > Añadir trigger > seleccionar `procesarCorreos` > Event source = Time-driven > Intervalo deseado.
- Crear trigger instalable onEdit para `onEditInstallable`:
  - En Apps Script: Triggers > Añadir trigger > seleccionar `onEditInstallable` > Event source = From spreadsheet > Event type = On edit.

Control de concurrencia y reprocesos
- LockService evita ejecuciones simultáneas.
- Etiqueta temporal "Procesando" se aplica a hilos en tratamiento.
- Persistencia de `lastIncrement` en PropertiesService para mantener numeración secuencial.

Reglas de adjuntos
- No se guardan imágenes (tipo o extensión).
- Nombres sanitizados; prefijo con timestamp + ID.
- Adjuntos excesivamente grandes pueden ser omitidos (revisar implementación si aplica un límite).

Criterios de aceptación (resumido)
- Un único registro por hilo la primera vez; posteriores respuestas actualizan ese registro.
- ID incremental secuencial, persistente entre ejecuciones.
- Adjuntos no-imagen guardados en Drive con ruta documentada en la hoja.
- Estado = "Nuevo" al primer registro, "Actualización" en posteriores.
- Etiquetas de Gmail actualizadas: "Procesando" temporal y "Pendiente Revisión".

Casos de prueba (para QA)
- Caso 1 — Nuevo hilo sin adjuntos

Precondición: Crear email nuevo (no hilo existente), sin adjuntos, no leído en Inbox.
Pasos: Ejecutar procesarCorreos().
Esperado:
Nueva fila con nuevo incremental, ID_CORREO = threadId, Fecha_recepción = fecha primer mensaje, Cuerpo = texto, Periodo/Mes correctos, Ruta_Adjuntos = "Sin adjuntos", Nombre_Adjuntos vacío, Estado="Nuevo".
Hilo etiquetado con "Pendiente Revisión".
lastIncrement incrementado en PropertiesService.

- Caso 2 — Nuevo hilo con adjuntos válidos

Precondición: Email nuevo con 1-3 adjuntos no-imagenes.
Pasos: Ejecutar procesarCorreos().
Esperado:
Carpeta en Drive creada: BASE_FOLDER_NAME/Pendiente Revisión/{ID}_{Solicitante}.
Archivos guardados con nombres sanitizados y prefijo timestamp.
Columnas Ruta_Adjuntos y Nombre_Adjuntos llenas con URL y nombres.
Tipo_Adjunto contiene extensiones en mayúscula.
Estado="Nuevo".

- Caso 3 — Hilo existente recibe respuesta no leída (actualización)

Precondición: Hilo previamente procesado (fila existente). Enviar nueva respuesta al hilo y dejarla no leída.
Pasos: Ejecutar procesarCorreos().
Esperado:
Mismo registro (misma fila) actualizado: Cuerpo actualizado al último mensaje, Fecha_Procesamiento actualizada, Periodo/Mes recalculados si aplica, Estado="Actualización".
Nuevos adjuntos (no-imagenes) guardados en la carpeta existente; Nombre_Adjuntos ampliado con nuevos nombres.
Hilo etiquetado "Pendiente Revisión" (etiqueta anterior removida si era distinta).
No se crea un nuevo incremental.

- Caso 4 — Respuesta que solo añade imágenes

Precondición: Hilo existente; nueva respuesta con solo imágenes.
Pasos: Ejecutar procesarCorreos().
Esperado:
Fila actualizada con nuevo cuerpo y Fecha_Procesamiento, Estado="Actualización".
No se crean archivos en Drive (no se guardan imágenes).
Tipo_Adjunto y Nombre_Adjuntos sin cambios.

- Caso 5 — Concurrencia / Procesos paralelos

Precondición: Ejecutar dos instancias del script casi simultáneamente.
Pasos: Iniciar procesarCorreos() en ambas.
Esperado:
LockService evita colisiones (una instancia aborta con log WARN).
No se duplica incremental ni filas por los mismos hilos (etiqueta "Procesando" evita reprocesos).

- Caso 6 — Texto muy largo en cuerpo

Precondición: Email con cuerpo > 49,000 caracteres.
Pasos: Ejecutar procesarCorreos().
Esperado:
Cuerpo truncado a ~49000 y se adiciona nota "... (texto truncado por límite de celda)".
Script no falla al escribir en hoja.

- Caso 7 — Dominio interno vs externo

Precondición: Email desde remitente con dominio configurado DOMAIN_INTERNAL y otro externo.
Pasos: Ejecutar procesarCorreos() para ambos.
Esperado:
Columna Interno/Externo correctamente "Interno" o "Externo".

- Caso 8 — Error en Drive o Gmail (simular revocación de permisos)

Precondición: Revocar permiso Drive o Gmail temporalmente (si es posible).
Pasos: Ejecutar script.
Esperado:
Error capturado, registrado en LOGS y notificación por email si corresponde.
Lock liberado al final.

Flujo de trabajo para QA (pasos)
1. Preparación:
   - Clonar la Sheet de producción a ambiente de prueba.
   - Crear cuenta de prueba Gmail con mensajes simulados (nuevos hilos, respuestas, adjuntos).
   - Ajustar constantes en script (BASE_FOLDER_NAME, DOMAIN_INTERNAL).
   - Ejecutar `resetearId()` si necesitas empezar desde 0.
2. Ejecución:
   - Ejecutar `procesarCorreos()` manual o via trigger.
   - Registrar evidence (fila, URL Drive, etiquetas en Gmail).
3. Validación:
   - Comparar resultados con criterios de aceptación por caso.
   - Revisar hoja LOGS para errores.
4. Reporte:
   - Documentar PASS/FAIL y adjuntar capturas/URLs/logs.
5. Limpieza:
   - Borrar hilos de prueba o marcarlos leídos.
   - Borrar carpetas/archivos generados en Drive.
   - Restablecer lastIncrement si es necesario.

Troubleshooting y notas
- Si el script alcanza el tiempo máximo, reducir `MAX_THREADS_PER_RUN` o aumentar frecuencia de trigger.
- Si ves filas duplicadas por hilo: verificar que `ID_CORREO` contenga threadId y que existingIds se calcule correctamente.
- Mejoras recomendadas: almacenar ID de carpeta en columna separada para evitar parsing de URL; manejar límites de tamaño de adjunto con log en LOGS.
- Para reiniciar numeración: `resetearId()` — afecta la continuidad y puede crear duplicados si hay filas existentes; usar con cuidado.

Comandos útiles (local)
- Listar archivos del workspace (PowerShell):
  Get-ChildItem -Recurse -File -Name

Contacto / Admin
- EMAIL_ADMIN (configurable en el script) es quien recibe notificaciones en caso de alertas.

Historial de cambios
- Versión actual: seguimiento por hilo (threadId), columnas añadidas Estado/id_interno/Fecha_respuesta, PropertiesService para lastIncrement, LockService y etiquetas temporales.

---

Generar plantilla de pruebas (Google Sheet) o checklist en formato exportable: indicar si lo deseas y lo genero.