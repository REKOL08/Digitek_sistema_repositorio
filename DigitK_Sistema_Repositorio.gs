// ============================================================
//  SISTEMA DE CARGUE REPOSITORIO INSTITUCIONAL — DIGITK
//  Fundación Universitaria del Área Andina — Bibliotecas
//  Google Apps Script — Pegar en: Extensiones > Apps Script
// ============================================================

// ─── CONFIGURACIÓN GLOBAL ───────────────────────────────────
const CONFIG = {
  SHEET_ID: "TU_SHEET_ID_AQUI",           // ID del Google Sheet principal
  DRIVE_ROOT_FOLDER_ID: "TU_FOLDER_ID_AQUI", // ID de la carpeta raíz GR-GRI-REPOSITORIO en Drive
  CORREO_ADMIN: "bibliotecas@areandina.edu.co", // Correo del administrador general
  DOMINIO_INSTITUCIONAL: "@areandina.edu.co",
  NOMBRE_SISTEMA: "Sistema de Cargue Repositorio Digitk",
  NOMBRE_INSTITUCION: "Fundación Universitaria del Área Andina",

  // Hojas del Spreadsheet
  HOJAS: {
    SOLICITUDES: "Solicitudes",
    COLABORADORES: "Colaboradores",
    DASHBOARD: "Dashboard",
    CATALOGO: "Catálogo"
  },

  // Estados del flujo
  ESTADOS: {
    RECIBIDA: "Recibida",
    EN_REVISION: "En revisión",
    DOCUMENTOS_INCOMPLETOS: "Documentos incompletos",
    APROBADA: "Aprobada",
    RECHAZADA: "Rechazada",
    CARGADA_DIGITK: "Cargada en Digitk"
  }
};

// ─── SETUP INICIAL ──────────────────────────────────────────
// Ejecutar UNA SOLA VEZ para configurar todo el sistema
function setupInicial() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // 1. Crear hoja Solicitudes
  crearHojaSolicitudes(ss);

  // 2. Crear hoja Colaboradores
  crearHojaColaboradores(ss);

  // 3. Crear hoja Dashboard
  crearHojaDashboard(ss);

  // 4. Crear hoja Catálogo
  crearHojaCatalogo(ss);

  // 5. Instalar triggers
  instalarTriggers();

  Logger.log("✅ Setup completado exitosamente");
  SpreadsheetApp.getUi().alert("✅ Sistema Digitk configurado correctamente.\n\nAhora vincula el Google Form a la hoja 'Solicitudes'.");
}

function crearHojaSolicitudes(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJAS.SOLICITUDES);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJAS.SOLICITUDES);
  else hoja.clearContents();

  const encabezados = [
    "ID Solicitud", "Timestamp", "Estado", "Sede",
    // Datos del solicitante
    "Nombre Completo", "Correo Institucional", "Rol", "Facultad", "Programa",
    // Datos del material
    "Tipo de Material", "Título", "Autor(es)", "Director/Asesor",
    "Año", "ISBN/ISSN/DOI", "Palabras Clave", "Resumen",
    // Metadatos adicionales
    "Idioma", "Número de Páginas", "Ciudad de la Sede",
    // Control documental
    "Enlace Carpeta Drive", "PDF Cargado", "Autorización Cargada", "Documentos Completos",
    // Gestión biblioteca
    "Colaborador Asignado", "Fecha Revisión", "Observaciones Biblioteca",
    "Fecha Cargue Digitk", "Enlace Digitk", "Historial"
  ];

  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hoja.getRange(1, 1, 1, encabezados.length)
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10);

  hoja.setFrozenRows(1);
  hoja.setColumnWidth(1, 100);
  hoja.setColumnWidth(11, 300); // Título
  hoja.setColumnWidth(16, 400); // Resumen
  hoja.setColumnWidth(23, 250); // Enlace Carpeta Drive

  Logger.log("✅ Hoja Solicitudes creada");
}

function crearHojaColaboradores(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJAS.COLABORADORES);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJAS.COLABORADORES);
  else hoja.clearContents();

  const encabezados = ["Sede", "Nombre", "Correo", "Cargo", "Activo"];
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hoja.getRange(1, 1, 1, encabezados.length)
    .setBackground("#2E7D32")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  // Datos de ejemplo — REEMPLAZAR con los colaboradores reales
  const colaboradores = [
    ["Bogotá", "Nombre Colaborador Bogotá", "colaborador.bogota@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Pereira", "Nombre Colaborador Pereira", "colaborador.pereira@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Valledupar", "Nombre Colaborador Valledupar", "colaborador.valledupar@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Virtual", "Nombre Colaborador Virtual", "colaborador.virtual@areandina.edu.co", "Bibliotecólogo", "SÍ"]
  ];

  hoja.getRange(2, 1, colaboradores.length, encabezados.length).setValues(colaboradores);
  hoja.setFrozenRows(1);
  Logger.log("✅ Hoja Colaboradores creada — actualiza los correos reales");
}

function crearHojaDashboard(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJAS.DASHBOARD);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJAS.DASHBOARD);
  else hoja.clearContents();

  // Título
  hoja.getRange("A1").setValue("📊 DASHBOARD — SISTEMA DIGITK AREANDINA");
  hoja.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#1B5E20");

  // Contadores por estado
  hoja.getRange("A3").setValue("ESTADO");
  hoja.getRange("B3").setValue("CANTIDAD");
  hoja.getRange("A3:B3").setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold");

  const estados = [
    ["Recibida", `=COUNTIF(Solicitudes!C:C,"Recibida")`],
    ["En revisión", `=COUNTIF(Solicitudes!C:C,"En revisión")`],
    ["Documentos incompletos", `=COUNTIF(Solicitudes!C:C,"Documentos incompletos")`],
    ["Aprobada", `=COUNTIF(Solicitudes!C:C,"Aprobada")`],
    ["Rechazada", `=COUNTIF(Solicitudes!C:C,"Rechazada")`],
    ["Cargada en Digitk", `=COUNTIF(Solicitudes!C:C,"Cargada en Digitk")`],
    ["TOTAL", `=COUNTA(Solicitudes!A:A)-1`]
  ];

  hoja.getRange(4, 1, estados.length, 2).setValues(estados);
  hoja.getRange(10, 1, 1, 2).setBackground("#E8F5E9").setFontWeight("bold");

  // Por sede
  hoja.getRange("D3").setValue("SEDE");
  hoja.getRange("E3").setValue("SOLICITUDES");
  hoja.getRange("D3:E3").setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold");

  const sedes = [
    ["Bogotá", `=COUNTIF(Solicitudes!D:D,"Bogotá")`],
    ["Pereira", `=COUNTIF(Solicitudes!D:D,"Pereira")`],
    ["Valledupar", `=COUNTIF(Solicitudes!D:D,"Valledupar")`],
    ["Virtual", `=COUNTIF(Solicitudes!D:D,"Virtual")`]
  ];
  hoja.getRange(4, 4, sedes.length, 2).setValues(sedes);

  // Por tipo de material
  hoja.getRange("G3").setValue("TIPO DE MATERIAL");
  hoja.getRange("H3").setValue("CANTIDAD");
  hoja.getRange("G3:H3").setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold");

  const tipos = [
    ["Trabajo de grado (pregrado)", `=COUNTIF(Solicitudes!J:J,"Trabajo de grado (pregrado)")`],
    ["Trabajo de grado (posgrado)", `=COUNTIF(Solicitudes!J:J,"Trabajo de grado (posgrado)")`],
    ["Artículo de revista", `=COUNTIF(Solicitudes!J:J,"Artículo de revista")`],
    ["Libro / capítulo de libro", `=COUNTIF(Solicitudes!J:J,"Libro / capítulo de libro")`],
    ["Ponencia / congreso", `=COUNTIF(Solicitudes!J:J,"Ponencia / congreso")`],
    ["Informe de investigación", `=COUNTIF(Solicitudes!J:J,"Informe de investigación")`],
    ["Recurso audiovisual", `=COUNTIF(Solicitudes!J:J,"Recurso audiovisual")`],
    ["Otro", `=COUNTIF(Solicitudes!J:J,"Otro")`]
  ];
  hoja.getRange(4, 7, tipos.length, 2).setValues(tipos);

  Logger.log("✅ Hoja Dashboard creada");
}

function crearHojaCatalogo(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJAS.CATALOGO);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJAS.CATALOGO);
  else hoja.clearContents();

  const encabezados = [
    "ID", "Título", "Autor(es)", "Tipo", "Facultad", "Programa",
    "Sede", "Año", "ISBN/ISSN/DOI", "Enlace Digitk", "Fecha Cargue"
  ];
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hoja.getRange(1, 1, 1, encabezados.length)
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  hoja.setFrozenRows(1);
  Logger.log("✅ Hoja Catálogo creada");
}

function instalarTriggers() {
  // Eliminar triggers existentes para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger: al enviar el formulario
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(CONFIG.SHEET_ID)
    .onFormSubmit()
    .create();

  // Trigger: revisión diaria a las 8am
  ScriptApp.newTrigger("revisionDiaria")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log("✅ Triggers instalados");
}

// ─── TRIGGER: NUEVA SOLICITUD ────────────────────────────────
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const hojaSolicitudes = ss.getSheetByName(CONFIG.HOJAS.SOLICITUDES);
    const ultimaFila = hojaSolicitudes.getLastRow();

    // Leer datos de la respuesta del formulario
    const datos = hojaSolicitudes.getRange(ultimaFila, 1, 1, hojaSolicitudes.getLastColumn()).getValues()[0];

    // Mapeo de columnas (ajustar según el orden real del Form)
    const solicitud = parsearSolicitud(datos, ultimaFila);

    // Validar dominio institucional
    if (!solicitud.correo.endsWith(CONFIG.DOMINIO_INSTITUCIONAL)) {
      Logger.log(`❌ Correo no institucional rechazado: ${solicitud.correo}`);
      enviarCorreoRechazoCorreo(solicitud);
      return;
    }

    // Generar ID único
    const idSolicitud = generarID(solicitud.sede, solicitud.tipoMaterial);
    hojaSolicitudes.getRange(ultimaFila, 1).setValue(idSolicitud);

    // Establecer estado inicial
    hojaSolicitudes.getRange(ultimaFila, 3).setValue(CONFIG.ESTADOS.RECIBIDA);

    // Crear carpeta en Drive con estructura correcta
    const carpetaSolicitud = crearCarpetaDrive(solicitud, idSolicitud);
    hojaSolicitudes.getRange(ultimaFila, 23).setValue(carpetaSolicitud.url); // Enlace Carpeta Drive

    // Registrar historial
    const historial = `[${formatearFecha(new Date())}] Solicitud recibida por ${solicitud.nombre}`;
    hojaSolicitudes.getRange(ultimaFila, hojaSolicitudes.getLastColumn()).setValue(historial);

    // Encontrar colaborador de la sede
    const colaborador = obtenerColaborador(ss, solicitud.sede);

    // Guardar colaborador asignado
    hojaSolicitudes.getRange(ultimaFila, 24).setValue(colaborador ? colaborador.nombre : "Sin asignar");

    // Notificar al solicitante
    enviarCorreoConfirmacionSolicitante(solicitud, idSolicitud, carpetaSolicitud.url);

    // Notificar al colaborador de biblioteca
    if (colaborador) {
      enviarCorreoColaborador(solicitud, idSolicitud, colaborador, carpetaSolicitud.url, ultimaFila);
    }

    // Notificar al admin general
    enviarCorreoAdmin(solicitud, idSolicitud, colaborador);

    Logger.log(`✅ Solicitud ${idSolicitud} procesada correctamente`);

  } catch (error) {
    Logger.log(`❌ Error en onFormSubmit: ${error.toString()}`);
    MailApp.sendEmail(CONFIG.CORREO_ADMIN, `[ERROR] Sistema Digitk - ${new Date()}`, error.toString());
  }
}

// ─── PARSEO DE DATOS DEL FORMULARIO ─────────────────────────
// IMPORTANTE: Ajustar los índices según el orden real de preguntas del Form
function parsearSolicitud(datos, fila) {
  return {
    fila: fila,
    timestamp: datos[1] || new Date(),        // Col B
    nombre: datos[4] || "",                    // Col E
    correo: datos[5] || "",                    // Col F
    rol: datos[6] || "",                       // Col G
    facultad: datos[7] || "",                  // Col H
    programa: datos[8] || "",                  // Col I
    tipoMaterial: datos[9] || "",              // Col J
    titulo: datos[10] || "",                   // Col K
    autores: datos[11] || "",                  // Col L
    director: datos[12] || "",                 // Col M
    anio: datos[13] || "",                     // Col N
    identificador: datos[14] || "",            // Col O (ISBN/ISSN/DOI)
    palabrasClave: datos[15] || "",            // Col P
    resumen: datos[16] || "",                  // Col Q
    idioma: datos[17] || "Español",            // Col R
    paginas: datos[18] || "",                  // Col S
    sede: datos[19] || ""                      // Col T (Ciudad de la Sede)
  };
}

// ─── GENERADOR DE ID ─────────────────────────────────────────
function generarID(sede, tipo) {
  const prefSede = { "Bogotá": "BOG", "Pereira": "PER", "Valledupar": "VDR", "Virtual": "VIR" };
  const prefTipo = {
    "Trabajo de grado (pregrado)": "TGP",
    "Trabajo de grado (posgrado)": "TGM",
    "Artículo de revista": "ART",
    "Libro / capítulo de libro": "LIB",
    "Ponencia / congreso": "PON",
    "Informe de investigación": "INF",
    "Recurso audiovisual": "AUD",
    "Otro": "OTR"
  };
  const fecha = Utilities.formatDate(new Date(), "America/Bogota", "yyyyMMdd");
  const random = Math.floor(Math.random() * 9000) + 1000;
  const ps = prefSede[sede] || "GEN";
  const pt = prefTipo[tipo] || "OTR";
  return `DGT-${ps}-${pt}-${fecha}-${random}`;
}

// ─── CREAR CARPETA EN DRIVE ──────────────────────────────────
function crearCarpetaDrive(solicitud, idSolicitud) {
  try {
    const raiz = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);

    // Navegar/crear: Facultad
    const carpetaFacultad = obtenerOCrearCarpeta(raiz, solicitud.facultad);

    // Navegar/crear: Programa
    const carpetaPrograma = obtenerOCrearCarpeta(carpetaFacultad, solicitud.programa);

    // Navegar/crear: Ciudad/Sede
    const carpetaSede = obtenerOCrearCarpeta(carpetaPrograma, solicitud.sede);

    // Crear carpeta del material con ID + título truncado
    const nombreCarpeta = `[${idSolicitud}] ${solicitud.titulo.substring(0, 60)}`;
    const carpetaMaterial = carpetaSede.createFolder(nombreCarpeta);

    // Crear subcarpetas estándar
    carpetaMaterial.createFolder("📄 Documento Principal PDF");
    carpetaMaterial.createFolder("✅ Autorización Publicación");
    carpetaMaterial.createFolder("📎 Documentos Complementarios");

    // Crear archivo de metadatos
    const metadatos = generarMetadatos(solicitud, idSolicitud);
    const archivo = DriveApp.createFile(`METADATOS_${idSolicitud}.txt`, metadatos, MimeType.PLAIN_TEXT);
    archivo.moveTo(carpetaMaterial);

    Logger.log(`✅ Carpeta creada: ${carpetaMaterial.getUrl()}`);
    return { url: carpetaMaterial.getUrl(), nombre: nombreCarpeta };

  } catch (error) {
    Logger.log(`❌ Error creando carpeta Drive: ${error.toString()}`);
    return { url: "Error al crear carpeta", nombre: "Error" };
  }
}

function obtenerOCrearCarpeta(padre, nombre) {
  const iter = padre.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre.createFolder(nombre);
}

function generarMetadatos(solicitud, id) {
  return `SISTEMA DE CARGUE REPOSITORIO DIGITK — AREANDINA
========================================================
ID Solicitud: ${id}
Fecha: ${formatearFecha(new Date())}

SOLICITANTE
-----------
Nombre: ${solicitud.nombre}
Correo: ${solicitud.correo}
Rol: ${solicitud.rol}
Facultad: ${solicitud.facultad}
Programa: ${solicitud.programa}
Sede: ${solicitud.sede}

MATERIAL
--------
Tipo: ${solicitud.tipoMaterial}
Título: ${solicitud.titulo}
Autor(es): ${solicitud.autores}
Director/Asesor: ${solicitud.director}
Año: ${solicitud.anio}
ISBN/ISSN/DOI: ${solicitud.identificador}
Idioma: ${solicitud.idioma}
Páginas: ${solicitud.paginas}
Palabras clave: ${solicitud.palabrasClave}

RESUMEN
-------
${solicitud.resumen}

CARPETAS REQUERIDAS
-------------------
📄 Documento Principal PDF — Archivo PDF del trabajo completo
✅ Autorización Publicación — Carta firmada de autorización
📎 Documentos Complementarios — Anexos, cartas adicionales

ESTADO: Pendiente de carga de documentos
`;
}

// ─── OBTENER COLABORADOR POR SEDE ────────────────────────────
function obtenerColaborador(ss, sede) {
  const hoja = ss.getSheetByName(CONFIG.HOJAS.COLABORADORES);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === sede && datos[i][4] === "SÍ") {
      return {
        sede: datos[i][0],
        nombre: datos[i][1],
        correo: datos[i][2],
        cargo: datos[i][3]
      };
    }
  }
  return null;
}

// ─── CORREOS AUTOMÁTICOS ─────────────────────────────────────

function enviarCorreoConfirmacionSolicitante(solicitud, idSolicitud, urlCarpeta) {
  const asunto = `[Digitk Areandina] Solicitud recibida — ${idSolicitud}`;
  const cuerpo = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
  <div style="background: #1B5E20; padding: 24px; text-align: center;">
    <h1 style="color: #ffffff; margin: 0; font-size: 20px;">📚 Repositorio Institucional Digitk</h1>
    <p style="color: #A5D6A7; margin: 8px 0 0; font-size: 14px;">Fundación Universitaria del Área Andina</p>
  </div>
  <div style="padding: 28px;">
    <p style="color: #333; font-size: 15px;">Hola <strong>${solicitud.nombre}</strong>,</p>
    <p style="color: #555; line-height: 1.6;">Tu solicitud de cargue al repositorio institucional ha sido recibida exitosamente.</p>

    <div style="background: #E8F5E9; border-left: 4px solid #2E7D32; padding: 16px; border-radius: 4px; margin: 20px 0;">
      <p style="margin: 0 0 8px; color: #1B5E20; font-weight: bold;">📋 Detalles de tu solicitud</p>
      <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #333;">
        <tr><td style="padding: 4px 0; color: #666;">ID Solicitud:</td><td style="font-weight: bold;">${idSolicitud}</td></tr>
        <tr><td style="padding: 4px 0; color: #666;">Título:</td><td>${solicitud.titulo}</td></tr>
        <tr><td style="padding: 4px 0; color: #666;">Tipo:</td><td>${solicitud.tipoMaterial}</td></tr>
        <tr><td style="padding: 4px 0; color: #666;">Sede:</td><td>${solicitud.sede}</td></tr>
        <tr><td style="padding: 4px 0; color: #666;">Estado:</td><td><span style="background: #FFF9C4; padding: 2px 8px; border-radius: 12px; color: #F57F17;">⏳ Recibida</span></td></tr>
      </table>
    </div>

    <div style="background: #FFF3E0; border-left: 4px solid #E65100; padding: 16px; border-radius: 4px; margin: 20px 0;">
      <p style="margin: 0 0 8px; color: #E65100; font-weight: bold;">📁 Acción requerida — Cargar tus documentos</p>
      <p style="color: #555; font-size: 14px; margin: 0 0 12px;">Se ha creado una carpeta en Drive exclusiva para tu solicitud. Por favor carga los siguientes documentos:</p>
      <ul style="color: #555; font-size: 14px; margin: 0; padding-left: 20px; line-height: 1.8;">
        <li><strong>📄 Documento Principal PDF</strong> — Archivo PDF completo del trabajo</li>
        <li><strong>✅ Autorización Publicación</strong> — Carta de autorización firmada</li>
        <li><strong>📎 Documentos Complementarios</strong> — Anexos (si aplica)</li>
      </ul>
      <a href="${urlCarpeta}" style="display: inline-block; margin-top: 14px; background: #2E7D32; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold;">📂 Ir a mi carpeta en Drive →</a>
    </div>

    <p style="color: #555; font-size: 14px; line-height: 1.6;">El equipo de Biblioteca revisará tus documentos y recibirás una notificación con el resultado. Si tienes dudas, responde a este correo.</p>
  </div>
  <div style="background: #F5F5F5; padding: 16px; text-align: center;">
    <p style="color: #999; font-size: 12px; margin: 0;">Bibliotecas Areandina · ${solicitud.sede} · ${CONFIG.NOMBRE_SISTEMA}</p>
  </div>
</div>`;

  MailApp.sendEmail({ to: solicitud.correo, subject: asunto, htmlBody: cuerpo });
}

function enviarCorreoColaborador(solicitud, idSolicitud, colaborador, urlCarpeta, fila) {
  const asunto = `[Digitk] Nueva solicitud para revisar — ${idSolicitud}`;

  // Generar links de acción directa
  const urlScript = ScriptApp.getService().getUrl();
  const linkAprobar = `${urlScript}?accion=aprobar&fila=${fila}&id=${idSolicitud}`;
  const linkIncompleto = `${urlScript}?accion=incompleto&fila=${fila}&id=${idSolicitud}`;
  const linkRechazar = `${urlScript}?accion=rechazar&fila=${fila}&id=${idSolicitud}`;

  const cuerpo = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
  <div style="background: #1B5E20; padding: 24px;">
    <h1 style="color: #ffffff; margin: 0; font-size: 18px;">📬 Nueva solicitud de cargue al repositorio</h1>
    <p style="color: #A5D6A7; margin: 8px 0 0; font-size: 13px;">Sede ${solicitud.sede} — Requiere tu revisión</p>
  </div>
  <div style="padding: 28px;">
    <p style="color: #333;">Hola <strong>${colaborador.nombre}</strong>,</p>
    <p style="color: #555;">Se ha recibido una nueva solicitud de cargue al Repositorio Digitk que requiere tu revisión.</p>

    <div style="background: #E8F5E9; border-left: 4px solid #2E7D32; padding: 16px; border-radius: 4px; margin: 20px 0;">
      <p style="margin: 0 0 12px; color: #1B5E20; font-weight: bold;">📋 Datos de la solicitud</p>
      <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #333;">
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666; width: 40%;">ID:</td><td style="font-weight: bold;">${idSolicitud}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Solicitante:</td><td>${solicitud.nombre}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Correo:</td><td>${solicitud.correo}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Rol:</td><td>${solicitud.rol}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Tipo de material:</td><td>${solicitud.tipoMaterial}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Título:</td><td>${solicitud.titulo}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Autor(es):</td><td>${solicitud.autores}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Facultad:</td><td>${solicitud.facultad}</td></tr>
        <tr style="border-bottom: 1px solid #C8E6C9;"><td style="padding: 6px 0; color: #666;">Programa:</td><td>${solicitud.programa}</td></tr>
        <tr><td style="padding: 6px 0; color: #666;">Año:</td><td>${solicitud.anio}</td></tr>
      </table>
    </div>

    <a href="${urlCarpeta}" style="display: inline-block; margin-bottom: 20px; background: #1565C0; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; font-size: 14px;">📂 Ver carpeta con documentos →</a>

    <p style="color: #333; font-weight: bold; margin-bottom: 12px;">Una vez revisados los documentos, selecciona la acción:</p>
    <div style="display: flex; gap: 12px; flex-wrap: wrap;">
      <a href="${linkAprobar}" style="background: #2E7D32; color: white; padding: 12px 20px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold;">✅ Aprobar — documentos completos</a>
      <a href="${linkIncompleto}" style="background: #E65100; color: white; padding: 12px 20px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold;">⚠️ Documentos incompletos</a>
      <a href="${linkRechazar}" style="background: #B71C1C; color: white; padding: 12px 20px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold;">❌ Rechazar solicitud</a>
    </div>

    <p style="color: #999; font-size: 12px; margin-top: 20px;">También puedes cambiar el estado directamente en el Google Sheet.</p>
  </div>
  <div style="background: #F5F5F5; padding: 16px; text-align: center;">
    <p style="color: #999; font-size: 12px; margin: 0;">Bibliotecas Areandina — ${CONFIG.NOMBRE_SISTEMA}</p>
  </div>
</div>`;

  MailApp.sendEmail({ to: colaborador.correo, subject: asunto, htmlBody: cuerpo });
}

function enviarCorreoAdmin(solicitud, idSolicitud, colaborador) {
  const asunto = `[Digitk Admin] Nueva solicitud: ${idSolicitud} — ${solicitud.sede}`;
  const cuerpo = `Nueva solicitud registrada en el Sistema Digitk.\n\nID: ${idSolicitud}\nSolicitante: ${solicitud.nombre} (${solicitud.correo})\nTipo: ${solicitud.tipoMaterial}\nTítulo: ${solicitud.titulo}\nSede: ${solicitud.sede}\nColaborador asignado: ${colaborador ? colaborador.nombre + " (" + colaborador.correo + ")" : "Sin asignar"}`;
  MailApp.sendEmail({ to: CONFIG.CORREO_ADMIN, subject: asunto, body: cuerpo });
}

function enviarCorreoRechazoCorreo(solicitud) {
  const asunto = "[Digitk Areandina] Solicitud no procesada — correo no institucional";
  const cuerpo = `Hola,\n\nTu solicitud de cargue al Repositorio Digitk no pudo ser procesada porque el correo utilizado (${solicitud.correo}) no corresponde a un dominio institucional de Areandina.\n\nPor favor vuelve a diligenciar el formulario usando tu correo ${CONFIG.DOMINIO_INSTITUCIONAL}.\n\nBibliotecas Areandina`;
  try {
    MailApp.sendEmail({ to: solicitud.correo, subject: asunto, body: cuerpo });
  } catch (e) {
    Logger.log("No se pudo notificar correo no institucional");
  }
}

// ─── WEB APP: ACCIONES DESDE CORREO ─────────────────────────
// Publica este script como Web App para que los botones del correo funcionen
function doGet(e) {
  const accion = e.parameter.accion;
  const fila = parseInt(e.parameter.fila);
  const id = e.parameter.id;
  const observacion = e.parameter.obs || "";

  if (!accion || !fila || !id) {
    return HtmlService.createHtmlOutput("<h2>❌ Parámetros inválidos</h2>");
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.HOJAS.SOLICITUDES);

  try {
    switch (accion) {
      case "aprobar":
        return procesarAccion(hoja, fila, id, CONFIG.ESTADOS.APROBADA, ss);
      case "incompleto":
        return procesarAccion(hoja, fila, id, CONFIG.ESTADOS.DOCUMENTOS_INCOMPLETOS, ss);
      case "rechazar":
        return procesarAccion(hoja, fila, id, CONFIG.ESTADOS.RECHAZADA, ss);
      case "digitk":
        return procesarAccion(hoja, fila, id, CONFIG.ESTADOS.CARGADA_DIGITK, ss);
      default:
        return HtmlService.createHtmlOutput("<h2>❌ Acción no reconocida</h2>");
    }
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>❌ Error: ${error.toString()}</h2>`);
  }
}

function procesarAccion(hoja, fila, id, nuevoEstado, ss) {
  const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];

  // Verificar que el ID coincida (seguridad)
  if (datos[0] !== id) {
    return HtmlService.createHtmlOutput("<h2>❌ ID no coincide. Acción no autorizada.</h2>");
  }

  const solicitud = parsearSolicitud(datos, fila);
  const estadoAnterior = datos[2];

  // Actualizar estado
  hoja.getRange(fila, 3).setValue(nuevoEstado);
  hoja.getRange(fila, 25).setValue(formatearFecha(new Date())); // Fecha revisión

  // Actualizar historial
  const historialActual = datos[hoja.getLastColumn() - 1] || "";
  const nuevoHistorial = historialActual + `\n[${formatearFecha(new Date())}] Estado cambiado a: ${nuevoEstado}`;
  hoja.getRange(fila, hoja.getLastColumn()).setValue(nuevoHistorial);

  // Aplicar formato de color según estado
  aplicarColorEstado(hoja, fila, nuevoEstado);

  // Notificar al solicitante según el nuevo estado
  notificarCambioEstado(solicitud, id, nuevoEstado, estadoAnterior);

  // Si se cargó en Digitk, actualizar hoja Catálogo
  if (nuevoEstado === CONFIG.ESTADOS.CARGADA_DIGITK) {
    actualizarCatalogo(ss, datos, id, solicitud);
  }

  const mensajes = {
    [CONFIG.ESTADOS.APROBADA]: "✅ Solicitud APROBADA. El solicitante ha sido notificado.",
    [CONFIG.ESTADOS.DOCUMENTOS_INCOMPLETOS]: "⚠️ Se notificó al solicitante que sus documentos están incompletos.",
    [CONFIG.ESTADOS.RECHAZADA]: "❌ Solicitud RECHAZADA. El solicitante ha sido notificado.",
    [CONFIG.ESTADOS.CARGADA_DIGITK]: "🎉 Material marcado como CARGADO EN DIGITK. El solicitante fue notificado."
  };

  return HtmlService.createHtmlOutput(`
    <html><head><style>
      body{font-family:Arial,sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5;}
      .card{background:white;padding:32px;border-radius:12px;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1);max-width:400px;}
      h2{color:#1B5E20;margin-bottom:12px;}
      p{color:#555;line-height:1.6;}
      .id{font-size:12px;color:#999;margin-top:16px;}
    </style></head>
    <body><div class="card">
      <h2>${mensajes[nuevoEstado] || "Acción registrada"}</h2>
      <p>Solicitud <strong>${id}</strong> actualizada en el sistema.</p>
      <p class="id">Puedes cerrar esta ventana.</p>
    </div></body></html>
  `);
}

function notificarCambioEstado(solicitud, id, nuevoEstado, estadoAnterior) {
  const configuraciones = {
    [CONFIG.ESTADOS.APROBADA]: {
      asunto: `[Digitk Areandina] ✅ Solicitud aprobada — ${id}`,
      mensaje: `Tus documentos han sido revisados y están completos. Tu material será cargado al Repositorio Digitk próximamente.`,
      color: "#2E7D32",
      emoji: "✅"
    },
    [CONFIG.ESTADOS.DOCUMENTOS_INCOMPLETOS]: {
      asunto: `[Digitk Areandina] ⚠️ Documentos incompletos — ${id}`,
      mensaje: `El equipo de Biblioteca revisó tu carpeta en Drive y encontró que faltan documentos. Por favor completa la información y vuelve a notificar al bibliotecario de tu sede.`,
      color: "#E65100",
      emoji: "⚠️"
    },
    [CONFIG.ESTADOS.RECHAZADA]: {
      asunto: `[Digitk Areandina] Solicitud no aprobada — ${id}`,
      mensaje: `Tu solicitud de cargue no fue aprobada en esta ocasión. Para mayor información sobre los motivos, comunícate con la Biblioteca de tu sede.`,
      color: "#B71C1C",
      emoji: "❌"
    },
    [CONFIG.ESTADOS.CARGADA_DIGITK]: {
      asunto: `[Digitk Areandina] 🎉 Material publicado en el repositorio — ${id}`,
      mensaje: `¡Felicitaciones! Tu material ha sido cargado exitosamente al Repositorio Institucional Digitk y ya está disponible para la comunidad académica.`,
      color: "#1565C0",
      emoji: "🎉"
    }
  };

  const cfg = configuraciones[nuevoEstado];
  if (!cfg) return;

  const cuerpo = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">
  <div style="background:${cfg.color};padding:24px;text-align:center;">
    <h1 style="color:#fff;margin:0;font-size:20px;">${cfg.emoji} Actualización de solicitud</h1>
    <p style="color:rgba(255,255,255,0.8);margin:8px 0 0;font-size:13px;">${CONFIG.NOMBRE_SISTEMA}</p>
  </div>
  <div style="padding:28px;">
    <p style="color:#333;">Hola <strong>${solicitud.nombre}</strong>,</p>
    <p style="color:#555;line-height:1.6;">${cfg.mensaje}</p>
    <div style="background:#F5F5F5;padding:16px;border-radius:6px;margin:20px 0;font-size:14px;color:#333;">
      <strong>ID:</strong> ${id}<br>
      <strong>Título:</strong> ${solicitud.titulo}<br>
      <strong>Estado:</strong> ${nuevoEstado}
    </div>
    <p style="color:#555;font-size:14px;">Para consultas adicionales, contacta a la Biblioteca de tu sede (${solicitud.sede}).</p>
  </div>
  <div style="background:#F5F5F5;padding:16px;text-align:center;">
    <p style="color:#999;font-size:12px;margin:0;">Bibliotecas Areandina · ${CONFIG.NOMBRE_SISTEMA}</p>
  </div>
</div>`;

  MailApp.sendEmail({ to: solicitud.correo, subject: cfg.asunto, htmlBody: cuerpo });
}

// ─── ACTUALIZAR CATÁLOGO ─────────────────────────────────────
function actualizarCatalogo(ss, datos, id, solicitud) {
  const hoja = ss.getSheetByName(CONFIG.HOJAS.CATALOGO);
  const urlDigitk = datos[28] || ""; // Columna de enlace Digitk
  hoja.appendRow([
    id,
    solicitud.titulo,
    solicitud.autores,
    solicitud.tipoMaterial,
    solicitud.facultad,
    solicitud.programa,
    solicitud.sede,
    solicitud.anio,
    solicitud.identificador,
    urlDigitk,
    formatearFecha(new Date())
  ]);
}

// ─── FORMATO CONDICIONAL POR ESTADO ─────────────────────────
function aplicarColorEstado(hoja, fila, estado) {
  const colores = {
    [CONFIG.ESTADOS.RECIBIDA]: "#FFF9C4",
    [CONFIG.ESTADOS.EN_REVISION]: "#E3F2FD",
    [CONFIG.ESTADOS.DOCUMENTOS_INCOMPLETOS]: "#FFF3E0",
    [CONFIG.ESTADOS.APROBADA]: "#E8F5E9",
    [CONFIG.ESTADOS.RECHAZADA]: "#FFEBEE",
    [CONFIG.ESTADOS.CARGADA_DIGITK]: "#E0F7FA"
  };
  const color = colores[estado] || "#FFFFFF";
  const lastCol = hoja.getLastColumn();
  hoja.getRange(fila, 1, 1, lastCol).setBackground(color);
}

// ─── REVISIÓN DIARIA ─────────────────────────────────────────
function revisionDiaria() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.HOJAS.SOLICITUDES);
  const datos = hoja.getDataRange().getValues();
  const hoy = new Date();

  for (let i = 1; i < datos.length; i++) {
    const estado = datos[i][2];
    const timestamp = new Date(datos[i][1]);
    const diasTranscurridos = Math.floor((hoy - timestamp) / (1000 * 60 * 60 * 24));

    // Recordatorio a los 5 días si sigue en "Recibida" sin documentos
    if (estado === CONFIG.ESTADOS.RECIBIDA && diasTranscurridos >= 5) {
      const solicitud = parsearSolicitud(datos[i], i + 1);
      const urlCarpeta = datos[i][22] || "";
      enviarRecordatorioPendiente(solicitud, datos[i][0], diasTranscurridos, urlCarpeta);
    }

    // Alerta al colaborador si lleva más de 3 días "En revisión"
    if (estado === CONFIG.ESTADOS.EN_REVISION && diasTranscurridos >= 3) {
      const colaborador = obtenerColaborador(ss, datos[i][3]);
      if (colaborador) {
        enviarAlertaColaborador(colaborador, datos[i][0], datos[i][10], diasTranscurridos);
      }
    }
  }
}

function enviarRecordatorioPendiente(solicitud, id, dias, urlCarpeta) {
  const asunto = `[Digitk Areandina] Recordatorio — Documentos pendientes: ${id}`;
  const cuerpo = `Hola ${solicitud.nombre},\n\nTu solicitud ${id} lleva ${dias} días sin recibir los documentos requeridos.\n\nPor favor carga los archivos en tu carpeta de Drive:\n${urlCarpeta}\n\nSi tienes algún inconveniente, contacta a la Biblioteca de tu sede.\n\nBibliotecas Areandina`;
  MailApp.sendEmail({ to: solicitud.correo, subject: asunto, body: cuerpo });
}

function enviarAlertaColaborador(colaborador, id, titulo, dias) {
  const asunto = `[Digitk Admin] Alerta revisión pendiente: ${id}`;
  const cuerpo = `Hola ${colaborador.nombre},\n\nLa solicitud ${id} (${titulo}) lleva ${dias} días en estado "En revisión" sin respuesta.\n\nPor favor verifica el estado de esta solicitud.\n\n${CONFIG.NOMBRE_SISTEMA}`;
  MailApp.sendEmail({ to: colaborador.correo, subject: asunto, body: cuerpo });
}

// ─── UTILIDADES ──────────────────────────────────────────────
function formatearFecha(fecha) {
  return Utilities.formatDate(fecha, "America/Bogota", "dd/MM/yyyy HH:mm");
}

// ─── MENÚ PERSONALIZADO EN SHEETS ───────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Sistema Digitk")
    .addItem("🔧 Ejecutar setup inicial", "setupInicial")
    .addItem("📊 Actualizar Dashboard", "actualizarDashboardManual")
    .addSeparator()
    .addItem("📋 Ver solicitudes pendientes", "verPendientes")
    .addItem("✅ Marcar seleccionada como aprobada", "aprobarDesdeSheet")
    .addItem("📤 Marcar seleccionada como cargada en Digitk", "marcarDigitkDesdeSheet")
    .addToUi();
}

function actualizarDashboardManual() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Dashboard actualizado — las fórmulas se calculan automáticamente.", "📊 Dashboard", 3);
}

function verPendientes() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.HOJAS.SOLICITUDES);
  const datos = hoja.getDataRange().getValues();
  let pendientes = 0;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] === CONFIG.ESTADOS.RECIBIDA || datos[i][2] === CONFIG.ESTADOS.EN_REVISION) {
      pendientes++;
    }
  }

  SpreadsheetApp.getUi().alert(`📋 Solicitudes pendientes de revisión: ${pendientes}`);
}

function aprobarDesdeSheet() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const fila = hoja.getActiveCell().getRow();
  if (fila <= 1) { SpreadsheetApp.getUi().alert("Selecciona una fila de solicitud válida."); return; }

  const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const id = datos[0];
  const solicitud = parsearSolicitud(datos, fila);

  hoja.getRange(fila, 3).setValue(CONFIG.ESTADOS.APROBADA);
  aplicarColorEstado(hoja, fila, CONFIG.ESTADOS.APROBADA);
  notificarCambioEstado(solicitud, id, CONFIG.ESTADOS.APROBADA, datos[2]);

  SpreadsheetApp.getUi().alert(`✅ Solicitud ${id} aprobada. El solicitante fue notificado.`);
}

function marcarDigitkDesdeSheet() {
  const hoja = SpreadsheetApp.getActiveSheet();
  const fila = hoja.getActiveCell().getRow();
  if (fila <= 1) { SpreadsheetApp.getUi().alert("Selecciona una fila de solicitud válida."); return; }

  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt("Enlace en Digitk", "Ingresa el URL del material en el Repositorio Digitk:", ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;

  const urlDigitk = respuesta.getResponseText();
  const datos = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const id = datos[0];
  const solicitud = parsearSolicitud(datos, fila);
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  hoja.getRange(fila, 3).setValue(CONFIG.ESTADOS.CARGADA_DIGITK);
  hoja.getRange(fila, 29).setValue(urlDigitk); // Col AC — Enlace Digitk
  hoja.getRange(fila, 28).setValue(formatearFecha(new Date())); // Col AB — Fecha Cargue Digitk
  aplicarColorEstado(hoja, fila, CONFIG.ESTADOS.CARGADA_DIGITK);
  notificarCambioEstado(solicitud, id, CONFIG.ESTADOS.CARGADA_DIGITK, datos[2]);
  actualizarCatalogo(ss, hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0], id, solicitud);

  ui.alert(`🎉 Material ${id} marcado como cargado en Digitk. El solicitante fue notificado.`);
}
