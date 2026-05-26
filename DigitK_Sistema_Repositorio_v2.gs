// ============================================================
//  SISTEMA DE CARGUE REPOSITORIO INSTITUCIONAL DIGITK
//  Fundación Universitaria del Área Andina — Bibliotecas
//  Google Apps Script v2.0 — Individual + Masivo
//  Autor: Bibliotecas Areandina / REKOL08
// ============================================================

const CONFIG = {
  SHEET_ID: "1p2sdrLu4AGZZsCvih1smbkINULuRF1Qjpa53kFkIu7E",
  DRIVE_ROOT_FOLDER_ID: "1m_K0ji740CoiL5jEiLRU2mVs6Mpen9HX",   // GR-GRI-Opciones de grado
  CORREO_ADMIN: "alabrada3@areandina.edu.co",
  DOMINIO_INSTITUCIONAL: "@areandina.edu.co",
  NOMBRE_SISTEMA: "Sistema de Cargue Repositorio Digitk",
  NOMBRE_INSTITUCION: "Fundación Universitaria del Área Andina",
  NOMBRE_CARPETA_TERMINADOS: "Proceso terminado",
  DIAS_RECORDATORIO_SOLICITANTE: 5,
  DIAS_ALERTA_BIBLIOTECARIO: 3
};

// ──────────────────────────────────────────────
//  ESTADOS DEL SISTEMA
// ──────────────────────────────────────────────
const ESTADOS = {
  PENDIENTE_ARCHIVOS: "Pendiente de archivos",
  EN_REVISION: "En revisión",
  INCOMPLETO: "Incompleto",
  APROBADO: "Aprobado — pendiente Digitk",
  CARGADO_DIGITK: "Cargado en Digitk",
  RECHAZADO: "Rechazado"
};

// ============================================================
//  SETUP INICIAL — ejecutar UNA sola vez
// ============================================================
function setupInicial() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  _crearHojaSolicitudes(ss);
  _crearHojaColaboradores(ss);
  _crearHojaDashboard(ss);
  _crearHojaCatalogo(ss);
  _crearHojaPlantillaMasiva(ss);

  _instalarTriggers();

  Logger.log("✅ Setup inicial completado exitosamente.");
}

function _crearHojaSolicitudes(ss) {
  let hoja = ss.getSheetByName("Solicitudes");
  if (!hoja) hoja = ss.insertSheet("Solicitudes");
  hoja.clearContents();

  const encabezados = [
    "ID", "Fecha solicitud", "Tipo carga", "Nombre solicitante", "Correo solicitante",
    "Rol", "Facultad", "Programa", "Sede", "Tipo de material",
    "Título", "Autor(es)", "Director/Asesor", "Año", "ISBN/ISSN/DOI",
    "Palabras clave", "Resumen", "Idioma", "Páginas",
    "Link carpeta Drive", "Estado", "Fecha estado", "Observaciones bibliotecario",
    "Fecha cargue Digitk", "Link Digitk", "Días en estado actual"
  ];

  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  hoja.getRange(1, 1, 1, encabezados.length)
    .setBackground("#1a5276").setFontColor("#ffffff").setFontWeight("bold");
  hoja.setFrozenRows(1);

  // Formato condicional por estado
  const reglas = [];
  const colEstado = 21; // columna U

  const colores = [
    [ESTADOS.PENDIENTE_ARCHIVOS, "#fef9c3", "#7d6608"],
    [ESTADOS.EN_REVISION, "#dbeafe", "#1e40af"],
    [ESTADOS.INCOMPLETO, "#fee2e2", "#991b1b"],
    [ESTADOS.APROBADO, "#dcfce7", "#166534"],
    [ESTADOS.CARGADO_DIGITK, "#d1fae5", "#064e3b"],
    [ESTADOS.RECHAZADO, "#f3f4f6", "#374151"]
  ];

  colores.forEach(([estado, bg, fg]) => {
    const regla = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(estado)
      .setBackground(bg).setFontColor(fg)
      .setRanges([hoja.getRange(2, colEstado, 1000, 1)])
      .build();
    reglas.push(regla);
  });
  hoja.setConditionalFormatRules(reglas);

  Logger.log("✅ Hoja Solicitudes creada.");
}

function _crearHojaColaboradores(ss) {
  let hoja = ss.getSheetByName("Colaboradores");
  if (!hoja) hoja = ss.insertSheet("Colaboradores");
  hoja.clearContents();

  const encabezados = ["Sede", "Nombre", "Correo", "Cargo", "Activo"];
  hoja.getRange(1, 1, 1, 5).setValues([encabezados])
    .setBackground("#1a5276").setFontColor("#ffffff").setFontWeight("bold");

  const datos = [
    ["Bogotá", "Colaborador Bogotá", "colaborador.bogota@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Pereira", "Colaborador Pereira", "alabrada3@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Valledupar", "Colaborador Valledupar", "colaborador.valledupar@areandina.edu.co", "Bibliotecólogo", "SÍ"],
    ["Virtual", "Colaborador Virtual", "colaborador.virtual@areandina.edu.co", "Bibliotecólogo", "SÍ"]
  ];
  hoja.getRange(2, 1, datos.length, 5).setValues(datos);
  Logger.log("✅ Hoja Colaboradores creada.");
}

function _crearHojaDashboard(ss) {
  let hoja = ss.getSheetByName("Dashboard");
  if (!hoja) hoja = ss.insertSheet("Dashboard");
  hoja.clearContents();

  hoja.getRange("A1").setValue("DASHBOARD — SISTEMA DIGITK")
    .setFontSize(16).setFontWeight("bold").setFontColor("#1a5276");
  hoja.getRange("A2").setValue("Actualizado automáticamente");

  const metricas = [
    ["", "ESTADO", "TOTAL"],
    ["", ESTADOS.PENDIENTE_ARCHIVOS, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.PENDIENTE_ARCHIVOS}")`],
    ["", ESTADOS.EN_REVISION, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.EN_REVISION}")`],
    ["", ESTADOS.INCOMPLETO, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.INCOMPLETO}")`],
    ["", ESTADOS.APROBADO, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.APROBADO}")`],
    ["", ESTADOS.CARGADO_DIGITK, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.CARGADO_DIGITK}")`],
    ["", ESTADOS.RECHAZADO, `=COUNTIF(Solicitudes!U:U,"${ESTADOS.RECHAZADO}")`],
    ["", "TOTAL SOLICITUDES", "=COUNTA(Solicitudes!A:A)-1"]
  ];
  hoja.getRange(4, 1, metricas.length, 3).setValues(metricas);
  hoja.getRange(4, 2, 1, 2).setBackground("#1a5276").setFontColor("#ffffff").setFontWeight("bold");

  Logger.log("✅ Hoja Dashboard creada.");
}

function _crearHojaCatalogo(ss) {
  let hoja = ss.getSheetByName("Catálogo");
  if (!hoja) hoja = ss.insertSheet("Catálogo");
  hoja.clearContents();

  const encabezados = ["Facultad", "Programa", "Sedes disponibles", "ID Carpeta Drive"];
  hoja.getRange(1, 1, 1, 4).setValues([encabezados])
    .setBackground("#1a5276").setFontColor("#ffffff").setFontWeight("bold");
  Logger.log("✅ Hoja Catálogo creada.");
}

function _crearHojaPlantillaMasiva(ss) {
  let hoja = ss.getSheetByName("Plantilla Masiva");
  if (!hoja) hoja = ss.insertSheet("Plantilla Masiva");
  hoja.clearContents();

  hoja.getRange("A1").setValue("PLANTILLA CARGA MASIVA — SISTEMA DIGITK")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1a5276");
  hoja.getRange("A2").setValue(
    "Instrucciones: Llene una fila por documento. No modifique los encabezados. " +
    "Al terminar, descargue como .xlsx y adjúntelo en el formulario de carga masiva."
  ).setFontColor("#7f8c8d").setWrap(true);

  const encabezados = [
    "Nombre solicitante*", "Correo institucional*", "Rol*",
    "Facultad*", "Programa*", "Sede*", "Tipo de material*",
    "Título*", "Autor(es)*", "Director/Asesor", "Año*",
    "ISBN/ISSN/DOI", "Palabras clave*", "Resumen*", "Idioma*", "Páginas"
  ];
  hoja.getRange(4, 1, 1, encabezados.length).setValues([encabezados])
    .setBackground("#2e86c1").setFontColor("#ffffff").setFontWeight("bold");

  // Fila de ejemplo
  const ejemplo = [
    "García López, Juan", "jgarcia@areandina.edu.co", "Docente",
    "Facultad de Ciencias Administrativas, Económicas y Financieras",
    "Administración de Empresas", "Bogotá", "Trabajo de grado (pregrado)",
    "Gestión del talento humano en pymes", "García López, Juan; Martínez, Ana",
    "Dr. Pérez Ruiz", "2024", "978-3-16-148410-0",
    "talento humano, pymes, gestión", "Este trabajo analiza...", "Español", "120"
  ];
  hoja.getRange(5, 1, 1, ejemplo.length).setValues([ejemplo])
    .setBackground("#eaf4fb").setFontColor("#7f8c8d").setFontStyle("italic");

  Logger.log("✅ Hoja Plantilla Masiva creada.");
}

function _instalarTriggers() {
  // Eliminar triggers existentes para no duplicar
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger para nuevo envío de formulario individual
  ScriptApp.newTrigger("onFormSubmitIndividual")
    .forSpreadsheet(CONFIG.SHEET_ID)
    .onFormSubmit()
    .create();

  // Trigger diario para revisión de vencimientos
  ScriptApp.newTrigger("revisionDiaria")
    .timeBased().everyDays(1).atHour(8).create();

  Logger.log("✅ Triggers instalados.");
}

// ============================================================
//  PROCESAMIENTO FORMULARIO INDIVIDUAL
// ============================================================
function onFormSubmitIndividual(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const hojaSolicitudes = ss.getSheetByName("Solicitudes");
    const respuestas = e.namedValues;

    const correo = (respuestas["Correo institucional"] || [""])[0].trim();
    if (!correo.endsWith(CONFIG.DOMINIO_INSTITUCIONAL)) {
      Logger.log("❌ Correo no institucional rechazado: " + correo);
      return;
    }

    const id = _generarID();
    const sede = (respuestas["Sede"] || [""])[0].trim();
    const facultad = (respuestas["Facultad"] || [""])[0].trim();
    const programa = (respuestas["Programa académico"] || [""])[0].trim();
    const autor = (respuestas["Autor(es)"] || [""])[0].trim();
    const anio = (respuestas["Año"] || [""])[0].trim();
    const titulo = (respuestas["Título del material"] || [""])[0].trim();
    const nombreSolicitante = (respuestas["Nombre completo"] || [""])[0].trim();
    const tipoMaterial = (respuestas["Tipo de material"] || [""])[0].trim();

    const nombreCarpeta = _formatearNombreCarpeta(autor, anio);
    const carpetaAutor = _crearCarpetaEnDrive(facultad, programa, sede, nombreCarpeta);
    const linkCarpeta = carpetaAutor ? carpetaAutor.getUrl() : "Error al crear carpeta";

    const fila = [
      id,
      new Date(),
      "Individual",
      nombreSolicitante,
      correo,
      (respuestas["Rol"] || [""])[0],
      facultad,
      programa,
      sede,
      tipoMaterial,
      titulo,
      autor,
      (respuestas["Director / Asesor (si aplica)"] || [""])[0],
      anio,
      (respuestas["ISBN / ISSN / DOI (si aplica)"] || [""])[0],
      (respuestas["Palabras clave"] || [""])[0],
      (respuestas["Resumen"] || [""])[0],
      (respuestas["Idioma"] || [""])[0],
      (respuestas["Número de páginas"] || [""])[0],
      linkCarpeta,
      ESTADOS.PENDIENTE_ARCHIVOS,
      new Date(),
      "",
      "",
      "",
      ""
    ];

    hojaSolicitudes.appendRow(fila);

    // Correo al solicitante con link de Drive
    _enviarCorreoSolicitante(correo, nombreSolicitante, titulo, linkCarpeta, id);

    // Correo al colaborador de la sede
    const correoColaborador = _obtenerCorreoColaborador(ss, sede);
    if (correoColaborador) {
      _enviarNotificacionColaborador(correoColaborador, nombreSolicitante, titulo, sede, programa, id, linkCarpeta);
    }

    Logger.log("✅ Solicitud individual procesada: " + id);
  } catch (err) {
    Logger.log("❌ Error en onFormSubmitIndividual: " + err.message);
    GmailApp.sendEmail(CONFIG.CORREO_ADMIN, "Error sistema Digitk", "Error: " + err.message);
  }
}

// ============================================================
//  PROCESAMIENTO CARGA MASIVA
//  Llamar manualmente desde el menú con el Excel adjunto ya
//  disponible en Drive en una carpeta conocida.
// ============================================================
function procesarCargaMasiva() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    "Carga Masiva",
    "Pega el ID de la carpeta de Drive donde está el Excel de carga masiva:",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() !== ui.Button.OK) return;

  const carpetaId = respuesta.getResponseText().trim();
  const carpeta = DriveApp.getFolderById(carpetaId);
  const archivos = carpeta.getFilesByType(MimeType.MICROSOFT_EXCEL);

  if (!archivos.hasNext()) {
    ui.alert("No se encontró un archivo Excel en esa carpeta.");
    return;
  }

  const archivo = archivos.next();
  const blob = archivo.getBlob();
  const ssTemp = SpreadsheetApp.openById(
    Drive.Files.insert(
      { title: "temp_masiva_" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
      blob
    ).id
  );

  const hojaDatos = ssTemp.getSheets()[0];
  const datos = hojaDatos.getDataRange().getValues();

  // Fila 1 = título, fila 2 = instrucciones, fila 3 = vacía, fila 4 = encabezados, desde fila 5 = datos
  // Si la plantilla viene directamente con encabezados en fila 1:
  let filaInicio = 1;
  // Detectar si hay encabezado buscando "Nombre solicitante"
  for (let i = 0; i < Math.min(5, datos.length); i++) {
    if (String(datos[i][0]).toLowerCase().includes("nombre solicitante")) {
      filaInicio = i + 1;
      break;
    }
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hojaSolicitudes = ss.getSheetByName("Solicitudes");
  let procesados = 0;
  let errores = 0;
  const resumen = [];

  for (let i = filaInicio; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0] || !fila[1]) continue; // fila vacía

    const nombreSolicitante = String(fila[0]).trim();
    const correo = String(fila[1]).trim();
    const rol = String(fila[2]).trim();
    const facultad = String(fila[3]).trim();
    const programa = String(fila[4]).trim();
    const sede = String(fila[5]).trim();
    const tipoMaterial = String(fila[6]).trim();
    const titulo = String(fila[7]).trim();
    const autor = String(fila[8]).trim();
    const director = String(fila[9]).trim();
    const anio = String(fila[10]).trim();
    const isbn = String(fila[11]).trim();
    const palabrasClave = String(fila[12]).trim();
    const resumenDoc = String(fila[13]).trim();
    const idioma = String(fila[14]).trim();
    const paginas = String(fila[15]).trim();

    if (!correo.endsWith(CONFIG.DOMINIO_INSTITUCIONAL)) {
      errores++;
      resumen.push({ nombre: nombreSolicitante, estado: "ERROR: correo no institucional" });
      continue;
    }

    try {
      const id = _generarID();
      const nombreCarpeta = _formatearNombreCarpeta(autor, anio);
      const carpetaAutor = _crearCarpetaEnDrive(facultad, programa, sede, nombreCarpeta);
      const linkCarpeta = carpetaAutor ? carpetaAutor.getUrl() : "Error";

      hojaSolicitudes.appendRow([
        id, new Date(), "Masiva",
        nombreSolicitante, correo, rol,
        facultad, programa, sede, tipoMaterial,
        titulo, autor, director, anio, isbn,
        palabrasClave, resumenDoc, idioma, paginas,
        linkCarpeta, ESTADOS.PENDIENTE_ARCHIVOS, new Date(),
        "", "", "", ""
      ]);

      _enviarCorreoSolicitante(correo, nombreSolicitante, titulo, linkCarpeta, id);
      procesados++;
      resumen.push({ nombre: nombreSolicitante, estado: "✅ OK — " + linkCarpeta });
      Utilities.sleep(200); // Evitar límites de Gmail
    } catch (err) {
      errores++;
      resumen.push({ nombre: nombreSolicitante, estado: "ERROR: " + err.message });
    }
  }

  // Limpiar archivo temporal
  DriveApp.getFileById(ssTemp.getId()).setTrashed(true);

  // Enviar resumen consolidado al admin/bibliotecario
  _enviarResumenMasivo(procesados, errores, resumen);

  ui.alert(`Carga masiva completada.\n✅ Procesados: ${procesados}\n❌ Errores: ${errores}\nRevisa tu correo para el resumen completo.`);
}

// ============================================================
//  GESTIÓN DE ESTADOS DESDE CORREO (Web App)
// ============================================================
function doGet(e) {
  const accion = e.parameter.accion;
  const id = e.parameter.id;
  const obs = e.parameter.obs || "";

  if (!accion || !id) {
    return HtmlService.createHtmlOutput("<h2>Parámetros inválidos</h2>");
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName("Solicitudes");
  const datos = hoja.getDataRange().getValues();

  let filaIdx = -1;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === id) { filaIdx = i + 1; break; }
  }

  if (filaIdx === -1) {
    return HtmlService.createHtmlOutput("<h2>Solicitud no encontrada: " + id + "</h2>");
  }

  let nuevoEstado = "";
  let mensaje = "";

  switch (accion) {
    case "aprobar":
      nuevoEstado = ESTADOS.APROBADO;
      mensaje = "✅ Solicitud APROBADA";
      break;
    case "incompleto":
      nuevoEstado = ESTADOS.INCOMPLETO;
      mensaje = "⚠️ Marcada como INCOMPLETA";
      break;
    case "rechazar":
      nuevoEstado = ESTADOS.RECHAZADO;
      mensaje = "❌ Solicitud RECHAZADA";
      break;
    case "digitk":
      nuevoEstado = ESTADOS.CARGADO_DIGITK;
      mensaje = "🎉 Marcada como CARGADA EN DIGITK";
      _moverAProcesoTerminado(datos[filaIdx - 1]);
      break;
    default:
      return HtmlService.createHtmlOutput("<h2>Acción desconocida</h2>");
  }

  hoja.getRange(filaIdx, 21).setValue(nuevoEstado);
  hoja.getRange(filaIdx, 22).setValue(new Date());
  if (obs) hoja.getRange(filaIdx, 23).setValue(obs);

  // Notificar al solicitante
  const correoSolicitante = datos[filaIdx - 1][4];
  const nombreSolicitante = datos[filaIdx - 1][3];
  const titulo = datos[filaIdx - 1][10];
  _notificarCambioEstado(correoSolicitante, nombreSolicitante, titulo, nuevoEstado, obs);

  return HtmlService.createHtmlOutput(
    `<html><body style="font-family:Arial;padding:40px;text-align:center;">
    <h2 style="color:#1a5276;">${mensaje}</h2>
    <p>ID: <strong>${id}</strong></p>
    <p>Solicitud: <strong>${titulo}</strong></p>
    ${obs ? `<p>Observación: ${obs}</p>` : ""}
    <p style="color:#7f8c8d;">Puedes cerrar esta ventana.</p>
    </body></html>`
  );
}

// ============================================================
//  MOVER CARPETA A "PROCESO TERMINADO"
// ============================================================
function _moverAProcesoTerminado(filaData) {
  try {
    const linkCarpeta = filaData[19];
    if (!linkCarpeta || linkCarpeta === "Error") return;

    const idCarpeta = _extraerIdDriveLink(linkCarpeta);
    if (!idCarpeta) return;

    const carpetaAutor = DriveApp.getFolderById(idCarpeta);
    const carpetaPadre = carpetaAutor.getParents().next(); // Ciudad
    const carpetaPrograma = carpetaPadre.getParents().next(); // Programa

    // Buscar o crear "Proceso terminado" dentro del Programa
    let carpetaTerminado;
    const iter = carpetaPrograma.getFoldersByName(CONFIG.NOMBRE_CARPETA_TERMINADOS);
    if (iter.hasNext()) {
      carpetaTerminado = iter.next();
    } else {
      carpetaTerminado = carpetaPrograma.createFolder(CONFIG.NOMBRE_CARPETA_TERMINADOS);
    }

    carpetaAutor.moveTo(carpetaTerminado);
    Logger.log("✅ Carpeta movida a Proceso terminado: " + carpetaAutor.getName());
  } catch (err) {
    Logger.log("⚠️ No se pudo mover carpeta: " + err.message);
  }
}

// ============================================================
//  CREACIÓN DE CARPETAS EN DRIVE
// ============================================================
function _crearCarpetaEnDrive(facultad, programa, sede, nombreAutor) {
  const raiz = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);

  const carpetaFacultad = _buscarOCrearCarpeta(raiz, facultad);
  const carpetaPrograma = _buscarOCrearCarpeta(carpetaFacultad, programa);
  const carpetaSede = _buscarOCrearCarpeta(carpetaPrograma, sede);
  const carpetaAutor = _buscarOCrearCarpeta(carpetaSede, nombreAutor);

  return carpetaAutor;
}

function _buscarOCrearCarpeta(padre, nombre) {
  const iter = padre.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return padre.createFolder(nombre);
}

function _formatearNombreCarpeta(autor, anio) {
  // Formato: Apellido, Nombre Año
  const partes = autor.split(",");
  const apellido = partes[0] ? partes[0].trim() : autor.trim();
  const nombre = partes[1] ? partes[1].trim().split(";")[0].trim() : "";
  return nombre ? `${apellido}, ${nombre} ${anio}` : `${apellido} ${anio}`;
}

function _extraerIdDriveLink(link) {
  const match = link.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// ============================================================
//  COLABORADORES
// ============================================================
function _obtenerCorreoColaborador(ss, sede) {
  const hoja = ss.getSheetByName("Colaboradores");
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === sede && datos[i][4] === "SÍ") {
      return datos[i][2];
    }
  }
  return CONFIG.CORREO_ADMIN;
}

// ============================================================
//  CORREOS AUTOMÁTICOS
// ============================================================
function _enviarCorreoSolicitante(correo, nombre, titulo, linkCarpeta, id) {
  const asunto = `[Digitk] Tu solicitud fue recibida — ${titulo}`;
  const cuerpo = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">
  <div style="background:#1a5276;padding:24px;border-radius:8px 8px 0 0;">
    <h2 style="color:#fff;margin:0;">📚 ${CONFIG.NOMBRE_SISTEMA}</h2>
    <p style="color:#aed6f1;margin:4px 0 0;">${CONFIG.NOMBRE_INSTITUCION}</p>
  </div>
  <div style="border:1px solid #d5d8dc;border-top:none;padding:28px;border-radius:0 0 8px 8px;">
    <p>Hola <strong>${nombre}</strong>,</p>
    <p>Tu solicitud de cargue al repositorio <strong>Digitk</strong> fue recibida exitosamente.</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0;">
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;width:35%;">ID Solicitud</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${id}</td></tr>
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;">Título</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${titulo}</td></tr>
    </table>
    <div style="background:#fef9c3;border-left:4px solid #f39c12;padding:16px;margin:16px 0;border-radius:4px;">
      <strong>📂 Próximo paso — Subir tus archivos</strong>
      <p style="margin:8px 0 4px;">Entra a tu carpeta asignada en Drive y sube los siguientes documentos:</p>
      <ul style="margin:0;padding-left:20px;">
        <li>📄 Documento PDF completo</li>
        <li>✅ Autorización de publicación firmada</li>
        <li>📎 Documentos complementarios (si aplica)</li>
      </ul>
    </div>
    <div style="text-align:center;margin:24px 0;">
      <a href="${linkCarpeta}" style="background:#1a5276;color:#fff;padding:12px 28px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;">
        📂 Ir a mi carpeta en Drive
      </a>
    </div>
    <p style="color:#7f8c8d;font-size:13px;">Una vez subidos los archivos, el equipo de Bibliotecas revisará tu solicitud y te notificará por este medio.</p>
    <hr style="border:none;border-top:1px solid #d5d8dc;margin:20px 0;">
    <p style="color:#7f8c8d;font-size:12px;text-align:center;">
      ${CONFIG.NOMBRE_INSTITUCION} — Bibliotecas Areandina<br>
      Sistema Repositorio Digitk
    </p>
  </div>
</body>
</html>`;
  GmailApp.sendEmail(correo, asunto, "", { htmlBody: cuerpo });
}

function _enviarNotificacionColaborador(correoColab, solicitante, titulo, sede, programa, id, linkCarpeta) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const urlAprobar = `${webAppUrl}?accion=aprobar&id=${id}`;
  const urlIncompleto = `${webAppUrl}?accion=incompleto&id=${id}`;
  const urlRechazar = `${webAppUrl}?accion=rechazar&id=${id}`;

  const asunto = `[Digitk] Nueva solicitud pendiente — ${sede} — ${titulo}`;
  const cuerpo = `
<!DOCTYPE html>
<html>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">
  <div style="background:#1a5276;padding:24px;border-radius:8px 8px 0 0;">
    <h2 style="color:#fff;margin:0;">📋 Nueva solicitud de cargue</h2>
    <p style="color:#aed6f1;margin:4px 0 0;">Sede: ${sede}</p>
  </div>
  <div style="border:1px solid #d5d8dc;border-top:none;padding:28px;border-radius:0 0 8px 8px;">
    <p>Hay una nueva solicitud de cargue al repositorio Digitk:</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0;">
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;width:35%;">ID</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${id}</td></tr>
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;">Solicitante</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${solicitante}</td></tr>
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;">Título</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${titulo}</td></tr>
      <tr><td style="padding:8px;background:#eaf4fb;font-weight:bold;">Programa</td>
          <td style="padding:8px;border:1px solid #d5d8dc;">${programa}</td></tr>
    </table>
    <p><strong>Verifica que el solicitante haya subido los 3 archivos a la carpeta, luego toma una acción:</strong></p>
    <div style="text-align:center;margin:24px 0;">
      <a href="${linkCarpeta}" style="background:#2e86c1;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;margin:4px;">
        📂 Ver carpeta Drive
      </a>
    </div>
    <div style="display:flex;gap:8px;justify-content:center;margin:16px 0;text-align:center;">
      <a href="${urlAprobar}" style="background:#27ae60;color:#fff;padding:10px 24px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;margin:4px;">
        ✅ Aprobar
      </a>
      <a href="${urlIncompleto}" style="background:#f39c12;color:#fff;padding:10px 24px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;margin:4px;">
        ⚠️ Incompleto
      </a>
      <a href="${urlRechazar}" style="background:#e74c3c;color:#fff;padding:10px 24px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;margin:4px;">
        ❌ Rechazar
      </a>
    </div>
    <hr style="border:none;border-top:1px solid #d5d8dc;margin:20px 0;">
    <p style="color:#7f8c8d;font-size:12px;text-align:center;">
      Sistema Repositorio Digitk — ${CONFIG.NOMBRE_INSTITUCION}
    </p>
  </div>
</body>
</html>`;
  GmailApp.sendEmail(correoColab, asunto, "", { htmlBody: cuerpo });
}

function _notificarCambioEstado(correo, nombre, titulo, estado, observacion) {
  const iconos = {
    [ESTADOS.APROBADO]: "✅",
    [ESTADOS.INCOMPLETO]: "⚠️",
    [ESTADOS.RECHAZADO]: "❌",
    [ESTADOS.CARGADO_DIGITK]: "🎉"
  };
  const icono = iconos[estado] || "📋";
  const asunto = `[Digitk] ${icono} Actualización de tu solicitud — ${titulo}`;

  const mensajeExtra = estado === ESTADOS.INCOMPLETO
    ? `<div style="background:#fef3cd;border-left:4px solid #f39c12;padding:12px;margin:12px 0;border-radius:4px;">
        <strong>Acción requerida:</strong> ${observacion || "Por favor revisa los archivos en tu carpeta de Drive y completa los documentos faltantes."}</div>`
    : estado === ESTADOS.CARGADO_DIGITK
    ? `<div style="background:#d4edda;border-left:4px solid #27ae60;padding:12px;margin:12px 0;border-radius:4px;">
        Tu material ya está publicado en el repositorio institucional Digitk. ¡Gracias por tu aporte académico!</div>`
    : observacion
    ? `<div style="background:#f8d7da;border-left:4px solid #e74c3c;padding:12px;margin:12px 0;border-radius:4px;">
        <strong>Observación:</strong> ${observacion}</div>`
    : "";

  const cuerpo = `
<!DOCTYPE html>
<html>
<body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#2c3e50;">
  <div style="background:#1a5276;padding:24px;border-radius:8px 8px 0 0;">
    <h2 style="color:#fff;margin:0;">📚 ${CONFIG.NOMBRE_SISTEMA}</h2>
  </div>
  <div style="border:1px solid #d5d8dc;border-top:none;padding:28px;border-radius:0 0 8px 8px;">
    <p>Hola <strong>${nombre}</strong>,</p>
    <p>Tu solicitud <strong>"${titulo}"</strong> tiene una actualización:</p>
    <div style="text-align:center;font-size:22px;font-weight:bold;padding:20px;background:#eaf4fb;border-radius:8px;margin:16px 0;">
      ${icono} ${estado}
    </div>
    ${mensajeExtra}
    <hr style="border:none;border-top:1px solid #d5d8dc;margin:20px 0;">
    <p style="color:#7f8c8d;font-size:12px;text-align:center;">
      ${CONFIG.NOMBRE_INSTITUCION} — Bibliotecas Areandina
    </p>
  </div>
</body>
</html>`;
  GmailApp.sendEmail(correo, asunto, "", { htmlBody: cuerpo });
}

function _enviarResumenMasivo(procesados, errores, resumen) {
  const filas = resumen.map(r =>
    `<tr>
      <td style="padding:8px;border:1px solid #d5d8dc;">${r.nombre}</td>
      <td style="padding:8px;border:1px solid #d5d8dc;">${r.estado}</td>
    </tr>`
  ).join("");

  const cuerpo = `
<!DOCTYPE html>
<html>
<body style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#2c3e50;">
  <div style="background:#1a5276;padding:24px;border-radius:8px 8px 0 0;">
    <h2 style="color:#fff;margin:0;">📋 Resumen Carga Masiva — Digitk</h2>
    <p style="color:#aed6f1;">Procesamiento completado</p>
  </div>
  <div style="border:1px solid #d5d8dc;border-top:none;padding:28px;border-radius:0 0 8px 8px;">
    <div style="display:flex;gap:16px;margin-bottom:24px;">
      <div style="flex:1;background:#d4edda;padding:16px;border-radius:8px;text-align:center;">
        <div style="font-size:32px;font-weight:bold;color:#155724;">${procesados}</div>
        <div>✅ Procesados</div>
      </div>
      <div style="flex:1;background:#f8d7da;padding:16px;border-radius:8px;text-align:center;">
        <div style="font-size:32px;font-weight:bold;color:#721c24;">${errores}</div>
        <div>❌ Errores</div>
      </div>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#1a5276;color:#fff;">
        <th style="padding:10px;text-align:left;">Solicitante</th>
        <th style="padding:10px;text-align:left;">Estado</th>
      </tr>
      ${filas}
    </table>
    <p style="color:#7f8c8d;font-size:12px;margin-top:20px;text-align:center;">
      Los solicitantes procesados ya recibieron su correo con el link a su carpeta Drive.<br>
      Revisa el Sheet de Solicitudes para ver el estado completo.
    </p>
  </div>
</body>
</html>`;

  GmailApp.sendEmail(
    CONFIG.CORREO_ADMIN,
    `[Digitk] Carga masiva completada — ${procesados} procesados, ${errores} errores`,
    "",
    { htmlBody: cuerpo }
  );
}

// ============================================================
//  REVISIÓN DIARIA — RECORDATORIOS Y VENCIMIENTOS
// ============================================================
function revisionDiaria() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName("Solicitudes");
  const datos = hoja.getDataRange().getValues();
  const ahora = new Date();

  for (let i = 1; i < datos.length; i++) {
    const estado = datos[i][20];
    const fechaEstado = new Date(datos[i][21]);
    const dias = Math.floor((ahora - fechaEstado) / (1000 * 60 * 60 * 24));
    const correoSolicitante = datos[i][4];
    const nombreSolicitante = datos[i][3];
    const titulo = datos[i][10];
    const id = datos[i][0];

    if (estado === ESTADOS.PENDIENTE_ARCHIVOS && dias === CONFIG.DIAS_RECORDATORIO_SOLICITANTE) {
      const linkCarpeta = datos[i][19];
      _enviarRecordatorioSolicitante(correoSolicitante, nombreSolicitante, titulo, linkCarpeta, id, dias);
    }

    if (estado === ESTADOS.EN_REVISION && dias >= CONFIG.DIAS_ALERTA_BIBLIOTECARIO) {
      const sede = datos[i][8];
      const correoColab = _obtenerCorreoColaborador(ss, sede);
      GmailApp.sendEmail(
        correoColab,
        `[Digitk] ⏰ Solicitud sin revisar hace ${dias} días — ${titulo}`,
        `La solicitud ID ${id} de ${nombreSolicitante} lleva ${dias} días en estado "En revisión". Por favor revísala.`
      );
    }
  }
}

function _enviarRecordatorioSolicitante(correo, nombre, titulo, linkCarpeta, id, dias) {
  GmailApp.sendEmail(
    correo,
    `[Digitk] ⏰ Recordatorio — Aún no has subido tus archivos`,
    "",
    {
      htmlBody: `<p>Hola <strong>${nombre}</strong>,</p>
      <p>Han pasado <strong>${dias} días</strong> desde que enviaste tu solicitud para <em>"${titulo}"</em> (ID: ${id}), 
      pero aún no detectamos archivos en tu carpeta.</p>
      <p><a href="${linkCarpeta}" style="background:#1a5276;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px;">
      📂 Ir a mi carpeta</a></p>
      <p style="color:#7f8c8d;font-size:12px;">Si ya los subiste, ignora este mensaje.</p>`
    }
  );
}

// ============================================================
//  MENÚ PERSONALIZADO EN EL SHEET
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Sistema Digitk")
    .addItem("⚙️ Ejecutar Setup Inicial", "setupInicial")
    .addSeparator()
    .addItem("📦 Procesar Carga Masiva (Excel)", "procesarCargaMasiva")
    .addSeparator()
    .addItem("🔄 Actualizar Dashboard", "actualizarDashboard")
    .addItem("📊 Ver solicitudes pendientes", "verPendientes")
    .addSeparator()
    .addItem("📋 Descargar Plantilla Masiva", "abrirPlantillaMasiva")
    .addToUi();
}

function actualizarDashboard() {
  // El dashboard usa fórmulas, se actualiza solo. Este botón es para forzar recalcular.
  SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName("Dashboard").getDataRange().getValue();
  SpreadsheetApp.getUi().alert("Dashboard actualizado.");
}

function verPendientes() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hoja = ss.getSheetByName("Solicitudes");
  const datos = hoja.getDataRange().getValues();
  let lista = "";
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][20] === ESTADOS.PENDIENTE_ARCHIVOS || datos[i][20] === ESTADOS.EN_REVISION) {
      lista += `\n• [${datos[i][0]}] ${datos[i][10]} (${datos[i][20]})`;
    }
  }
  SpreadsheetApp.getUi().alert("Solicitudes activas:" + (lista || "\nNinguna."));
}

function abrirPlantillaMasiva() {
  const url = SpreadsheetApp.openById(CONFIG.SHEET_ID).getUrl() + "#gid=" +
    SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName("Plantilla Masiva").getSheetId();
  SpreadsheetApp.getUi().alert("Abre esta URL para ver la plantilla:\n" + url);
}

// ============================================================
//  UTILIDADES
// ============================================================
function _generarID() {
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `DIGITK-${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}-${Math.random().toString(36).substr(2,5).toUpperCase()}`;
}
