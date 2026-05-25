# 📚 Sistema de Cargue Repositorio Institucional — Digitk Areandina

> Sistema de automatización para la gestión de solicitudes de publicación en el Repositorio Institucional Digitk de la Fundación Universitaria del Área Andina.

---

## ¿Qué hace el sistema?

Automatiza todo el flujo de cargue al repositorio Digitk, **desde la solicitud hasta la notificación de publicación**, dejando el paso de cargue manual en Digitk en manos del bibliotecario.

### Flujo completo

```
Docente/Investigador llena el Form
         ↓
Apps Script valida dominio @areandina.edu.co
         ↓
Crea carpeta en Drive (Facultad → Programa → Sede)
         ↓
Notifica al solicitante con link de la carpeta
         ↓
Notifica al colaborador de biblioteca de la sede
         ↓
Colaborador revisa documentos y aprueba/rechaza (un clic desde el correo)
         ↓
Notificación automática al solicitante con resultado
         ↓
Bibliotecario carga manualmente en Digitk  ← (paso manual, no se toca)
         ↓
Marca como "Cargada en Digitk" desde el Sheet
         ↓
Notificación final al solicitante ✅
```

---

## Requisitos previos

- Cuenta Google institucional (`@areandina.edu.co`)
- Google Sheets, Forms, Drive y Gmail activos
- Carpeta raíz `GR-GRI-REPOSITORIO` ya creada en Google Drive

---

## Instalación paso a paso

### 1. Preparar el Google Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea un nuevo Spreadsheet
2. Copia el **ID** de la URL: `docs.google.com/spreadsheets/d/**ESTE_ID**/edit`
3. Renombra el Sheet como `Sistema Repositorio Digitk — Areandina`

### 2. Pegar el código en Apps Script

1. En el Sheet: **Extensiones → Apps Script**
2. Borra el código por defecto
3. Pega todo el contenido de `DigitK_Sistema_Repositorio.gs`
4. Reemplaza en la sección `CONFIG`:
   - `TU_SHEET_ID_AQUI` → el ID del Sheet
   - `TU_FOLDER_ID_AQUI` → el ID de la carpeta raíz en Drive
   - `bibliotecas@areandina.edu.co` → el correo real del administrador

### 3. Ejecutar el setup inicial

1. En Apps Script, selecciona la función `setupInicial`
2. Clic en **Ejecutar** (concede los permisos solicitados)
3. Esto crea automáticamente las 4 hojas del sistema y los triggers

### 4. Registrar los colaboradores

Ve a la hoja **Colaboradores** y reemplaza los datos de ejemplo con los reales:

| Sede | Nombre | Correo | Cargo | Activo |
|------|--------|--------|-------|--------|
| Bogotá | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Pereira | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Valledupar | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Virtual | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |

### 5. Crear el Google Form

Crea un formulario con las siguientes preguntas **en este orden exacto**:

| # | Pregunta | Tipo |
|---|----------|------|
| 1 | Nombre completo | Texto corto |
| 2 | Correo institucional | Texto corto |
| 3 | Rol | Opción múltiple: Docente / Investigador / Administrativo / Área de Investigaciones |
| 4 | Facultad | Lista desplegable |
| 5 | Programa académico | Texto corto |
| 6 | Tipo de material | Opción múltiple (ver lista abajo) |
| 7 | Título del material | Texto largo |
| 8 | Autor(es) | Texto largo |
| 9 | Director / Asesor (si aplica) | Texto corto |
| 10 | Año de publicación/elaboración | Texto corto |
| 11 | ISBN / ISSN / DOI (si aplica) | Texto corto |
| 12 | Palabras clave | Texto corto |
| 13 | Resumen | Párrafo |
| 14 | Idioma | Opción múltiple: Español / Inglés / Otro |
| 15 | Número de páginas | Texto corto |
| 16 | Sede | Opción múltiple: Bogotá / Pereira / Valledupar / Virtual |

**Tipos de material:**
- Trabajo de grado (pregrado)
- Trabajo de grado (posgrado / maestría)
- Artículo de revista
- Libro / capítulo de libro
- Ponencia / congreso
- Informe de investigación
- Recurso audiovisual
- Otro

**Vincular el Form al Sheet:**
En el Form: Respuestas → ícono de Sheets → Seleccionar hoja existente → hoja `Solicitudes`

### 6. Publicar como Web App (para botones de acción en correo)

1. En Apps Script: **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Ejecutar como: **Yo**
4. Acceso: **Cualquier persona**
5. Clic en **Implementar** → copia la URL generada

---

## Estructura de Drive generada automáticamente

```
GR-GRI-REPOSITORIO/
└── [Facultad]/
    └── [Programa]/
        └── [Sede: Bogotá | Pereira | Valledupar | Virtual]/
            └── [DGT-BOG-TGP-20260525-4821] Título del trabajo.../
                ├── 📄 Documento Principal PDF/
                ├── ✅ Autorización Publicación/
                ├── 📎 Documentos Complementarios/
                └── METADATOS_DGT-BOG-TGP-20260525-4821.txt
```

---

## Estados del flujo

| Estado | Significado | Color en Sheet |
|--------|-------------|----------------|
| ⏳ Recibida | Solicitud registrada, pendiente de documentos | Amarillo |
| 🔵 En revisión | Bibliotecario revisando documentos | Azul claro |
| ⚠️ Documentos incompletos | Faltan documentos en la carpeta Drive | Naranja |
| ✅ Aprobada | Documentos completos, listos para Digitk | Verde claro |
| ❌ Rechazada | Solicitud no aprobada | Rojo claro |
| 🎉 Cargada en Digitk | Material publicado en el repositorio | Cian |

---

## Automatizaciones incluidas

| Automatización | Cuándo se activa |
|----------------|-----------------|
| Creación de carpeta en Drive | Al enviar el formulario |
| Correo de confirmación al solicitante | Al enviar el formulario |
| Correo al colaborador de biblioteca | Al enviar el formulario |
| Correo de resultado al solicitante | Al aprobar / rechazar / marcar incompleto |
| Recordatorio al solicitante | 5 días sin cargar documentos |
| Alerta al colaborador | 3 días sin respuesta en revisión |
| Actualización del Catálogo | Al marcar como "Cargada en Digitk" |

---

## Acciones rápidas desde el Sheet

El menú **📚 Sistema Digitk** en Google Sheets permite:
- ✅ Aprobar solicitud seleccionada con un clic
- 📤 Marcar como cargada en Digitk (pide el URL)
- 📋 Ver cuántas solicitudes están pendientes

---

## Costos

| Componente | Costo |
|------------|-------|
| Google Forms | Gratuito |
| Google Sheets | Gratuito |
| Google Apps Script | Gratuito (100 correos/día cuenta personal, 1.500 con Workspace) |
| Google Drive | Gratuito hasta 15 GB |
| **Total** | **$0** |

---

*Sistema desarrollado para Bibliotecas Areandina — Bogotá · Pereira · Valledupar · Virtual*
