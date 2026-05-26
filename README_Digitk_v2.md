# 📚 Sistema de Cargue Repositorio Institucional Digitk v2.0
**Fundación Universitaria del Área Andina — Bibliotecas Areandina**

![Estado](https://img.shields.io/badge/Estado-En%20desarrollo-yellow)
![Stack](https://img.shields.io/badge/Stack-Google%20Suite-4285F4?logo=google)
![Versión](https://img.shields.io/badge/Versión-2.0-blue)

---

## ¿Qué hace este sistema?

Automatiza el flujo completo de solicitud, organización y seguimiento de documentos académicos para su cargue al repositorio institucional Digitk. Elimina la revisión manual carpeta por carpeta y reemplaza el Excel manual por un proceso notificado y trazable.

**Flujo antes:** Alguien crea una carpeta manualmente → sube PDFs → el bibliotecario recorre todas las carpetas buscando documentos → llena el Excel manualmente.

**Flujo con el sistema:** Solicitante llena un formulario → el sistema crea la carpeta en Drive automáticamente → notifica al solicitante con el link → el bibliotecario recibe alerta y aprueba con un clic → el Excel se llena solo.

---

## Estructura de carpetas en Drive

```
GR-GRI-Opciones de grado
└── Facultad
    └── Programa
        └── Sede (Bogotá / Pereira / Valledupar / Virtual)
            ├── Apellido, Nombre Año      ← solicitud activa
            └── Proceso terminado
                └── Apellido, Nombre Año  ← ya cargado en Digitk
```

---

## Dos modos de carga

### Modo Individual
Para docentes, investigadores o administrativos que suben 1–5 documentos.
- Llenan el Google Form
- El sistema crea la carpeta automáticamente en la ruta correcta
- Reciben el link por correo para subir sus 3 PDFs
- El bibliotecario de su sede recibe notificación con botones de acción

### Modo Masivo
Para administrativos de facultad que recopilan 20–200 trabajos al cierre de semestre.
- Descargan la **Plantilla Masiva** desde la hoja del mismo nombre en el Sheet
- Llenan una fila por documento
- Desde el menú **📚 Sistema Digitk → Procesar Carga Masiva**, suben el Excel
- El sistema procesa todos los registros, crea todas las carpetas y envía todos los correos
- El bibliotecario recibe un **resumen consolidado** con el estado de cada documento

---

## Instalación paso a paso

### Paso 1 — Crear el Google Sheet
1. Ve a [sheets.google.com](https://sheets.google.com)
2. Crea un Spreadsheet nuevo
3. Copia el ID de la URL: `docs.google.com/spreadsheets/d/**ESTE_ID**/edit`

### Paso 2 — Pegar el código en Apps Script
1. En el Sheet: **Extensiones → Apps Script**
2. Borra el código por defecto
3. Pega todo el contenido de `DigitK_Sistema_Repositorio_v2.gs`
4. Actualiza el bloque `CONFIG` con tus IDs reales:

```javascript
const CONFIG = {
  SHEET_ID: "ID_DE_TU_SHEET",
  DRIVE_ROOT_FOLDER_ID: "ID_DE_CARPETA_GR-GRI-REPOSITORIO",
  CORREO_ADMIN: "tu.correo@areandina.edu.co",
  ...
};
```

5. Guarda con `Ctrl+S`

### Paso 3 — Ejecutar el setup inicial
1. En Apps Script, selecciona la función `setupInicial`
2. Clic en **Ejecutar ▶️**
3. Acepta los permisos cuando los solicite
4. Verifica que se crearon las 5 hojas: **Solicitudes, Colaboradores, Dashboard, Catálogo, Plantilla Masiva**

### Paso 4 — Llenar la hoja Colaboradores
Reemplaza los datos de ejemplo con los correos reales de cada sede:

| Sede | Nombre | Correo | Cargo | Activo |
|------|--------|--------|-------|--------|
| Bogotá | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Pereira | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Valledupar | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |
| Virtual | Nombre real | correo@areandina.edu.co | Bibliotecólogo | SÍ |

### Paso 5 — Crear el Google Form (Modo Individual)
Crea un formulario con estas preguntas en orden exacto:

| # | Pregunta | Tipo |
|---|----------|------|
| 1 | Nombre completo | Texto corto |
| 2 | Correo institucional | Texto corto |
| 3 | Rol | Opción múltiple: Docente / Investigador / Administrativo / Área de Investigaciones |
| 4 | Facultad | Texto corto |
| 5 | Programa académico | Texto corto |
| 6 | Tipo de material | Opción múltiple (ver lista abajo) |
| 7 | Título del material | Párrafo |
| 8 | Autor(es) | Párrafo |
| 9 | Director / Asesor (si aplica) | Texto corto |
| 10 | Año | Texto corto |
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

**Vincula el Form al Sheet:**
- En el Form: Respuestas → ícono Sheets → selecciona el Sheet existente → hoja "Solicitudes"

### Paso 6 — Publicar como Web App (activa los botones del correo)
1. En Apps Script: **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Ejecutar como: **Yo**
4. Acceso: **Cualquier persona**
5. Clic en **Implementar** — copia la URL generada
6. Esa URL ya está hardcodeada en el script con `ScriptApp.getService().getUrl()`, no necesitas pegarla manualmente

---

## Uso del Modo Masivo

1. En el Sheet, ve a la hoja **Plantilla Masiva**
2. Descárgala como `.xlsx` (**Archivo → Descargar → Microsoft Excel**)
3. Llena una fila por documento (respeta los encabezados, no los modifiques)
4. Sube el Excel a una carpeta en Drive y copia su ID desde la URL
5. En el Sheet: menú **📚 Sistema Digitk → Procesar Carga Masiva**
6. Pega el ID de la carpeta cuando lo solicite
7. El sistema procesa todo y te envía el resumen por correo

---

## Flujo de estados

```
PENDIENTE DE ARCHIVOS
        ↓ (solicitante sube PDFs)
   EN REVISIÓN
        ↓ (bibliotecario verifica)
   ┌────────────────────────────┐
   ↓              ↓             ↓
APROBADO      INCOMPLETO    RECHAZADO
   ↓
CARGADO EN DIGITK
   ↓
[Carpeta movida a "Proceso terminado"]
```

---

## Límites del sistema (Google gratuito)

| Recurso | Límite | Para Digitk |
|---------|--------|-------------|
| Correos/día (personal) | 100 | Suficiente para uso normal |
| Correos/día (Workspace) | 1,500 | Recomendado para carga masiva |
| Tiempo ejecución script | 6 min | Masiva de ~100 docs: ~3 min |
| Triggers | 20 | Se usan 2 |

> **Nota:** Para cargas masivas de más de 80 documentos en un solo lote, se recomienda usar una cuenta Google Workspace institucional para no alcanzar el límite de correos diarios.

---

## Archivos del repositorio

| Archivo | Descripción |
|---------|-------------|
| `DigitK_Sistema_Repositorio_v2.gs` | Código Apps Script completo |
| `README_Digitk_v2.md` | Esta guía |

---

## Créditos

Desarrollado por **Bibliotecas Areandina / REKOL08**  
Fundación Universitaria del Área Andina  
Sistema compatible con Google Workspace Educación
