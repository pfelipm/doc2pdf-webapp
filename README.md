# Doc2PDF WebApp

Microservicio web ligero e interno para **Google Workspace** desarrollado íntegramente en **Google Apps Script** (sin bibliotecas ni servidores externos) que permite convertir documentos de escritorio a **PDF** desde una interfaz web moderna, controlada desde una hoja de cálculo de administración.

---

## Características principales

- **100 % nativo y sin dependencias externas:**
  - Backend exclusivamente con Apps Script, **Drive API v3** (servicio avanzado nativo) y `UrlFetchApp`.
  - Frontend y diálogos HTML5 + CSS3 + Vanilla JS + iconos/gráficos SVG embebidos (sin CDNs ni librerías JS externas).
- **Soporte completo de documentos de escritorio (21 extensiones + nativos Workspace):**
  - **Documentos:** `.docx`, `.doc`, `.dot`, `.dotx`, `.odt`, `.rtf`
  - **Hojas de cálculo:** `.xlsx`, `.xls`, `.xlt`, `.xltx`, `.xlsm`, `.ods`, `.csv`, `.tsv` *(con selector de papel, orientación, ajuste al ancho, cuadrícula y márgenes)*
  - **Presentaciones:** `.pptx`, `.ppt`, `.odp`
  - **Texto y marcado:** `.txt`, `.md`, `.html`, `.htm`
  - **Enlaces de Google Drive:** conversión directa desde URL/ID compartido dentro del dominio (tanto archivos de escritorio en Drive como Google Docs, Sheets y Slides nativos).
- **Superación de límites prácticos de Apps Script:**
  - Exportación PDF vía endpoints oficiales de `docs.google.com/.../export` con `UrlFetchApp` (evita el límite de 10 MB de `Drive.Files.export`).
  - Subida en dos velocidades: directa Base64 (`<= 6 MB`) y fragmentada por bloques de 2 MB sobre **Resumable Upload de Drive API v3** (`> 6 MB`).
  - Descarga fragmentada automática cuando el PDF resultante supera `7 MB`.
- **Privacidad, control y cumplimiento RGPD:**
  - Modelo **Opción A**: ejecutado como administrador / cuenta funcional del centro (`executeAs: "USER_DEPLOYING"`) y restringido al dominio Workspace (`access: "DOMAIN"`). La Google Sheet de control permanece **100 % privada**.
  - Borrado **permanente** inmediato de temporales (`Drive.Files.remove`) en bloque `finally`, más trigger diario de purga para huérfanos.
  - Protección contra *Formula Injection* en `Auditoría` y política configurable de registro de nombres (`COMPLETO`, `ANONIMIZADO`, `NINGUNO`).
- **Panel administrativo en Google Sheets:**
  - Solo 2 pestañas en la hoja: `Ajustes` y `Auditoría`.
  - Diálogo modal HTML rico **Resumen y analítica** (KPIs, actividad diaria SVG, donut por familias, top extensiones, usuarios activos y últimos trabajos).
  - Modal **Acerca de Doc2PDF** de atribución desde el menú personalizado.

---

## Arquitectura de seguridad y defensa en profundidad

Siguiendo un diseño *Zero Trust* respecto al navegador y mínimo privilegio en Apps Script:

1. **Mínimo privilegio en hojas de cálculo (`@OnlyCurrentDoc` + `spreadsheets.currentonly`):** el manifiesto restringe el acceso de `SpreadsheetApp` exclusivamente al libro de control contenedor, impidiendo leer o alterar otras hojas del desplegador.
2. **Encapsulación estricta de funciones internas (sufijo `_`):** todas las funciones internas de configuración, caché, lectura/purga de auditoría y conversión llevan sufijo `_` para impedir su invocación directa desde `google.script.run` en consola DevTools.
3. **Control de rol Administrador y tokens efímeros de diálogo (`assertSpreadsheetAdmin_` + `dialogToken`):** los manejadores de menú y endpoints RPC de los diálogos de Sheets validan en el servidor que la cuenta activa pertenezca al propietario/editores de la hoja de cálculo y presentan un `dialogToken` efímero emitido en servidor al abrir el diálogo.
4. **Defensa anti-*Confused Deputy* en conversión por enlace de Drive (`assertDriveFileSharedWithRequester_`):** aunque la WebApp ejecuta como cuenta administradora, `convertDriveUrl_` bloquea explícitamente el acceso a la propia hoja de control, a la carpeta temporal interna y a cualquier archivo privado del administrador que no esté compartido explícitamente con el solicitante, con su dominio Workspace o de forma pública.
5. **Política *Fail-Closed* e integridad de sesiones:** cualquier petición sin email institucional verificado (`NO_SESSION_EMAIL`) es rechazada; las sesiones fragmentadas de subida y descarga (`uploadFileChunk_`, `finalizeChunkedConversion_`, `abortChunkedUpload_`, `fetchPdfDownloadChunk_`, `completePdfDownload_`) verifican estrictamente que el solicitante sea el propietario de la sesión.
6. **Protección de hoja de cálculo:** escritura en `Auditoría` serializada con `LockService`, sanitización contra *Formula Injection* (`=`, `+`, `-`, `@`) y modos de registro `COMPLETO` / `ANONIMIZADO` / `NINGUNO`.

---

## Estructura del repositorio (`src/`)

La raíz del repositorio Git y de `clasp` vive dentro de `src/`, dejando libre el directorio padre (`doc2pdf-webapp/`) para recursos auxiliares locales o de difusión fuera del control de versiones:

```text
src/
├── .clasp.json.example      # Plantilla de configuración para clasp (rootDir: ".")
├── .claspignore
├── .gitignore
├── LICENSE                  # GNU GPL v3.0
├── README.md
├── appsscript.json          # Manifiesto V8, Drive API v3, scopes mínimos y acceso DOMAIN
├── Config.gs                # Ajustes cacheados, catálogo de formatos y helpers
├── Setup.gs                 # Menú personalizado, inicialización de pestañas y diálogos
├── Auth.gs                  # Identidad Workspace, listas de acceso y cuota diaria
├── Audit.gs                 # Auditoría concurrente (LockService), retención y analítica
├── Cleanup.gs               # Carpeta temporal y purga permanente inmediata/programada
├── Converter.gs             # Motor de conversión (directo, chunked y enlace Drive)
├── WebApp.gs                # doGet() y endpoints RPC para google.script.run
├── Index.html               # Estructura de la WebApp (tema claro/oscuro + <dialog>)
├── Styles.html              # CSS adaptativo con soporte light-dark() y fallbacks
├── ClientScript.html        # Cola, drag & drop, chunking y vista previa PDF
├── DashboardDialog.html     # Diálogo modal de Resumen y analítica (SVG nativo)
├── AboutDialog.html         # Diálogo modal "Acerca de Doc2PDF"
├── InitDialog.html          # Confirmación de reinicialización de pestañas
└── DeployDialog.html        # Asistente paso a paso para guardar la URL de WebApp
```

---

## Instalación y despliegue con `clasp`

1. **Crear la hoja de cálculo de control**
   - Desde la cuenta administradora o **cuenta funcional del centro** (recomendado, p. ej. `servicios-edtech@tucentro.edu`), crea una Google Sheet nueva (por ejemplo, `Doc2PDF · Control`).
   - Abre **Extensiones → Apps Script**, ve a **Configuración del proyecto** (icono de engranaje) y copia el **ID del script**.

2. **Vincular y subir el código con `clasp`**
   Sitúate dentro de la carpeta del repositorio (`src/`) y ejecuta:
   ```bash
   cd src
   cp .clasp.json.example .clasp.json
   # Edita .clasp.json y sustituye REEMPLAZA_CON_EL_SCRIPT_ID_DE_TU_HOJA por tu scriptId
   clasp login
   clasp push
   ```

3. **Inicializar la hoja y el trigger de mantenimiento**
   - Recarga la hoja de cálculo y espera unos segundos hasta que aparezca el menú **`📄 Doc2PDF`**.
   - Ejecuta **`📄 Doc2PDF → ✨ Inicializar / reparar hoja`** (autoriza los permisos solicitados la primera vez).
   - Ejecuta **`📄 Doc2PDF → 🛠️ Mantenimiento → ⏱️ Instalar trigger diario de limpieza`**.

4. **Desplegar como WebApp en Google Workspace**
   - En el editor de Apps Script, haz clic en **Implementar → Nueva implementación → Aplicación web**.
   - **Ejecutar como:** `Yo` *(la cuenta administradora/funcional propietaria de la hoja)*.
   - **Quién tiene acceso:** `Cualquier usuario de [tu dominio Workspace]`.
   - Copia la URL generada terminada en `/exec`.
   - En la hoja de cálculo, abre **`📄 Doc2PDF → 🚀 Configurar URL de WebApp`**, pega la URL y pulsa **Guardar URL**.

---

## Ajustes configurables (`Ajustes`)

| Clave | Valor inicial | Descripción |
| :--- | :--- | :--- |
| `NOMBRE_APP` | `Doc2PDF` | Título principal de la WebApp |
| `SUBTITULO_CENTRO` | `Conversor interno de documentos a PDF` | Subtítulo en cabecera |
| `COLOR_ACENTO` | `#1a73e8` | Color de acento hexadecimal (`#RRGGBB`) |
| `LOGO_URL` | *(vacío)* | URL HTTPS opcional del logotipo |
| `SERVICIO_ACTIVO` | `SI` | Activa (`SI`) o pausa en mantenimiento (`NO`) |
| `MENSAJE_MANTENIMIENTO` | *(mensaje estándar)* | Aviso mostrado cuando `SERVICIO_ACTIVO = NO` |
| `DOMINIOS_PERMITIDOS` | *(vacío)* | Dominios Workspace separados por comas (vacío = dominio del desplegador) |
| `LISTA_BLANCA_USUARIOS` | *(vacío)* | Correos permitidos (vacío = todo el dominio) |
| `LISTA_NEGRA_USUARIOS` | *(vacío)* | Correos bloqueados |
| `LIMITE_DIARIO_USUARIO` | `30` | Conversiones máximas por usuario al día (`0` = sin límite) |
| `TAMANO_MAX_MB` | `20` | Peso máximo por archivo (`1` a `45` MB) |
| `MAX_ARCHIVOS_LOTE` | `5` | Máximo de archivos simultáneos por lote (`1` a `20`) |
| `FAMILIAS_HABILITADAS` | `DOCUMENTOS,HOJAS,PRESENTACIONES,TEXTO` | Familias de formatos habilitadas |
| `PERMITIR_ENLACE_DRIVE` | `SI` | Habilita la pestaña de conversión por enlace/ID de Drive |
| `REGISTRAR_NOMBRE_ARCHIVO` | `COMPLETO` | `COMPLETO`, `ANONIMIZADO` (hash SHA-256) o `NINGUNO` |
| `RETENCION_AUDITORIA_DIAS` | `180` | Días de retención automática (`0` = indefinido) |
| `CARPETA_TEMP_ID` | *(automático)* | ID de `_Doc2PDF_Temp` en Drive |

> **Nota sobre caché:** los ajustes se cachean durante **5 minutos**. Tras editar la pestaña `Ajustes`, puedes aplicar los cambios al instante ejecutando **`📄 Doc2PDF → 🛠️ Mantenimiento → 🔄 Limpiar caché de ajustes`**.

---

## Licencia y autoría

Desarrollado por **Pablo Felip** ([@pfelipm](https://pablofelip.online/sobre-mi/)) bajo licencia **GNU GPL v3.0** ([LICENSE](LICENSE)).
