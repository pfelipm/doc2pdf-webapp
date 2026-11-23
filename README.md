<p align="center">
  <img src="assets/banner.jpg" alt="Doc2PDF WebApp — Conversor nativo de documentos a PDF para Google Workspace" width="100%" />
</p>

# Doc2PDF WebApp

Microservicio web ligero e interno para **Google Workspace** desarrollado íntegramente en **Google Apps Script** (sin bibliotecas ni servidores externos). Está diseñado para que **un administrador (o cuenta funcional) despliegue una única instancia de la WebApp y ofrezca conversión centralizada de documentos a PDF a todos los usuarios de su dominio**, evitando que el personal o el alumnado tenga que instalar software adicional, mantener infraestructura propia o subir archivos institucionales a conversores online de terceros, mientras mantiene las políticas de acceso, cuotas y auditoría gobernadas desde una Google Sheet privada.

---

## Características principales

- **100 % nativo y sin dependencias externas por defecto:**
  - Backend exclusivamente con Apps Script, **Drive API v3** (servicio avanzado nativo) y `UrlFetchApp`.
  - Frontend y diálogos HTML5 + CSS3 + Vanilla JS + iconos SVG embebidos, incluyendo empaquetador **`.zip` nativo** en navegador para descargar lotes completos sin librerías externas.
  - Visor PDF embebido opcional (`VISOR_PDF_EMBEBIDO = SI`) sobre `<canvas>` con **PDF.js** cargado bajo demanda (si se mantiene en `NO`, abre el PDF directamente en una nueva pestaña con cero dependencias externas).
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
  - Carpeta temporal dedicada `_Doc2PDF_Temp` creada automáticamente dentro de la carpeta donde reside la Google Sheet (con salvaguarda anti-borrado en purga sobre la propia hoja y subcarpetas).
  - Borrado **permanente** inmediato de temporales (`Drive.Files.remove`) en bloque `finally`, más trigger diario de purga para huérfanos.
  - Protección contra *Formula Injection* en `Auditoría` y política configurable de registro de nombres (`COMPLETO`, `ANONIMIZADO`, `NINGUNO`).
- **Panel administrativo en Google Sheets:**
  - Solo 2 pestañas optimizadas: `Ajustes` (con casillas `SI`/`NO`, desplegables, validaciones numéricas estrictas, ajuste multilínea y autoanchura) y `Auditoría` (filtros activos en todas las columnas y cabecera calibrada).
  - Recorte automático de filas y columnas vacías sobrantes al inicializar o reparar (avisando explícitamente del restablecimiento de `Ajustes` y respetando filas/columnas con datos en `Auditoría`).
  - Diálogo modal HTML rico **Resumen y analítica** (KPIs, actividad diaria SVG, donut por familias, top extensiones, usuarios activos y últimos trabajos).
  - Modal **Acerca de Doc2PDF** de atribución con banner gráfico WebP embebido en Base64 desde el menú personalizado.

---

## Arquitectura de seguridad y defensa en profundidad

Siguiendo un diseño *Zero Trust* respecto al navegador y mínimo privilegio en Apps Script:

1. **Mínimo privilegio en hojas de cálculo (`@OnlyCurrentDoc` + `spreadsheets.currentonly`):** el manifiesto restringe el acceso de `SpreadsheetApp` exclusivamente al libro de control contenedor, impidiendo leer o alterar otras hojas del desplegador.
2. **Encapsulación estricta de funciones internas (sufijo `_`):** todas las funciones internas de configuración, caché, lectura/purga de auditoría y conversión llevan sufijo `_` para impedir su invocación directa desde `google.script.run` en consola DevTools.
3. **Control de rol Administrador y tokens efímeros de diálogo (`assertSpreadsheetAdmin_` + `dialogToken`):** a diferencia de un backend web tradicional (donde la app pública y el panel de administración se aíslan por rutas o sesiones distintas), en proyectos Apps Script vinculados a una hoja que además publican una WebApp `google.script.run` aplana todas las funciones globales sin `_` en un único espacio RPC compartido entre la WebApp y los modales de Sheets. Mientras `assertSpreadsheetAdmin_` bloquea a cualquier usuario que no sea propietario o editor de la hoja, el `dialogToken` efímero emitido al abrir cada diálogo añade una capa ligera (~20 líneas, análoga a un token anti-CSRF de contexto) de defensa en profundidad frente a una eventual escalada lateral *WebApp → funciones administrativas de Sheets* bajo la sesión del propio administrador.
4. **Defensa anti-*Confused Deputy* en conversión por enlace de Drive (`assertDriveFileSharedWithRequester_`):** aunque la WebApp ejecuta como cuenta administradora, `convertDriveUrl_` bloquea explícitamente el acceso a la propia hoja de control, a la carpeta temporal interna y a cualquier archivo privado del administrador que no esté compartido explícitamente con el solicitante, con su dominio Workspace o de forma pública.
5. **Política *Fail-Closed* e integridad de sesiones:** cualquier petición sin email institucional verificado (`NO_SESSION_EMAIL`) es rechazada; las sesiones fragmentadas de subida y descarga (`uploadFileChunk_`, `finalizeChunkedConversion_`, `abortChunkedUpload_`, `fetchPdfDownloadChunk_`, `completePdfDownload_`) verifican estrictamente que el solicitante sea el propietario de la sesión.
6. **Protección de hoja de cálculo:** escritura en `Auditoría` serializada con `LockService`, sanitización contra *Formula Injection* (`=`, `+`, `-`, `@`) y modos de registro `COMPLETO` / `ANONIMIZADO` / `NINGUNO`.

---

## Estructura del repositorio

La raíz del repositorio coincide con el directorio raíz de `clasp` (`rootDir: "."`):

```text
.
├── assets/
│   └── banner.jpg           # Banner del repositorio (1376×688, proporción 2:1)
├── .clasp.json.example      # Plantilla de configuración para clasp (rootDir: ".")
├── .claspignore             # Excluye documentación, assets y metadatos al hacer clasp push
├── .gitignore
├── LICENSE                  # GNU GPL v3.0
├── README.md
├── appsscript.json          # Manifiesto V8, Drive API v3, scopes mínimos y acceso DOMAIN
├── Config.gs                # Ajustes cacheados, catálogo de formatos, límites y helpers
├── Setup.gs                 # Menú personalizado, inicialización/recorte de pestañas y diálogos
├── Auth.gs                  # Identidad Workspace, listas de acceso y cuota diaria
├── Audit.gs                 # Auditoría concurrente (LockService), filtros, retención y analítica
├── Cleanup.gs               # Carpeta temporal _Doc2PDF_Temp junto a la hoja y purga segura
├── Converter.gs             # Motor de conversión (directo, chunked y enlace Drive)
├── WebApp.gs                # doGet() y endpoints RPC para google.script.run
├── Index.html               # Estructura de la WebApp (tema claro/oscuro + <dialog>)
├── Styles.html              # CSS adaptativo con soporte light-dark() y fallbacks
├── ClientScript.html        # Cola, drag & drop, chunking y vista previa PDF
├── DashboardDialog.html     # Diálogo modal de Resumen y analítica (SVG nativo)
├── AboutDialog.html         # Diálogo modal "Acerca de Doc2PDF" con banner Base64 WebP
├── InitDialog.html          # Confirmación de reinicialización con aviso de borrado/recorte
└── DeployDialog.html        # Asistente paso a paso para guardar la URL de WebApp
```

---

## Instalación y despliegue con `clasp`

1. **Crear la hoja de cálculo de control**
   - Desde la cuenta administradora o **cuenta funcional del centro** (recomendado, p. ej. `servicios-edtech@tucentro.edu`), crea una Google Sheet nueva dentro de una **carpeta privada** del administrador (por ejemplo, `Doc2PDF · Control`).
   - Abre **Extensiones → Apps Script**, ve a **Configuración del proyecto** (icono de engranaje) y copia el **ID del script**.

2. **Clonar, vincular y subir el código con `clasp`**
   En la raíz del repositorio clonado ejecuta:
   ```bash
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

La pestaña `Ajustes` incorpora validaciones estrictas de datos sobre sus 21 parámetros:

| Clave | Valor inicial | Tipo / Validación | Descripción |
| :--- | :--- | :--- | :--- |
| `NOMBRE_APP` | `Doc2PDF` | Texto (máx. 60 car.) | Título principal de la WebApp y pestaña |
| `SUBTITULO_CENTRO` | `Conversor interno de documentos a PDF` | Texto (máx. 120 car.) | Subtítulo institucional en cabecera |
| `COLOR_ACENTO` | `#1a73e8` | Color Hex (`#RRGGBB`) | Color principal de botones, iconos y realces |
| `LOGO_URL` | *(vacío)* | URL `https://` | URL HTTPS opcional del logotipo |
| `SERVICIO_ACTIVO` | `SI` | Casilla (`SI` / `NO`) | Activa (`SI`) o pausa en mantenimiento (`NO`) |
| `MENSAJE_MANTENIMIENTO` | *(mensaje estándar)* | Texto (máx. 300 car.) | Aviso mostrado cuando `SERVICIO_ACTIVO = NO` |
| `DOMINIOS_PERMITIDOS` | *(vacío)* | Lista por comas | Dominios Workspace permitidos (vacío = dominio del desplegador) |
| `LISTA_BLANCA_USUARIOS` | *(vacío)* | Lista por comas | Correos autorizados (vacío = todo el dominio) |
| `LISTA_NEGRA_USUARIOS` | *(vacío)* | Lista por comas | Correos bloqueados |
| `LIMITE_DIARIO_USUARIO` | `30` | Numérico (`0 – 10 000`) | Conversiones máximas por usuario al día (`0` = sin límite; sugerido) |
| `TAMANO_MAX_MB` | `20` | Numérico (`1 – 45`) | Peso máximo por archivo en MB *(límite técnico seguro bajo tope de 50 MB de Apps Script)* |
| `MAX_ARCHIVOS_LOTE` | `5` | Numérico (`1 – 20`) | Máximo de archivos admitidos por lote al arrastrar o seleccionar *(sugerido)* |
| `PERMITIR_DOCUMENTOS` | `SI` | Casilla (`SI` / `NO`) | Habilita documentos de texto (`.docx`, `.doc`, `.dot`, `.dotx`, `.odt`, `.rtf`) |
| `PERMITIR_HOJAS` | `SI` | Casilla (`SI` / `NO`) | Habilita hojas de cálculo (`.xlsx`, `.xls`, `.xlt`, `.xltx`, `.xlsm`, `.ods`, `.csv`, `.tsv`) |
| `PERMITIR_PRESENTACIONES` | `SI` | Casilla (`SI` / `NO`) | Habilita presentaciones (`.pptx`, `.ppt`, `.odp`) |
| `PERMITIR_TEXTO` | `SI` | Casilla (`SI` / `NO`) | Habilita texto plano y marcado (`.txt`, `.md`, `.html`, `.htm`) |
| `PERMITIR_ENLACE_DRIVE` | `SI` | Casilla (`SI` / `NO`) | Habilita la pestaña de conversión por enlace/ID de Drive |
| `VISOR_PDF_EMBEBIDO` | `NO` | Casilla (`SI` / `NO`) | Activa (`SI`) el visor modal en `<canvas>` con PDF.js bajo demanda; si es `NO`, abre el PDF en pestaña nueva sin librerías externas |
| `REGISTRAR_NOMBRE_ARCHIVO` | `COMPLETO` | Desplegable (`COMPLETO` / `ANONIMIZADO` / `NINGUNO`) | Nivel de privacidad del nombre registrado en `Auditoría` |
| `RETENCION_AUDITORIA_DIAS` | `180` | Numérico (`0 – 3 650`) | Días de retención automática (`0` = indefinido; sugerido) |
| `CARPETA_TEMP_ID` | *(automático)* | Texto (ID Drive) | ID de la subcarpeta `_Doc2PDF_Temp` creada junto a esta hoja |

> **Notas importantes:**
> - **Caché:** los ajustes se cachean durante **5 minutos**. Tras editar la pestaña `Ajustes`, puedes aplicar los cambios al instante ejecutando **`📄 Doc2PDF → 🛠️ Mantenimiento → 🔄 Limpiar caché de ajustes`**.
> - **Equipo administrador y modo mantenimiento:** el rol de administración se determina por la cuenta desplegadora del script más el **propietario y los editores explícitos** de la Google Sheet de control (cacheado 60 s). Quienes pertenezcan a este grupo mantienen acceso de pruebas con un banner de aviso cuando `SERVICIO_ACTIVO = NO` (mientras el resto de usuarios queda bloqueado) y quedan exentos de `LISTA_BLANCA_USUARIOS` y `DOMINIOS_PERMITIDOS`.
> - **Concurrencia de conversión:** `MAX_ARCHIVOS_LOTE` limita cuántos archivos entran en cola en cada envío; la WebApp los procesa de **2 en 2** en paralelo (`DIRECT_CONCURRENCY = 2`) para subidas directas (`<= 6 MB`) y enlaces de Drive, y de **1 en 1** en subidas fragmentadas (`> 6 MB`), protegiendo así el cupo compartido de 30 ejecuciones simultáneas de Apps Script.
> - **Carpeta temporal y privacidad:** `_Doc2PDF_Temp` se crea automáticamente como subcarpeta en la misma carpeta donde reside la Google Sheet (heredando sus permisos). Se recomienda alojar la hoja de control en una **carpeta privada** del administrador o cuenta funcional.

---

## Limitaciones conocidas

Doc2PDF prioriza un despliegue simple, sin servidores intermedios ni mantenimiento de infraestructura externa, asumiendo a cambio las restricciones propias de Google Apps Script y del motor de conversión de Google Workspace.

### Calidad y fidelidad de la conversión

- **Motor basado en importación/exportación de Google Workspace:** Doc2PDF no utiliza Microsoft Office ni LibreOffice. Para convertir, importa temporalmente el documento al formato nativo equivalente de Google (Docs, Sheets o Slides) y lo exporta a PDF. Es una solución práctica para el día a día, pero no replica al 100 % el renderizado del programa de escritorio original.
- **Maquetación compleja y tipografías:** en archivos `.docx` o `.pptx` que emplean fuentes no disponibles en Workspace, macros/VBA (`.xlsm`, que se omiten en la conversión), controles ActiveX, objetos SmartArt, fórmulas matemáticas complejas, cuadros de texto flotantes o encabezados y pies avanzados, pueden producirse sustituciones tipográficas, saltos de página distintos o ligeros desajustes visuales.
- **Hojas de cálculo (`.xlsx`, `.ods`, `.csv`):** aunque la WebApp permite ajustar tamaño de papel, orientación, ajuste al ancho, cuadrícula y márgenes, las áreas de impresión complejas o saltos de página manuales definidos en Excel pueden diferir respecto a la exportación nativa.
- **Texto y marcado (`.md`, `.html`, `.txt`):** la conversión mediante Drive aplica un renderizado básico; no ejecuta JavaScript, no resuelve hojas de estilo o recursos externos complejos ni interpreta extensiones avanzadas (como diagramas Mermaid o bloques LaTeX).
- **Archivos cifrados o dañados:** documentos protegidos con contraseña, cifrado IRM/DRM o estructura corrupta fallarán durante la importación temporal en Drive.
- **Uso recomendado:** pensado para tramitación interna y difusión ágil en centros educativos u organizaciones; cuando se requiera fidelidad tipográfica estricta o calidad de imprenta, conviene generar el PDF directamente desde la aplicación de origen.

### Limitaciones de Google Apps Script y Workspace

- **Límite de tamaño (`<= 45 MB`):** aunque la subida y descarga fragmentadas evitan los límites de `google.script.run` y el tope de `10 MB` de `Drive.Files.export`, Google Apps Script impone un máximo duro de **`50 MB`** en `UrlFetchApp` y memoria de ejecución. Tanto el archivo original como el PDF generado deben quedar por debajo de ese umbral (`TAMANO_MAX_MB` configurable entre `1` y `45 MB`).
- **Tiempo máximo de ejecución y latencia:** la documentación oficial de Apps Script fija un límite de **6 minutos por ejecución** —aunque en la práctica en entornos Google Workspace se hayan observado márgenes superiores (hasta 30 minutos) no garantizados contractualmente, y las llamadas interactivas `google.script.run` siguen sujetas también a los tiempos de espera del navegador y `UrlFetchApp`—; además, cada conversión depende de los arranques en frío (*cold starts*) y del tiempo de importación/exportación en Drive ante documentos muy extensos o cargados de imágenes pesadas.
- **Concurrencia y cuotas compartidas (`USER_DEPLOYING`):** al ejecutarse bajo la cuenta administradora o funcional que despliega la WebApp, las cuotas de Apps Script (por ejemplo, el límite de **30 ejecuciones simultáneas por usuario** y los cupos diarios de `UrlFetchApp` y Drive API) se comparten entre todo el personal del dominio. El proyecto está dimensionado para uso interno departamental o de centro, no para tráfico concurrente masivo.
- **Escalado de `Auditoría` en Google Sheets:** las escrituras se serializan con `LockService` (espera máxima de 15 s) para evitar colisiones, y la hoja de cálculo perderá fluidez si acumula decenas de miles de registros sin purga periódica (`RETENCION_AUDITORIA_DIAS`).
- **Entorno *iframe* y multicuenta de Google:** las WebApps de Apps Script corren en un *iframe* aislado (`script.google.com`), por lo que sesiones del navegador con múltiples cuentas de Google abiertas a la vez pueden fallar si la cuenta principal del navegador no pertenece al dominio Workspace. Del mismo modo, las cabeceras de seguridad (`CSP` / `X-Frame-Options`) impiden incrustar un PDF nativo mediante `<iframe>` en ciertos navegadores (motivo por el que el visor embebido opcional usa PDF.js sobre `<canvas>` y el modo sin dependencias abre el PDF en una pestaña nueva).
- **Endpoints de exportación de Docs/Sheets/Slides:** para evitar el límite de 10 MB de `Drive.Files.export`, la exportación utiliza las URLs de `docs.google.com/.../export` autenticadas con el token OAuth del script (`ScriptApp.getOAuthToken()`). Es una vía consolidada en Workspace, pero no forma parte de un contrato REST versionado como Drive API v3 y podría verse afectada si Google cambiara esos endpoints.

---

## Licencia y autoría

Desarrollado por **Pablo Felip** ([@pfelipm](https://pablofelip.online/sobre-mi/)) bajo licencia **GNU GPL v3.0** ([LICENSE](LICENSE)).
