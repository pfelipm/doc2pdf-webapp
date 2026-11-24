<p align="center">
  <img src="assets/banner.jpg" alt="Doc2PDF WebApp — Conversor nativo de documentos a PDF para Google Workspace" width="100%" />
</p>

# Doc2PDF WebApp

![Versión v1.0](assets/badges/version.svg)
[![Licencia GNU GPL v3.0](assets/badges/license.svg)](LICENSE)
![Google Apps Script V8](assets/badges/apps-script.svg)
![Drive API v3](assets/badges/drive-api.svg)

Microservicio web ligero e interno para **Google Workspace** desarrollado íntegramente en **Google Apps Script** (sin bibliotecas ni servidores externos). Está diseñado para que **un administrador (o cuenta funcional) despliegue una única instancia de la WebApp y ofrezca conversión centralizada de documentos a PDF a todos los usuarios de su dominio**, evitando que el personal o el alumnado tenga que instalar software adicional, mantener infraestructura propia o subir archivos institucionales a conversores online de terceros, mientras mantiene las políticas de acceso, cuotas y auditoría gobernadas desde una Google Sheet privada.

---

## Características principales

- **100 % nativo y sin dependencias externas por defecto:**
  - Backend exclusivamente con Apps Script, **Drive API v3** (servicio avanzado nativo) y `UrlFetchApp`.
  - Frontend y diálogos HTML5 + CSS3 + Vanilla JS + iconos SVG embebidos, incluyendo empaquetador **`.zip` nativo** en navegador para descargar lotes completos sin librerías externas.
  - Visor PDF embebido opcional (`VISOR_PDF_EMBEBIDO`, casilla desmarcada por defecto) sobre `<canvas>` con **PDF.js** cargado bajo demanda; si se mantiene desmarcada (`NO`), abre el PDF directamente en una nueva pestaña con cero dependencias externas.
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
  - Control de acceso flexible mediante listas mixtas de **cuentas individuales y Google Groups** (`LISTA_NEGRA_USUARIOS` y `LISTA_BLANCA_USUARIOS`), con autohibridación o prefijo estricto `g:grupo@dominio` y precedencia de denegación (*deny overrides allow*).
  - Carpeta temporal dedicada `_Doc2PDF_Temp` creada automáticamente dentro de la carpeta donde reside la Google Sheet (con salvaguarda anti-borrado en purga sobre la propia hoja y subcarpetas).
  - Borrado **permanente** inmediato de temporales (`Drive.Files.remove`) en bloque `finally`, más trigger diario de purga para huérfanos.
  - Protección contra *Formula Injection* en `Auditoría` y política configurable de registro de nombres (`COMPLETO`, `ANONIMIZADO`, `NINGUNO`).
- **Panel administrativo en Google Sheets:**
  - Solo 2 pestañas optimizadas: `Ajustes` (con casillas `SI`/`NO`, desplegables, validaciones numéricas estrictas, ajuste multilínea y autoanchura) y `Auditoría` (filtros activos en todas las columnas y cabecera calibrada).
  - Recorte automático de filas y columnas vacías sobrantes al inicializar o reparar (avisando explícitamente del restablecimiento de `Ajustes` y respetando filas/columnas con datos en `Auditoría`).
  - Diálogo modal HTML rico **Resumen y analítica** (KPIs, actividad diaria SVG, donut por familias, top extensiones, usuarios activos clicables y tabla reciente con opción para desplegar toda la hoja `Auditoría`, buscar, filtrar por usuario/familia/estado/origen, paginar y exportar a **CSV**).
  - Modales HTML de mantenimiento con estética unificada: **Trigger diario de mantenimiento** (comprueba estado, activa o desactiva), **Purgar archivos temporales ahora** (huérfanos &gt; 30 min o todos), **Verificar formatos Drive API** (acordeón por familias) y **Verificar listas de acceso y grupos** (resolución en vivo, leyenda `g:`, detección *Fail-Closed* y solapamientos).
  - Modal **Acerca de Doc2PDF** de atribución con banner gráfico WebP embebido en Base64 desde el menú personalizado.

---

## Arquitectura de seguridad y defensa en profundidad

Siguiendo un diseño *Zero Trust* respecto al navegador y mínimo privilegio en Apps Script:

1. **Mínimo privilegio en hojas de cálculo (`@OnlyCurrentDoc` + `spreadsheets.currentonly`):** el manifiesto restringe el acceso de `SpreadsheetApp` exclusivamente al libro de control contenedor, impidiendo leer o alterar otras hojas del desplegador.
2. **Encapsulación estricta de funciones internas (sufijo `_`):** todas las funciones internas de configuración, caché, lectura/purga de auditoría y conversión llevan sufijo `_` para impedir su invocación directa desde `google.script.run` en consola DevTools.
3. **Control de rol Administrador y tokens efímeros de diálogo (`assertSpreadsheetAdmin_` + `dialogToken`):** a diferencia de un backend web tradicional (donde la app pública y el panel de administración se aíslan por rutas o sesiones distintas), en proyectos Apps Script vinculados a una hoja que además publican una WebApp `google.script.run` aplana todas las funciones globales sin `_` en un único espacio RPC compartido entre la WebApp y los modales de Sheets. Mientras `assertSpreadsheetAdmin_` bloquea a cualquier usuario que no sea propietario o editor de la hoja, el `dialogToken` efímero emitido al abrir cada diálogo añade una capa ligera (~20 líneas, análoga a un token anti-CSRF de contexto) de defensa en profundidad frente a una eventual escalada lateral *WebApp → funciones administrativas de Sheets* bajo la sesión del propio administrador.
4. **Defensa anti-*Confused Deputy* en conversión por enlace de Drive (`assertDriveFileSharedWithRequester_`):** aunque la WebApp ejecuta como cuenta administradora, `convertDriveUrl_` bloquea explícitamente el acceso a la propia hoja de control, a la carpeta temporal interna y a cualquier archivo privado del administrador que no esté compartido explícitamente con el solicitante, con su dominio Workspace o de forma pública.
5. **Política *Fail-Closed* e integridad de sesiones:** cualquier petición sin email institucional verificado (`NO_SESSION_EMAIL`) es rechazada. En control de acceso se aplica precedencia estricta `LISTA_NEGRA_USUARIOS` → `LISTA_BLANCA_USUARIOS` → `DOMINIOS_PERMITIDOS`, y cualquier grupo marcado con `g:` inaccesible o sin permiso para consultar miembros (`BLACKLIST_GROUP_ERROR` / `WHITELIST_GROUP_ERROR`) deniega el acceso preventivamente. Las sesiones fragmentadas de subida y descarga (`uploadFileChunk_`, `finalizeChunkedConversion_`, `abortChunkedUpload_`, `fetchPdfDownloadChunk_`, `completePdfDownload_`) verifican estrictamente que el solicitante sea el propietario de la sesión.
6. **Protección de hoja de cálculo:** escritura en `Auditoría` serializada con `LockService`, sanitización contra *Formula Injection* (`=`, `+`, `-`, `@`) y modos de registro `COMPLETO` / `ANONIMIZADO` / `NINGUNO`.

---

## Estructura del repositorio

La raíz del repositorio coincide con el directorio raíz de `clasp` (`rootDir: "."`):

```text
.
├── assets/
│   ├── badges/
│   │   ├── apps-script.svg      # Badge Google Apps Script V8
│   │   ├── drive-api.svg        # Badge Drive API v3
│   │   ├── license.svg          # Badge GNU GPL v3.0
│   │   └── version.svg          # Badge versión del proyecto
│   └── banner.jpg               # Banner del repositorio (1376×688, proporción 2:1)
├── .clasp.json.example          # Plantilla de configuración para clasp (rootDir: ".")
├── .claspignore                 # Excluye documentación, assets y metadatos al hacer clasp push
├── .gitignore
├── LICENSE                      # GNU GPL v3.0
├── README.md
├── appsscript.json              # Manifiesto V8, Drive API v3, scopes mínimos y acceso DOMAIN
├── Config.gs                    # Ajustes cacheados, catálogo de formatos, límites y helpers
├── Setup.gs                     # Menú personalizado, inicialización/recorte de pestañas y diálogos
├── Auth.gs                      # Identidad Workspace, listas de acceso (usuarios/Google Groups) y cuota diaria
├── Audit.gs                     # Auditoría concurrente (LockService), retención, analítica y consulta completa
├── Cleanup.gs                   # Carpeta temporal _Doc2PDF_Temp junto a la hoja y purga segura
├── Converter.gs                 # Motor de conversión (directo, chunked y enlace Drive)
├── WebApp.gs                    # doGet() y endpoints RPC para google.script.run
├── Index.html                   # Estructura de la WebApp (tema claro/oscuro + <dialog>)
├── Styles.html                  # CSS adaptativo con soporte light-dark() y fallbacks
├── ClientScript.html            # Cola, drag & drop, chunking y vista previa PDF
├── DashboardDialog.html         # Resumen y analítica (SVG nativo) + explorador filtrable y CSV de Auditoría
├── DialogStyles.html            # Hoja de estilos compartida para diálogos de mantenimiento
├── AboutDialog.html             # Diálogo modal "Acerca de Doc2PDF" con banner Base64 WebP
├── InitDialog.html              # Confirmación de reinicialización con aviso de borrado/recorte
├── AccessCheckDialog.html       # Diagnóstico visual de listas negra/blanca, Google Groups y leyenda g:
├── FormatsCheckDialog.html      # Verificación desplegable por familias de formatos en Drive API v3
├── TriggerDialog.html           # Estado, activación y desactivación del trigger diario
├── PurgeTempDialog.html         # Purga selectiva de temporales (huérfanos > 30 min o total)
└── DeployDialog.html            # Asistente paso a paso para guardar la URL de WebApp
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
   - Ejecuta **`📄 Doc2PDF → 🛠️ Mantenimiento → ⏱️ Trigger diario de mantenimiento`** *(y pulsa **Activar trigger diario** en el diálogo)*.

4. **Desplegar como WebApp en Google Workspace**
   - En el editor de Apps Script, haz clic en **Implementar → Nueva implementación → Aplicación web**.
   - **Ejecutar como:** `Yo` *(la cuenta administradora/funcional propietaria de la hoja)*.
   - **Quién tiene acceso:** `Cualquier usuario de [tu dominio Workspace]`.
   - Copia la URL generada terminada en `/exec`.
   - En la hoja de cálculo, abre **`📄 Doc2PDF → 🚀 Configurar URL de WebApp`**, pega la URL y pulsa **Guardar URL**.

### Pruebas con impersonación (solo desarrollo)

Para validar la experiencia de un usuario estándar (modo mantenimiento, listas blancas/negras, restricción de dominio o cuota diaria) sin necesitar una segunda cuenta en el dominio, [Auth.gs](Auth.gs) incluye dos constantes de simulación apagadas por defecto (`DEV_IMPERSONATE_ENABLED_ = false` y `DEV_IMPERSONATE_EMAIL_ = ''`):

- **Activación temporal en entorno `/dev`:** cambia ambas constantes en tu copia local / editor de Apps Script hacia un correo de prueba de tu dominio, abre la implementación de pruebas (`/dev`) y **devuélvelas a `false` / `''` antes de publicar `/exec` o confirmar en Git** (la suite de verificación falla si quedan activas).
- **Salvaguardas integradas:**
  - **Exclusivo de la cuenta desplegadora:** solo surte efecto si la sesión real coincide con `Session.getEffectiveUser()` (los co-editores de la hoja y usuarios normales no pueden suplantar identidades).
  - **Banner permanente y trazabilidad:** la WebApp muestra un aviso rojo superior en todo momento (incluso en pantallas de bloqueo) y `Auditoría` registra la fila como `<simulado> [simulado por <desplegador>]`.
  - **Sin impacto en cuotas de terceros:** las conversiones realizadas en modo simulación no incrementan la cuota diaria del usuario simulado ni su recuento en `Auditoría`; las operaciones del menú en Google Sheets siguen validando la identidad real.

---

## Ajustes configurables (`Ajustes`)

La pestaña `Ajustes` incorpora validaciones estrictas de datos sobre sus 21 parámetros:

| Clave | Valor inicial | Tipo / Validación | Descripción |
| :--- | :--- | :--- | :--- |
| `NOMBRE_APP` | `Doc2PDF` | Texto (máx. 60 car.) | Título principal de la WebApp y pestaña |
| `SUBTITULO_CENTRO` | `Conversor interno de documentos a PDF` | Texto (máx. 120 car.) | Subtítulo institucional en cabecera |
| `COLOR_ACENTO` | `#1a73e8` | Color Hex (`#RRGGBB`) | Color principal de botones, iconos y realces |
| `LOGO_URL` | *(vacío)* | URL `https://` | URL HTTPS opcional del logotipo |
| `SERVICIO_ACTIVO` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Activa (marcada) o pausa en modo mantenimiento (desmarcada) el conversor |
| `MENSAJE_MANTENIMIENTO` | *(mensaje estándar)* | Texto (máx. 300 car.) | Aviso mostrado cuando `SERVICIO_ACTIVO` está desmarcado |
| `DOMINIOS_PERMITIDOS` | *(vacío)* | Lista por comas | Dominios Workspace permitidos (vacío = dominio del desplegador) |
| `LISTA_NEGRA_USUARIOS` | *(vacío)* | Lista por comas | Opcional *(se evalúa primero y prevalece sobre lista blanca)*: correos de usuario o Google Groups bloqueados *(usa `g:grupo@dominio` para modo estricto)* |
| `LISTA_BLANCA_USUARIOS` | *(vacío)* | Lista por comas | Opcional *(vacío = todo el dominio)*: correos de usuario o Google Groups autorizados *(usa `g:grupo@dominio` para modo estricto; exentos admins de la hoja)* |
| `LIMITE_DIARIO_USUARIO` | `30` | Numérico (`0 – 10 000`) | Conversiones máximas por usuario al día (`0` = sin límite; sugerido) |
| `TAMANO_MAX_MB` | `20` | Numérico (`1 – 45`) | Peso máximo por archivo en MB *(límite técnico seguro bajo tope de 50 MB de Apps Script)* |
| `MAX_ARCHIVOS_LOTE` | `5` | Numérico (`1 – 20`) | Máximo de archivos admitidos por lote al arrastrar o seleccionar *(sugerido)* |
| `PERMITIR_DOCUMENTOS` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Habilita documentos de texto (`.docx`, `.doc`, `.dot`, `.dotx`, `.odt`, `.rtf`) |
| `PERMITIR_HOJAS` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Habilita hojas de cálculo (`.xlsx`, `.xls`, `.xlt`, `.xltx`, `.xlsm`, `.ods`, `.csv`, `.tsv`) |
| `PERMITIR_PRESENTACIONES` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Habilita presentaciones (`.pptx`, `.ppt`, `.odp`) |
| `PERMITIR_TEXTO` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Habilita texto plano y marcado (`.txt`, `.md`, `.html`, `.htm`) |
| `PERMITIR_ENLACE_DRIVE` | `SI` *(marcada)* | Casilla (`SI` / `NO`) | Habilita la pestaña de conversión por enlace/ID de Drive |
| `VISOR_PDF_EMBEBIDO` | `NO` *(desmarcada)* | Casilla (`SI` / `NO`) | Si está marcada, usa el visor modal en `<canvas>` con PDF.js bajo demanda; desmarcada, abre el PDF en pestaña nueva sin librerías externas |
| `REGISTRAR_NOMBRE_ARCHIVO` | `COMPLETO` | Desplegable (`COMPLETO` / `ANONIMIZADO` / `NINGUNO`) | Nivel de privacidad del nombre registrado en `Auditoría` |
| `RETENCION_AUDITORIA_DIAS` | `180` | Numérico (`0 – 3 650`) | Días de retención automática (`0` = indefinido; sugerido) |
| `CARPETA_TEMP_ID` | *(automático)* | Texto (ID Drive) | ID de la subcarpeta `_Doc2PDF_Temp` creada junto a esta hoja |

> [!NOTE]
> - **Caché:** los ajustes y la resolución/pertenencia a Google Groups se cachean durante **5 minutos** (`ScriptCache`). Tras editar la pestaña `Ajustes` o modificar miembros de un grupo, puedes aplicar los cambios al instante ejecutando **`📄 Doc2PDF → 🛠️ Mantenimiento → 🔄 Limpiar caché de ajustes`**.
> - **Equipo administrador y modo mantenimiento:** el rol de administración se determina por la cuenta desplegadora del script más el **propietario y los editores explícitos** de la Google Sheet de control (cacheado 60 s). Quienes pertenezcan a este grupo mantienen acceso de pruebas con un banner de aviso cuando `SERVICIO_ACTIVO` está desmarcado (`NO`) —mientras el resto de usuarios queda bloqueado— y quedan exentos de `LISTA_BLANCA_USUARIOS`, `DOMINIOS_PERMITIDOS` y `LIMITE_DIARIO_USUARIO` (`Sin límite (Admin)` en la cabecera, conservando el registro en `Auditoría`), aunque **nunca están exentos de `LISTA_NEGRA_USUARIOS`**.
> - **Cuota diaria y caché:** `LIMITE_DIARIO_USUARIO` se calcula mediante caché (`ScriptCache`, TTL máximo de 6 h) respaldada sobre `Auditoría` en la zona horaria de la hoja de cálculo, contabilizando únicamente conversiones finalizadas con estado `OK`.
> - **Concurrencia de conversión:** `MAX_ARCHIVOS_LOTE` limita cuántos archivos entran en cola en cada envío; la WebApp los procesa de **2 en 2** en paralelo (`DIRECT_CONCURRENCY = 2`) para subidas directas (`<= 6 MB`) y enlaces de Drive, y de **1 en 1** en subidas fragmentadas (`> 6 MB`), protegiendo así el cupo compartido de 30 ejecuciones simultáneas de Apps Script.
> - **Carpeta temporal y privacidad:** `_Doc2PDF_Temp` se crea automáticamente como subcarpeta en la misma carpeta donde reside la Google Sheet (heredando sus permisos). Se recomienda alojar la hoja de control en una **carpeta privada** del administrador o cuenta funcional.

### Control de acceso: listas de usuarios y grupos

Tanto `LISTA_NEGRA_USUARIOS` como `LISTA_BLANCA_USUARIOS` aceptan indistintamente **cuentas individuales de usuario y direcciones de Google Groups** en la misma celda, separadas por comas, puntos y coma o saltos de línea (por ejemplo: `alumnado@centro.edu, g:exalumnos@centro.edu, becario@centro.edu`).

1. **Orden de evaluación y precedencia (*deny overrides allow*):**
   Doc2PDF evalúa siempre las reglas en el mismo orden en que figuran en `Ajustes`:
   1. **`LISTA_NEGRA_USUARIOS` (máxima prioridad):** si la cuenta coincide directamente o pertenece a algún grupo de esta lista, el acceso se **deniega siempre** (`USER_BLOCKED`), **prevaleciendo sobre cualquier permiso en `LISTA_BLANCA_USUARIOS`, `DOMINIOS_PERMITIDOS` o el rol Administrador**.
   2. **`LISTA_BLANCA_USUARIOS`:** si contiene al menos una entrada, solo pueden acceder los usuarios o miembros de grupos incluidos en ella (más propietarios/editores explícitos de la hoja); si está vacía, se admite al conjunto del dominio.
   3. **`DOMINIOS_PERMITIDOS`:** valida el dominio de la cuenta (salvo para administradores de la hoja).

   | Caso de solapamiento | Resultado |
   | :--- | :--- |
   | Usuario en lista negra **y** en lista blanca | **Bloqueado** (`LISTA_NEGRA_USUARIOS` prevalece) |
   | Grupo en lista negra **y** usuario individual en lista blanca | **Bloqueado** si el usuario pertenece al grupo negro |
   | Grupo en lista blanca **y** usuario o grupo bloqueado en lista negra | **Bloqueado** para quien esté en lista negra; **permitido** para el resto del grupo blanco |

2. **Cómo comprueba las entradas y para qué sirve el prefijo `g:` (*Fail-Closed*):**
   - **Comprobación en dos pasos:** primero comprueba en memoria si el correo del usuario conectado aparece escrito literalmente en la lista (entre las entradas sin prefijo `g:`). Si coincide, resuelve el acceso al instante sin hacer ninguna consulta a Google Groups, ahorrando latencia y cuota diaria de `GroupsApp`. Solo cuando no hay coincidencia directa consulta los posibles grupos (cacheando el resultado 5 minutos en `ScriptCache`).
   - **Por qué existe `g:` (autodetección frente a modo estricto *Fail-Closed*):**
     - Si escribes una dirección **sin prefijo** (`alumnado@centro.edu`) y Google Groups no la encuentra, Doc2PDF no puede distinguir si se trata de una **cuenta de usuario individual** o de un **grupo mal escrito, borrado u oculto** para la cuenta desplegadora. Por ello asume que es una cuenta individual (cacheándola 5 minutos para no repetir consultas). El riesgo es que, en `LISTA_NEGRA_USUARIOS`, una simple errata o un cambio de visibilidad dejaría de bloquear a los miembros del grupo en silencio.
     - Al anteponer **`g:`** (`g:alumnado@centro.edu`, o su alias `grupo:alumnado@centro.edu`), declaras explícitamente que esa entrada **debe existir y poder leerse como un Google Group**. Si falla por errata, borrado o falta de permisos, Doc2PDF no la interpreta como correo individual: **bloquea preventivamente el acceso (*Fail-Closed*)** y registra el grupo afectado en `Auditoría` (`BLACKLIST_GROUP_ERROR` / `WHITELIST_GROUP_ERROR`) para que ninguna regla quede desactivada sin aviso.

   | Formato en celda | Comportamiento si no resuelve como grupo | Comportamiento si el grupo existe pero oculta miembros | Uso recomendado |
   | :--- | :--- | :--- | :--- |
   | `correo@dominio` *(sin prefijo)* | Lo trata como cuenta de usuario individual (cacheado 5 min para evitar reconsultas) | **Bloqueo *Fail-Closed*** (`BLACKLIST_GROUP_ERROR` / `WHITELIST_GROUP_ERROR`) | Cuentas individuales o listas mixtas simples donde prefieres omitir prefijos |
   | `g:grupo@dominio` *(o `grupo:grupo@dominio`)* | **Bloqueo *Fail-Closed*** (error explícito de grupo inexistente o sin permiso) | **Bloqueo *Fail-Closed*** (`BLACKLIST_GROUP_ERROR` / `WHITELIST_GROUP_ERROR`) | **Recomendado para todos los Google Groups** (especialmente en `LISTA_NEGRA_USUARIOS`): impide que una errata o falta de permisos silencie una regla |

3. **Membresía directa y requisitos de permisos (Google Groups):**
   - La comprobación usa el servicio nativo `GroupsApp.getGroupByEmail(correo).hasUser(email)` bajo el scope `https://www.googleapis.com/auth/groups`, sin requerir **Admin SDK / Directory API** ni privilegios de Superadministrador sobre toda la consola Workspace.
   - **Solo verifica membresía directa:** si un grupo A contiene anidado un grupo B, la pertenencia a B **no** se hereda automáticamente; añade explícitamente el grupo B si necesitas incluirlo.
   - **Visibilidad obligatoria para la cuenta desplegadora:** para que `GroupsApp` pueda evaluar un grupo sin errores, la cuenta desplegadora (`executeAs: "USER_DEPLOYING"`) debe poder localizar el grupo en el directorio y tener permiso de lectura sobre su lista de miembros (**Google Groups → Configuración del grupo → *Quién puede ver a los miembros***).
   - **Diagnóstico en Sheets:** tras configurar grupos o cuentas, ejecuta **`📄 Doc2PDF → 🛠️ Mantenimiento → 🔍 Verificar listas de acceso y grupos`** para validar en vivo qué entradas resuelven como grupo, detectar faltas de permisos o advertir sobre solapamientos antes de publicar cambios.

#### Flujo de evaluación y optimizaciones

El diagrama siguiente resume cómo Doc2PDF aplica **cortocircuito por precedencia**, **resolución directa en memoria sin llamadas a `GroupsApp`** y **caché de 5 minutos en `ScriptCache`** (`GTYPE` / `GMEM`) para minimizar latencia y consumo de cuota incluso con listas mixtas amplias:

```mermaid
---
config:
  flowchart:
    wrappingWidth: 360
    padding: 18
---
flowchart TD
  REQ(["Petición entrante<br/>correo Workspace"]) --> BL_EMPTY

  subgraph BL["1. LISTA_NEGRA_USUARIOS (prioridad máxima · admins)"]
    BL_EMPTY{"¿Lista negra<br/>vacía?"}
    BL_DIRECT{"¿Correo literal sin g:<br/>coincide en memoria?"}
    BL_GRP["Evaluar grupos<br/>caché 5 min (GTYPE/GMEM)<br/>y GroupsApp"]
    BL_BLOCK(["Denegado: USER_BLOCKED<br/>corta lista blanca"])
    BL_ERR(["Denegado Fail-Closed<br/>BLACKLIST_GROUP_ERROR"])

    BL_EMPTY -->|"No"| BL_DIRECT
    BL_DIRECT -->|"Sí · 0 llamadas<br/>GroupsApp"| BL_BLOCK
    BL_DIRECT -->|"No"| BL_GRP
    BL_GRP -->|"Miembro de grupo"| BL_BLOCK
    BL_GRP -->|"Error g: o sin<br/>permiso miembros"| BL_ERR
  end

  BL_EMPTY -->|"Sí"| WL_SKIP
  BL_GRP -->|"Sin coincidencia"| WL_SKIP

  subgraph WL["2. LISTA_BLANCA_USUARIOS (exentos admins)"]
    WL_SKIP{"¿Lista blanca vacía<br/>o usuario Admin?"}
    WL_DIRECT{"¿Correo literal sin g:<br/>coincide en memoria?"}
    WL_GRP["Evaluar grupos<br/>caché 5 min (GTYPE/GMEM)<br/>y GroupsApp"]
    WL_ERR(["Denegado Fail-Closed<br/>WHITELIST_GROUP_ERROR"])
    WL_DENY(["Denegado<br/>NOT_IN_WHITELIST"])

    WL_SKIP -->|"No"| WL_DIRECT
    WL_DIRECT -->|"No"| WL_GRP
    WL_GRP -->|"Error g: o sin<br/>permiso miembros"| WL_ERR
    WL_GRP -->|"Sin coincidencia"| WL_DENY
  end

  WL_SKIP -->|"Sí · 0 llamadas<br/>GroupsApp"| DOM_CHK
  WL_DIRECT -->|"Sí · 0 llamadas<br/>GroupsApp"| DOM_CHK
  WL_GRP -->|"Miembro de grupo"| DOM_CHK

  subgraph DOM["3. DOMINIOS_PERMITIDOS (exentos admins)"]
    DOM_CHK{"¿Admin o dominio<br/>autorizado?"}
    ALLOW_OK(["Acceso autorizado<br/>pasa a cuota diaria"])
    DOM_DENY(["Denegado<br/>DOMAIN_NOT_ALLOWED"])

    DOM_CHK -->|"Sí"| ALLOW_OK
    DOM_CHK -->|"No"| DOM_DENY
  end

  classDef deny stroke:#d93025,stroke-width:2px;
  classDef allow stroke:#188038,stroke-width:2px;
  classDef fast stroke:#1a73e8,stroke-width:2px;

  class BL_BLOCK,BL_ERR,WL_ERR,WL_DENY,DOM_DENY deny;
  class ALLOW_OK allow;
  class BL_DIRECT,WL_SKIP,WL_DIRECT fast;
```

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
