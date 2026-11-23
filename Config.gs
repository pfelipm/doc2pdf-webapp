/**
 * @OnlyCurrentDoc
 *
 * Config.gs
 * Configuración centralizada, caché de ajustes, catálogo de formatos nativos
 * y normalización de valores para Doc2PDF.
 */

const APP_NAME = 'Doc2PDF';
const APP_VERSION = 'v1.0 (noviembre 2026)';
const APP_REPO_URL = 'https://github.com/pfelipm/doc2pdf-webapp';
const APP_AUTHOR_NAME = 'Pablo Felip';
const APP_AUTHOR_HANDLE = '@pfelipm';
const APP_AUTHOR_URL = 'https://pablofelip.online/sobre-mi/';

const SHEET_CONFIG_NAME = 'Ajustes';
const SHEET_AUDIT_NAME = 'Auditoría';

const CONFIG_CACHE_KEY = 'DOC2PDF_CONFIG_V1';
const CONFIG_CACHE_TTL_SEC = 300;

const PROP_WEBAPP_URL = 'WEBAPP_URL';
const PROP_TEMP_FOLDER_ID = 'TEMP_FOLDER_ID';
const TEMP_FOLDER_DEFAULT_NAME = '_Doc2PDF_Temp';

const DIRECT_UPLOAD_MAX_BYTES = 6 * 1024 * 1024; // 6 MB
const CHUNK_SIZE_BYTES = 2 * 1024 * 1024;        // 2 MB (múltiplo exacto de 256 KB)
const DIRECT_DOWNLOAD_MAX_BYTES = 7 * 1024 * 1024; // 7 MB
const DOWNLOAD_CHUNK_BYTES = 3 * 1024 * 1024;    // 3 MB

const AUDIT_MODE_COMPLETO = 'COMPLETO';
const AUDIT_MODE_ANONIMIZADO = 'ANONIMIZADO';
const AUDIT_MODE_NINGUNO = 'NINGUNO';

const FAMILY_ORDER = ['DOCUMENTOS', 'HOJAS', 'PRESENTACIONES', 'TEXTO'];

const FAMILY_META = {
  DOCUMENTOS: { label: 'Documentos', color: '#1a73e8' },
  HOJAS: { label: 'Hojas de cálculo', color: '#188038' },
  PRESENTACIONES: { label: 'Presentaciones', color: '#e37400' },
  TEXTO: { label: 'Texto y marcado', color: '#7c3aed' }
};

/**
 * Límites min/max compartidos entre la validación de la hoja Ajustes y la normalización del servidor.
 */
const CONFIG_NUMERIC_LIMITS = {
  LIMITE_DIARIO_USUARIO: {
    min: 0,
    max: 10000,
    help: 'Introduce un número entre 0 (sin límite) y 10000.'
  },
  TAMANO_MAX_MB: {
    min: 1,
    max: 45,
    help: 'Introduce un valor entre 1 y 45 MB (límite técnico seguro bajo el tope de 50 MB de Apps Script / UrlFetchApp).'
  },
  MAX_ARCHIVOS_LOTE: {
    min: 1,
    max: 20,
    help: 'Introduce un valor entre 1 y 20 archivos por lote.'
  },
  RETENCION_AUDITORIA_DIAS: {
    min: 0,
    max: 3650,
    help: 'Introduce un valor entre 0 (indefinido) y 3650 días (10 años).'
  }
};

/**
 * Definición ordenada de los ajustes disponibles en la pestaña "Ajustes".
 * Formato: [clave, valorPorDefecto, descripcion, tipoValidacion]
 */
const DEFAULT_CONFIG_DEFS = [
  ['NOMBRE_APP', 'Doc2PDF', 'Nombre visible de la aplicación web y título de pestaña', 'text'],
  ['SUBTITULO_CENTRO', 'Conversor interno de documentos a PDF', 'Subtítulo institucional mostrado en la cabecera', 'text'],
  ['COLOR_ACENTO', '#1a73e8', 'Color principal en hexadecimal (#RRGGBB)', 'text'],
  ['LOGO_URL', '', 'URL HTTPS opcional del logotipo del centro (vacío para usar icono integrado)', 'text'],
  ['SERVICIO_ACTIVO', 'SI', 'Activa (marcada) o pausa en modo mantenimiento (desmarcada) el conversor', 'checkbox'],
  ['MENSAJE_MANTENIMIENTO', 'El servicio se encuentra temporalmente en mantenimiento. Inténtalo de nuevo más tarde.', 'Mensaje mostrado cuando SERVICIO_ACTIVO está desmarcado', 'text'],
  ['DOMINIOS_PERMITIDOS', '', 'Dominios Workspace separados por comas (vacío = mismo dominio que la cuenta ejecutora)', 'text'],
  ['LISTA_BLANCA_USUARIOS', '', 'Opcional: correos autorizados separados por comas (vacío = todo el dominio)', 'text'],
  ['LISTA_NEGRA_USUARIOS', '', 'Opcional: correos bloqueados separados por comas', 'text'],
  ['LIMITE_DIARIO_USUARIO', 30, 'Máximo de conversiones al día por usuario (0 = sin límite)', 'number'],
  ['TAMANO_MAX_MB', 20, 'Tamaño máximo permitido por archivo en MB (rango seguro: 1 a 45)', 'number'],
  ['MAX_ARCHIVOS_LOTE', 5, 'Máximo de archivos admitidos por lote al arrastrar o seleccionar (1 a 20)', 'number'],
  ['PERMITIR_DOCUMENTOS', 'SI', 'Habilita conversión de documentos de texto (.docx, .doc, .dot, .dotx, .odt, .rtf)', 'checkbox'],
  ['PERMITIR_HOJAS', 'SI', 'Habilita conversión de hojas de cálculo (.xlsx, .xls, .xlt, .xltx, .xlsm, .ods, .csv, .tsv)', 'checkbox'],
  ['PERMITIR_PRESENTACIONES', 'SI', 'Habilita conversión de presentaciones (.pptx, .ppt, .odp)', 'checkbox'],
  ['PERMITIR_TEXTO', 'SI', 'Habilita conversión de texto plano y marcado (.txt, .md, .html, .htm)', 'checkbox'],
  ['PERMITIR_ENLACE_DRIVE', 'SI', 'Permite convertir documentos pegando una URL compartida de Google Drive', 'checkbox'],
  ['VISOR_PDF_EMBEBIDO', 'NO', 'Si está marcada, usa el visor PDF embebido en modal con PDF.js (CDN externo); desmarcada, abre el PDF en nueva pestaña sin dependencias externas', 'checkbox'],
  ['REGISTRAR_NOMBRE_ARCHIVO', 'COMPLETO', 'Nivel de detalle del nombre en Auditoría: COMPLETO, ANONIMIZADO o NINGUNO', 'audit_mode'],
  ['RETENCION_AUDITORIA_DIAS', 180, 'Días de conservación en Auditoría antes de purga automática (0 = indefinido)', 'number'],
  ['CARPETA_TEMP_ID', '', 'ID de la carpeta temporal en Drive (se crea automáticamente como subcarpeta _Doc2PDF_Temp junto a esta hoja si se deja vacío)', 'text']
];

/**
 * Catálogo completo de formatos de documentos de escritorio soportados de forma nativa
 * por Google Drive API v3 sin bibliotecas externas.
 */
const SUPPORTED_FORMATS = {
  docx: {
    ext: 'docx',
    label: 'Microsoft Word (.docx)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  doc: {
    ext: 'doc',
    label: 'Microsoft Word 97-2003 (.doc)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/msword',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  dot: {
    ext: 'dot',
    label: 'Plantilla Word 97-2003 (.dot)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/msword',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  dotx: {
    ext: 'dotx',
    label: 'Plantilla Word (.dotx)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  odt: {
    ext: 'odt',
    label: 'Texto OpenDocument (.odt)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/vnd.oasis.opendocument.text',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  rtf: {
    ext: 'rtf',
    label: 'Texto enriquecido (.rtf)',
    family: 'DOCUMENTOS',
    sourceMime: 'application/rtf',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  txt: {
    ext: 'txt',
    label: 'Texto plano (.txt)',
    family: 'TEXTO',
    sourceMime: 'text/plain',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  md: {
    ext: 'md',
    label: 'Markdown (.md)',
    family: 'TEXTO',
    sourceMime: 'text/markdown',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  html: {
    ext: 'html',
    label: 'Documento HTML (.html)',
    family: 'TEXTO',
    sourceMime: 'text/html',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  htm: {
    ext: 'htm',
    label: 'Documento HTML (.htm)',
    family: 'TEXTO',
    sourceMime: 'text/html',
    targetGoogleMime: 'application/vnd.google-apps.document',
    exportType: 'document'
  },
  xlsx: {
    ext: 'xlsx',
    label: 'Microsoft Excel (.xlsx)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  xls: {
    ext: 'xls',
    label: 'Microsoft Excel 97-2003 (.xls)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.ms-excel',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  xlt: {
    ext: 'xlt',
    label: 'Plantilla Excel 97-2003 (.xlt)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.ms-excel',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  xltx: {
    ext: 'xltx',
    label: 'Plantilla Excel (.xltx)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  xlsm: {
    ext: 'xlsm',
    label: 'Excel con macros (.xlsm)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.ms-excel.sheet.macroenabled.12',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  ods: {
    ext: 'ods',
    label: 'Hoja OpenDocument (.ods)',
    family: 'HOJAS',
    sourceMime: 'application/vnd.oasis.opendocument.spreadsheet',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  csv: {
    ext: 'csv',
    label: 'Valores separados por comas (.csv)',
    family: 'HOJAS',
    sourceMime: 'text/csv',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  tsv: {
    ext: 'tsv',
    label: 'Valores separados por tabulador (.tsv)',
    family: 'HOJAS',
    sourceMime: 'text/tab-separated-values',
    targetGoogleMime: 'application/vnd.google-apps.spreadsheet',
    exportType: 'spreadsheets'
  },
  pptx: {
    ext: 'pptx',
    label: 'Microsoft PowerPoint (.pptx)',
    family: 'PRESENTACIONES',
    sourceMime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    targetGoogleMime: 'application/vnd.google-apps.presentation',
    exportType: 'presentation'
  },
  ppt: {
    ext: 'ppt',
    label: 'Microsoft PowerPoint 97-2003 (.ppt)',
    family: 'PRESENTACIONES',
    sourceMime: 'application/vnd.ms-powerpoint',
    targetGoogleMime: 'application/vnd.google-apps.presentation',
    exportType: 'presentation'
  },
  odp: {
    ext: 'odp',
    label: 'Presentación OpenDocument (.odp)',
    family: 'PRESENTACIONES',
    sourceMime: 'application/vnd.oasis.opendocument.presentation',
    targetGoogleMime: 'application/vnd.google-apps.presentation',
    exportType: 'presentation'
  }
};

/**
 * Formatos nativos de Google Workspace soportados al convertir por enlace de Drive.
 */
const NATIVE_GOOGLE_FORMATS = {
  'application/vnd.google-apps.document': {
    ext: 'gdoc',
    family: 'DOCUMENTOS',
    exportType: 'document',
    label: 'Google Docs'
  },
  'application/vnd.google-apps.spreadsheet': {
    ext: 'gsheet',
    family: 'HOJAS',
    exportType: 'spreadsheets',
    label: 'Google Sheets'
  },
  'application/vnd.google-apps.presentation': {
    ext: 'gslides',
    family: 'PRESENTACIONES',
    exportType: 'presentation',
    label: 'Google Slides'
  }
};

/**
 * Obtiene la hoja de cálculo contenedora de forma perezosa.
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Convierte valores variados de hojas de cálculo a booleano.
 */
function parseBoolean_(val, fallback) {
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (val === null || val === undefined) return fallback;
  const normalized = String(val).trim().toUpperCase();
  if (!normalized) return fallback;
  if (['SI', 'SÍ', 'TRUE', 'VERDADERO', 'ACTIVADO', '1', 'YES', 'ON'].indexOf(normalized) !== -1) {
    return true;
  }
  if (['NO', 'FALSE', 'FALSO', 'DESACTIVADO', '0', 'OFF'].indexOf(normalized) !== -1) {
    return false;
  }
  return fallback;
}

/**
 * Valida estrictamente un color hexadecimal #RRGGBB.
 */
function sanitizeHexColor_(val, fallback) {
  const defaultHex = fallback || '#1a73e8';
  if (!val) return defaultHex;
  let candidate = String(val).trim();
  if (candidate.charAt(0) !== '#') {
    candidate = '#' + candidate;
  }
  return /^#[0-9a-fA-F]{6}$/.test(candidate) ? candidate.toLowerCase() : defaultHex;
}

/**
 * Valida que una URL de logotipo utilice exclusivamente HTTPS.
 */
function sanitizeHttpsUrl_(val) {
  if (!val) return '';
  const trimmed = String(val).trim();
  return /^https:\/\/[^\s"'<>]+$/i.test(trimmed) ? trimmed : '';
}

/**
 * Parsea un entero acotado en un rango seguro.
 */
function parseIntClamped_(val, fallback, minVal, maxVal) {
  if (val === '' || val === null || val === undefined) return fallback;
  const parsed = parseInt(val, 10);
  if (isNaN(parsed)) return fallback;
  let result = parsed;
  if (typeof minVal === 'number' && result < minVal) result = minVal;
  if (typeof maxVal === 'number' && result > maxVal) result = maxVal;
  return result;
}

/**
 * Divide una lista separada por comas, punto y coma o saltos de línea.
 * @return {Array<string>}
 */
function parseListSetting_(val, toUpper) {
  if (!val) return [];
  return String(val)
    .split(/[,;\s\n]+/)
    .map(function (item) {
      const cleaned = item.trim();
      return toUpper ? cleaned.toUpperCase() : cleaned.toLowerCase();
    })
    .filter(function (item) {
      return item.length > 0;
    });
}

/**
 * Construye el mapa base con valores por defecto sin leer de la hoja.
 */
function buildDefaultRawConfig_() {
  const raw = {};
  for (let i = 0; i < DEFAULT_CONFIG_DEFS.length; i++) {
    raw[DEFAULT_CONFIG_DEFS[i][0]] = DEFAULT_CONFIG_DEFS[i][1];
  }
  return raw;
}

/**
 * Normaliza un diccionario crudo de configuración aplicando validaciones y límites.
 */
function normalizeConfigObject_(rawInput) {
  const defaults = buildDefaultRawConfig_();
  const raw = Object.assign({}, defaults, rawInput || {});

  const allowDocumentos = parseBoolean_(raw.PERMITIR_DOCUMENTOS, true);
  const allowHojas = parseBoolean_(raw.PERMITIR_HOJAS, true);
  const allowPresentaciones = parseBoolean_(raw.PERMITIR_PRESENTACIONES, true);
  const allowTexto = parseBoolean_(raw.PERMITIR_TEXTO, true);

  const familiesList = [];
  if (allowDocumentos) familiesList.push('DOCUMENTOS');
  if (allowHojas) familiesList.push('HOJAS');
  if (allowPresentaciones) familiesList.push('PRESENTACIONES');
  if (allowTexto) familiesList.push('TEXTO');

  const rawAuditMode = String(raw.REGISTRAR_NOMBRE_ARCHIVO || AUDIT_MODE_COMPLETO).trim().toUpperCase();
  const auditNameMode = [AUDIT_MODE_COMPLETO, AUDIT_MODE_ANONIMIZADO, AUDIT_MODE_NINGUNO].indexOf(rawAuditMode) !== -1
    ? rawAuditMode
    : AUDIT_MODE_COMPLETO;

  return {
    NOMBRE_APP: String(raw.NOMBRE_APP || defaults.NOMBRE_APP).trim().slice(0, 60) || defaults.NOMBRE_APP,
    SUBTITULO_CENTRO: String(raw.SUBTITULO_CENTRO || '').trim().slice(0, 120),
    COLOR_ACENTO: sanitizeHexColor_(raw.COLOR_ACENTO, defaults.COLOR_ACENTO),
    LOGO_URL: sanitizeHttpsUrl_(raw.LOGO_URL),
    SERVICIO_ACTIVO: parseBoolean_(raw.SERVICIO_ACTIVO, true),
    MENSAJE_MANTENIMIENTO: String(raw.MENSAJE_MANTENIMIENTO || defaults.MENSAJE_MANTENIMIENTO).trim().slice(0, 300),
    DOMINIOS_PERMITIDOS: parseListSetting_(raw.DOMINIOS_PERMITIDOS, false),
    LISTA_BLANCA_USUARIOS: parseListSetting_(raw.LISTA_BLANCA_USUARIOS, false),
    LISTA_NEGRA_USUARIOS: parseListSetting_(raw.LISTA_NEGRA_USUARIOS, false),
    LIMITE_DIARIO_USUARIO: parseIntClamped_(
      raw.LIMITE_DIARIO_USUARIO,
      30,
      CONFIG_NUMERIC_LIMITS.LIMITE_DIARIO_USUARIO.min,
      CONFIG_NUMERIC_LIMITS.LIMITE_DIARIO_USUARIO.max
    ),
    TAMANO_MAX_MB: parseIntClamped_(
      raw.TAMANO_MAX_MB,
      20,
      CONFIG_NUMERIC_LIMITS.TAMANO_MAX_MB.min,
      CONFIG_NUMERIC_LIMITS.TAMANO_MAX_MB.max
    ),
    MAX_ARCHIVOS_LOTE: parseIntClamped_(
      raw.MAX_ARCHIVOS_LOTE,
      5,
      CONFIG_NUMERIC_LIMITS.MAX_ARCHIVOS_LOTE.min,
      CONFIG_NUMERIC_LIMITS.MAX_ARCHIVOS_LOTE.max
    ),
    PERMITIR_DOCUMENTOS: allowDocumentos,
    PERMITIR_HOJAS: allowHojas,
    PERMITIR_PRESENTACIONES: allowPresentaciones,
    PERMITIR_TEXTO: allowTexto,
    FAMILIAS_HABILITADAS: familiesList,
    PERMITIR_ENLACE_DRIVE: parseBoolean_(raw.PERMITIR_ENLACE_DRIVE, true),
    VISOR_PDF_EMBEBIDO: parseBoolean_(raw.VISOR_PDF_EMBEBIDO, false),
    REGISTRAR_NOMBRE_ARCHIVO: auditNameMode,
    RETENCION_AUDITORIA_DIAS: parseIntClamped_(
      raw.RETENCION_AUDITORIA_DIAS,
      180,
      CONFIG_NUMERIC_LIMITS.RETENCION_AUDITORIA_DIAS.min,
      CONFIG_NUMERIC_LIMITS.RETENCION_AUDITORIA_DIAS.max
    ),
    CARPETA_TEMP_ID: String(raw.CARPETA_TEMP_ID || '').trim()
  };
}

/**
 * Lee y normaliza la configuración desde la pestaña "Ajustes" usando CacheService.
 * @param {boolean=} forceRefresh Si es true, omite y renueva la caché.
 * @return {Object}
 */
function getConfig_(forceRefresh) {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    if (!forceRefresh && cache) {
      const cachedJson = cache.get(CONFIG_CACHE_KEY);
      if (cachedJson) {
        return JSON.parse(cachedJson);
      }
    }
  } catch (e) {
    console.warn('No se pudo leer la caché de ajustes:', e.message);
  }

  const raw = buildDefaultRawConfig_();
  try {
    const ss = getSpreadsheet_();
    const sheet = ss ? ss.getSheetByName(SHEET_CONFIG_NAME) : null;
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      for (let i = 0; i < values.length; i++) {
        const key = String(values[i][0] || '').trim();
        if (key && Object.prototype.hasOwnProperty.call(raw, key)) {
          raw[key] = values[i][1];
        }
      }
    }
  } catch (e) {
    console.warn('No se pudo leer la pestaña Ajustes; usando valores por defecto:', e.message);
  }

  const normalized = normalizeConfigObject_(raw);
  try {
    if (cache) {
      cache.put(CONFIG_CACHE_KEY, JSON.stringify(normalized), CONFIG_CACHE_TTL_SEC);
    }
  } catch (e) {
    console.warn('No se pudo almacenar la caché de ajustes:', e.message);
  }
  return normalized;
}

/**
 * Invalida la caché de configuración.
 */
function clearConfigCache_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache) {
      cache.remove(CONFIG_CACHE_KEY);
    }
  } catch (e) {
    console.warn('Error limpiando caché de ajustes:', e.message);
  }
}

/**
 * Actualiza un valor concreto en la pestaña "Ajustes" e invalida la caché.
 */
function setConfigValue_(key, value) {
  const ss = getSpreadsheet_();
  if (!ss) return false;
  const sheet = ss.getSheetByName(SHEET_CONFIG_NAME);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      clearConfigCache_();
      return true;
    }
  }
  return false;
}

/**
 * Devuelve la lista de formatos habilitados según FAMILIAS_HABILITADAS.
 * @return {Array<Object>}
 */
function getEnabledFormats_(config) {
  const cfg = config || getConfig_();
  const enabledFamilies = cfg.FAMILIAS_HABILITADAS || FAMILY_ORDER;
  const result = [];
  const keys = Object.keys(SUPPORTED_FORMATS);

  for (let i = 0; i < keys.length; i++) {
    const fmt = SUPPORTED_FORMATS[keys[i]];
    if (enabledFamilies.indexOf(fmt.family) !== -1) {
      result.push(fmt);
    }
  }
  return result;
}

/**
 * Obtiene la definición de formato por extensión si su familia está habilitada.
 */
function getFormatByExtension_(ext, config) {
  if (!ext) return null;
  const cleanExt = String(ext).trim().replace(/^\./, '').toLowerCase();
  const fmt = SUPPORTED_FORMATS[cleanExt];
  if (!fmt) return null;

  const cfg = config || getConfig_();
  const enabledFamilies = cfg.FAMILIAS_HABILITADAS || FAMILY_ORDER;
  return enabledFamilies.indexOf(fmt.family) !== -1 ? fmt : null;
}

/**
 * Obtiene la definición de formato de escritorio a partir del MIME type si está habilitado.
 */
function getFormatByMimeType_(mime, config) {
  if (!mime) return null;
  const cleanMime = String(mime).trim().toLowerCase();
  const cfg = config || getConfig_();
  const enabledFamilies = cfg.FAMILIAS_HABILITADAS || FAMILY_ORDER;

  // Orden determinista para MIME compartidos (p. ej. doc/dot, xls/xlt, html/htm)
  const preferredOrder = [
    'docx', 'doc', 'dotx', 'dot', 'odt', 'rtf',
    'xlsx', 'xls', 'xltx', 'xlt', 'xlsm', 'ods', 'csv', 'tsv',
    'pptx', 'ppt', 'odp',
    'md', 'txt', 'html', 'htm'
  ];

  for (let i = 0; i < preferredOrder.length; i++) {
    const fmt = SUPPORTED_FORMATS[preferredOrder[i]];
    if (fmt && fmt.sourceMime.toLowerCase() === cleanMime && enabledFamilies.indexOf(fmt.family) !== -1) {
      return fmt;
    }
  }
  return null;
}

/**
 * Obtiene la definición de formato si el archivo ya es nativo de Google Workspace.
 */
function getNativeFormatByMime_(mime, config) {
  if (!mime) return null;
  const cleanMime = String(mime).trim().toLowerCase();
  const nativeFmt = NATIVE_GOOGLE_FORMATS[cleanMime];
  if (!nativeFmt) return null;

  const cfg = config || getConfig_();
  const enabledFamilies = cfg.FAMILIAS_HABILITADAS || FAMILY_ORDER;
  return enabledFamilies.indexOf(nativeFmt.family) !== -1 ? nativeFmt : null;
}

/**
 * Construye el objeto de configuración seguro para enviar al cliente web.
 * Nunca expone listas de usuarios ni IDs internos de carpetas.
 */
function getPublicClientConfig_(userContext, config) {
  const cfg = config || getConfig_();
  const enabledFormats = getEnabledFormats_(cfg);
  const familiesSummary = FAMILY_ORDER
    .filter(function (fam) {
      return cfg.FAMILIAS_HABILITADAS.indexOf(fam) !== -1;
    })
    .map(function (fam) {
      const exts = enabledFormats
        .filter(function (f) { return f.family === fam; })
        .map(function (f) { return f.ext; });
      return {
        id: fam,
        label: FAMILY_META[fam].label,
        color: FAMILY_META[fam].color,
        extensions: exts
      };
    });

  return {
    appName: cfg.NOMBRE_APP,
    subtitle: cfg.SUBTITULO_CENTRO,
    accentColor: cfg.COLOR_ACENTO,
    logoUrl: cfg.LOGO_URL,
    version: APP_VERSION,
    serviceActive: cfg.SERVICIO_ACTIVO,
    maintenanceMessage: cfg.MENSAJE_MANTENIMIENTO,
    maxSizeMB: cfg.TAMANO_MAX_MB,
    maxSizeBytes: cfg.TAMANO_MAX_MB * 1024 * 1024,
    maxBatchFiles: cfg.MAX_ARCHIVOS_LOTE,
    dailyLimit: cfg.LIMITE_DIARIO_USUARIO,
    directUploadMaxBytes: DIRECT_UPLOAD_MAX_BYTES,
    chunkSizeBytes: CHUNK_SIZE_BYTES,
    allowDriveLink: cfg.PERMITIR_ENLACE_DRIVE,
    embeddedPdfViewer: cfg.VISOR_PDF_EMBEBIDO,
    locale: getSpreadsheetLocaleBcp47_(getSpreadsheet_()),
    auditNameMode: cfg.REGISTRAR_NOMBRE_ARCHIVO,
    enabledFamilies: cfg.FAMILIAS_HABILITADAS,
    familiesSummary: familiesSummary,
    enabledFormats: enabledFormats.map(function (f) {
      return {
        ext: f.ext,
        label: f.label,
        family: f.family,
        sourceMime: f.sourceMime
      };
    }),
    user: userContext
  };
}

/**
 * Obtiene la URL guardada de la WebApp en ScriptProperties.
 */
function getSavedWebAppUrl_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_WEBAPP_URL) || '';
}

/**
 * Valida y guarda la URL desplegada de la WebApp en ScriptProperties.
 */
function saveWebAppUrl_(url) {
  try {
    const trimmedUrl = String(url || '').trim();
    if (!trimmedUrl) {
      PropertiesService.getScriptProperties().deleteProperty(PROP_WEBAPP_URL);
      return { success: true, url: '' };
    }

    const isValidMacros = trimmedUrl.indexOf('https://script.google.com/macros/s/') === 0;
    const isValidDomainMacros = trimmedUrl.indexOf('https://script.google.com/a/') === 0;
    if (!isValidMacros && !isValidDomainMacros) {
      return {
        success: false,
        error: 'La URL debe comenzar por https://script.google.com/macros/s/ o https://script.google.com/a/...'
      };
    }

    PropertiesService.getScriptProperties().setProperty(PROP_WEBAPP_URL, trimmedUrl);
    return { success: true, url: trimmedUrl };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
